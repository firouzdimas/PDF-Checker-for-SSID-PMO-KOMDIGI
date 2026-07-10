# ================================================================
# SSID VALIDATOR v3.1 — DRY RUN (GROQ API + LLAMA VISION)
# ================================================================
# Panduan:
# 1. Jalankan di Google Colab
# 2. Sel pertama: install dependensi
# 3. Sel kedua: salin seluruh kode ini
# 4. Isi GROQ_API_KEYS dengan API Key dari console.groq.com
# 5. Jalankan — hasil print summary per-kriteria
#
# METODE: Groq Free API (Llama 4 Scout Vision)
# STRATEGI GAMBAR: 3 Overlapping Scan Strips + DPI=300
# ================================================================

# ===================== SEL 1 (JALANKAN DULU) =====================
# !apt-get install -y poppler-utils
# !pip install pdf2image Pillow pandas openpyxl gspread google-auth requests google-api-python-client pypdf
# =================================================================

import base64
import json
import re
import io
import time
import requests
import os
import difflib
import pandas as pd
from datetime import datetime

# Mencoba mengimpor library Google Colab
try:
    from google.colab import auth
    from google.colab import files as colab_files
    IN_COLAB = True
except ImportError:
    IN_COLAB = False

# Auto-install dependensi
try:
    import pdf2image
    import pypdf
except ImportError:
    if IN_COLAB:
        print("[*] Menginstall dependensi otomatis di Colab...")
        import subprocess
        import sys
        try:
            subprocess.check_call(["apt-get", "update", "-y"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.check_call(["apt-get", "install", "-y", "poppler-utils"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pdf2image", "Pillow", "pypdf"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print("[✅] Instalasi dependensi sukses!")
        except Exception as e:
            print(f"[❌] Gagal install: {e}")

from google.auth import default
from googleapiclient.discovery import build
import gspread

# =================== KONFIGURASI ===================

GROQ_API_KEYS = [
    "MASUKKAN_API_KEY_GROQ_DI_SINI",
]

if IN_COLAB:
    try:
        from google.colab import userdata
        colab_key = userdata.get("GROQ_API_KEY")
        if colab_key and colab_key.strip():
            if not GROQ_API_KEYS or any("MASUKKAN" in k for k in GROQ_API_KEYS):
                GROQ_API_KEYS = [colab_key.strip()]
                print("[🔑] API Key dimuat dari Secrets Colab.")
    except Exception:
        pass

GROQ_MODEL = "meta-llama/llama-4-scout-17b-16e-instruct"
INPUT_SPREADSHEET_ID = "1KgeP2G4B6EQfX4CenmPNCqkno-1ZS8myNkgSYJopdFQ"
OUTPUT_SPREADSHEET_ID = "1P9jqL-ukharkBa24V3qEDyT_28GutMP6Npqn0rc7i3E"
MAX_IMAGE_WIDTH = 900

TARGET_SITES = {
    32: {"site_id": "AM16208465011216N", "nama": "MTSN HUMBANG HASUNDUTAN"},
    35: {"site_id": "AM16208457991225N", "nama": "KANTOR KEPALA DESA SIROMBU"},
    3: {"site_id": "AM16208461781214N", "nama": "RSUD LUKAS"},
    16: {"site_id": "AM16208465031214N", "nama": "LEMBAGA PEMASYARAKATAN KELAS III TELUK DALAM"},
    34: {"site_id": "AM16208461771214N", "nama": "KANTOR KEPALA DESA BOTOHILI SORAKE"},
    40: {"site_id": "AM16208457991214N", "nama": "KANTOR DESA BAWOOTALUA"},
    41: {"site_id": "AM16208461221214N", "nama": "KANTOR KEPALA DESA LAGUNDRI"},
    36: {"site_id": "AM16208459311224N", "nama": "SMA NEGERI 1 LOTU"},
    22: {"site_id": "AM16208465041221N", "nama": "MADRASAH ALIYAH NEGERI 3 PADANG LAWAS"},
}


def clean_site_id(s):
    if not s:
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).strip().upper())


TARGET_NOS = set(TARGET_SITES.keys())
TARGET_SITE_IDS = {v["site_id"] for v in TARGET_SITES.values()}
TARGET_SITE_IDS_CLEAN = {clean_site_id(sid) for sid in TARGET_SITE_IDS}
TARGET_SITE_NAME_MAP = {clean_site_id(v["site_id"]): v["nama"] for v in TARGET_SITES.values()}

# =================== END KONFIGURASI ===================


def get_folder_id(url):
    if not url:
        return None
    match = re.search(r'folders/([a-zA-Z0-9-_]+)', url)
    return match.group(1) if match else None


def get_pdf_from_folder(drive_service, folder_id):
    query = f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false"
    results = drive_service.files().list(
        q=query, fields="files(id, name, createdTime)", orderBy="createdTime desc"
    ).execute()
    files_list = results.get('files', [])
    if not files_list:
        return None, None, None, 0
    pdf_count = len(files_list)
    target_file = files_list[0]
    file_id = target_file['id']
    file_name = target_file['name']
    created_time = target_file.get('createdTime', '')
    request = drive_service.files().get_media(fileId=file_id)
    file_content = request.execute()
    return file_name, file_content, created_time, pdf_count


def get_images_from_folder(drive_service, folder_id, max_width=1280):
    from PIL import Image
    mime_types = ["image/png", "image/jpeg", "image/jpg", "image/webp", "image/bmp", "image/tiff"]
    mime_query = " or ".join([f"mimeType='{m}'" for m in mime_types])
    query = f"'{folder_id}' in parents and ({mime_query}) and trashed=false"
    results = drive_service.files().list(
        q=query, fields="files(id, name, createdTime, mimeType)", orderBy="name asc"
    ).execute()
    files_list = results.get('files', [])
    if not files_list:
        return []
    image_data_list = []
    for f in files_list:
        try:
            request = drive_service.files().get_media(fileId=f['id'])
            content = request.execute()
            img = Image.open(io.BytesIO(content))
            if img.mode not in ('RGB', 'L'):
                img = img.convert('RGB')
            if img.width > max_width:
                ratio = max_width / img.width
                img = img.resize((max_width, int(img.height * ratio)), Image.LANCZOS)
            buffer = io.BytesIO()
            img.save(buffer, format="PNG", optimize=True)
            img_base64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
            image_data_list.append({
                "name": f['name'], "base64": img_base64,
                "created": f.get('createdTime', ''), "mime": f.get('mimeType', '')
            })
        except Exception as e:
            print(f"    [⚠️] Gagal memproses gambar '{f['name']}': {e}")
    return image_data_list


# =================== IMAGE STRATEGY: 3 OVERLAPPING SCAN STRIPS ===================

def pil_to_base64(pil_img, max_width=1000):
    """Konversi PIL Image ke base64 JPEG string (kompresi tinggi)."""
    from PIL import Image
    if pil_img.width > max_width:
        ratio = max_width / pil_img.width
        pil_img = pil_img.resize((max_width, int(pil_img.height * ratio)), Image.LANCZOS)
    buffer = io.BytesIO()
    if pil_img.mode != 'RGB':
        pil_img = pil_img.convert('RGB')
    pil_img.save(buffer, format="JPEG", quality=65)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def create_scan_strips(page_pil_img):
    """Bagi halaman menjadi 3 strip horizontal tumpang tindih untuk scanning BAKTI AKSI.
    Strip 1: 0%–45% (atas)
    Strip 2: 30%–75% (tengah, overlap atas-bawah)
    Strip 3: 55%–100% (bawah)
    Returns: list of base64 strings (3 strips).
    """
    w, h = page_pil_img.size
    strips = []

    # Strip 1: Atas (0% - 45%)
    strip1 = page_pil_img.crop((0, 0, w, int(h * 0.45)))
    strips.append(pil_to_base64(strip1, max_width=1000))

    # Strip 2: Tengah (30% - 75%)
    strip2 = page_pil_img.crop((0, int(h * 0.30), w, int(h * 0.75)))
    strips.append(pil_to_base64(strip2, max_width=1000))

    # Strip 3: Bawah (55% - 100%)
    strip3 = page_pil_img.crop((0, int(h * 0.55), w, h))
    strips.append(pil_to_base64(strip3, max_width=1000))

    return strips


def parse_pdf_layout_deterministically(pdf_bytes):
    """Menganalisis PDF menggunakan pypdf secara deterministik.
    Mengembalikan: (before_idx, after_idx, traffic_idx, layout_type)
    - layout_type: 'stacked' (Before+After di 1 halaman) atau 'single' (terpisah)
    """
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        num_pages = len(reader.pages)
        before_idx = None
        after_idx = None
        traffic_idx = None
        layout_type = "single"

        # Scan setiap halaman untuk keyword, prioritaskan halaman awal
        page_texts = []
        for idx in range(num_pages):
            text = (reader.pages[idx].extract_text() or "").strip().upper()
            page_texts.append(text)

        # Pass 1: Cari halaman yang mengandung BOTH "BEFORE" dan "AFTER" (stacked layout)
        for idx, text in enumerate(page_texts):
            # Hanya cari di halaman 1-2 (indeks 0-1), bukan halaman belakang
            if idx <= 1 and "BEFORE" in text and "AFTER" in text:
                before_idx = idx
                after_idx = idx
                layout_type = "stacked"
                break

        # Pass 2: Jika belum ketemu stacked, cari halaman terpisah
        if layout_type != "stacked":
            for idx, text in enumerate(page_texts):
                if "BEFORE" in text and before_idx is None:
                    before_idx = idx
                elif "AFTER" in text and after_idx is None:
                    after_idx = idx

        # Pass 3: Cari halaman traffic (mengandung "AP1" atau "TRAFFIC" atau "INBOUND") - Scan dari belakang
        for idx in reversed(range(num_pages)):
            if idx != before_idx and idx != after_idx:
                text = page_texts[idx]
                if any(kw in text for kw in ["AP1", "TRAFFIC", "INBOUND", "MONITORING"]):
                    traffic_idx = idx
                    break

        # Fallback jika keyword tidak ditemukan
        if before_idx is None and after_idx is None:
            if num_pages == 1:
                before_idx = 0
                after_idx = 0
                layout_type = "stacked"
            elif num_pages == 2:
                before_idx = 0
                after_idx = 0
                layout_type = "stacked"
                traffic_idx = 1
            elif num_pages == 3:
                before_idx = 0
                after_idx = 1
                traffic_idx = 2
                layout_type = "single"
            else:  # 4+ halaman
                before_idx = 0
                after_idx = 1
                traffic_idx = num_pages - 1  # Halaman terakhir biasanya traffic
                layout_type = "single"
        elif after_idx is None:
            # Before ditemukan tapi After tidak
            after_idx = before_idx + 1 if before_idx + 1 < num_pages else before_idx
            if after_idx == before_idx:
                layout_type = "stacked"

        # Traffic fallback: halaman terakhir jika belum ditemukan
        if traffic_idx is None:
            last_page = num_pages - 1
            if last_page != before_idx and last_page != after_idx:
                traffic_idx = last_page
            elif num_pages > 2:
                traffic_idx = num_pages - 1

        return before_idx, after_idx, traffic_idx, layout_type
    except Exception:
        return 0, 0, None, "stacked"


def _encode_pil_image(pil_img, max_width=1280):
    """Helper: encode PIL image ke base64 JPEG string dengan resize (JPEG format untuk optimalisasi token & bandwidth)."""
    from PIL import Image
    if pil_img.width > max_width:
        ratio = max_width / pil_img.width
        pil_img = pil_img.resize((max_width, int(pil_img.height * ratio)), Image.LANCZOS)
    buffer = io.BytesIO()
    if pil_img.mode != 'RGB':
        pil_img = pil_img.convert('RGB')
    pil_img.save(buffer, format="JPEG", quality=65)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def pdf_to_images(pdf_bytes, max_width=1280):
    """Konversi PDF ke list gambar base64 + 3 scan strips dari halaman After.
    DPI=300.
    Untuk PDF >3 halaman: hanya kirim 3 halaman terpenting (Before, After, Traffic)
    + 3 scan strips = MAKSIMAL 6 gambar agar tidak melebihi batas Groq API.
    Returns: (list_base64, layout_type, num_full_pages)
    """
    from pdf2image import convert_from_bytes
    from PIL import Image

    images = convert_from_bytes(pdf_bytes, dpi=300)
    total_pdf_pages = len(images)

    before_idx, after_idx, traffic_idx, layout_type = parse_pdf_layout_deterministically(pdf_bytes)

    print(f"    [🔍] PDF {total_pdf_pages} halaman | Before: hal.{before_idx+1} | After: hal.{after_idx+1} | "
          f"Traffic: {'hal.'+str(traffic_idx+1) if traffic_idx is not None else 'N/A'} | Layout: {layout_type}")
    print(f"    [🔍] Resolusi PDF: {images[after_idx].width}x{images[after_idx].height}px (DPI=300)")

    # Untuk PDF ≤3 halaman: kirim semua halaman
    # Untuk PDF >3 halaman: hanya kirim Before, After, Traffic (max 3 halaman)
    if total_pdf_pages <= 3:
        encoded_images = [_encode_pil_image(img, max_width) for img in images]
        num_full_pages = len(encoded_images)
        print(f"    [📄] Mengirim semua {num_full_pages} halaman")
    else:
        # Pilih hanya halaman terpenting
        selected_indices = []
        selected_labels = []

        if layout_type == "stacked":
            # Before+After di 1 halaman
            selected_indices.append(after_idx)
            selected_labels.append(f"Before+After (hal.{after_idx+1})")
        else:
            # Before dan After terpisah
            if before_idx is not None:
                selected_indices.append(before_idx)
                selected_labels.append(f"Before (hal.{before_idx+1})")
            if after_idx is not None and after_idx not in selected_indices:
                selected_indices.append(after_idx)
                selected_labels.append(f"After (hal.{after_idx+1})")

        if traffic_idx is not None and traffic_idx not in selected_indices:
            selected_indices.append(traffic_idx)
            selected_labels.append(f"Traffic (hal.{traffic_idx+1})")

        encoded_images = [_encode_pil_image(images[i], max_width) for i in selected_indices]
        num_full_pages = len(encoded_images)
        print(f"    [📄] PDF >3 halaman → seleksi {num_full_pages} halaman terpenting: {', '.join(selected_labels)}")

    # Tambahkan 3 scan strips dari halaman After
    try:
        strips = create_scan_strips(images[after_idx])
        encoded_images.extend(strips)
        print(f"    [🔍] +3 scan strips tumpang tindih (0-45%, 30-75%, 55-100%) dari halaman {after_idx+1}")
    except Exception as e:
        print(f"    [⚠️] Gagal membuat scan strips: {e}")

    return encoded_images, layout_type, num_full_pages


# =================== PROMPT AI ===================

ANALYSIS_PROMPT = """Anda adalah sistem QA PMO otomatis untuk memeriksa laporan perubahan SSID AP1 proyek BAKTI AKSI.

INSTRUKSI — Periksa SEMUA poin berikut dari gambar-gambar yang dikirimkan:

1. IDENTIFIKASI DOKUMEN:
   - Cari "SITE ID" dan "NAMA LOKASI" di bagian atas/header dokumen.

2. CAPTURE BEFORE (biasanya gambar/halaman pertama di PDF):
   - Apakah URL/web address bar browser terlihat?
   - Baca field "* Name" atau "Name()" → ini adalah SSID sebelum perubahan (ssid_before).
   - Apakah nama site di capture before identik/serupa dengan NAMA LOKASI dokumen?

3. CAPTURE AFTER (biasanya gambar/halaman kedua di PDF) — 🚨 PENGECEKAN UTAMA:
   - Apakah URL/web address bar browser terlihat?
   - Apakah nama site di capture after identik/serupa dengan NAMA LOKASI dokumen?
   - Baca field "* Name" → ini adalah ssid_after.
   - 🚨 HITUNG BAGAN "BAKTI AKSI":
     Periksa SELURUH capture AFTER. Hitung berapa BAGAN/AREA TERPISAH yang mengandung kata "BAKTI AKSI" atau "BAKTI AKS".
     Contoh bagan: breadcrumb/header (Wi-Fi > BAKTI AKSI...), field input * Name, judul halaman, dll.
     GUNAKAN GAMBAR SCAN STRIP RESOLUSI TINGGI untuk membaca teks dengan presisi!
     Isi "jumlah_bagan_bakti_aksi" (integer) dan "daftar_bagan_bakti_aksi" (list string).

4. CAPTURE TRAFFIC (biasanya gambar/halaman terakhir di PDF):
   - Apakah judul/label grafik mengandung kata "AP1"? (traffic_contains_ap1)
   - Apakah nama site di judul grafik identik/serupa dengan NAMA LOKASI? (traffic_site_name_identik)
   - Apakah ada PANAH penunjuk pada grafik? (traffic_panah_ada)
   - Ekstrak tanggal dari dekat panah (traffic_tanggal_eksekusi, format DD-MM-YYYY).
   - Apakah ada GRAFIK/data traffic SEBELUM titik panah? (traffic_grafik_before_ada)
     → true jika ada garis/area grafik yang menunjukkan aktivitas sebelum panah.
     → false jika grafik kosong/flatline/tidak ada data sebelum panah.
   - Apakah ada GRAFIK/data traffic SESUDAH titik panah? (traffic_grafik_after_ada)
     → true jika ada garis/area grafik yang menunjukkan aktivitas sesudah panah.
     → false jika grafik kosong/flatline/tidak ada data sesudah panah.
   - Tingkat anomali traffic: bandingkan grafik sebelum vs sesudah panah.
     TIDAK ADA = normal, RENDAH = sedikit turun, SEDANG = drop >50%, TINGGI = mati total.

WAJIB kembalikan output dalam format JSON mentah (tanpa pembungkus markdown):
{
  "site_id": "string",
  "nama_lokasi": "string",
  "ssid_before": "string — teks field Name pada capture Before",
  "ssid_after": "string — teks field Name pada capture After",
  "web_address_before_visible": true/false,
  "web_address_after_visible": true/false,
  "nama_site_cocok_before": true/false,
  "nama_site_cocok_after": true/false,
  "nama_site_cocok_traffic": true/false,
  "nama_site_before_terbaca": "string — nama site yang terbaca di capture Before",
  "nama_site_after_terbaca": "string — nama site yang terbaca di capture After",
  "nama_site_traffic_terbaca": "string — nama site yang terbaca di judul grafik traffic",
  "jumlah_bagan_bakti_aksi": 0,
  "daftar_bagan_bakti_aksi": [],
  "traffic_contains_ap1": true/false,
  "traffic_site_name_identik": true/false,
  "traffic_panah_ada": true/false,
  "traffic_tanggal_eksekusi": "DD-MM-YYYY",
  "traffic_grafik_before_ada": true/false,
  "traffic_grafik_after_ada": true/false,
  "traffic_anomali_level": "TIDAK ADA/RENDAH/SEDANG/TINGGI",
  "catatan_ai": "string penjelasan singkat dari AI"
}"""


def analyze_with_groq(image_base64_list, api_key, layout_type="single", num_full_pages=2,
                      max_retries=5, initial_delay=25, target_site_id="", target_nama=""):
    """Kirim gambar ke Groq API dengan deskripsi struktur dinamis + konteks target."""
    content = []
    for img_b64 in image_base64_list:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}
        })

    num_imgs = len(image_base64_list)
    num_strips = num_imgs - num_full_pages  # biasanya 3

    # ── Injeksi Konteks Target ──
    # Mencegah AI salah membaca header proyek (misal "2024 SL BAKTI") sebagai site_id/nama_lokasi
    target_context = ""
    if target_site_id or target_nama:
        target_context = (
            f"🎯 KONTEKS TARGET — Dokumen ini untuk site berikut:\n"
            f"   SITE ID TARGET    : {target_site_id}\n"
            f"   NAMA LOKASI TARGET: {target_nama}\n"
            f"   ⚠️ PENTING: Gunakan nilai di atas sebagai site_id dan nama_lokasi dalam JSON output.\n"
            f"   Jangan keliru dengan label proyek (misal '2024 SL BAKTI') yang ada di header PDF.\n"
            f"   Untuk field 'nama_site_cocok_*', bandingkan nama site yang TERBACA di gambar\n"
            f"   dengan NAMA LOKASI TARGET di atas ('{target_nama}').\n"
            f"\n{'=' * 50}\n\n"
        )

    structure_desc = f"🚨 STRUKTUR GAMBAR — Total {num_imgs} gambar:\n"

    # Deskripsikan halaman penuh
    if layout_type == "stacked":
        if num_full_pages == 1:
            structure_desc += "- Gambar 1: Halaman penuh (BEFORE + AFTER digabung, Before di atas, After di bawah)\n"
        elif num_full_pages == 2:
            structure_desc += (
                "- Gambar 1: Halaman penuh (BEFORE + AFTER digabung dalam 1 halaman)\n"
                "- Gambar 2: Halaman TRAFFIC MONITORING\n"
            )
        else:
            for i in range(num_full_pages):
                structure_desc += f"- Gambar {i+1}: Halaman PDF ke-{i+1}\n"
    else:  # single / multi-page
        page_labels = ["BEFORE", "AFTER", "TRAFFIC MONITORING"]
        for i in range(num_full_pages):
            label = page_labels[i] if i < len(page_labels) else f"Halaman {i+1}"
            structure_desc += f"- Gambar {i+1}: Screenshot {label}\n"

    # Deskripsikan scan strips
    if num_strips >= 3:
        s = num_full_pages + 1
        structure_desc += (
            f"- Gambar {s}: 🔍 SCAN STRIP ATAS (0-45%) — resolusi tinggi dari halaman After\n"
            f"- Gambar {s+1}: 🔍 SCAN STRIP TENGAH (30-75%) — resolusi tinggi dari halaman After\n"
            f"- Gambar {s+2}: 🔍 SCAN STRIP BAWAH (55-100%) — resolusi tinggi dari halaman After\n"
            f"\n🚨 GUNAKAN SCAN STRIP (Gambar {s}-{s+2}) UNTUK MEMBACA TEKS SSID DAN BAKTI AKSI!\n"
            f"Strip ini jauh lebih tajam dari halaman penuh. Baca teks dari sini!\n"
        )

    structure_desc += "\n" + "=" * 50 + "\n\n"
    full_prompt = target_context + structure_desc + ANALYSIS_PROMPT

    content.append({"type": "text", "text": full_prompt})

    payload = {
        "model": GROQ_MODEL,
        "messages": [{"role": "user", "content": content}],
        "temperature": 0.1,
        "max_tokens": 2048,
    }

    delay = initial_delay
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=payload, timeout=120
            )
            if response.status_code == 429:
                print(f"    [!] Rate limit (429). Menunggu {delay}s ({attempt}/{max_retries})...")
                time.sleep(delay)
                delay = min(delay * 2, 120)
                continue
            if response.status_code in [500, 503]:
                print(f"    [!] Server error ({response.status_code}). Menunggu {delay}s ({attempt}/{max_retries})...")
                time.sleep(delay)
                delay = min(delay * 2, 120)
                continue
            response.raise_for_status()
            result_json = response.json()
            if "choices" in result_json and result_json["choices"]:
                response_text = result_json["choices"][0]["message"]["content"]
                clean_text = response_text.strip()
                if clean_text.startswith("```"):
                    clean_text = re.sub(r'^```(?:json)?\s*', '', clean_text)
                    clean_text = re.sub(r'\s*```$', '', clean_text)
                try:
                    return json.loads(clean_text)
                except json.JSONDecodeError:
                    json_match = re.search(r'\{[\s\S]*\}', response_text)
                    if json_match:
                        try:
                            return json.loads(json_match.group())
                        except json.JSONDecodeError:
                            return None
            return None
        except Exception as e:
            if attempt == max_retries:
                print(f"    [❌] Groq API Error pada percobaan terakhir: {e}")
                if 'response' in locals() and response is not None:
                    print(f"    [❌] Status Code: {response.status_code}")
                    print(f"    [❌] Response Body: {response.text}")
                return None
            time.sleep(delay)
            delay = min(delay * 2, 120)
    return None


# =================== BUSINESS RULES (6 KRITERIA) ===================

def _bool(val):
    """Normalisasi boolean dari berbagai format AI response."""
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        return val.strip().upper() in ("YA", "TRUE", "YES", "1")
    return bool(val)


def apply_business_rules(result, file_name="", kolom_lengkap=True, kolom_kosong_list=None,
                         kolom_ad_ah=None, target_nama="", target_site_id=""):
    """Terapkan 6 kriteria validasi. Returns result dict with updated status & notes."""
    if not result:
        return result

    if kolom_kosong_list is None:
        kolom_kosong_list = []
    if kolom_ad_ah is None:
        kolom_ad_ah = {}

    # ── Override site_id dan nama_lokasi dari AI dengan data target yang diketahui ──
    # AI sering salah membaca header proyek (misal "2024 SL BAKTI") sebagai site_id/nama_lokasi
    if target_site_id:
        ai_site_id = result.get("site_id", "")
        if ai_site_id and clean_site_id(ai_site_id) != clean_site_id(target_site_id):
            print(f"    [🔄] Override site_id: AI='{ai_site_id}' → Target='{target_site_id}'")
        result["site_id"] = target_site_id
    if target_nama:
        ai_nama = result.get("nama_lokasi", "")
        if ai_nama and clean_site_id(ai_nama) != clean_site_id(target_nama):
            print(f"    [🔄] Override nama_lokasi: AI='{ai_nama}' → Target='{target_nama}'")
        result["nama_lokasi"] = target_nama

    fail_reasons = []
    notes = []

    # ── ① Nama File PDF vs Nama Site Target (catatan saja) ──
    if file_name and target_nama:
        clean_fname = re.sub(r'[^A-Z0-9]', '', file_name.upper())
        clean_target = re.sub(r'[^A-Z0-9]', '', target_nama.upper())
        fname_similarity = difflib.SequenceMatcher(None, clean_fname, clean_target).ratio()
        if clean_target in clean_fname or fname_similarity >= 0.50:
            notes.append(f"① Nama File PDF: ✅ Identik ('{file_name}')")
        else:
            notes.append(f"① Nama File PDF: ⚠️ Tidak identik — file: '{file_name}' vs target: '{target_nama}'")
    else:
        notes.append(f"① Nama File PDF: ℹ️ '{file_name or '-'}'")

    # ── ② Kolom AD-AH Spreadsheet (FAIL jika ada yang kosong) ──
    if kolom_lengkap:
        kolom_summary = " | ".join([f"{k}: {v}" for k, v in kolom_ad_ah.items()])
        notes.append(f"② Kolom AD-AH: ✅ Lengkap ({kolom_summary})")
    else:
        fail_reasons.append(f"Kolom spreadsheet kosong: {', '.join(kolom_kosong_list)}")
        filled = {k: v for k, v in kolom_ad_ah.items() if v}
        filled_str = " | ".join([f"{k}: {v}" for k, v in filled.items()]) if filled else "-"
        notes.append(f"② Kolom AD-AH: ❌ Kosong: [{', '.join(kolom_kosong_list)}] | Terisi: {filled_str}")

    # ── ③ Web Address Before & After (catatan saja) ──
    web_before = _bool(result.get("web_address_before_visible", False))
    web_after = _bool(result.get("web_address_after_visible", False))
    notes.append(f"③ Web Address: Before {'✅' if web_before else '❌'} | After {'✅' if web_after else '❌'}")

    # ── ④ Nama Site Identik di Before, After, Traffic (catatan saja) ──
    cocok_before = _bool(result.get("nama_site_cocok_before", False))
    cocok_after = _bool(result.get("nama_site_cocok_after", False))
    cocok_traffic = _bool(result.get("nama_site_cocok_traffic", False))

    site_before = result.get("nama_site_before_terbaca", result.get("ssid_before", "-"))
    site_after = result.get("nama_site_after_terbaca", result.get("ssid_after", "-"))
    site_traffic = result.get("nama_site_traffic_terbaca", "-")

    notes.append(
        f"④ Nama Site: Before {'✅' if cocok_before else '⚠️'}({site_before}) | "
        f"After {'✅' if cocok_after else '⚠️'}({site_after}) | "
        f"Traffic {'✅' if cocok_traffic else '⚠️'}({site_traffic})"
    )

    # ── ⑤ BAKTI AKSI ≥2 bagan di capture After (FAIL jika <2) ──
    jumlah_bagan = result.get("jumlah_bagan_bakti_aksi", 0)
    if not isinstance(jumlah_bagan, int):
        try:
            jumlah_bagan = int(jumlah_bagan)
        except (ValueError, TypeError):
            jumlah_bagan = 0

    # Fallback: jika AI lapor 0 tapi ssid_after jelas mengandung BAKTI
    ssid_after_val = str(result.get("ssid_after", "")).upper()
    if jumlah_bagan == 0 and "BAKTI" in re.sub(r'[^A-Z0-9]', '', ssid_after_val):
        jumlah_bagan = 1
        result["jumlah_bagan_bakti_aksi"] = 1

    daftar_bagan = result.get("daftar_bagan_bakti_aksi", [])
    if not isinstance(daftar_bagan, list):
        daftar_bagan = []

    bagan_detail = ", ".join(daftar_bagan) if daftar_bagan else "-"
    if jumlah_bagan >= 2:
        notes.append(f"⑤ BAKTI AKSI: ✅ {jumlah_bagan} bagan ({bagan_detail})")
    else:
        fail_reasons.append(f"BAKTI AKSI hanya {jumlah_bagan} bagan di capture After (minimal 2)")
        notes.append(f"⑤ BAKTI AKSI: ❌ Hanya {jumlah_bagan} bagan ({bagan_detail})")

    # ── ⑥ Traffic: AP1, Panah, Grafik, Anomali ──
    ap1_ada = _bool(result.get("traffic_contains_ap1", False))
    panah_ada = _bool(result.get("traffic_panah_ada", False))
    grafik_before = _bool(result.get("traffic_grafik_before_ada", False))
    grafik_after = _bool(result.get("traffic_grafik_after_ada", False))
    site_match = _bool(result.get("traffic_site_name_identik", False))
    anomali = str(result.get("traffic_anomali_level", "TIDAK ADA")).upper()
    tgl_eks = result.get("traffic_tanggal_eksekusi", "-")

    # 6a: FAIL jika AP1 tidak ada ATAU panah tidak ada
    if not ap1_ada:
        fail_reasons.append("Traffic tidak mengandung kata AP1")
    if not panah_ada:
        fail_reasons.append("Panah penunjuk tidak ada pada grafik traffic")

    # 6b: FAIL jika tidak ada grafik sebelum atau sesudah panah
    if not grafik_before:
        fail_reasons.append("Tidak ada grafik/data traffic SEBELUM titik panah")
    if not grafik_after:
        fail_reasons.append("Tidak ada grafik/data traffic SESUDAH titik panah")

    notes.append(
        f"⑥ Traffic: AP1 {'✅' if ap1_ada else '❌'} | "
        f"Site {'✅' if site_match else '⚠️'} | "
        f"Panah {'✅' if panah_ada else '❌'} | "
        f"Grafik Before/After {'✅' if grafik_before else '❌'}/{'✅' if grafik_after else '❌'} | "
        f"Anomali: {anomali} | Tgl: {tgl_eks}"
    )

    # ── Tentukan Status Final ──
    if fail_reasons:
        result["status_check"] = "TIDAK SESUAI"
    else:
        result["status_check"] = "SESUAI"

    # Gabungkan catatan
    ai_catatan = result.get("catatan_ai", result.get("catatan", ""))
    all_notes = "\n".join(notes)
    if fail_reasons:
        fail_text = " | ".join(fail_reasons)
        result["catatan"] = f"[GAGAL: {fail_text}]\n{all_notes}"
        if ai_catatan and ai_catatan != "-":
            result["catatan"] += f"\n[AI: {ai_catatan}]"
    else:
        result["catatan"] = all_notes
        if ai_catatan and ai_catatan != "-":
            result["catatan"] += f"\n[AI: {ai_catatan}]"

    result["_notes_list"] = notes
    result["_fail_reasons"] = fail_reasons
    return result


# =================== OUTPUT FUNCTIONS ===================

def parse_to_output_dict(no_val, idx_in, site_id_in, result, tanggal_submit, kolom_lengkap, kolom_kosong_list):
    return {
        "No": no_val,
        "Baris Sheet": idx_in,
        "Site ID": result.get("site_id", site_id_in),
        "Nama Lokasi": result.get("nama_lokasi", "-"),
        "Format File": "PDF",
        "SSID Before": result.get("ssid_before", "-"),
        "SSID After": result.get("ssid_after", "-"),
        "Tgl Eksekusi": result.get("traffic_tanggal_eksekusi", "-"),
        "Tgl Submit": tanggal_submit,
        "Kolom AD-AH": "Lengkap" if kolom_lengkap else f"Kosong: {', '.join(kolom_kosong_list)}",
        "Bagan BAKTI AKSI": result.get("jumlah_bagan_bakti_aksi", 0),
        "Traffic AP1": "✅" if _bool(result.get("traffic_contains_ap1", False)) else "❌",
        "Traffic Panah": "✅" if _bool(result.get("traffic_panah_ada", False)) else "❌",
        "Traffic Grafik": f"B:{'✅' if _bool(result.get('traffic_grafik_before_ada', False)) else '❌'} A:{'✅' if _bool(result.get('traffic_grafik_after_ada', False)) else '❌'}",
        "Hasil Verifikasi PMO": result.get("status_check", "TIDAK SESUAI"),
        "Catatan": result.get("catatan", "-"),
    }


def create_result_row(no_val, baris, site_id, nama_lokasi, status, catatan, kolom_lengkap=False, format_file="-"):
    return {
        "No": no_val, "Baris Sheet": baris, "Site ID": site_id,
        "Nama Lokasi": nama_lokasi, "Hasil Verifikasi PMO": status, "Catatan": catatan,
    }


def print_result(result, file_name="", kolom_lengkap=True, kolom_kosong_list=None):
    """Print summary per 6 kriteria."""
    status = result.get("status_check", "?")
    icon = "✅" if status == "SESUAI" else "❌"

    print(f"\n    {'═' * 55}")
    print(f"    📋 HASIL ANALISIS — {result.get('site_id', '-')}")
    print(f"    {'═' * 55}")
    print(f"    Nama Lokasi : {result.get('nama_lokasi', '-')}")
    print(f"    SSID Before : {result.get('ssid_before', '-')}")
    print(f"    SSID After  : {result.get('ssid_after', '-')}")
    print(f"    {'─' * 55}")

    # Print per-kriteria notes
    notes_list = result.get("_notes_list", [])
    for note in notes_list:
        print(f"    {note}")

    print(f"    {'─' * 55}")
    fail_reasons = result.get("_fail_reasons", [])
    if fail_reasons:
        print(f"    {icon} STATUS: {status}")
        for fr in fail_reasons:
            print(f"    ⛔ {fr}")
    else:
        print(f"    {icon} STATUS: {status}")

    ai_note = result.get("catatan_ai", result.get("catatan", ""))
    # Only print AI note if it's a simple string (not the full compiled notes)
    if ai_note and ai_note != "-" and "\n" not in str(ai_note):
        print(f"    💬 AI: {ai_note}")

    print(f"    {'═' * 55}")


# ================= FUNGSI UTAMA =================

def run_dry_run():
    total_sites = len(TARGET_SITES)
    print("=" * 70)
    print("  SSID VALIDATOR v3.1 — COLAB EDITION (GROQ VISION)")
    print(f"  {total_sites} Site | DPI=300 | 3 Scan Strips | 6 Kriteria Validasi")
    print("=" * 70)

    api_key_active = None
    for k in GROQ_API_KEYS:
        if k and k.strip() and "MASUKKAN" not in k:
            api_key_active = k.strip()
            break
    if not api_key_active:
        print("\n[❌] Silakan isi GROQ_API_KEYS")
        return

    print(f"[*] Model: {GROQ_MODEL}")
    sleep_time = 25.0
    print(f"[*] Delay: {sleep_time}s antar request")

    if IN_COLAB:
        print("\n[*] Otentikasi Google Colab...")
        auth.authenticate_user()
        creds, _ = default()
    else:
        print("\n[*] Application Default Credentials...")
        creds, _ = default()

    gc = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)

    try:
        sh_in = gc.open_by_key(INPUT_SPREADSHEET_ID)
        sheet_in = sh_in.get_worksheet(0)
        print(f"[✅] Sheet Input: '{sh_in.title}'")
    except Exception as e:
        print(f"[❌] Gagal buka Spreadsheet Input: {e}")
        return

    try:
        sh_out = gc.open_by_key(OUTPUT_SPREADSHEET_ID)
        sheet_out = sh_out.get_worksheet(0)
        data_out = sheet_out.get_all_values()
        print(f"[✅] Sheet Output: '{sh_out.title}'")
    except Exception as e:
        print(f"[❌] Gagal buka Spreadsheet Output: {e}")
        return

    # Buat map Site ID -> baris di sheet output (Kolom H = indeks 7, Kolom I = indeks 8)
    col_h_idx = 7  # Kolom H (Site ID)
    col_i_idx = 8  # Kolom I (Nama Lokasi)

    out_site_map = {}  # Map: Cleaned Site ID (H) -> Row index
    out_name_map = {}  # Map: Cleaned Nama Lokasi (I) -> Row index

    for r_idx, r_data in enumerate(data_out, start=1):
        # Site ID (Kolom H)
        if len(r_data) > col_h_idx:
            sid_val = r_data[col_h_idx].strip()
            if sid_val:
                out_site_map[clean_site_id(sid_val)] = r_idx
        # Nama Lokasi (Kolom I)
        if len(r_data) > col_i_idx:
            name_val = r_data[col_i_idx].strip()
            if name_val:
                out_name_map[clean_site_id(name_val)] = r_idx

    data_in = sheet_in.get_all_values()
    headers = data_in[1] if len(data_in) > 1 else data_in[0]

    no_col_idx = 0
    for i, h in enumerate(headers):
        if h.strip().lower() in ["no", "no.", "nomor"]:
            no_col_idx = i
            break

    site_col_idx = 6   # Kolom G
    link_col_idx = 35  # Kolom AJ

    results_list = []
    processed_count = 0
    total_targets = len(TARGET_SITES)
    cells_to_update = []

    print(f"\n{'=' * 70}")
    print(f"  MULAI ANALISIS — {total_targets} site")
    print(f"{'=' * 70}")

    for idx_in, row_in in enumerate(data_in[2:], start=3):
        no_val = row_in[no_col_idx].strip() if len(row_in) > no_col_idx else ""
        site_id_in = row_in[site_col_idx].strip() if len(row_in) > site_col_idx else ""

        site_id_in_clean = clean_site_id(site_id_in)
        if not (site_id_in_clean and site_id_in_clean in TARGET_SITE_IDS_CLEAN):
            continue

        processed_count += 1
        target_nama = TARGET_SITE_NAME_MAP.get(site_id_in_clean, "")

        print(f"\n{'─' * 70}")
        print(f"  [{processed_count}/{total_targets}] No: {no_val} | {site_id_in} | {target_nama}")
        print(f"{'─' * 70}")

        # Helper untuk mencari baris output secara robust
        def find_row_output():
            # 1. Coba pencocokan Site ID persis (cleaned)
            if site_id_in_clean in out_site_map:
                return out_site_map[site_id_in_clean]

            # 2. Coba pencocokan Site ID 14-char prefix
            for k_sid, r_idx in out_site_map.items():
                if len(k_sid) >= 14 and len(site_id_in_clean) >= 14:
                    if k_sid[:14] == site_id_in_clean[:14]:
                        print(f"    [⚠️] Site ID cocok via prefix: '{site_id_in}' -> row '{r_idx}' di output (Kolom H)")
                        return r_idx

            # 3. Coba pencocokan Nama Lokasi persis (cleaned)
            clean_target_name = clean_site_id(target_nama)
            if clean_target_name in out_name_map:
                r_idx = out_name_map[clean_target_name]
                print(f"    [⚠️] Site ID '{site_id_in}' cocok via Nama Lokasi: '{target_nama}' -> row '{r_idx}' di output (Kolom I)")
                return r_idx

            # 4. Coba pencocokan Nama Lokasi parsial/longgar
            for k_name, r_idx in out_name_map.items():
                if len(k_name) > 8 and len(clean_target_name) > 8:
                    if k_name in clean_target_name or clean_target_name in k_name:
                        print(f"    [⚠️] Site ID '{site_id_in}' cocok via nama parsial: '{target_nama}' -> row '{r_idx}' di output (Kolom I)")
                        return r_idx

            return None

        # Helper untuk writeback list
        def queue_writeback(status_am_val, fail_notes_val):
            row_out_idx = find_row_output()
            if row_out_idx:
                date_str = datetime.now().strftime("%d-%b-%y")  # Format: 10-Jul-26
                cells_to_update.append(gspread.Cell(row=row_out_idx, col=37, value=date_str))
                cells_to_update.append(gspread.Cell(row=row_out_idx, col=38, value="AMEL"))
                cells_to_update.append(gspread.Cell(row=row_out_idx, col=39, value=status_am_val))
                cells_to_update.append(gspread.Cell(row=row_out_idx, col=40, value=fail_notes_val))
                print(f"    [📝] Menyiapkan write-back ke baris {row_out_idx}: AK={date_str}, AL=AMEL, AM={status_am_val}, AN={fail_notes_val[:40]}")
            else:
                print(f"    [⚠️] Site ID '{site_id_in}' tidak ditemukan di sheet output (Kolom H & I). Skip write-back.")

        # Baca kolom AD-AH
        kolom_ad_ah = {}
        col_labels = {29: "AD", 30: "AE", 31: "AF", 32: "AG", 33: "AH"}
        kolom_lengkap = True
        kolom_kosong_list = []
        for ci, label in col_labels.items():
            val = row_in[ci].strip() if len(row_in) > ci else ""
            kolom_ad_ah[label] = val
            if not val:
                kolom_lengkap = False
                kolom_kosong_list.append(label)

        link_dokumen = row_in[link_col_idx].strip() if len(row_in) > link_col_idx else ""
        if not link_dokumen:
            results_list.append(create_result_row(no_val, idx_in, site_id_in, target_nama, "TIDAK SESUAI", "Link GDrive kosong"))
            queue_writeback("NOK", "[FAIL: Link GDrive kosong]")
            continue

        folder_id = get_folder_id(link_dokumen)
        if not folder_id:
            results_list.append(create_result_row(no_val, idx_in, site_id_in, target_nama, "TIDAK SESUAI", "Format URL GDrive tidak valid"))
            queue_writeback("NOK", "[FAIL: Format URL GDrive tidak valid]")
            continue

        try:
            file_name, file_content, created_time, pdf_count = get_pdf_from_folder(drive_service, folder_id)

            if not file_content:
                print(f"    [⚠️] PDF tidak ditemukan. Cari gambar...")
                image_files = get_images_from_folder(drive_service, folder_id, max_width=MAX_IMAGE_WIDTH)
                if not image_files:
                    results_list.append(create_result_row(no_val, idx_in, site_id_in, target_nama, "TIDAK SESUAI", "Tidak ada file di folder GDrive"))
                    queue_writeback("NOK", "[FAIL: Tidak ada file di folder GDrive]")
                    continue

                img_b64s = [img['base64'] for img in image_files]
                created_time = image_files[0].get('created', '') if image_files else ''
                result = analyze_with_groq(img_b64s, api_key_active, layout_type="single",
                                           num_full_pages=len(img_b64s), max_retries=5,
                                           target_site_id=site_id_in, target_nama=target_nama)
                if result:
                    result = apply_business_rules(result, file_name=image_files[0].get('name', ''),
                                                  kolom_lengkap=kolom_lengkap, kolom_kosong_list=kolom_kosong_list,
                                                  kolom_ad_ah=kolom_ad_ah, target_nama=target_nama,
                                                  target_site_id=site_id_in)
                    # Override: gambar terpisah selalu FAIL
                    result["status_check"] = "TIDAK SESUAI"
                    result["_fail_reasons"] = result.get("_fail_reasons", []) + ["File berupa gambar terpisah, bukan 1 PDF"]

                    status_am = "NOK"
                    fail_notes = f"[FAIL: {', '.join(result.get('_fail_reasons', []))}]"
                    ai_msg = result.get("catatan_ai", "")
                    if ai_msg and ai_msg != "-":
                        fail_notes += f" {ai_msg}"

                    tanggal_submit = created_time[:10] if created_time else "-"
                    if tanggal_submit != "-" and "-" in tanggal_submit:
                        parts = tanggal_submit.split("-")
                        tanggal_submit = f"{parts[2]}-{parts[1]}-{parts[0]}"
                    print_result(result, file_name=image_files[0].get('name', ''))
                    results_list.append(parse_to_output_dict(no_val, idx_in, site_id_in, result, tanggal_submit, kolom_lengkap, kolom_kosong_list))
                    queue_writeback(status_am, fail_notes)
                else:
                    results_list.append(create_result_row(no_val, idx_in, site_id_in, target_nama, "ERROR", "Groq API gagal"))
                    queue_writeback("NOK", "[ERROR: Groq API gagal menganalisis gambar]")
                continue

            print(f"    [📄] PDF: '{file_name}' ({pdf_count} file)")
            pdf_pages, detected_layout, num_full_pages = pdf_to_images(file_content, max_width=MAX_IMAGE_WIDTH)
            print(f"    [📄] Total gambar: {len(pdf_pages)} ({num_full_pages} halaman + {len(pdf_pages)-num_full_pages} strips)")

            result = analyze_with_groq(pdf_pages, api_key_active, layout_type=detected_layout,
                                        num_full_pages=num_full_pages,
                                        target_site_id=site_id_in, target_nama=target_nama)

            if result:
                result = apply_business_rules(result, file_name=file_name,
                                              kolom_lengkap=kolom_lengkap, kolom_kosong_list=kolom_kosong_list,
                                              kolom_ad_ah=kolom_ad_ah, target_nama=target_nama,
                                              target_site_id=site_id_in)

                status_check = result.get("status_check", "TIDAK SESUAI")
                status_am = "OK" if status_check == "SESUAI" else "NOK"
                fail_notes = ""
                if status_am == "NOK":
                    fail_reasons = result.get("_fail_reasons", [])
                    ai_msg = result.get("catatan_ai", "")
                    if fail_reasons:
                        fail_notes = f"[FAIL: {', '.join(fail_reasons)}]"
                        if ai_msg and ai_msg != "-":
                            fail_notes += f" {ai_msg}"
                    else:
                        fail_notes = ai_msg if ai_msg != "-" else "Gagal verifikasi"

                tanggal_submit = created_time[:10] if created_time else "-"
                if tanggal_submit != "-":
                    parts = tanggal_submit.split("-")
                    tanggal_submit = f"{parts[2]}-{parts[1]}-{parts[0]}"

                print_result(result, file_name=file_name)
                results_list.append(parse_to_output_dict(no_val, idx_in, site_id_in, result, tanggal_submit, kolom_lengkap, kolom_kosong_list))
                queue_writeback(status_am, fail_notes)
            else:
                results_list.append(create_result_row(no_val, idx_in, site_id_in, target_nama, "ERROR", "Groq API gagal"))
                queue_writeback("NOK", "[ERROR: Groq API gagal menganalisis PDF]")

        except Exception as e:
            results_list.append(create_result_row(no_val, idx_in, site_id_in, target_nama, "ERROR", f"Error: {str(e)}"))
            queue_writeback("NOK", f"[ERROR: {str(e)}]")

        if processed_count < total_targets:
            print(f"    [⏳] Delay {sleep_time}s...")
            time.sleep(sleep_time)

    # ── CEK TARGET SITE YANG TIDAK DITEMUKAN ──
    processed_site_ids = {clean_site_id(r.get("Site ID", "")) for r in results_list if r.get("Site ID")}
    missing_sites = []
    for tgt_no, tgt_info in TARGET_SITES.items():
        tgt_clean = clean_site_id(tgt_info["site_id"])
        if tgt_clean not in processed_site_ids:
            missing_sites.append(tgt_info)

    if missing_sites:
        print(f"\n{'=' * 70}")
        print(f"  ⚠️ TARGET SITE TIDAK DITEMUKAN DI SPREADSHEET INPUT")
        print(f"{'=' * 70}")
        for ms in missing_sites:
            print(f"  ❓ {ms['site_id']} — {ms['nama']}")
        print(f"  ℹ️ Kemungkinan: Site ID tidak ada di spreadsheet input, atau format berbeda.")
        print(f"{'─' * 70}")

    # ── RINGKASAN AKHIR ──
    if results_list:
        sesuai = sum(1 for r in results_list if r.get("Hasil Verifikasi PMO") == "SESUAI")
        tidak = sum(1 for r in results_list if r.get("Hasil Verifikasi PMO") == "TIDAK SESUAI")
        error = sum(1 for r in results_list if r.get("Hasil Verifikasi PMO") == "ERROR")

        print(f"\n{'=' * 70}")
        print(f"  📊 RINGKASAN — ✅ {sesuai} OK | ❌ {tidak} NOK | ⚠️ {error} ERROR")
        print(f"{'=' * 70}")
        for r in results_list:
            s_id = r.get("Site ID", "-")
            s_name = r.get("Nama Lokasi", "-")
            pmo = r.get("Hasil Verifikasi PMO", "?")
            icon = "✅" if pmo == "SESUAI" else ("❌" if pmo == "TIDAK SESUAI" else "⚠️")
            print(f"  {icon} {s_id} — {s_name}")
        print(f"{'─' * 70}")

    # ── BATCH UPDATE GOOGLE SHEET ──
    if cells_to_update:
        print(f"\n{'=' * 70}")
        print(f"  📝 WRITE-BACK KE GOOGLE SPREADSHEET")
        print(f"{'=' * 70}")
        print(f"    [*] Mengupdate {len(cells_to_update) // 4} baris data ke Google Sheet...")
        try:
            sheet_out.update_cells(cells_to_update)
            print("    [✅] Write-back ke Google Sheet sukses!")
        except Exception as e:
            print(f"    [❌] Gagal write-back: {e}")

    print(f"\n{'=' * 70}")
    print(f"  SELESAI — {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    run_dry_run()
