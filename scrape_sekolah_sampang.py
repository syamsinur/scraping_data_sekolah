"""
Scraper sekolah Kabupaten Sampang melalui API Sekolah Kita Kemendikdasmen.

Instalasi:
    python -m pip install requests pandas openpyxl

Jalankan:
    python scrape_sekolah_sampang.py

Keamanan/resume:
- Folder output selalu berada di sebelah file script, bukan bergantung pada folder terminal.
- Daftar ID disimpan setiap selesai mengambil satu halaman.
- Setelah semua halaman listing selesai, listing_complete.json dibuat.
- Detail setiap sekolah langsung ditulis ke detail_checkpoint.jsonl.
- JSONL di-flush dan fsync setelah setiap sekolah berhasil.
- CSV dan XLSX disimpan secara atomik dengan file sementara yang berekstensi benar.
- Saat dijalankan ulang, detail yang sudah berhasil di checkpoint akan dilewati.
"""

from __future__ import annotations

import json
import os
import random
import time
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ---------------------------------------------------------------------------
# KONFIGURASI
# ---------------------------------------------------------------------------

BASE_URL = "https://sekolah.data.kemendikdasmen.go.id"
SEARCH_URL = f"{BASE_URL}/v1/sekolah-service/sekolah/cari-sekolah"
DETAIL_URL = f"{BASE_URL}/v1/sekolah-service/sekolah/full-detail"

START_PAGE = 0
END_PAGE = 54                 # inklusif: halaman 0 sampai 54
PAGE_SIZE = 48
EXPECTED_TOTAL_IDS = 2606  # jumlah yang dilaporkan API saat scraper dibuat
AUTOSAVE_EVERY = 10            # autosave tabel setiap 10 detail baru
REQUEST_DELAY_SECONDS = 0.8
TIMEOUT_SECONDS = 45

# Penting: lokasi output tetap di sebelah file .py ini.
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "hasil_scraping_sampang"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

IDS_FILE = OUTPUT_DIR / "sekolah_ids.json"
LISTING_COMPLETE_FILE = OUTPUT_DIR / "listing_complete.json"
CHECKPOINT_FILE = OUTPUT_DIR / "detail_checkpoint.jsonl"
CSV_FILE = OUTPUT_DIR / "data_sekolah_sampang.csv"
XLSX_FILE = OUTPUT_DIR / "data_sekolah_sampang.xlsx"
ERRORS_FILE = OUTPUT_DIR / "gagal_scraping.jsonl"

PAYLOAD_TEMPLATE = {
    "page": 0,
    "size": PAGE_SIZE,
    "keyword": "",
    "kabupaten_kota": "Kab. Sampang",
    "bentuk_pendidikan": (
        "SD, SMA, SMK, SMP, MA, MI, MTs, Pondok Pesantren, RA, "
        "Sekolah Rakyat, TK"
    ),
    "status_sekolah": "",
}


# ---------------------------------------------------------------------------
# SESSION HTTP DENGAN RETRY
# ---------------------------------------------------------------------------

def build_session() -> requests.Session:
    session = requests.Session()

    retry = Retry(
        total=5,
        connect=5,
        read=5,
        status=5,
        backoff_factor=1.2,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "POST"}),
        respect_retry_after_header=True,
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=4,
        pool_maxsize=4,
    )
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/130.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Referer": f"{BASE_URL}/sekolah",
    })
    return session


# ---------------------------------------------------------------------------
# FUNGSI PENYIMPANAN YANG LEBIH AMAN
# ---------------------------------------------------------------------------

def append_jsonl_durable(path: Path, obj: dict[str, Any]) -> None:
    """Tambahkan satu record JSONL, lalu flush dan fsync melalui file writable."""
    path.parent.mkdir(parents=True, exist_ok=True)
    line = json.dumps(obj, ensure_ascii=False, default=str) + "\n"

    # Mode 'a' adalah mode tulis; jangan gunakan 'rb' untuk fsync.
    with path.open("a", encoding="utf-8", newline="") as file:
        file.write(line)
        file.flush()
        os.fsync(file.fileno())


def atomic_write_text(path: Path, content: str) -> None:
    """Tulis file sementara, sinkronkan, lalu ganti file tujuan secara atomik."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(path.name + ".tmp")

    with temp_path.open("w", encoding="utf-8", newline="") as file:
        file.write(content)
        file.flush()
        os.fsync(file.fileno())

    os.replace(temp_path, path)


def atomic_write_json(path: Path, obj: Any) -> None:
    content = json.dumps(obj, ensure_ascii=False, indent=2, default=str)
    atomic_write_text(path, content)


def fsync_existing_file(path: Path) -> None:
    """
    Sinkronkan file yang sudah dibuat library (Pandas/openpyxl).
    Gunakan r+b agar descriptor bukan read-only, untuk menghindari EBADF di Windows.
    """
    with path.open("r+b") as file:
        file.flush()
        os.fsync(file.fileno())


def save_table_files(records: list[dict[str, Any]]) -> None:
    """
    Simpan CSV dan XLSX lewat file sementara.
    Suffix sementara tetap diakhiri .csv/.xlsx agar Pandas mengenali formatnya.
    """
    if not records:
        print("  Autosave dilewati: belum ada detail berhasil.")
        return

    df = pd.DataFrame(records)

    # Ekstensi akhir harus .csv dan .xlsx, bukan .csv.tmp / .xlsx.tmp.
    csv_tmp = CSV_FILE.with_name(f"{CSV_FILE.stem}.tmp.csv")
    xlsx_tmp = XLSX_FILE.with_name(f"{XLSX_FILE.stem}.tmp.xlsx")

    try:
        df.to_csv(csv_tmp, index=False, encoding="utf-8-sig")
        fsync_existing_file(csv_tmp)
        os.replace(csv_tmp, CSV_FILE)

        df.to_excel(xlsx_tmp, index=False, engine="openpyxl")
        fsync_existing_file(xlsx_tmp)
        os.replace(xlsx_tmp, XLSX_FILE)

        print(f"  Autosave berhasil: {len(records)} sekolah -> CSV dan XLSX")
    finally:
        for temp_path in (csv_tmp, xlsx_tmp):
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except OSError as exc:
                print(f"  Peringatan: tidak bisa menghapus file sementara {temp_path.name}: {exc}")


def read_jsonl_records(path: Path) -> list[dict[str, Any]]:
    """
    Baca JSONL. Baris terakhir yang terpotong karena listrik/program mati
    akan dilewati, sementara baris valid sebelumnya tetap digunakan.
    """
    records: list[dict[str, Any]] = []
    if not path.exists():
        return records

    with path.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
                if isinstance(obj, dict):
                    records.append(obj)
            except json.JSONDecodeError:
                print(
                    f"PERINGATAN: baris JSONL tidak valid dilewati: "
                    f"{path.name}:{line_number}"
                )
    return records


def get_existing_successes() -> dict[str, dict[str, Any]]:
    """Checkpoint JSONL menjadi sumber utama untuk melanjutkan detail."""
    successes: dict[str, dict[str, Any]] = {}

    for record in read_jsonl_records(CHECKPOINT_FILE):
        school_id = record.get("sekolah_id")
        if school_id and record.get("status_scraping") == "berhasil":
            # Jika ada ID duplikat, record valid terakhir dipakai.
            successes[str(school_id)] = record

    return successes


def load_saved_listing() -> list[dict[str, Any]]:
    if not IDS_FILE.exists():
        return []

    try:
        loaded = json.loads(IDS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        print(f"Peringatan: {IDS_FILE.name} tidak bisa dibaca: {exc}")
        return []

    if not isinstance(loaded, list):
        print("Peringatan: format sekolah_ids.json bukan list; daftar dibangun ulang.")
        return []

    by_id: dict[str, dict[str, Any]] = {}
    for item in loaded:
        if isinstance(item, dict) and item.get("sekolah_id"):
            school_id = str(item["sekolah_id"])
            by_id[school_id] = {**item, "sekolah_id": school_id}

    return list(by_id.values())


# ---------------------------------------------------------------------------
# AMBIL DAFTAR ID SEKOLAH
# ---------------------------------------------------------------------------

def extract_school_ids(session: requests.Session) -> list[dict[str, Any]]:
    # Jika listing sebelumnya selesai, tidak perlu meminta halaman 0-54 lagi.
    saved_listing = load_saved_listing()
    if IDS_FILE.exists() and LISTING_COMPLETE_FILE.exists() and saved_listing:
        print(
            f"Listing sebelumnya sudah selesai. Memakai {len(saved_listing)} "
            "ID dari checkpoint; tidak mengulang listing."
        )
        return saved_listing

    # Kompatibilitas dengan checkpoint versi lama yang belum punya marker.
    # Jika jumlah ID sudah mencapai total yang diketahui dari API, jangan ulangi listing.
    if saved_listing and len(saved_listing) >= EXPECTED_TOTAL_IDS:
        atomic_write_json(LISTING_COMPLETE_FILE, {
            "selesai": True,
            "start_page": START_PAGE,
            "end_page": END_PAGE,
            "page_size": PAGE_SIZE,
            "jumlah_id_unik": len(saved_listing),
            "dipulihkan_dari_checkpoint_lama": True,
            "waktu_unix": time.time(),
        })
        print(
            f"Checkpoint lama berisi {len(saved_listing)} ID (>= {EXPECTED_TOTAL_IDS}); "
            "listing dilewati agar tidak mengulang permintaan."
        )
        return saved_listing

    by_id: dict[str, dict[str, Any]] = {
        str(item["sekolah_id"]): item for item in saved_listing
        if item.get("sekolah_id")
    }

    for page_number in range(START_PAGE, END_PAGE + 1):
        payload = dict(PAYLOAD_TEMPLATE)
        payload["page"] = page_number

        print(f"\n[LISTING] Meminta page={page_number}, size={PAGE_SIZE}")

        try:
            response = session.post(
                SEARCH_URL,
                json=payload,
                timeout=TIMEOUT_SECONDS,
            )
            response.raise_for_status()
            result = response.json()
        except (requests.RequestException, ValueError) as exc:
            # ID dari halaman yang telah tersimpan tetap aman.
            raise RuntimeError(
                f"Gagal mengambil listing page={page_number}. "
                f"ID dari halaman sebelumnya sudah disimpan. Detail: {exc}"
            ) from exc

        api_status = result.get("status_code")
        if api_status not in (None, 200):
            raise RuntimeError(
                f"API listing page={page_number} mengembalikan "
                f"status_code={api_status}: {result.get('message')}"
            )

        data = result.get("data", [])
        if isinstance(data, dict):
            data = data.get("data", data.get("sekolah", []))

        if not isinstance(data, list):
            raise RuntimeError(
                f"Format data listing page={page_number} tidak valid "
                f"({type(data).__name__})."
            )

        page_id_count = 0
        for school in data:
            if not isinstance(school, dict):
                continue
            school_id = school.get("sekolah_id")
            if not school_id:
                continue

            school_id = str(school_id)
            by_id[school_id] = {**school, "sekolah_id": school_id}
            page_id_count += 1

        # Simpan setelah setiap halaman. Kalau program mati, ID sebelumnya aman.
        atomic_write_json(IDS_FILE, list(by_id.values()))

        print(
            f"  ID ditemukan di halaman ini: {page_id_count}; "
            f"total ID unik tersimpan: {len(by_id)}"
        )

        time.sleep(REQUEST_DELAY_SECONDS + random.uniform(0, 0.3))

    # Penanda ini hanya ditulis setelah semua halaman selesai tanpa error.
    atomic_write_json(LISTING_COMPLETE_FILE, {
        "selesai": True,
        "start_page": START_PAGE,
        "end_page": END_PAGE,
        "page_size": PAGE_SIZE,
        "jumlah_id_unik": len(by_id),
        "waktu_unix": time.time(),
    })

    return list(by_id.values())


# ---------------------------------------------------------------------------
# DETAIL SEKOLAH DAN FLATTEN DATA
# ---------------------------------------------------------------------------

def flatten_detail(
    school_id: str,
    listing_record: dict[str, Any],
    response_json: dict[str, Any],
) -> dict[str, Any]:
    api_data = response_json.get("data", {})
    if not isinstance(api_data, dict):
        api_data = {}

    sekolah_list = api_data.get("sekolah", [])
    main_school: dict[str, Any] = {}
    if (
        isinstance(sekolah_list, list)
        and sekolah_list
        and isinstance(sekolah_list[0], dict)
    ):
        main_school = sekolah_list[0]

    row: dict[str, Any] = {
        "sekolah_id": school_id,
        "nama_sekolah": main_school.get("nama") or listing_record.get("nama", ""),
        "npsn": main_school.get("npsn") or listing_record.get("npsn", ""),
        "url_profil": f"{BASE_URL}/profil-sekolah/{school_id}",
        "status_code_api": response_json.get("status_code"),
        "message_api": response_json.get("message"),
    }

    # Kolom data sekolah utama.
    for key, value in main_school.items():
        row[f"sekolah_{key}"] = value

    # Bagian API lain tetap disimpan sebagai JSON string jika bertingkat.
    for key, value in api_data.items():
        if key == "sekolah":
            continue
        if isinstance(value, (dict, list)):
            row[key] = json.dumps(value, ensure_ascii=False, default=str)
        else:
            row[key] = value

    row["status_scraping"] = "berhasil"
    row["error"] = ""
    return row


def fetch_detail(
    session: requests.Session,
    school_id: str,
    listing_record: dict[str, Any],
) -> dict[str, Any]:
    url = f"{DETAIL_URL}/{school_id}"
    response = session.get(url, timeout=TIMEOUT_SECONDS)
    response.raise_for_status()
    result = response.json()

    api_status = result.get("status_code")
    if api_status not in (None, 200):
        raise RuntimeError(
            f"API detail status_code={api_status}: {result.get('message')}"
        )

    if not isinstance(result.get("data"), dict):
        raise RuntimeError("Respons detail tidak memiliki objek 'data' yang valid.")

    record = flatten_detail(school_id, listing_record, result)

    # Arsip lengkap API disimpan juga agar struktur aslinya tidak hilang.
    record["detail_json_lengkap"] = json.dumps(
        result,
        ensure_ascii=False,
        separators=(",", ":"),
        default=str,
    )
    return record


# ---------------------------------------------------------------------------
# PROGRAM UTAMA
# ---------------------------------------------------------------------------

def main() -> None:
    session = build_session()

    print("=== 1. Memuat/mengambil daftar ID sekolah ===")
    school_records = extract_school_ids(session)
    print(f"Jumlah ID sekolah unik: {len(school_records)}")

    print("\n=== 2. Membaca checkpoint detail ===")
    successes = get_existing_successes()
    print(f"Detail berhasil yang sudah ada: {len(successes)}")

    pending = [
        item for item in school_records
        if str(item.get("sekolah_id")) not in successes
    ]
    print(f"Detail yang masih perlu diambil: {len(pending)}")

    if not pending:
        print("Semua ID listing sudah memiliki detail berhasil.")
        save_table_files(list(successes.values()))
        print(f"Folder output: {OUTPUT_DIR}")
        return

    processed_since_save = 0

    for index, listing_record in enumerate(pending, start=1):
        school_id = str(listing_record["sekolah_id"])
        print(f"\n[DETAIL {index}/{len(pending)}] {school_id}")

        try:
            record = fetch_detail(session, school_id, listing_record)

            # Satu sekolah berhasil = langsung masuk checkpoint permanen.
            append_jsonl_durable(CHECKPOINT_FILE, record)
            successes[school_id] = record
            processed_since_save += 1

            print(f"  BERHASIL: {record.get('nama_sekolah', '')}")

        except Exception as exc:
            error_record = {
                "sekolah_id": school_id,
                "url_detail_api": f"{DETAIL_URL}/{school_id}",
                "status_scraping": "gagal",
                "error": f"{type(exc).__name__}: {exc}",
                "waktu_unix": time.time(),
            }

            try:
                append_jsonl_durable(ERRORS_FILE, error_record)
            except OSError as log_exc:
                print(f"  PERINGATAN: gagal menulis log error: {log_exc}")

            print(f"  GAGAL: {error_record['error']}")

        if processed_since_save >= AUTOSAVE_EVERY:
            save_table_files(list(successes.values()))
            processed_since_save = 0

        time.sleep(REQUEST_DELAY_SECONDS + random.uniform(0, 0.4))

    # Simpan data terakhir walaupun jumlahnya kurang dari interval autosave.
    save_table_files(list(successes.values()))

    total_success = len(successes)
    total_missing = max(0, len(school_records) - total_success)

    print("\n=== SELESAI ===")
    print(f"ID listing tersimpan : {len(school_records)}")
    print(f"Detail berhasil      : {total_success}")
    print(f"Detail belum berhasil: {total_missing}")
    print(f"Folder output        : {OUTPUT_DIR}")
    print(f"Checkpoint           : {CHECKPOINT_FILE}")
    print(f"CSV                  : {CSV_FILE}")
    print(f"Excel                : {XLSX_FILE}")
    print(f"Log kegagalan        : {ERRORS_FILE}")
    print("\nJika program terputus, jalankan ulang perintah yang sama.")
    print("Detail yang sudah berhasil di checkpoint akan dilewati.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(
            "\nProgram dihentikan oleh pengguna. Data yang sudah masuk "
            "checkpoint tetap tersimpan. Jalankan ulang untuk melanjutkan."
        )
    except Exception as exc:
        print(f"\nPROGRAM BERHENTI: {type(exc).__name__}: {exc}")
        print(
            "Data yang sudah masuk checkpoint tetap tersedia. "
            "Periksa koneksi/API lalu jalankan ulang untuk melanjutkan."
        )
