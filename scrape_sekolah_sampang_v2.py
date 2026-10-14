"""
Persiapan:
    pip install playwright pandas openpyxl
    playwright install chromium

Jalankan:
    python scrape_sekolah_sampang.py

Catatan:
- Ubah variabel page untuk menentukan page pertama, ubah MAX_PAGES menjadi None jika ingin mengambil semua page.
- Gunakan jeda yang wajar dan sesuaikan dengan ketentuan penggunaan situs.
"""

from __future__ import annotations

import json
import re
import time
from urllib.parse import parse_qs, urlencode, urlsplit, urlunsplit

import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


START_URL = (
    "https://sekolah.data.kemendikdasmen.go.id/sekolah?keyword="
    "&kabupaten_kota=Kab.%20Sampang"
    "&bentuk_pendidikan=SD,%20SMA,%20SMK,%20SMP,%20MA,%20MI,%20MTs,%20Pondok%20Pesantren,%20RA,%20Sekolah%20Rakyat,%20TK"
    "&status_sekolah=&page=2&size=48"
)

START_PAGE = 0
MAX_PAGES = 1  
HEADLESS = False  # True jika ingin browser berjalan tanpa tampilan
DELAY_SECONDS = 1.0  # jeda antarprofil agar tidak membebani situs

OUTPUT_CSV = "data_sekolah_sampang.csv"
OUTPUT_EXCEL = "data_sekolah_sampang.xlsx"
OUTPUT_JSON = "data_sekolah_sampang.json"

PROFILE_LINK_SELECTOR = 'a[href*="/profil-sekolah/"]'

# Label yang terlihat di halaman detail sekolah. Nilai umumnya berada di baris setelah label.
FIELD_LABELS = {
    "akreditasi": ["Akreditasi"],
    "status_sekolah": ["Status Sekolah"],
    "npsn": ["NPSN"],
    "bentuk_pendidikan": ["Bentuk Pendidikan"],
    "telepon": ["Telepon", "Nomor Telepon"],
    "email": ["Email", "E-mail"],
    "website_sekolah": ["Website Sekolah", "Website"],
    "yayasan": ["Yayasan"],
    "guru": ["Guru"],
    "siswa_laki_laki": ["Siswa Laki-Laki", "Siswa Laki-Laki ", "Siswa Laki Laki"],
    "siswa_perempuan": ["Siswa Perempuan"],
    "rombongan_belajar": ["Rombongan Belajar"],
    "daya_tampung": ["Daya Tampung", "Daya Tampung Sekolah"],
    "ruang_kelas": ["Ruang Kelas"],
    "laboratorium": ["Laboratorium"],
    "perpustakaan": ["Perpustakaan"],
    "kurikulum": ["Kurikulum"],
    "akses_internet": ["Akses Internet"],
    "penyelenggaraan": ["Penyelenggaraan"],
    "sumber_listrik": ["Sumber Listrik"],
    "semester_data": ["Semester Data"],
    "daya_listrik": ["Daya Listrik"],
    "luas_tanah": ["Luas Tanah"],
    "rasio_siswa_rombel": ["Rasio Siswa Rombel"],
    "rasio_rombel_ruang_kelas": ["Rasio Rombel Ruang Kelas"],
    "rasio_siswa_guru": ["Rasio Siswa Guru"],
    "persentase_guru_kualifikasi": [
        "Persentase Guru Kualifikasi",
        "Persentase Guru Kualifikasi Akademik",
    ],
}


def normalize_text(value: str) -> str:
    """Normalisasi teks untuk membandingkan label dengan lebih tahan variasi."""
    value = value.replace("\u00a0", " ").strip().lower()
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def make_page_url(base_url: str, page_number: int) -> str:
    """Ganti parameter page pada URL, dengan mempertahankan filter lainnya."""
    parts = urlsplit(base_url)
    query = parse_qs(parts.query, keep_blank_values=True)
    query["page"] = [str(page_number)]
    return urlunsplit(
        (parts.scheme, parts.netloc, parts.path, urlencode(query, doseq=True), parts.fragment)
    )


def get_profile_links(page) -> list[str]:
    """Ambil URL profil jika situs memang menyediakan tautan <a href> langsung."""
    links = page.locator('a[href*="/profil-sekolah/"], a[href*="/index.php/chome/profil/"]').evaluate_all(
        "elements => elements.map(el => el.href).filter(Boolean)"
    )
    return list(dict.fromkeys(links))


def get_lihat_mode_and_count(page) -> tuple[str, int]:
    """Cari kontrol 'Lihat'. Utamakan per kartu <article>, lalu cari teks global."""
    articles = page.locator("article")
    article_count = articles.count()
    if article_count:
        # Di beberapa versi situs setiap kartu sekolah dibungkus elemen <article>.
        usable = 0
        for i in range(article_count):
            card = articles.nth(i)
            if card.get_by_text("Lihat", exact=True).count() > 0 or card.get_by_role("button", name=re.compile(r"Lihat", re.I)).count() > 0:
                usable += 1
        if usable:
            return "article", article_count

    # Fallback bila kartu tidak menggunakan elemen <article>.
    text_count = page.get_by_text("Lihat", exact=True).count()
    if text_count:
        return "text", text_count
    return "none", 0


def click_lihat(page, mode: str, index: int) -> str:
    """Klik tombol Lihat ke-index lalu kembalikan URL profil yang dibuka."""
    if mode == "article":
        card = page.locator("article").nth(index)
        target = card.get_by_text("Lihat", exact=True)
        if target.count() == 0:
            target = card.get_by_role("button", name=re.compile(r"Lihat", re.I))
        if target.count() == 0:
            raise RuntimeError(f"Tombol Lihat tidak ditemukan pada kartu ke-{index + 1}")
        target.first.click(timeout=15_000)
    elif mode == "text":
        target = page.get_by_text("Lihat", exact=True).nth(index)
        target.click(timeout=15_000)
    else:
        raise RuntimeError("Tidak ada tombol Lihat yang bisa diklik")

    try:
        page.wait_for_function(
            "() => location.pathname.includes('/profil-sekolah/') || location.pathname.includes('/index.php/chome/profil/')",
            timeout=15_000,
        )
    except PlaywrightTimeoutError:
        # Cetak alamat saat ini agar mudah didiagnosis jika tombol tidak menavigasi.
        raise RuntimeError(f"Klik Lihat tidak membuka URL profil. URL saat ini: {page.url}")

    return page.url


def scroll_to_load_more_content(page, max_rounds: int = 10) -> None:
    """Gulir ke bawah untuk memicu pemuatan konten yang bersifat lazy-load."""
    previous_height = -1
    stable_rounds = 0

    for _ in range(max_rounds):
        height = page.evaluate("document.body.scrollHeight")
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(350)
        new_height = page.evaluate("document.body.scrollHeight")

        if new_height == previous_height or new_height == height:
            stable_rounds += 1
        else:
            stable_rounds = 0

        previous_height = new_height
        if stable_rounds >= 2:
            break

    page.evaluate("window.scrollTo(0, 0)")


def extract_label_values(lines: list[str]) -> dict[str, str]:
    """Ambil nilai untuk label umum dari teks halaman; data lengkap tetap disimpan terpisah."""
    normalized_labels: dict[str, tuple[str, list[str]]] = {}
    all_label_names: set[str] = set()

    for column, aliases in FIELD_LABELS.items():
        for alias in aliases:
            normalized_alias = normalize_text(alias)
            normalized_labels[normalized_alias] = (column, aliases)
            all_label_names.add(normalized_alias)

    result: dict[str, str] = {column: "" for column in FIELD_LABELS}

    for i, original_line in enumerate(lines):
        normalized_line = normalize_text(original_line)
        matched_column = None
        inline_value = ""

        # Kasus umum: label pada baris tersendiri.
        if normalized_line in normalized_labels:
            matched_column = normalized_labels[normalized_line][0]
        else:
            # Kasus alternatif: "NPSN: 12345678" atau label disertai tanda baca.
            for normalized_alias, (column, aliases) in sorted(
                normalized_labels.items(), key=lambda item: len(item[0]), reverse=True
            ):
                if normalized_line.startswith(normalized_alias + " "):
                    matched_column = column
                    # Ambil bagian setelah alias dari teks asli, jika memungkinkan.
                    alias = next(
                        (a for a in aliases if normalize_text(original_line).startswith(normalize_text(a))),
                        aliases[0],
                    )
                    inline_value = original_line[len(alias):].strip(" :-–—\t")
                    break

        if not matched_column or result.get(matched_column):
            continue

        if inline_value:
            result[matched_column] = inline_value
            continue

        # Nilai biasanya berada di baris berikutnya. Lewati baris kosong (sudah dibuang)
        # atau baris yang juga merupakan label lain.
        j = i + 1
        while j < len(lines) and normalize_text(lines[j]) in all_label_names:
            j += 1

        if j < len(lines):
            candidate = lines[j].strip()
            # Hindari mengisi nilai dengan label navigasi/heading kosong.
            if normalize_text(candidate) not in {"sekolah", "statistik sekolah"}:
                result[matched_column] = candidate

    return result


def scrape_detail(page, url: str, already_open: bool = False) -> dict:
    """Buka satu profil dan ambil field terstruktur serta seluruh teks halaman."""
    if not already_open or page.url != url:
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)

    try:
        page.wait_for_load_state("networkidle", timeout=8_000)
    except PlaywrightTimeoutError:
        # Situs dengan request latar belakang yang terus aktif mungkin tidak pernah idle.
        pass

    page.wait_for_timeout(1_000)
    try:
        page.locator("body").wait_for(state="visible", timeout=10_000)
    except PlaywrightTimeoutError:
        pass

    scroll_to_load_more_content(page)

    body_text = page.locator("body").inner_text(timeout=15_000)
    lines = [line.strip() for line in body_text.splitlines() if line.strip()]

    try:
        name = page.locator("h1").first.inner_text(timeout=3_000).strip()
    except PlaywrightTimeoutError:
        name = ""

    # Di halaman profil, alamat singkat biasanya tampil tepat di bawah nama sekolah.
    address_short = ""
    if name:
        target_name = normalize_text(name)
        for idx, line in enumerate(lines):
            if normalize_text(line) == target_name and idx + 1 < len(lines):
                address_short = lines[idx + 1]
                break

    fields = extract_label_values(lines)
    return {
        "nama_sekolah": name,
        "alamat_ringkas": address_short,
        **fields,
        "url_profil": url,
        "teks_lengkap_halaman": body_text,
        "status_scraping": "berhasil",
        "error": "",
    }


def main() -> None:
    records: list[dict] = []
    visited_profile_urls: set[str] = set()
    listing_page = START_PAGE
    listing_pages_done = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page(viewport={"width": 1440, "height": 1000})
        page.set_default_timeout(15_000)

        while MAX_PAGES is None or listing_pages_done < MAX_PAGES:
            listing_url = make_page_url(START_URL, listing_page)
            print(f"\nMembuka halaman daftar page={listing_page}: {listing_url}")
            page.goto(listing_url, wait_until="domcontentloaded", timeout=60_000)

            try:
                page.wait_for_load_state("networkidle", timeout=8_000)
            except PlaywrightTimeoutError:
                pass
            page.wait_for_timeout(2_500)
            scroll_to_load_more_content(page)

            profile_urls = get_profile_links(page)
            new_urls = [url for url in profile_urls if url not in visited_profile_urls]
            print(f"Tautan profil langsung ditemukan: {len(profile_urls)}; belum diproses: {len(new_urls)}")

            # Jalur 1: jika URL profil tersedia sebagai tautan biasa.
            if profile_urls:
                for url in new_urls:
                    visited_profile_urls.add(url)
                    number = len(records) + 1
                    print(f"[{number}] Mengambil profil: {url}")
                    try:
                        record = scrape_detail(page, url)
                    except Exception as exc:
                        record = {
                            "nama_sekolah": "",
                            "alamat_ringkas": "",
                            **{column: "" for column in FIELD_LABELS},
                            "url_profil": url,
                            "teks_lengkap_halaman": "",
                            "status_scraping": "gagal",
                            "error": f"{type(exc).__name__}: {exc}",
                        }
                        print(f"   Gagal mengambil profil: {record['error']}")
                    records.append(record)
                    time.sleep(DELAY_SECONDS)

            # Jalur 2: tombol Lihat berupa aksi JavaScript tanpa href langsung.
            else:
                mode, count = get_lihat_mode_and_count(page)
                print(f"Tautan <a href> tidak ditemukan. Mode tombol: {mode}; jumlah kandidat: {count}")

                if count == 0:
                    # File debug membantu melihat apa yang sebenarnya diterima browser.
                    try:
                        page.screenshot(path="debug_sekolah_sampang.png", full_page=True)
                        with open("debug_sekolah_sampang.html", "w", encoding="utf-8") as f:
                            f.write(page.content())
                        print("Tidak menemukan tombol Lihat. Debug disimpan ke debug_sekolah_sampang.png dan debug_sekolah_sampang.html")
                        print("Judul halaman:", page.title())
                        print("Teks awal halaman:", page.locator("body").inner_text()[:500].replace("\n", " | "))
                    except Exception as debug_exc:
                        print(f"Gagal menyimpan file debug: {debug_exc}")
                    print("Proses dihentikan karena kartu sekolah/tombol Lihat tidak ditemukan.")
                    break

                for i in range(count):
                    # Buka ulang daftar sebelum tiap klik agar indeks kartu tidak bergeser.
                    if i > 0:
                        page.goto(listing_url, wait_until="domcontentloaded", timeout=60_000)
                        try:
                            page.wait_for_load_state("networkidle", timeout=8_000)
                        except PlaywrightTimeoutError:
                            pass
                        page.wait_for_timeout(1_800)
                        scroll_to_load_more_content(page)

                    try:
                        url = click_lihat(page, mode, i)
                        if url in visited_profile_urls:
                            print(f"[{i + 1}/{count}] Profil sudah pernah diproses, dilewati: {url}")
                            continue
                        visited_profile_urls.add(url)
                        number = len(records) + 1
                        print(f"[{number}] Mengambil profil dari tombol Lihat: {url}")
                        record = scrape_detail(page, url, already_open=True)
                    except Exception as exc:
                        print(f"[{i + 1}/{count}] Gagal membuka/mengambil kartu sekolah: {type(exc).__name__}: {exc}")
                        continue

                    records.append(record)
                    time.sleep(DELAY_SECONDS)

            listing_pages_done += 1
            listing_page += 1

        browser.close()

    if not records:
        print("Belum ada data yang berhasil direkam.")
        print("Periksa file debug_sekolah_sampang.png/html untuk memastikan halaman daftar benar-benar termuat.")
        return

    # Simpan tiga format. JSON berguna untuk mempertahankan teks lengkap tanpa kehilangan struktur.
    pd.DataFrame(records).to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    pd.DataFrame(records).to_excel(OUTPUT_EXCEL, index=False)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as file:
        json.dump(records, file, ensure_ascii=False, indent=2)

    success_count = sum(row.get("status_scraping") == "berhasil" for row in records)
    print("\nSelesai.")
    print(f"Profil tercatat: {len(records)}; berhasil: {success_count}; gagal: {len(records) - success_count}")
    print(f"CSV  : {OUTPUT_CSV}")
    print(f"Excel: {OUTPUT_EXCEL}")
    print(f"JSON : {OUTPUT_JSON}")


if __name__ == "__main__":
    main()
