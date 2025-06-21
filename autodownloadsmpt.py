import os
import time
import requests
from urllib.parse import urljoin, urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select


# === KONFIGURASI ===
EMAIL = "199108272020122026@kemenag.go.id"  # Ganti dengan email login kamu
PASSWORD = "elok01"  # Ganti dengan password login kamu
FOLDER_PATH_DOWNLOAD = r"C:\Users\Rival\Downloads\SPMT_CPNS_Signed"  # Folder tujuan download
TTE_LOGIN = "https://tte.kemenag.go.id"

# === SETUP DRIVER ===
options = webdriver.EdgeOptions()
options.use_chromium = True

# Konfigurasi download untuk bypass save as dialog
prefs = {
    "download.default_directory": FOLDER_PATH_DOWNLOAD,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True,
    "profile.default_content_setting_values.notifications": 2,
    "profile.default_content_settings.popups": 0,
    "profile.managed_default_content_settings.images": 2,
    "profile.content_settings.plugin_whitelist.adobe-flash-player": 1,
    "profile.content_settings.exceptions.plugins.*,*.per_resource.adobe-flash-player": 1,
    "PluginsAllowedForUrls": "https://tte.kemenag.go.id",
    "profile.default_content_setting_values.automatic_downloads": 1
}
options.add_experimental_option("prefs", prefs)

# Tambahan opsi untuk bypass dialog
options.add_argument("--disable-extensions")
options.add_argument("--disable-plugins-discovery")
options.add_argument("--disable-web-security")
options.add_argument("--allow-running-insecure-content")

# Pastikan folder download ada
import os
if not os.path.exists(FOLDER_PATH_DOWNLOAD):
    os.makedirs(FOLDER_PATH_DOWNLOAD)
    print(f"📁 Folder download dibuat: {FOLDER_PATH_DOWNLOAD}")

driver = webdriver.Edge(options=options)
driver.get(TTE_LOGIN)

# Fungsi untuk download file langsung menggunakan requests
def download_file_direct(url, folder_path, cookies=None, headers=None):
    """
    Download file langsung ke folder tanpa dialog save as
    """
    try:
        # Setup session dengan cookies dari selenium
        session = requests.Session()

        if cookies:
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'])

        if headers:
            session.headers.update(headers)

        # Tambahkan headers yang umum digunakan browser
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36 Edg/91.0.864.59',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        })

        print(f"🔄 Memulai download dari: {url}")

        # Download file
        response = session.get(url, stream=True, timeout=30)
        response.raise_for_status()

        # Tentukan nama file
        filename = None
        if 'content-disposition' in response.headers:
            import re
            cd = response.headers['content-disposition']
            filename_match = re.findall('filename=(.+)', cd)
            if filename_match:
                filename = filename_match[0].strip('"')

        if not filename:
            # Ambil nama file dari URL
            parsed_url = urlparse(url)
            filename = os.path.basename(parsed_url.path)
            if not filename or '.' not in filename:
                filename = f"document_{int(time.time())}.pdf"

        # Pastikan ekstensi file
        if not filename.lower().endswith(('.pdf', '.doc', '.docx', '.xls', '.xlsx')):
            filename += '.pdf'

        filepath = os.path.join(folder_path, filename)

        # Tulis file
        with open(filepath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        file_size = os.path.getsize(filepath)
        print(f"✅ File berhasil didownload: {filename} ({file_size} bytes)")
        return True, filename

    except Exception as e:
        print(f"❌ Error download langsung: {str(e)}")
        return False, None

# === LOGIN ===
wait = WebDriverWait(driver, 20)
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div/div[1]/div/div/div[2]/form/div[1]/div/div[3]/input'))).click()
wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/section/div/div[1]/div/div/div[2]/form/div[4]/div/input'))).send_keys(EMAIL)
wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/section/div/div[1]/div/div/div[2]/form/div[5]/div/input'))).send_keys(PASSWORD)
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div/div[1]/div/div/div[2]/form/div[6]/div[2]/button'))).click()

# Tunggu login selesai
wait.until(EC.presence_of_element_located((By.XPATH, '//nav')))  # Tunggu hingga elemen navigasi muncul
print("✅ Login berhasil")

# === Akses Halaman Dokumen ===
driver.get("https://tte.kemenag.go.id/satker/dokumen/naskah/index/unggah")
wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="example2_length"]')))  # Tunggu elemen filter muncul

# === Pilih filter "100" data menggunakan XPath yang benar ===
filter_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div/section/div[1]/div/section/div/div[2]/div/div[1]/div[1]/div/label/select')))
select = Select(filter_dropdown)
select.select_by_visible_text("100")  # Pilih "100"

# Tunggu refresh tabel selesai (beberapa detik agar tabel terupdate)
print("🔄 Menunggu pembaruan tabel setelah filter diterapkan...")
wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/section/div/section/div[1]/div/section/div/div[2]/div/div[2]/div/table/tbody')))  # Tunggu tabel terupdate

print("✅ Tabel tbody ditemukan, menunggu 5 detik untuk memastikan data dimuat...")
time.sleep(5)  # Tunggu 5 detik seperti yang diminta

# === Looping untuk mencari dokumen "SPMT CPNS T.A 2024" dengan status "FINAL" ===
page_number = 1
while True:
    print(f"🔍 Memeriksa halaman {page_number}...")

    # Debug: Cek apakah tabel ada dan struktur header
    try:
        table_exists = driver.find_element(By.XPATH, "/html/body/section/div/section/div[1]/div/section/div/div[2]/div/div[2]/div/table")
        print("✅ Tabel ditemukan")

        # Debug: Cek header tabel untuk memastikan urutan kolom
        headers = driver.find_elements(By.XPATH, "/html/body/section/div/section/div[1]/div/section/div/div[2]/div/div[2]/div/table/thead/tr/th")
        print(f"📋 Jumlah kolom header: {len(headers)}")
        perihal_col = None
        unduh_col = None
        for i, header in enumerate(headers):
            # Ambil teks dari elemen dan sub-elemen untuk menangani struktur HTML yang kompleks
            header_text = header.text.strip()
            header_html = header.get_attribute('innerHTML')
            print(f"   Kolom {i+1}: '{header_text}' (HTML: {header_html[:50]}...)")

            # Cek berbagai variasi teks untuk kolom Perihal
            if any(keyword in header_text.upper() for keyword in ["PERIHAL", "DOKUMEN"]):
                perihal_col = i + 1
                print(f"   ✅ Kolom Perihal Dokumen ditemukan di posisi {i+1}")

            # Cek berbagai variasi teks untuk kolom Unduh
            if any(keyword in header_text.upper() for keyword in ["UNDUH", "DOWNLOAD"]):
                unduh_col = i + 1
                print(f"   ✅ Kolom Unduh ditemukan di posisi {i+1}")

    except Exception as e:
        print(f"❌ Tabel tidak ditemukan atau error: {str(e)}")
        break

    # Debug: Cek jumlah baris dalam tabel dan isi beberapa baris pertama
    try:
        rows = driver.find_elements(By.XPATH, "/html/body/section/div/section/div[1]/div/section/div/div[2]/div/div[2]/div/table/tbody/tr")
        print(f"📊 Jumlah baris dalam tabel: {len(rows)}")

        # Debug: Tampilkan isi 3 baris pertama untuk verifikasi
        for i, row in enumerate(rows[:3]):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 3:
                    # Berdasarkan deteksi header, kolom perihal adalah kolom ke-3
                    perihal_text = cells[2].text if perihal_col == 3 else cells[1].text  # Kolom ke-3 (index 2) atau ke-2 (index 1)
                    print(f"🔍 Baris {i+1} - Perihal: '{perihal_text}'")

                    # Tampilkan semua kolom untuk debug
                    print(f"   📋 Semua kolom: {[cell.text[:20] + '...' if len(cell.text) > 20 else cell.text for cell in cells]}")

                    if len(cells) >= 8:
                        unduh_cell = cells[7]  # Kolom ke-8 (index 7)
                        final_links = unduh_cell.find_elements(By.TAG_NAME, "a")
                        print(f"   🔗 Jumlah link di kolom unduh: {len(final_links)}")
                        for link_idx, link in enumerate(final_links):
                            link_text = link.text.strip()
                            link_href = link.get_attribute("href")
                            print(f"   📎 Link {link_idx + 1}: '{link_text}' -> {link_href[:50]}...")
            except Exception as e:
                print(f"❌ Error membaca baris {i+1}: {str(e)}")
    except:
        print("❌ Tidak dapat menghitung baris tabel")

    # Cari elemen yang mengandung "SPMT CPNS T.A 2024" di kolom perihal dan status "FINAL" di kolom unduhan
    # Menggunakan posisi kolom yang dinamis berdasarkan header yang ditemukan
    dokumen_elements = []

    # Jika deteksi kolom gagal, gunakan asumsi berdasarkan screenshot (kolom 3 untuk Perihal, kolom 8 untuk Unduh)
    if not perihal_col or not unduh_col:
        print("⚠️ Deteksi kolom gagal, menggunakan asumsi berdasarkan screenshot:")
        perihal_col = 3  # Berdasarkan screenshot
        unduh_col = 8    # Berdasarkan screenshot
        print(f"   Menggunakan kolom Perihal: {perihal_col}, kolom Unduh: {unduh_col}")

    if perihal_col and unduh_col:
        print(f"🔍 Mencari dokumen menggunakan kolom Perihal: {perihal_col}, kolom Unduh: {unduh_col}")
        dokumen_elements = driver.find_elements(By.XPATH, f"/html/body/section/div/section/div[1]/div/section/div/div[2]/div/div[2]/div/table/tbody/tr[td[{perihal_col}][contains(text(),'SPMT CPNS T.A 2024 Update Unit Kerja')]]/td[{unduh_col}]//a[contains(text(),'FINAL')]")

        # Jika tidak ditemukan dengan XPath spesifik, coba dengan XPath yang lebih fleksibel
        if len(dokumen_elements) == 0:
            print("🔄 Mencoba dengan XPath alternatif 1...")
            dokumen_elements = driver.find_elements(By.XPATH, f"//table//tr[td[{perihal_col}][contains(text(),'SPMT CPNS T.A 2024 Update Unit Kerja')]]//a[contains(text(),'FINAL')]")

        # Jika masih tidak ditemukan, coba XPath yang lebih umum
        if len(dokumen_elements) == 0:
            print("🔄 Mencoba dengan XPath alternatif 2...")
            dokumen_elements = driver.find_elements(By.XPATH, "//tr[td[contains(text(),'SPMT CPNS T.A 2024 Update Unit Kerja')]]//a[contains(text(),'FINAL')]")

    # Jika masih tidak ditemukan atau kolom tidak teridentifikasi, coba mencari semua link FINAL dan filter manual
    if len(dokumen_elements) == 0:
        print("🔄 Mencoba dengan pendekatan manual...")
        all_rows = driver.find_elements(By.XPATH, "/html/body/section/div/section/div[1]/div/section/div/div[2]/div/div[2]/div/table/tbody/tr")
        dokumen_elements = []
        for row_index, row in enumerate(all_rows):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                # Cari di semua kolom untuk teks "SPMT CPNS T.A 2024"
                found_spmt = False
                for cell in cells:
                    if "SPMT CPNS T.A 2024 Update Unit Kerja" in cell.text:
                        found_spmt = True
                        print(f"✅ Ditemukan 'SPMT CPNS T.A 2024 Update Unit Kerja' di baris {row_index + 1}: {cell.text}")
                        break

                if found_spmt:
                    # Cari semua link di kolom unduh (kolom ke-8)
                    if len(cells) >= 8:
                        unduh_cell = cells[7]  # Index 7 untuk kolom ke-8
                        all_links = unduh_cell.find_elements(By.TAG_NAME, "a")

                        print(f"🔍 Ditemukan {len(all_links)} link di kolom unduh:")
                        for link_index, link in enumerate(all_links):
                            link_text = link.text.strip()
                            link_href = link.get_attribute("href")
                            print(f"   Link {link_index + 1}: '{link_text}' -> {link_href}")

                            # Cari link yang mengandung "FINAL" atau semua link jika tidak ada teks
                            if "FINAL" in link_text.upper() or link_text == "" or "FINAL" in str(link_href).upper():
                                dokumen_elements.append(link)
                                print(f"   ✅ Link ini akan didownload: '{link_text}' atau href mengandung FINAL")

                        # Jika tidak ada link dengan teks FINAL, ambil semua link di kolom unduh
                        if not any("FINAL" in link.text.upper() for link in all_links) and all_links:
                            print("   ⚠️ Tidak ada link dengan teks 'FINAL', mengambil semua link di kolom unduh")
                            dokumen_elements.extend(all_links)

            except Exception as e:
                print(f"❌ Error saat memproses baris {row_index + 1}: {str(e)}")
                continue

    print(f"📋 Ditemukan {len(dokumen_elements)} dokumen dengan kriteria yang dicari")

    if len(dokumen_elements) == 0:
        print("❌ Tidak ada dokumen 'SPMT CPNS T.A 2024 Update Unit Kerja' dengan status 'FINAL' yang ditemukan pada halaman ini.")

        # Cek apakah halaman berikutnya ada
        next_page_button = driver.find_elements(By.XPATH, '//a[@class="page-link" and contains(text(), "Next")]')
        if next_page_button and next_page_button[0].is_enabled():
            next_page_button[0].click()  # Klik tombol "Next" untuk halaman berikutnya
            print("➡️ Beralih ke halaman berikutnya...")
            page_number += 1
            time.sleep(3)  # Tunggu halaman dimuat
            continue
        else:
            print("✅ Semua halaman telah diperiksa. Tidak ada dokumen yang ditemukan.")
            break

    # Download dokumen yang ditemukan
    for i, dokumen in enumerate(dokumen_elements):
        try:
            doc_url = dokumen.get_attribute("href")
            print(f"✅ Dokumen ke-{i+1} ditemukan: {doc_url}")

            # Method 1: Download langsung menggunakan requests (bypass save as dialog)
            if doc_url and doc_url.startswith('http'):
                print(f"🔗 Link URL: {doc_url}")
                print("🔄 Mencoba download langsung dengan requests...")

                # Ambil cookies dari selenium untuk autentikasi
                cookies = driver.get_cookies()

                # Coba download langsung
                success, filename = download_file_direct(doc_url, FOLDER_PATH_DOWNLOAD, cookies)

                if success:
                    print(f"✅ Download langsung berhasil: {filename}")
                else:
                    print("❌ Download langsung gagal, mencoba metode browser...")

                    # Fallback: Method 2 - Klik dengan browser (akan otomatis download ke folder yang sudah dikonfigurasi)
                    print("🔄 Mencoba klik dengan browser...")
                    try:
                        # Scroll ke elemen agar terlihat
                        driver.execute_script("arguments[0].scrollIntoView(true);", dokumen)
                        time.sleep(1)

                        # Klik menggunakan JavaScript untuk menghindari masalah overlay
                        driver.execute_script("arguments[0].click();", dokumen)
                        print("✅ Klik berhasil dengan JavaScript")

                        # Tunggu download dimulai (browser akan otomatis download ke folder yang dikonfigurasi)
                        time.sleep(8)  # Tunggu lebih lama untuk memastikan download selesai

                    except Exception as click_error:
                        print(f"❌ Klik JavaScript gagal: {click_error}")

                        # Fallback: Method 3 - Navigasi langsung
                        print("🔄 Mencoba navigasi langsung...")
                        try:
                            current_url = driver.current_url
                            driver.get(doc_url)
                            time.sleep(8)  # Tunggu download
                            driver.get(current_url)  # Kembali ke halaman asal
                            time.sleep(3)  # Tunggu halaman dimuat
                            print("✅ Navigasi langsung berhasil")

                        except Exception as nav_error:
                            print(f"❌ Navigasi langsung gagal: {nav_error}")

            else:
                print(f"⚠️ Link tidak valid atau JavaScript: {doc_url}")
                # Untuk link JavaScript, coba klik langsung
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", dokumen)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", dokumen)
                    time.sleep(8)  # Tunggu download
                    print("✅ Klik JavaScript berhasil")
                except Exception as js_error:
                    print(f"❌ Klik JavaScript gagal: {js_error}")

            # Verifikasi apakah file berhasil didownload
            print("🔍 Memeriksa apakah file berhasil didownload...")

            # Cek file terbaru di folder download
            try:
                # Tunggu sebentar lagi untuk memastikan download selesai
                time.sleep(3)

                if os.path.exists(FOLDER_PATH_DOWNLOAD):
                    files = [f for f in os.listdir(FOLDER_PATH_DOWNLOAD) if os.path.isfile(os.path.join(FOLDER_PATH_DOWNLOAD, f))]
                    if files:
                        # Cari file terbaru
                        latest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(FOLDER_PATH_DOWNLOAD, x)))
                        file_time = os.path.getmtime(os.path.join(FOLDER_PATH_DOWNLOAD, latest_file))
                        current_time = time.time()

                        if current_time - file_time < 120:  # File dibuat dalam 120 detik terakhir
                            file_size = os.path.getsize(os.path.join(FOLDER_PATH_DOWNLOAD, latest_file))
                            print(f"✅ File terbaru: {latest_file} (Size: {file_size} bytes)")

                            # Cek apakah file masih dalam proses download
                            if not latest_file.endswith(('.crdownload', '.tmp', '.part')):
                                print("✅ Download berhasil!")
                            else:
                                print("🔄 File masih dalam proses download...")
                                # Tunggu sampai download selesai
                                max_wait = 30  # Maksimal 30 detik
                                while max_wait > 0 and latest_file.endswith(('.crdownload', '.tmp', '.part')):
                                    time.sleep(1)
                                    max_wait -= 1
                                    files = [f for f in os.listdir(FOLDER_PATH_DOWNLOAD) if os.path.isfile(os.path.join(FOLDER_PATH_DOWNLOAD, f))]
                                    if files:
                                        latest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(FOLDER_PATH_DOWNLOAD, x)))

                                if not latest_file.endswith(('.crdownload', '.tmp', '.part')):
                                    print("✅ Download selesai!")
                                else:
                                    print("⚠️ Download mungkin masih berlangsung...")
                        else:
                            print("⚠️ Tidak ada file baru yang terdeteksi dalam 120 detik terakhir")
                    else:
                        print("⚠️ Folder download kosong")
                else:
                    print("❌ Folder download tidak ditemukan")

            except Exception as verify_error:
                print(f"❌ Error saat verifikasi download: {verify_error}")

            print(f"📥 Proses download dokumen ke-{i+1} selesai")

        except Exception as e:
            print(f"❌ Error saat mengunduh dokumen ke-{i+1}: {str(e)}")
            # Pastikan kembali ke tab utama jika terjadi error
            try:
                if len(driver.window_handles) > 1:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
            except:
                pass

    # Menunggu beberapa detik sebelum mencari dokumen lagi (untuk menghindari terlalu cepat mengulang)
    time.sleep(2)

    # Cek apakah halaman berikutnya ada setelah download
    next_page_button = driver.find_elements(By.XPATH, '//a[@class="page-link" and contains(text(), "Next")]')
    if next_page_button and next_page_button[0].is_enabled():
        next_page_button[0].click()  # Klik tombol "Next" untuk halaman berikutnya
        print("➡️ Beralih ke halaman berikutnya...")
        page_number += 1
        time.sleep(3)
    else:
        print("✅ Semua dokumen telah diunduh.")
        break

# === Tutup Browser ===
driver.quit()