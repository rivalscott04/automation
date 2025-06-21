import os, time, pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains 
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys

# === KONFIGURASI ===
EMAIL = "199108272020122026@kemenag.go.id"
PASSWORD = "elok01"
FOLDER_PATH = r"C:\Users\rival\Documents\SPMT_CPNS"
TTE_LOGIN = "https://tte.kemenag.go.id"
UPLOAD_CREATE = "https://tte.kemenag.go.id/satker/dokumen/naskah/create"

# === TRACKING VARIABLES ===
processed_files = []
failed_files = []
success_count = 0
failed_count = 0

def safe_click(driver, selectors, description="elemen"):
    """Fungsi untuk mencoba beberapa selector dan metode klik"""
    clicked = False
    
    for selector in selectors:
        try:
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, selector)))
            
            # Scroll ke elemen
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            time.sleep(0.2)
            
            # Coba klik normal
            element.click()
            print(f"✅ {description} berhasil diklik dengan selector: {selector}")
            clicked = True
            break
            
        except TimeoutException:
            print(f"⚠️ Timeout untuk selector: {selector}")
            continue
        except Exception as e:
            print(f"⚠️ Error dengan selector {selector}: {e}")
            # Coba JavaScript click
            try:
                element = driver.find_element(By.XPATH, selector)
                driver.execute_script("arguments[0].click();", element)
                print(f"✅ {description} berhasil diklik dengan JavaScript: {selector}")
                clicked = True
                break
            except:
                continue
    
    return clicked

def select_dropdown_option(driver, dropdown_selectors, option_text, description="dropdown"):
    """Fungsi untuk memilih opsi dari dropdown Select2 dengan optimasi JavaScript"""
    wait = WebDriverWait(driver, 6)  # Mengurangi timeout
    
    # === METODE JAVASCRIPT CEPAT ===
    if "anchor" in description.lower():
        try:
            result = driver.execute_script(f"""
                // Coba set value langsung ke select element
                var selects = document.querySelectorAll('select[name="jenis_ttd"], #jenis_ttd');
                for (var i = 0; i < selects.length; i++) {{
                    var select = selects[i];
                    if (select) {{
                        select.value = '{option_text}';
                        select.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        return true;
                    }}
                }}
                return false;
            """)
            
            if result:
                print(f"✅ {description} berhasil dipilih dengan JavaScript: {option_text}")
                return True
        except Exception as e:
            print(f"⚠️ JavaScript method gagal: {e}")
    
    # === METODE SELENIUM FALLBACK ===
    dropdown_clicked = False
    
    # Buka dropdown
    for selector in dropdown_selectors:
        try:
            dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
            driver.execute_script("arguments[0].click();", dropdown)  # JavaScript click lebih cepat
            print(f"✅ {description} terbuka dengan selector: {selector}")
            dropdown_clicked = True
            break
        except Exception as e:
            print(f"⚠️ Gagal membuka {description} dengan selector {selector}: {e}")
            continue
    
    if not dropdown_clicked:
        print(f"🚫 Gagal membuka {description}!")
        return False
    
    time.sleep(0.1)  # Mengurangi delay
    
    # Pilih opsi - untuk anchor, langsung pilih # tanpa mencari posisi spesifik
    if "anchor" in description.lower():
        option_selectors = [
            f'//li[contains(@class, "select2-results__option") and text()="{option_text}"]',
            f'//li[contains(@class, "select2-results__option") and contains(text(), "{option_text}")]',
            '//*[@class="select2-results__options"]//li[1]'  # Pilih opsi pertama jika ada
        ]
    else:
        option_selectors = [
            f'//li[contains(@class, "select2-results__option") and contains(text(), "{option_text}")]',
            f'//*[@class="select2-results__options"]//li[contains(text(), "{option_text}")]',
            f'//ul[@class="select2-results__options"]/li[contains(text(), "{option_text}")]',
            '//*[@class="select2-results__options"]//li[1]'  # Fallback: pilih opsi pertama
        ]
    
    # Coba metode JavaScript untuk klik opsi
    try:
        js_result = driver.execute_script(f"""
            var selectors = [
                'li.select2-results__option:contains("{option_text}")',
                '.select2-results__options li:contains("{option_text}")',
                '.select2-results__options li:first-child'
            ];
            
            for (var i = 0; i < selectors.length; i++) {{
                var element = document.querySelector(selectors[i]);
                if (element) {{
                    element.click();
                    return true;
                }}
            }}
            return false;
        """)
        
        if js_result:
            print(f"✅ Opsi '{option_text}' dipilih dari {description} dengan JavaScript")
            return True
    except Exception:
        pass  # Lanjut ke metode Selenium
    
    # Fallback ke metode Selenium
    for selector in option_selectors:
        try:
            option = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
            driver.execute_script("arguments[0].click();", option)  # JavaScript click lebih cepat
            print(f"✅ Opsi '{option_text}' dipilih dari {description}")
            return True
        except Exception:
            continue
    
    print(f"🚫 Gagal memilih opsi '{option_text}' dari {description}")
    return False

def select_bootstrap_anchor(driver, anchor_symbol="#"):
    """Fungsi efisien untuk memilih anchor dengan prioritas JavaScript method untuk kecepatan"""
    wait = WebDriverWait(driver, 5)  # Mengurangi timeout
    print(f"📍 Memilih anchor dokumen ({anchor_symbol})...")
    
    # === METODE 1: JAVASCRIPT SUPER CEPAT ===
    print("🔄 Metode 1: JavaScript Manual (Cepat)...")
    try:
        # Metode JavaScript paling cepat - langsung set semua select yang mungkin
        result = driver.execute_script(f"""
            // Coba semua kemungkinan selector
            var selectors = ['select[name="anchor[]"]', 'select#anchor', 'select[name*="anchor"]'];
            var found = false;
            
            for (var i = 0; i < selectors.length; i++) {{
                var elements = document.querySelectorAll(selectors[i]);
                if (elements.length > 0) {{
                    for (var j = 0; j < elements.length; j++) {{
                        elements[j].value = '{anchor_symbol}';
                        elements[j].dispatchEvent(new Event('change'));
                        
                        // Trigger Bootstrap refresh jika ada
                        if (window.jQuery && typeof jQuery(elements[j]).selectpicker === 'function') {{
                            jQuery(elements[j]).selectpicker('refresh');
                        }}
                        found = true;
                    }}
                }}
            }}
            return found;
        """)
        
        if result:
            print(f"✅ Anchor '{anchor_symbol}' berhasil dipilih dengan JavaScript manual")
            return True
        else:
            print("⚠️ Tidak ada select element anchor yang ditemukan")
                
    except Exception as e:
        print(f"⚠️ Metode 1 (JavaScript Super Cepat) gagal: {e}")
    
    # === METODE 2: BOOTSTRAP SELECTPICKER (FALLBACK) ===
    print("🔄 Metode 2: Bootstrap Selectpicker (Fallback)...")
    try:
        # Cari dan klik button Bootstrap Selectpicker untuk anchor
        anchor_button_selectors = [
            '//button[@data-id="anchor" and contains(@class, "dropdown-toggle")]',
            '//select[@name="anchor[]"]/..//button[contains(@class, "dropdown-toggle")]'
        ]
        
        anchor_button_clicked = False
        for selector in anchor_button_selectors:
            try:
                anchor_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                driver.execute_script("arguments[0].click();", anchor_button)  # JavaScript click lebih cepat
                print(f"✅ Bootstrap Selectpicker button diklik: {selector}")
                anchor_button_clicked = True
                break
            except Exception as e:
                print(f"⚠️ Gagal dengan selector {selector}: {e}")
                continue
        
        if anchor_button_clicked:
            time.sleep(0.2)  # Mengurangi delay
            
            # Pilih opsi anchor dari Bootstrap Selectpicker dropdown
            anchor_option_selectors = [
                f'//div[contains(@class, "dropdown-menu show")]//a[contains(@class, "dropdown-item") and .//span[@class="text" and text()="{anchor_symbol}"]]',
                f'//a[@role="option" and contains(@class, "dropdown-item") and contains(., "{anchor_symbol}")]'
            ]
            
            for selector in anchor_option_selectors:
                try:
                    anchor_option = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    driver.execute_script("arguments[0].click();", anchor_option)  # JavaScript click lebih cepat
                    print(f"✅ Anchor '{anchor_symbol}' berhasil dipilih dengan Bootstrap Selectpicker")
                    return True
                except Exception:
                    continue
    
    except Exception as e:
        print(f"⚠️ Metode 2 (Bootstrap Selectpicker) gagal: {e}")
    
    # Jika kedua metode gagal
    print(f"🚫 Gagal memilih anchor '{anchor_symbol}' dengan kedua metode!")
    return False

def upload_file_robust(driver, file_path, file_name):
    """
    Fungsi robust untuk upload file dengan multiple fallback methods
    """
    print(f"📁 Mengupload file: {file_name}")
    
    # Method 1: Cari input file tersembunyi dan set value langsung
    try:
        print("🔄 Method 1: Mencari input file tersembunyi...")
        file_input_selectors = [
            '//input[@type="file"]',
            '//input[@name="file"]',
            '//input[@accept=".pdf"]',
            '//input[contains(@accept, "pdf")]',
            '//*[@type="file" and not(@style="display: none")]'
        ]
        
        for selector in file_input_selectors:
            try:
                file_input = driver.find_element(By.XPATH, selector)
                if file_input.is_displayed() or not file_input.get_attribute("style"):
                    # Langsung set file path ke input
                    file_input.send_keys(file_path)
                    print(f"✅ Method 1 berhasil dengan selector: {selector}")
                    time.sleep(0.5)
                    return True
            except Exception as e:
                print(f"⚠️ Method 1 selector {selector} gagal: {e}")
                continue
                
    except Exception as e:
        print(f"⚠️ Method 1 gagal: {e}")
    
    # Method 2: Klik tombol dan gunakan dialog dengan retry mechanism
    try:
        print("🔄 Method 2: Menggunakan dialog file dengan retry...")
        
        # Klik tombol cari file
        cari_file_selectors = [
            '/html/body/section/div/section/div/div/section/form/div[1]/div[5]/div/div[1]/div/button',
            '//button[contains(text(), "Cari File")]',
            '//button[contains(text(), "Browse")]',
            '//button[contains(text(), "Pilih File")]',
            '//*[@type="button" and contains(@class, "btn") and (contains(text(), "Cari") or contains(text(), "Browse"))]'
        ]
        
        button_clicked = False
        for selector in cari_file_selectors:
            try:
                button = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, selector)))
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                time.sleep(0.3)
                button.click()
                print(f"✅ Tombol file berhasil diklik dengan selector: {selector}")
                button_clicked = True
                break
            except Exception as e:
                print(f"⚠️ Gagal klik tombol dengan selector {selector}: {e}")
                continue
        
        if not button_clicked:
            print("🚫 Gagal mengklik tombol cari file!")
            return False
        
        # Tunggu dialog muncul
        print("⏳ Menunggu dialog file muncul...")
        time.sleep(1.5)
        
        # Gunakan pyautogui dengan retry mechanism
        max_retries = 3
        for attempt in range(max_retries):
            try:
                print(f"🔄 Attempt {attempt + 1}/{max_retries} - Mengetik path file...")
                
                # Clear field dulu (jika ada)
                pyautogui.hotkey('ctrl', 'a')  # Select all
                time.sleep(0.2)
                
                # Ketik path file
                pyautogui.write(file_path, interval=0.03)  # Lebih cepat
                time.sleep(0.5)
                
                # Tekan Enter
                pyautogui.press('enter')
                time.sleep(1)
                
                # Cek apakah file berhasil dipilih dengan melihat nama file di form
                file_selected = check_file_selected(driver, file_name)
                if file_selected:
                    print(f"✅ Method 2 berhasil pada attempt {attempt + 1}")
                    return True
                else:
                    print(f"⚠️ Attempt {attempt + 1} gagal - file belum terpilih")
                    if attempt < max_retries - 1:
                        # Coba buka dialog lagi
                        button.click()
                        time.sleep(1)
                        
            except Exception as e:
                print(f"⚠️ Method 2 attempt {attempt + 1} error: {e}")
                if attempt < max_retries - 1:
                    time.sleep(1)
                    
    except Exception as e:
        print(f"⚠️ Method 2 gagal: {e}")
    
    # Method 3: Gunakan JavaScript untuk inject file
    try:
        print("🔄 Method 3: Menggunakan JavaScript injection...")
        
        # Cari semua input file (termasuk yang hidden)
        file_inputs = driver.find_elements(By.XPATH, '//input[@type="file"]')
        
        for i, file_input in enumerate(file_inputs):
            try:
                # Buat input visible jika hidden
                driver.execute_script("arguments[0].style.display = 'block'; arguments[0].style.visibility = 'visible';", file_input)
                
                # Set file
                file_input.send_keys(file_path)
                print(f"✅ Method 3 berhasil dengan input ke-{i}")
                time.sleep(2)
                
                # Trigger change event
                driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", file_input)
                
                # Cek apakah berhasil
                if check_file_selected(driver, file_name):
                    return True
                    
            except Exception as e:
                print(f"⚠️ Method 3 input ke-{i} gagal: {e}")
                continue
                
    except Exception as e:
        print(f"⚠️ Method 3 gagal: {e}")
    
    # Method 4: Manual dengan koordinat (last resort)
    try:
        print("🔄 Method 4: Manual positioning dengan koordinat...")
        
        # Klik tombol cari file
        button = driver.find_element(By.XPATH, '/html/body/section/div/section/div/div/section/form/div[1]/div[5]/div/div[1]/div/button')
        button.click()
        time.sleep(3)
        
        # Coba beberapa cara untuk fokus ke dialog
        pyautogui.press('tab')  # Tab ke field file name
        time.sleep(0.5)
        
        # Ketik nama file saja (bukan full path)
        pyautogui.write(file_name, interval=0.1)
        time.sleep(1)
        
        pyautogui.press('enter')
        time.sleep(2)
        
        if check_file_selected(driver, file_name):
            print("✅ Method 4 berhasil")
            return True
            
    except Exception as e:
        print(f"⚠️ Method 4 gagal: {e}")
    
    print("🚫 Semua method gagal!")
    return False

def check_file_selected(driver, file_name):
    """
    Cek apakah file sudah terpilih dengan mencari indikator di halaman
    """
    try:
        # Cari indikator file terpilih
        file_indicators = [
            f'//*[contains(text(), "{file_name}")]',
            f'//*[contains(@value, "{file_name}")]',
            '//input[@type="file"][@value!=""]',
            '//*[@class="file-name" or @class="filename"]',
            '//*[contains(@class, "file") and contains(@class, "selected")]'
        ]
        
        for indicator in file_indicators:
            try:
                elements = driver.find_elements(By.XPATH, indicator)
                if elements and any(elem.text or elem.get_attribute('value') for elem in elements):
                    print(f"✅ File terdeteksi terpilih dengan indikator: {indicator}")
                    return True
            except:
                continue
        
        # Cek juga dengan JavaScript
        try:
            file_inputs = driver.find_elements(By.XPATH, '//input[@type="file"]')
            for file_input in file_inputs:
                value = file_input.get_attribute('value')
                if value and (file_name in value or value.endswith('.pdf')):
                    print("✅ File terdeteksi terpilih via JavaScript check")
                    return True
        except:
            pass
            
        return False
        
    except Exception as e:
        print(f"⚠️ Error saat cek file selection: {e}")
        return False

# Update fungsi process_single_file untuk menggunakan upload yang baru
def process_single_file(driver, file_path, file_name):
    """Memproses satu file PDF dengan upload yang diperbaiki"""
    global success_count, failed_count, processed_files, failed_files
    
    print(f"\n{'='*80}")
    print(f"📄 MEMPROSES FILE: {file_name}")
    print(f"{'='*80}")
    
    try:
        wait = WebDriverWait(driver, 20)
        
        # Buka halaman upload
        print("📄 Membuka halaman form upload...")
        driver.get(UPLOAD_CREATE)
        wait.until(EC.presence_of_element_located((By.XPATH, '//form')))
        print("📄 Halaman form upload terbuka")

        # Pilih jenis dokumen
        print("📋 Memilih jenis dokumen...")
        jenis_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[@class="select2-selection__rendered"]')))
        jenis_dropdown.click()
        time.sleep(1)
        
        # Pilih "Dokumen Lain-lain"
        dokumen_lain = wait.until(EC.element_to_be_clickable((By.XPATH, '//li[contains(text(), "Dokumen Lain-Lain")]')))
        dokumen_lain.click()
        print("✅ Dokumen Lain-Lain dipilih")

        # Isi Perihal Dokumen
        input_ket = '/html/body/section/div/section/div/div/section/form/div[1]/div[3]/div/input'
        wait.until(EC.presence_of_element_located((By.XPATH, input_ket))).clear()
        driver.find_element(By.XPATH, input_ket).send_keys("SPMT CPNS T.A 2024 Update Unit Kerja")

        # GUNAKAN FUNGSI UPLOAD YANG DIPERBAIKI
        file_uploaded = upload_file_robust(driver, file_path, file_name)
        
        if not file_uploaded:
            raise Exception("Gagal mengupload file")
        
        print(f"✅ File berhasil dipilih: {file_name}")
        time.sleep(0.5)

        # Lanjutkan dengan proses selanjutnya...

        # Klik tombol "Lanjut" 
        lanjut_selectors = [
            "//button[contains(text(), 'Lanjut')]",
            "//*[contains(@class, 'btn') and contains(text(), 'Lanjut')]",
            '/html/body/section/div/section/div/div/section/form/div[2]/button'
        ]
        
        lanjut_clicked = safe_click(driver, lanjut_selectors, "Tombol Lanjut")
        
        if not lanjut_clicked:
            print("🚫 Gagal mengklik tombol Lanjut!")
            driver.save_screenshot(f"debug_lanjut_error_{file_name}_{int(time.time())}.png")
            raise Exception("Gagal klik tombol Lanjut")

        time.sleep(2)
        print("🎉 Berhasil melanjutkan ke tahap berikutnya!")

        # Klik tombol "Lanjut" di halaman pemilihan pegawai
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div/section/div[2]/div/section/div[2]/a[2]'))).click()
        time.sleep(2)

        # === PEMILIHAN PENANDATANGAN ===
        print("🔍 Mencari dropdown pemilihan pegawai...")
        
        # Dropdown pegawai
        pegawai_selectors = [
            '//*[@id="select2-pegawai_id-container"]',
            '//span[contains(@class, "select2-selection__rendered")]',
            '//select[@name="pegawai_id"]/..//span[@class="select2-selection__rendered"]',
            '//*[contains(@class, "select2-container")]//span[@class="select2-selection__rendered"]'
        ]
        
        pegawai_clicked = False
        for selector in pegawai_selectors:
            try:
                pegawai_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                time.sleep(1)
                pegawai_dropdown.click()
                print(f"✅ Dropdown pegawai terbuka dengan selector: {selector}")
                pegawai_clicked = True
                break
            except Exception as e:
                print(f"⚠️ Gagal dengan selector {selector}: {e}")
                continue
        
        if not pegawai_clicked:
            print("🚫 Gagal membuka dropdown pegawai!")
            driver.save_screenshot(f"debug_pegawai_dropdown_{file_name}_{int(time.time())}.png")
            raise Exception("Gagal membuka dropdown pegawai")

        time.sleep(1)

        # Ketik nama pegawai untuk pencarian
        search_input_selectors = [
            '//input[@class="select2-search__field"]',
            '//input[contains(@class, "select2-search")]',
            '//*[@class="select2-container--open"]//input'
        ]
        
        search_typed = False
        for selector in search_input_selectors:
            try:
                search_input = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                search_input.clear()
                search_input.send_keys("zamroni aziz")
                print(f"✅ Nama pegawai berhasil diketik dengan selector: {selector}")
                search_typed = True
                break
            except Exception as e:
                print(f"⚠️ Gagal mengetik dengan selector {selector}: {e}")
                continue
        
        # Fallback dengan pyautogui jika gagal
        if not search_typed:
            print("🔄 Fallback: Menggunakan pyautogui untuk mengetik...")
            try:
                pyautogui.write("zamroni aziz", interval=0.05)
                search_typed = True
                print("✅ Nama pegawai berhasil diketik dengan pyautogui")
            except Exception as e:
                print(f"⚠️ Gagal dengan pyautogui: {e}")
        
        if not search_typed:
            print("🚫 Gagal mengetik nama pegawai!")
            driver.save_screenshot(f"debug_pegawai_search_{file_name}_{int(time.time())}.png")
            raise Exception("Gagal mengetik nama pegawai")
        
        time.sleep(0.3)  # Tunggu hasil pencarian

        # Tunggu dan klik suggestion yang muncul
        suggestion_selectors = [
            '//li[contains(@class, "select2-results__option") and contains(text(), "ZAMRONI AZIZ")]',
            '//li[contains(@class, "select2-results__option") and contains(text(), "zamroni aziz")]',
            '//*[@class="select2-results__options"]//li[1]'  # Pilih hasil pertama
        ]
        
        suggestion_clicked = False
        for selector in suggestion_selectors:
            try:
                suggestion = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, selector)))
                suggestion.click()
                print(f"✅ Suggestion pegawai diklik dengan selector: {selector}")
                suggestion_clicked = True
                break
            except Exception as e:
                print(f"⚠️ Gagal klik suggestion dengan selector {selector}: {e}")
                continue
        
        # Fallback: tekan Enter jika klik gagal
        if not suggestion_clicked:
            print("🔄 Fallback: Menekan Enter untuk memilih suggestion...")
            try:
                pyautogui.press('enter')
                suggestion_clicked = True
                print("✅ Suggestion dipilih dengan Enter")
            except Exception as e:
                print(f"⚠️ Gagal dengan Enter: {e}")
        
        if not suggestion_clicked:
            print("🚫 Gagal memilih pegawai!")
            driver.save_screenshot(f"debug_suggestion_{file_name}_{int(time.time())}.png")
            raise Exception("Gagal memilih pegawai")
        
        time.sleep(0.5)

        # === PILIH JENIS ANCHOR DOCUMENT ===
        print("📍 Memilih jenis anchor document...")
        
        # Metode cepat: Langsung set value dengan JavaScript jika ada select element
        jenis_selected = False
        try:
            print("🔄 Metode JavaScript cepat untuk jenis anchor...")
            select_element = driver.find_element(By.XPATH, '//select[@name="jenis_ttd"]')
            if select_element:
                # Set value langsung dengan JavaScript
                driver.execute_script("""
                    var select = arguments[0];
                    select.value = 'Digital';
                    select.dispatchEvent(new Event('change'));
                    // Trigger select2 update jika ada
                    if (window.jQuery && jQuery(select).data('select2')) {
                        jQuery(select).trigger('change');
                    }
                """, select_element)
                print("✅ Jenis anchor berhasil dipilih dengan JavaScript")
                jenis_selected = True
        except Exception as e:
            print(f"⚠️ Metode JavaScript gagal: {e}")
        
        # Fallback ke metode dropdown biasa jika JavaScript gagal
        if not jenis_selected:
            jenis_anchor_selectors = [
                '//select[@name="jenis_ttd"]/..//span[@class="select2-selection__rendered"]',
                '//*[contains(@class, "select2-container") and .//select[@name="jenis_ttd"]]//span[@class="select2-selection__rendered"]',
                '//span[contains(@class, "select2-selection__rendered") and contains(text(), "Pilih Jenis")]'
            ]
            
            jenis_selected = select_dropdown_option(
                driver, 
                jenis_anchor_selectors, 
                "Digital",
                "dropdown jenis anchor"
            )
        
        if not jenis_selected:
            print("🚫 Gagal memilih jenis anchor document!")
            driver.save_screenshot(f"debug_jenis_anchor_{file_name}_{int(time.time())}.png")
            # Lanjutkan proses meskipun gagal
        
        time.sleep(0.3)

        # === PILIH ANCHOR DOKUMEN MENGGUNAKAN BOOTSTRAP SELECTPICKER ===
        anchor_selected = select_bootstrap_anchor(driver, "#")
        
        if not anchor_selected:
            print("🚫 Gagal memilih anchor (#) menggunakan Bootstrap Selectpicker!")
            print("🔄 Mencoba metode manual...")
            
            # Metode manual sebagai fallback
            try:
                # Cari select element langsung
                select_element = driver.find_element(By.XPATH, '//select[@name="anchor[]" or @id="anchor"]')
                if select_element:
                    # Gunakan JavaScript untuk set value
                    driver.execute_script("arguments[0].value = '#'; arguments[0].dispatchEvent(new Event('change'));", select_element)
                    print("✅ Anchor (#) berhasil dipilih dengan metode JavaScript")
                    anchor_selected = True
                else:
                    print("🚫 Tidak dapat menemukan select element untuk anchor")
            except Exception as e:
                print(f"⚠️ Metode manual juga gagal: {e}")
                driver.save_screenshot(f"debug_bootstrap_anchor_{file_name}_{int(time.time())}.png")
        else:
            print("✅ Anchor (#) berhasil dipilih dari Bootstrap Selectpicker")

        time.sleep(1)

        # === TAMBAH PENANDATANGAN ===
        print("➕ Menambah penandatangan...")
        
        # Klik tombol "Tambah Penandatangan"
        tambah_selectors = [
            '//button[@type="submit" and contains(text(), "Tambah Penandatangan")]',
            '//button[contains(text(), "Tambah Penandatangan")]',
            '//*[contains(@class, "btn") and contains(text(), "Tambah")]',
            '//input[@type="submit" and contains(@value, "Tambah")]'
        ]
        
        tambah_clicked = safe_click(driver, tambah_selectors, "Tombol Tambah Penandatangan")
        
        if tambah_clicked:
            print("✅ Penandatangan berhasil ditambahkan!")
            time.sleep(0.8)
            
            # Cek apakah berhasil dengan mencari konfirmasi atau redirect
            try:
                # Tunggu halaman berubah atau muncul pesan sukses
                success_indicators = [
                    '//*[contains(text(), "berhasil")]',
                    '//*[contains(text(), "sukses")]',
                    '//*[contains(@class, "alert-success")]',
                    '//table//tr[contains(., "zamroni aziz")]'  # Cek apakah pegawai muncul di tabel
                ]
                
                for indicator in success_indicators:
                    try:
                        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, indicator)))
                        print("✅ Konfirmasi: Proses berhasil!")
                        break
                    except:
                        continue
                        
            except Exception as e:
                print(f"⚠️ Tidak dapat mengkonfirmasi keberhasilan: {e}")
            
            # === KLIK TOMBOL SELESAI ===
            print("🏁 Mengklik tombol Selesai...")
            
            selesai_selectors = [
                '//button[@type="submit" and contains(@class, "btn-success") and contains(text(), "Selesai")]',
                '//button[contains(@class, "btn-success") and contains(text(), "Selesai")]',
                '//button[contains(text(), "Selesai") and contains(@class, "float-right")]',
                '//button[@type="submit" and contains(text(), "Selesai")]',
                '//*[contains(@class, "btn") and contains(text(), "Selesai")]'
            ]
            
            selesai_clicked = safe_click(driver, selesai_selectors, "Tombol Selesai")
            
            if selesai_clicked:
                print("✅ Tombol Selesai berhasil diklik!")
                time.sleep(1.5)  # Tunggu proses selesai
                
                # Tunggu konfirmasi atau redirect setelah selesai
                try:
                    # Bisa jadi redirect ke halaman list atau muncul pesan sukses
                    WebDriverWait(driver, 6).until(
                        lambda d: d.current_url != driver.current_url or 
                                 len(d.find_elements(By.XPATH, '//*[contains(text(), "berhasil") or contains(text(), "sukses")]')) > 0
                    )
                    print("✅ Dokumen berhasil diselesaikan!")
                except Exception as e:
                    print(f"⚠️ Tidak dapat mengkonfirmasi penyelesaian: {e}")
            else:
                print("🚫 Gagal mengklik tombol Selesai!")
                driver.save_screenshot(f"debug_selesai_{file_name}_{int(time.time())}.png")
                # Tidak throw exception, karena mungkin dokumen sudah tersimpan
                
        else:
            print("🚫 Gagal menambah penandatangan!")
            driver.save_screenshot(f"debug_tambah_penandatangan_{file_name}_{int(time.time())}.png")
            raise Exception("Gagal menambah penandatangan")

        # Jika sampai sini berarti berhasil
        success_count += 1
        processed_files.append(file_name)
        print(f"✅ FILE BERHASIL DIPROSES: {file_name}")
        
        # Delay sejenak sebelum memproses file berikutnya
        time.sleep(1)
        
        return True
        
    except Exception as e:
        failed_count += 1
        failed_files.append({"file": file_name, "error": str(e)})
        print(f"❌ FILE GAGAL DIPROSES: {file_name}")
        print(f"❌ Error: {e}")
        driver.save_screenshot(f"debug_error_{file_name}_{int(time.time())}.png")
        return False

def print_final_report():
    """Cetak laporan akhir"""
    print(f"\n{'='*100}")
    print("📊 LAPORAN AKHIR PEMROSESAN BATCH")
    print(f"{'='*100}")
    print(f"✅ Berhasil diproses: {success_count} file")
    print(f"❌ Gagal diproses: {failed_count} file")
    print(f"📊 Total file: {success_count + failed_count}")
    
    if processed_files:
        print(f"\n✅ DAFTAR FILE BERHASIL:")
        for i, file in enumerate(processed_files, 1):
            print(f"   {i}. {file}")
    
    if failed_files:
        print(f"\n❌ DAFTAR FILE GAGAL:")
        for i, item in enumerate(failed_files, 1):
            print(f"   {i}. {item['file']} - Error: {item['error']}")
    
    print(f"\n{'='*100}")

# === MAIN PROCESS ===
try:
    # === SETUP EDGE ===
    options = webdriver.EdgeOptions()
    options.use_chromium = True
    driver = webdriver.Edge(options=options)
    driver.get(TTE_LOGIN)
    wait = WebDriverWait(driver, 20)

    # === LOGIN ADMIN SATKER ===
    print("🔐 Memulai proses login...")
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div/div[1]/div/div/div[2]/form/div[1]/div/div[3]/input'))).click()
    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/section/div/div[1]/div/div/div[2]/form/div[4]/div/input'))).send_keys(EMAIL)
    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/section/div/div[1]/div/div/div[2]/form/div[5]/div/input'))).send_keys(PASSWORD)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div/div[1]/div/div/div[2]/form/div[6]/div[2]/button'))).click()

    # Tunggu login selesai
    wait.until(EC.presence_of_element_located((By.XPATH, '//nav')))
    print("✅ Login berhasil")

    # === AMBIL DAFTAR FILE PDF ===
    print("📁 Mengambil daftar file PDF...")
    pdf_files = [f for f in os.listdir(FOLDER_PATH) if f.endswith('.pdf')]
    
    if not pdf_files:
        print("⚠️ Tidak ada file PDF di folder!")
        exit()
    
    print(f"📄 Ditemukan {len(pdf_files)} file PDF")
    
    # Konfirmasi sebelum memulai batch processing
    print(f"\n🚀 AKAN MEMPROSES {len(pdf_files)} FILE PDF")
    print("📝 Tekan Enter untuk melanjutkan atau Ctrl+C untuk membatalkan...")
    input()

    # === BATCH PROCESSING ===
    print(f"\n{'='*100}")
    print("🚀 MEMULAI BATCH PROCESSING")
    print(f"{'='*100}")
    
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"\n📋 Progress: {i}/{len(pdf_files)} ({(i/len(pdf_files)*100):.1f}%)")
        
        file_path = os.path.join(FOLDER_PATH, pdf_file)
        
        # Proses file
        result = process_single_file(driver, file_path, pdf_file)
        
        # Jeda sejenak antara file untuk menghindari overload
        if i < len(pdf_files):  # Tidak perlu delay setelah file terakhir
            print("⏳ Menunggu 5 detik sebelum file berikutnya...")
            time.sleep(5)

    # === CETAK LAPORAN AKHIR ===
    print_final_report()

except KeyboardInterrupt:
    print(f"\n⚠️ Proses dibatalkan oleh user!")
    print(f"✅ Berhasil diproses: {success_count} file")
    print(f"❌ Gagal diproses: {failed_count} file")
    print_final_report()
except Exception as e:
    print(f"❌ Error umum: {e}")
    driver.save_screenshot(f"debug_main_error_{int(time.time())}.png")
    print("📸 Screenshot error disimpan")
    print_final_report()
finally:
    print(f"\n{'='*100}")
    print("🏁 BATCH PROCESSING SELESAI")
    print_final_report()
    print(f"{'='*100}")
    
    # Uncomment jika ingin menutup browser otomatis
    # driver.quit()
    input("Tekan Enter untuk menutup browser...")
    driver.quit()