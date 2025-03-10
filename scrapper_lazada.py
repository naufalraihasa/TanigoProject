import time
import datetime
import pandas as pd
from tqdm import tqdm

# Import library Selenium untuk automasi browser
from selenium import webdriver as wb
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def scrolling(driver):
    """
    Melakukan scroll perlahan ke bawah halaman agar elemen-elemen termuat dengan baik.
    
    Parameter:
        driver (webdriver): Instance dari Selenium WebDriver.
    """
    scheight = 0.1
    while scheight < 9.9:
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight / {});".format(scheight)
        )
        time.sleep(0.3)
        scheight += 0.1

def reverse_scrolling(driver):
    """
    Melakukan scroll dengan cara menekan tombol PAGE_DOWN secara berulang agar elemen-elemen tampil sempurna.
    
    Parameter:
        driver (webdriver): Instance dari Selenium WebDriver.
    """
    body = driver.find_element(By.TAG_NAME, 'body')
    for _ in range(25):
        body.send_keys(Keys.PAGE_DOWN)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)

def extract_data(driver, product_data):
    """
    Mengekstrak data dari produk-produk yang tampil di halaman saat ini.
    Termasuk mengambil detail produk dengan membuka halaman produk secara terpisah.
    
    Parameter:
        driver (webdriver): Instance dari Selenium WebDriver.
        product_data (list): List untuk menyimpan data produk yang berhasil diambil.
    """
    # Mekanisme retry untuk memastikan elemen produk termuat dengan baik
    max_retries = 3
    for attempt in range(max_retries):
        try:
            scrolling(driver)
            data_items = wait(driver, 10).until(
                EC.visibility_of_all_elements_located((By.XPATH, '//div[contains(@class, "Bm3ON")]'))
            )
            break  # Jika berhasil, keluar dari loop retry
        except TimeoutException:
            if attempt == max_retries - 1:
                raise
            driver.refresh()
            time.sleep(3)

    for item in tqdm(data_items, desc="Memproses produk"):
        # Ambil nama produk
        try:
            name = item.find_element(By.XPATH, './/div[@class="RfADt"]').text
        except Exception:
            continue
        
        # Ambil harga produk
        try:
            price = item.find_element(By.XPATH, './/span[@class="ooOxS"]').text
        except Exception:
            continue
        
        # Ambil lokasi produk
        try:
            location = item.find_element(By.XPATH, './/span[@class="oa6ri "]').text
        except Exception:
            continue
        
        # Ambil link detail produk
        try:
            link_element = item.find_element(By.XPATH, './/a')
            details_link = link_element.get_attribute('href')
        except Exception:
            continue

        # Ambil jumlah produk yang terjual
        try:
            sold = item.find_element(By.XPATH, './/span[@class="_1cEkb"]').text
        except Exception:
            sold = None

        # --- BAGIAN TAMBAHAN: PENGAMBILAN DESKRIPSI, TOKO, BRAND, DAN RATING PRODUK ---
        description = None
        store = None
        brand = None
        rating = 0.0  # default rating 0
        
        # Dictionary pemetaan substring base64 ke nilai bintang
        BASE64_TO_STAR = {
            "ASUVORK5CYII=": 0.0,  # 0 bintang
            "ElFTkSuQmCC":   0.5,  # 0.5 bintang
            "BJRU5ErkJggg==": 1.0  # 1 bintang
        }

        main_window = driver.current_window_handle

        try:
            # Buka link detail produk di tab baru
            driver.execute_script(f"window.open('{details_link}','_blank');")
            driver.switch_to.window(driver.window_handles[-1])

            # Tunggu hingga elemen detail produk termuat (sesuaikan selector jika perlu)
            
            # description
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="html-content detail-content"]'))
            )
            
            # brand
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//a[@class="pdp-link pdp-link_size_s pdp-link_theme_blue pdp-product-brand__brand-link"]'))
            )
            
            # rating
            wait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'i._9-ogB.Dy1nx'))
            )
            
            # store
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//a[@class="pdp-link pdp-link_size_l pdp-link_theme_black seller-name__detail-name"]'))
            )

            # Ambil deskripsi produk
            try:
                description = driver.find_element(By.XPATH, '//div[@class="html-content detail-content"]').text
            except NoSuchElementException:
                pass

            # Ambil nama toko
            try:
                store = driver.find_element(By.XPATH, '//a[@class="pdp-link pdp-link_size_l pdp-link_theme_black seller-name__detail-name"]').text
            except NoSuchElementException:
                pass
            
            # Ambil nama brand
            try:
                brand = driver.find_element(By.XPATH, '//a[@class="pdp-link pdp-link_size_s pdp-link_theme_blue pdp-product-brand__brand-link"]').text
            except NoSuchElementException:
                pass
            
            # Ambil rating produk berbasis <i class="_9-ogB Dy1nx">
            try:
                rating_elements = driver.find_elements(By.CSS_SELECTOR, 'i._9-ogB.Dy1nx')
                total_stars = 0.0
                for rating_elem in rating_elements:
                    # Ambil properti background-image dari elemen rating
                    bg_img = rating_elem.value_of_css_property("background-image")
                    # bg_img diharapkan berbentuk: url("data:image/png;base64,...")
                    if bg_img.startswith('url("') and bg_img.endswith('")'):
                        data_url = bg_img[5:-2]  # hapus 'url("' dan '")'
                    else:
                        data_url = ""
                    matched = False
                    for base64_pattern, star_value in BASE64_TO_STAR.items():
                        if base64_pattern in data_url:
                            total_stars += star_value
                            matched = True
                            break
                rating = total_stars
            except Exception:
                rating = 0.0

        except (TimeoutException, NoSuchElementException):
            pass
        finally:
            driver.close()
            driver.switch_to.window(main_window)
        # --- SELESAI BAGIAN PENGAMBILAN DESKRIPSI, TOKO, BRAND, DAN RATING ---

        data = {
            'name': name,
            'price': price,
            'store': store,
            'brand': brand,
            'location': location,
            'sold': sold,
            'rating': rating,
            'details_link': details_link,
            'description': description
        }
        product_data.append(data)

def main():
    driver = wb.Chrome()
    driver.get('https://www.lazada.co.id/')
    driver.implicitly_wait(5)

    keywords = input("Keywords: ")
    pages = int(input("Pages: "))

    # Cari elemen input pencarian dan masukkan kata kunci
    search = driver.find_element(By.XPATH, '//input[@class="search-box__input--O34g"]')
    search.send_keys(keywords)
    search.send_keys(Keys.ENTER)
    driver.implicitly_wait(5)

    product_data = []

    for page in range(1, pages + 1):
        print(f"\n--- Halaman {page} ---")
        extract_data(driver, product_data)

        try:
            next_page = wait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[aria-label="right"]'))
            )
            next_page.click()
        except TimeoutException:
            print("Tombol halaman berikutnya tidak ditemukan, mencoba ulang...")
            driver.refresh()
            time.sleep(2)
            reverse_scrolling(driver)
            try:
                next_page = wait(driver, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, '[aria-label="Laman berikutnya"]'))
                )
                next_page.click()
            except:
                print("Tidak dapat berpindah ke halaman berikutnya.")
                break
        except:
            print("Terjadi kesalahan saat mencoba pindah halaman.")
            break

    driver.quit()

    df = pd.DataFrame(product_data)
    now = datetime.datetime.today().strftime('%d-%m-%Y')
    df.to_csv(f'Lazada_Moringa_{now}.csv', index=False)
    df.to_excel(f'Lazada_Moringa_{now}.xlsx', index=False)
    print(f"\nScraping selesai. File disimpan sebagai Lazada_{now}.csv dan Lazada_{now}.xlsx")

if __name__ == "__main__":
    main()
