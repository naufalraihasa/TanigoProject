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

    # Process only the first 5 products for testing
    for index, item in enumerate(tqdm(data_items, desc="Memproses produk")):
        if index >= 10:
            break
        
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
        # rating = 0.0  # default rating 0

        main_window = driver.current_window_handle

        try:
            # Buka link detail produk di tab baru
            driver.execute_script(f"window.open('{details_link}','_blank');")
            driver.switch_to.window(driver.window_handles[-1])

            # Tunggu hingga halaman utama benar-benar termuat
            wait(driver, 15).until(EC.presence_of_element_located(
                (By.TAG_NAME, 'body')))
            time.sleep(2)  # tambahan waktu tunggu ekstra jika perlu

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
            time.sleep(1)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)


            # Tunggu hingga elemen detail produk termuat (sesuaikan selector jika perlu)
            
            # description
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="pdp-product-detail"]'))
            )
            
            # Ambil deskripsi produk
            try:
                description = driver.find_element(By.XPATH, '//div[@class="pdp-product-detail"]').text
            except NoSuchElementException:
                pass
            
            # brand
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="pdp-product-brand"]'))
            )
            
            # Brand
            try:
                brand = wait(driver, 10).until(
                    EC.presence_of_element_located((
                        By.XPATH, '//div[contains(@class, "pdp-product-brand")]//a[contains(@class, "pdp-product-brand__brand-link")]'
                    ))
                ).text
            except (TimeoutException, NoSuchElementException):
                brand = None

            # Store
            try:
                store = wait(driver, 10).until(
                    EC.presence_of_element_located((
                        By.XPATH, '//div[contains(@class, "seller-name__detail")]//a[contains(@class, "seller-name__detail-name")]'
                    ))
                ).text
            except (TimeoutException, NoSuchElementException):
                store = None

            # Mapping baru sesuai gambar terbaru
            BASE64_TO_STAR = {
                "ASUVORK5CYII=": 0.0,   # 0-star
                "ElFTkSuQmCC":   0.5,   # 0.5-star
                "BJRU5ErkJggg==": 1.0   # 1-star
            }


            rating_elements = driver.find_elements(By.CSS_SELECTOR, 'div.pdp-review-summary img.star')
            total_stars = 0.0
            for img in rating_elements:
                img_src = img.get_attribute("src")
                for base64_key, star_value in BASE64_TO_STAR.items():
                    if base64_key in img_src:
                        total_stars += star_value
                        break
            rating = total_stars



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
