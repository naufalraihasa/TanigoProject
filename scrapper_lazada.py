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
        # Scroll ke posisi tertentu berdasarkan perhitungan tinggi halaman
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
            # Lakukan scroll untuk memuat elemen
            scrolling(driver)
            # Tunggu hingga semua elemen produk terlihat di halaman
            data_items = wait(driver, 10).until(
                EC.visibility_of_all_elements_located((By.XPATH, '//div[contains(@class, "Bm3ON")]'))
            )
            break  # Jika berhasil, keluar dari loop retry
        except TimeoutException:
            if attempt == max_retries - 1:
                raise
            driver.refresh()
            time.sleep(3)

    # Proses setiap produk yang ditemukan
    for item in tqdm(data_items, desc="Memproses produk"):
        # Ambil nama produk
        try:
            name = item.find_element(
                By.XPATH,
                './/div[@class="RfADt"]'
            ).text
        except Exception:
            continue
        
        # Ambil harga produk
        try:
            price = item.find_element(
                By.XPATH,
                './/span[@class="ooOxS"]'
            ).text
        except Exception:
            continue
        
        # Ambil lokasi produk
        try:
            location = item.find_element(
                By.XPATH,
                './/span[@class="oa6ri "]'
            ).text
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
            sold = item.find_element(
                By.XPATH, './/span[@class="_1cEkb"]'
            ).text
        except Exception:
            sold = None

        # --- BAGIAN TAMBAHAN: PENGAMBILAN DESKRIPSI, TOKO, DAN RATING PRODUK ---
        description = None
        store = None
        brand = None
        rating = 0.0  # default rating 0

        BASE64_TO_STAR = {
            "ASUVORK5CYII=": 0.0,  # base64 untuk 0 bintang
            "ElFTkSuQmCC":   0.5,  # base64 untuk 0.5 bintang
            "BJRU5ErkJggg==": 1.0  # base64 untuk 1 bintang
        }

        # Simpan handle tab utama agar bisa kembali setelah membuka tab baru
        main_window = driver.current_window_handle

        try:
            # Buka link detail produk di tab baru
            driver.execute_script(f"window.open('{details_link}','_blank');")
            driver.switch_to.window(driver.window_handles[-1])

            # Tunggu hingga elemen detail produk dan nama toko termuat
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="pdp-product-desc "]'))
            )
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="pdp-product-brand"]'))
            )
            wait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'img.star'))
                )
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="seller-name__detail"]'))
            )
            

            # Ambil teks deskripsi produk
            try:
                description = driver.find_element(By.XPATH, '//div[@class="pdp-product-desc "]').text
            except NoSuchElementException:
                pass

            # Ambil teks nama toko
            try:
                store = driver.find_element(By.XPATH, '//div[@class="seller-name__detail"]').text
            except NoSuchElementException:
                pass
            
            # Ambil teks nama brand
            try:
                brand = driver.find_element(By.XPATH, '//div[@class="pdp-product-brand"]').text
            except NoSuchElementException:
                pass
            
            # Ambil rating berbasis <img class="star">
            try:
                star_imgs = driver.find_elements(By.CSS_SELECTOR, 'img.star')
                total_stars = 0.0
                for star_img in star_imgs:
                    src_value = star_img.get_attribute("src")
                    matched = False
                    for base64_pattern, star_value in BASE64_TO_STAR.items():
                        if base64_pattern in src_value:
                            total_stars += star_value
                            matched = True
                            break
                    if not matched:
                        # Jika base64 tidak dikenali, anggap 0 atau skip
                        pass
                rating = total_stars
            except:
                rating = 0.0

        except (TimeoutException, NoSuchElementException):
            # Jika gagal mengambil data detail, biarkan nilainya None atau 0
            pass
        finally:
            # Tutup tab detail dan kembali ke tab utama
            driver.close()
            driver.switch_to.window(main_window)
        # --- SELESAI BAGIAN PENGAMBILAN DESKRIPSI, TOKO, DAN RATING ---

        # Simpan data produk ke dalam list
        data = {
            'name': name,
            'price': price,
            'store': store,
            'brand' : brand,
            'location': location,
            'sold': sold,
            'rating' : rating,
            'details_link': details_link,
            'description': description
        }
        product_data.append(data)

def main():
    """
    Fungsi utama yang menginisiasi proses scraping data dari Lazada.
    Melakukan inisiasi WebDriver, input kata kunci, navigasi antar halaman, dan menyimpan hasil scraping.
    """
    # Inisiasi Chrome WebDriver dan buka situs To Lazada
    driver = wb.Chrome()
    driver.get('https://www.lazada.co.id/')
    driver.implicitly_wait(5)

    # Input pencarian dari user
    keywords = input("Keywords: ")
    pages = int(input("Pages: "))

    # Cari elemen input pencarian dan masukkan kata kunci
    search = driver.find_element(
        By.XPATH, '//input[@class="search-box__input--O34g"]'
    )
    search.send_keys(keywords)
    search.send_keys(Keys.ENTER)


    driver.implicitly_wait(5)

    product_data = []

    # Looping untuk memproses setiap halaman hasil pencarian
    for page in range(1, pages + 1):
        print(f"\n--- Halaman {page} ---")
        extract_data(driver, product_data)

        try:
            next_page = wait(driver, 20).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, '[aria-label="right"]')
                )
            )
            next_page.click()
        except TimeoutException:
            print("Tombol halaman berikutnya tidak ditemukan, mencoba ulang...")
            driver.refresh()
            time.sleep(2)
            reverse_scrolling(driver)
            try:
                next_page = wait(driver, 20).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, '[aria-label="Laman berikutnya"]')
                    )
                )
                next_page.click()
            except:
                print("Tidak dapat berpindah ke halaman berikutnya.")
                break
        except:
            print("Terjadi kesalahan saat mencoba pindah halaman.")
            break


    # Tutup driver setelah proses scraping selesai
    driver.quit()

    # Simpan data yang telah di-scrape ke dalam file CSV dan Excel
    df = pd.DataFrame(product_data)
    now = datetime.datetime.today().strftime('%d-%m-%Y')
    df.to_csv(f'Lazada_Moringa_{now}.csv', index=False)
    df.to_excel(f'Lazada_Moringa_{now}.xlsx', index=False)
    print(f"\nScraping selesai. File disimpan sebagai Lazada_{now}.csv dan Lazada_{now}.xlsx")

if __name__ == "__main__":
    main()