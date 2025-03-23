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
                EC.visibility_of_all_elements_located((By.XPATH, '//li[contains(@class, "col-xs-2-4 shopee-search-item-result__item")]'))
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
                './/div[@class="line-clamp-2 break-words min-w-0 min-h-[2.5rem] text-sm"]'
            ).text
        except Exception:
            continue
        
        # Ambil harga produk
        try:
            price = item.find_element(
                By.XPATH,
                './/span[@class="font-medium text-base/5 truncate"]'
            ).text
        except Exception:
            continue
        
        # Ambil diskon produk
        try:
            discount = item.find_element(
                By.XPATH,
                './/div[@class="text-shopee-primary font-medium bg-shopee-pink py-0.5 px-1 text-sp10/3 h-4 rounded-[2px] shrink-0 mr-1"]'
            ).text
        except Exception:
            continue
        
        # Ambil lokasi produk
        try:
            location = item.find_element(
                By.XPATH,
                './/span[@class="ml-[3px] align-middle"]'
            ).text
        except Exception:
            continue
        
        # Ambil rating produk
        try:
            rating = item.find_element(
                By.XPATH,
                './/div[@class="text-shopee-black87 text-xs/sp14 flex-none"]'
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
                By.XPATH, './/span[@class="se8WAnkjbVXZNA8mT+Veuw=="]'
            ).text
        except Exception:
            sold = None

        # --- BAGIAN TAMBAHAN: PENGAMBILAN DESKRIPSI PRODUK ---
        description = None  # Inisialisasi default jika tidak ditemukan
        store = None

        # Simpan handle tab utama agar bisa kembali setelah membuka tab baru
        main_window = driver.current_window_handle

        try:
            # Buka link detail produk di tab baru
            driver.execute_script(f"window.open('{details_link}','_blank');")
            driver.switch_to.window(driver.window_handles[-1])

            # Tunggu hingga elemen detail produk dan nama toko termuat
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//p[@class="QN2lPu"]'))
            )
            
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="FV3T1n"]'))
            )

            # Ambil teks deskripsi produk
            description = driver.find_element(By.XPATH, '//p[@class="QN2lPu"]').text
            store = driver.find_element(By.XPATH, '//div[@class="FV3T1n"]').text

        except (TimeoutException, NoSuchElementException):
            # Jika gagal mengambil deskripsi, tetap biarkan nilainya None
            pass
        finally:
            # Tutup tab detail dan kembali ke tab utama
            driver.close()
            driver.switch_to.window(main_window)
        # --- SELESAI BAGIAN PENGAMBILAN DESKRIPSI ---

        # Simpan data produk ke dalam list
        data = {
            'name': name,
            'price': price,
            'discount' : discount,
            'store': store,
            'location': location,
            'sold': sold,
            'rating' : rating,
            'details_link': details_link,
            'description': description
        }
        product_data.append(data)

def main():
    """
    Fungsi utama yang menginisiasi proses scraping data dari Shopee.
    Melakukan inisiasi WebDriver, input kata kunci, navigasi antar halaman, dan menyimpan hasil scraping.
    """
    # Inisiasi Chrome WebDriver dan buka situs To Shopee
    driver = wb.Chrome()
    driver.get('https://shopee.co.id/')
    driver.implicitly_wait(5)

    # Input pencarian dari user
    keywords = input("Keywords: ")
    pages = int(input("Pages: "))

    # Cari elemen input pencarian dan masukkan kata kunci
    search = driver.find_element(
        By.XPATH, '//input[@class="shopee-searchbar-input__input"]'
    )
    search.send_keys(keywords)
    search.send_keys(Keys.ENTER)


    driver.implicitly_wait(5)

    product_data = []

    # Looping untuk memproses setiap halaman hasil pencarian
    for page in range(1, pages + 1):
        print(f"\n--- Halaman {page} ---")
        extract_data(driver, product_data)
        
        # Navigasi ke halaman berikutnya
        try:
            next_page = wait(driver, 20).until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    'nav.shopee-page-controller a.shopee-icon-button.shopee-icon-button--right[aria-disabled="false"]'
                ))
            )
            next_page.click()
        except TimeoutException:
            print("Tombol halaman berikutnya tidak ditemukan atau sedang disabled.")
            break
        except Exception as e:
            print(f"Terjadi kesalahan saat klik halaman berikutnya: {e}")
            break


    # Tutup driver setelah proses scraping selesai
    driver.quit()

    # Simpan data yang telah di-scrape ke dalam file CSV dan Excel
    df = pd.DataFrame(product_data)
    now = datetime.datetime.today().strftime('%d-%m-%Y')
    df.to_csv(f'Shopee_Moringa_capsule_{now}.csv', index=False)
    df.to_excel(f'Shopee_Moringa_capsule_{now}.xlsx', index=False)
    print(f"\nScraping selesai. File disimpan sebagai Shopee_{now}.csv dan Shopee_{now}.xlsx")

if __name__ == "__main__":
    main()
