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
                EC.visibility_of_all_elements_located((By.XPATH, '//div[contains(@class, "css-5wh65g")]'))
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
                By.XPATH, './/span[@class="_0T8-iGxMpV6NEsYEhwkqEg=="]'
            ).text
        except Exception:
            continue

        # Ambil informasi harga produk
        try:
            price_container = item.find_element(
                By.XPATH, './/div[contains(@class, "XvaCkHiisn2EZFq0THwVug==")]'
            )
            price_elements = price_container.find_elements(
                By.XPATH, './/div[contains(@class, "_67d6E1xDKIzw+i2D2L0tjw==")]'
            )

            if not price_elements:
                continue

            # Jika hanya ada satu elemen, maka dianggap sebagai harga
            if len(price_elements) == 1:
                price = price_elements[0].text
                original_price = price
                discount_price = None
            else:
                # Jika ada lebih dari satu elemen, identifikasi harga asli dan harga diskon
                discount_price = None
                original_price = None
                for p_el in price_elements:
                    class_attr = p_el.get_attribute("class")
                    if "t4jWW3NandT5hvCFAiotYg==" in class_attr:
                        discount_price = p_el.text  # Harga diskon
                    else:
                        original_price = p_el.text  # Harga asli
                price = discount_price if discount_price else original_price

        except Exception:
            continue

        # Ambil nama toko dan lakukan hover agar lokasi tampil
        try:
            store_element = wait(item, 5).until(
                EC.presence_of_element_located((By.XPATH, './/span[contains(@class,"T0rpy-LEwYNQifsgB-3SQw==")]'))
            )
            store = store_element.text
        except Exception:
            continue

        try:
            actions = ActionChains(driver)
            actions.move_to_element(store_element).perform()
            time.sleep(1)  # Jeda untuk menampilkan animasi flip (jika ada)
        except Exception:
            continue

        # Ambil lokasi toko
        try:
            location_element = item.find_element(By.XPATH, './/span[@class="pC8DMVkBZGW7-egObcWMFQ== flip"]')
            location = location_element.text
        except Exception:
            continue

        # Ambil jumlah produk yang terjual
        try:
            sold = item.find_element(
                By.XPATH, './/span[@class="se8WAnkjbVXZNA8mT+Veuw=="]'
            ).text
        except Exception:
            sold = None

        # Ambil link detail produk
        try:
            link_element = item.find_element(By.XPATH, './/a')
            details_link = link_element.get_attribute('href')
        except Exception:
            continue

        # --- BAGIAN TAMBAHAN: PENGAMBILAN DESKRIPSI PRODUK ---
        description = None  # Inisialisasi default jika tidak ditemukan

        # Simpan handle tab utama agar bisa kembali setelah membuka tab baru
        main_window = driver.current_window_handle

        try:
            # Buka link detail produk di tab baru
            driver.execute_script(f"window.open('{details_link}','_blank');")
            driver.switch_to.window(driver.window_handles[-1])

            # Tunggu hingga elemen detail produk termuat
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.css-1wa8o67'))
            )

            # Ambil teks deskripsi produk
            desc_element = driver.find_element(By.CSS_SELECTOR, 'div.css-1wa8o67 span.css-11oczh8.eytdjj00')
            description = desc_element.text

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
            'store': store,
            'location': location,
            'sold': sold,
            'details_link': details_link,
            'description': description
        }
        product_data.append(data)

def main():
    """
    Fungsi utama yang menginisiasi proses scraping data dari Tokopedia.
    Melakukan inisiasi WebDriver, input kata kunci, navigasi antar halaman, dan menyimpan hasil scraping.
    """
    # Inisiasi Chrome WebDriver dan buka situs Tokopedia
    driver = wb.Chrome()
    driver.get('https://www.tokopedia.com/')
    driver.implicitly_wait(5)

    # Input pencarian dari user
    keywords = input("Keywords: ")
    pages = int(input("Pages: "))

    # Cari elemen input pencarian dan masukkan kata kunci
    search = driver.find_element(
        By.XPATH,
        '//*[@id="header-main-wrapper"]/div[2]/div[2]/div/div/div/div/input'
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
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, '[aria-label="Laman berikutnya"]')
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
            except Exception:
                print("Tidak dapat berpindah ke halaman berikutnya.")
                break
        except Exception:
            print("Terjadi kesalahan saat mencoba pindah halaman.")
            break

    # Tutup driver setelah proses scraping selesai
    driver.quit()

    # Simpan data yang telah di-scrape ke dalam file CSV dan Excel
    df = pd.DataFrame(product_data)
    now = datetime.datetime.today().strftime('%d-%m-%Y')
    df.to_csv(f'Tokopedia_Moringa_Infusition_{now}.csv', index=False)
    df.to_excel(f'Tokopedia_Moringa_Infusition_{now}.xlsx', index=False)
    print(f"\nScraping selesai. File disimpan sebagai Tokopedia_{now}.csv dan Tokopedia_{now}.xlsx")

if __name__ == "__main__":
    main()
