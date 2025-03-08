import time
import datetime
import pandas as pd
from tqdm import tqdm

# Selenium Imports
from selenium import webdriver as wb
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def scrolling(driver):
    """
    Fungsi untuk melakukan scroll pada halaman web secara bertahap.
    Ditambahkan time.sleep(0.5) agar elemen sempat dimuat.
    """
    scheight = 0.1
    while scheight < 9.9:
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight / {});".format(scheight)
        )
        time.sleep(0.3)  # jeda 0.5 detik (sesuaikan sesuai kebutuhan Anda)
        scheight += 0.06

def reverse_scrolling(driver):
    """
    Fungsi untuk melakukan scroll ke bawah agar memicu pemuatan data baru.
    Ditambahkan time.sleep(0.5) di setiap loop agar situs sempat memuat konten.
    """
    body = driver.find_element(By.TAG_NAME, 'body')
    for _ in range(25):
        body.send_keys(Keys.PAGE_DOWN)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)

def extract_data(driver, product_data):
    """
    Fungsi utama untuk mengekstrak data dari setiap produk di halaman saat ini.
    Data yang diambil meliputi:
    - Nama produk
    - Harga (asli/diskon)
    - Nama toko
    - Jumlah produk terjual (jika ada)
    - Link detail produk
    """

    # Mekanisme retry untuk memastikan data benar-benar dimuat
    max_retries = 3
    for attempt in range(max_retries):
        try:
            # Scroll perlahan agar elemen dimuat
            scrolling(driver)

            # Tunggu hingga elemen card produk terlihat di halaman
            data_items = wait(driver, 20).until(
                EC.visibility_of_all_elements_located((By.XPATH, '//div[contains(@class, "css-5wh65g")]'))
            )

            # Jika berhasil mendapatkan data_items tanpa timeout, break dari retry-loop
            break
        except TimeoutException:
            # Jika masih percobaan terakhir, raise error
            if attempt == max_retries - 1:
                raise
            # Jika bukan percobaan terakhir, coba refresh & ulangi
            driver.refresh()
            time.sleep(3)  # beri jeda 3 detik sebelum scroll berikutnya

    # Loop untuk ekstraksi data produk yang ditemukan
    for item in tqdm(data_items, desc="Memproses produk"):

        # Ambil nama produk
        try:
            name = item.find_element(
                By.XPATH, './/span[@class="_0T8-iGxMpV6NEsYEhwkqEg=="]'
            ).text
        except:
            continue  # lompat ke item berikutnya

        # Ambil container harga
        try:
            # Mencari container harga
            price_container = item.find_element(
                By.XPATH, './/div[contains(@class, "XvaCkHiisn2EZFq0THwVug==")]'
            )
            
            # Mencari elemen-elemen terkait harga (bisa jadi satu atau beberapa)
            price_elements = price_container.find_elements(
                By.XPATH, './/div[contains(@class, "_67d6E1xDKIzw+i2D2L0tjw==")]'
            )

            # Kalau sama sekali tidak ada price_elements, skip item
            if not price_elements:
                continue

            # Jika hanya ada satu elemen, artinya tidak ada diskon
            if len(price_elements) == 1:
                price = price_elements[0].text
                original_price = price
                discount_price = None
            else:
                # Ada lebih dari satu elemen (harga diskon + harga asli)
                discount_price = None
                original_price = None
                for p_el in price_elements:
                    class_attr = p_el.get_attribute("class")
                    if "t4jWW3NandT5hvCFAiotYg==" in class_attr:
                        discount_price = p_el.text  # harga diskon
                    else:
                        original_price = p_el.text  # harga asli
                
                # Pilih harga diskon jika tersedia, jika tidak ya harga asli
                price = discount_price if discount_price else original_price

        except:
            # Jika container harga atau element harga tidak ditemukan
            # langsung skip item
            continue


        # Ambil nama toko dan daerah
        try:
            # Elemen store name (sebelum hover)
            store_element = wait(item, 5).until(
                EC.presence_of_element_located((By.XPATH, './/span[contains(@class,"T0rpy-LEwYNQifsgB-3SQw==")]'))
            )
            store = store_element.text
        except:
            continue

        # Lakukan hover agar teks “flip” muncul
        try:
            actions = ActionChains(driver)
            actions.move_to_element(store_element).perform()
            time.sleep(1)  # jeda singkat agar animasi flip muncul
        except:
            continue

        # Setelah hover, lokasi sering muncul di elemen berbeda atau teks sama berganti
        # Misalnya di screenshot, lokasi ada di <span class="pC8DMVkBZGW7-egObcWMFQ== flip">
        try:
            location_element = item.find_element(By.XPATH, './/span[@class="pC8DMVkBZGW7-egObcWMFQ== flip"]')
            location = location_element.text
        except:
            location = None


        # Ambil jumlah terjual (jika ada)
        try:
            sold = item.find_element(
                By.XPATH, './/span[@class="se8WAnkjbVXZNA8mT+Veuw=="]'
            ).text
        except:
            sold = None

        # Ambil link produk
        link_element = item.find_element(By.XPATH, './/a')
        details_link = link_element.get_attribute('href')

        # Simpan data ke dalam dictionary
        data = {
            'name': name,
            'price': price,
            'store': store,
            'location' : location,
            'sold': sold,
            'details_link': details_link
        }
        product_data.append(data)

def main():
    """
    Fungsi utama untuk menjalankan proses scraping:
    1. Membuka browser dan masuk ke halaman Tokopedia
    2. Memasukkan kata kunci pencarian & jumlah halaman
    3. Memulai proses ekstraksi data pada setiap halaman
    4. Menyimpan hasil ke CSV dan Excel
    """
    # Inisialisasi WebDriver & buka halaman Tokopedia
    driver = wb.Chrome()
    driver.get('https://www.tokopedia.com/')
    driver.implicitly_wait(5)

    # Meminta input dari pengguna
    keywords = input("Keywords: ")
    pages = int(input("Pages: "))

    # Inisialisasi pencarian
    search = driver.find_element(
        By.XPATH,
        '//*[@id="header-main-wrapper"]/div[2]/div[2]/div/div/div/div/input'
    )
    search.send_keys(keywords)
    search.send_keys(Keys.ENTER)

    driver.implicitly_wait(5)

    # List untuk menampung data produk
    product_data = []

    # Proses scraping per halaman
    for page in range(1, pages + 1):
        print(f"\n--- Halaman {page} ---")
        # Ekstraksi data di halaman saat ini
        extract_data(driver, product_data)

        # Klik tombol halaman berikutnya
        try:
            next_page = wait(driver, 20).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, '[aria-label="Laman berikutnya"]')
                )
            )
            next_page.click()
        except TimeoutException:
            # Jika tombol tidak ditemukan, coba refresh & scroll sekali lagi
            print("Tombol halaman berikutnya tidak ditemukan, mencoba ulang...")
            driver.refresh()
            time.sleep(2)
            reverse_scrolling(driver)
            # Coba temukan lagi
            try:
                next_page = wait(driver, 20).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, '[aria-label="Laman berikutnya"]')
                    )
                )
                next_page.click()
            except:
                # Jika tetap gagal, akhiri loop
                print("Tidak dapat berpindah ke halaman berikutnya.")
                break
        except:
            # Jika ada error lain, akhiri loop
            print("Terjadi kesalahan saat mencoba pindah halaman.")
            break

    # Tutup driver setelah scraping selesai
    driver.quit()

    # Membuat DataFrame dari list product_data
    df = pd.DataFrame(product_data)

    # Format tanggal untuk penamaan file
    now = datetime.datetime.today().strftime('%d-%m-%Y')

    # Ekspor data ke CSV dan Excel
    df.to_csv(f'Tokopedia_{now}.csv', index=False)
    df.to_excel(f'Tokopedia_{now}.xlsx', index=False)

    print(f"\nScraping selesai. File disimpan sebagai Tokopedia_{now}.csv dan Tokopedia_{now}.xlsx")

if __name__ == "__main__":
    main()
