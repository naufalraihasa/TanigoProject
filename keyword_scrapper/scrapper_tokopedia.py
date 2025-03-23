import time
import datetime
import pandas as pd
from tqdm import tqdm

from selenium import webdriver as wb
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def scrolling(driver):
    scheight = 0.1
    while scheight < 9.9:
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight / {});".format(scheight)
        )
        time.sleep(0.3)
        scheight += 0.1

def reverse_scrolling(driver):
    body = driver.find_element(By.TAG_NAME, 'body')
    for _ in range(25):
        body.send_keys(Keys.PAGE_DOWN)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)

def extract_data(driver, product_data):
    max_retries = 3
    for attempt in range(max_retries):
        try:
            scrolling(driver)
            data_items = wait(driver, 10).until(
                EC.visibility_of_all_elements_located(
                    (By.XPATH, '//div[contains(@class, "css-5wh65g")]')
                )
            )
            break
        except TimeoutException:
            if attempt == max_retries - 1:
                raise
            driver.refresh()
            time.sleep(3)

    for item in tqdm(data_items, desc="Memproses produk"):
        try:
            name = item.find_element(
                By.XPATH, './/span[@class="_0T8-iGxMpV6NEsYEhwkqEg=="]'
            ).text
        except:
            continue
        
        try:
            # Coba ambil harga diskon
            discounted_elem = item.find_element(
                By.XPATH,
                './/div[contains(@class, "_67d6E1xDKIzw+i2D2L0tjw== ") and contains(@class, "t4jWW3NandT5hvCFAiotYg==")]'
            )
            discounted_price = discounted_elem.text

            # Ambil harga normal dari elemen <span>
            original_elem = item.find_element(
                By.XPATH,
                './/span[@class="q6wH9+Ht7LxnxrEgD22BCQ=="]'
            )
            original_price = original_elem.text

            # Gunakan harga diskon sebagai harga utama
            price = discounted_price

        except NoSuchElementException:
            # Jika tidak ada diskon, ambil harga normal dari <div>
            normal_elem = item.find_element(
                By.XPATH,
                './/div[@class="_67d6E1xDKIzw+i2D2L0tjw== "]'
            )
            price = normal_elem.text
            original_price = price
            discounted_price = None
            
        try:
            discount = item.find_element(
                By.XPATH, './/span[@class="vRrrC5GSv6FRRkbCqM7QcQ=="]'
            ).text
        except:
            discount = None

        try:
            store_element = wait(item, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH, './/span[contains(@class,"T0rpy-LEwYNQifsgB-3SQw==")]')
                )
            )
            store = store_element.text
        except:
            continue

        try:
            actions = ActionChains(driver)
            actions.move_to_element(store_element).perform()
            time.sleep(1)
        except:
            continue

        try:
            location_element = item.find_element(
                By.XPATH, './/span[@class="pC8DMVkBZGW7-egObcWMFQ== flip"]'
            )
            location = location_element.text
        except:
            continue

        try:
            sold = item.find_element(
                By.XPATH, './/span[@class="se8WAnkjbVXZNA8mT+Veuw=="]'
            ).text
        except:
            sold = None

        try:
            link_element = item.find_element(By.XPATH, './/a')
            details_link = link_element.get_attribute('href')
        except:
            continue

        # --- BAGIAN BUKA DETAIL PRODUK ---
        description = None
        rating = "0"  # Default rating = 0 jika tidak ditemukan
        main_window = driver.current_window_handle

        try:
            driver.execute_script(f"window.open('{details_link}','_blank');")
            driver.switch_to.window(driver.window_handles[-1])

            # Tunggu elemen detail
            wait(driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.css-1wa8o67'))
            )

            # Ambil deskripsi
            try:
                desc_element = driver.find_element(
                    By.CSS_SELECTOR,
                    'div.css-1wa8o67 span.css-11oczh8.eytdjj00'
                )
                description = desc_element.text
            except (NoSuchElementException, TimeoutException):
                pass

            # Ambil rating
            try:
                rating_element = driver.find_element(
                    By.XPATH,
                    '//span[@class="main" and @data-testid="lblPDPDetailProductRatingNumber"]'
                )
                rating = rating_element.text
            except NoSuchElementException:
                rating = "0"

        except (TimeoutException, NoSuchElementException):
            # Gagal memuat halaman detail atau elemen rating
            pass
        finally:
            driver.close()
            driver.switch_to.window(main_window)
        # --- SELESAI BAGIAN DETAIL ---

        data = {
            'name': name,
            'original_price': original_price,
            'discounted_price' : discounted_price,
            'discount' : discount,
            'store': store,
            'location': location,
            'sold': sold,
            'details_link': details_link,
            'description': description,
            'rating': rating
        }
        product_data.append(data)

def main():
    driver = wb.Chrome()
    driver.get('https://www.tokopedia.com/')
    driver.implicitly_wait(5)

    keywords = input("Keywords: ")
    pages = int(input("Pages: "))

    search = driver.find_element(
        By.XPATH,
        '//*[@id="header-main-wrapper"]/div[2]/div[2]/div/div/div/div/input'
    )
    search.send_keys(keywords)
    search.send_keys(Keys.ENTER)
    driver.implicitly_wait(5)

    product_data = []

    for page in range(1, pages + 1):
        print(f"\n--- Halaman {page} ---")
        extract_data(driver, product_data)

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
            except:
                print("Tidak dapat berpindah ke halaman berikutnya.")
                break
        except:
            print("Terjadi kesalahan saat mencoba pindah halaman.")
            break

    driver.quit()

    df = pd.DataFrame(product_data)
    now = datetime.datetime.today().strftime('%d-%m-%Y')
    df.to_csv(f'Tokopedia_Moringa_{now}.csv', index=False)
    df.to_excel(f'Tokopedia_Moringa_{now}.xlsx', index=False)
    print(f"\nScraping selesai. File disimpan sebagai Tokopedia_Moringa_{now}.csv dan Tokopedia_Moringa_{now}.xlsx")

if __name__ == "__main__":
    main()
