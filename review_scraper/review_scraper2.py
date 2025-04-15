import time
import datetime
import pandas as pd
from tqdm import tqdm

from selenium import webdriver as wb
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def scrolling(driver):
    """
    Lakukan scroll secara bertahap untuk memicu lazy-loading.
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
    Lakukan reverse scrolling dengan mengirim perintah PAGE_DOWN agar memicu loading elemen secara natural.
    """
    body = driver.find_element(By.TAG_NAME, 'body')
    for _ in range(25):
        body.send_keys(Keys.PAGE_DOWN)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)

def load_product_links(csv_file='products.csv'):
    """
    Memuat daftar URL produk dari file CSV.
    Pastikan file CSV memiliki kolom 'url'.
    """
    try:
        df = pd.read_csv(csv_file)
        return df['url'].dropna().tolist()
    except Exception as e:
        print("Error reading product URLs:", e)
        return []

def scrape_product_reviews(driver, product_url):
    """
    Melakukan scraping review untuk satu produk berdasarkan URL yang diberikan.
    
    Proses:
      1. Buka URL produk.
      2. Ambil nama produk sekali di awal (jika elemen tersedia).
      3. Lakukan scroll dan reverse scrolling untuk memicu lazy-loading review.
      4. Ekstrak informasi review: rating, waktu review, dan teks review.
      5. Navigasi ke halaman review selanjutnya menggunakan tombol paginasi.
    """
    reviews = []
    driver.get(product_url)
    
    # 1) Ambil nama produk sekali di awal
    try:
        # Tunggu sampai elemen h1 dengan data-testid lblPDPDetailProductName muncul
        product_name_elem = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'h1.css-j63za0[data-testid="lblPDPDetailProductName"]')
            )
        )
        product_name = product_name_elem.text
    except (TimeoutException, NoSuchElementException):
        product_name = None

    # 2) Lakukan scroll agar review diload
    scrolling(driver)
    reverse_scrolling(driver)
    time.sleep(1)  # Delay untuk memastikan halaman sudah stabil

    page = 1
    while True:
        # 3) Ambil seluruh elemen review di halaman ini
        review_elements = driver.find_elements(By.CSS_SELECTOR, "article.css-15m2bcr")
        if not review_elements:
            print(f"Tidak ada review ditemukan pada {product_url} di halaman {page}.")
            break

        # 4) Looping setiap review dan ambil data
        for elem in review_elements:
            rating_value = None
            review_time = None
            review_text = None

            # Ambil rating (misal: "bintang 5")
            try:
                rating_elem = elem.find_element(By.CSS_SELECTOR, 'div[data-testid="icnStarRating"]')
                rating_aria = rating_elem.get_attribute("aria-label")
                if rating_aria and "bintang" in rating_aria:
                    # Contoh: "bintang 5" -> ambil angka "5"
                    parts = rating_aria.split()
                    if len(parts) > 1:
                        rating_value = parts[1]
            except NoSuchElementException:
                pass

            # Ambil waktu review (misal: "6 bulan lalu")
            try:
                time_elem = elem.find_element(By.CSS_SELECTOR, "p.css-vqrjg4-unf-heading.e1qvo2ff8")
                review_time = time_elem.text.strip()
            except NoSuchElementException:
                pass

            # Ambil teks review
            try:
                text_elem = elem.find_element(By.CSS_SELECTOR, "p.css-cvmev1-unf-heading.e1qvo2ff8")
                review_text = text_elem.text.strip()
            except NoSuchElementException:
                pass

            # Simpan data review
            if any([rating_value, review_time, review_text]):
                reviews.append({
                    'product_name': product_name,  # pakai nama produk yang diambil di awal
                    'page': page,
                    'rating': rating_value,
                    'review_time': review_time,
                    'review_text': review_text,
                    'product_url': product_url
                })

        # 5) Pagination untuk halaman review selanjutnya
        next_page = page + 1
        xpath_next = f"//*[contains(@class, 'unf-pagination-item')]//button[contains(@aria-label, 'Laman {next_page}')]"

        try:
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath_next))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
            time.sleep(0.3)
            driver.execute_script("arguments[0].click();", next_button)
            time.sleep(0.5)
            scrolling(driver)
            reverse_scrolling(driver)
            time.sleep(1)
            page += 1
        except Exception as e:
            print(f"Pagination selesai/terhenti untuk {product_url} di halaman {page}. Exception: {e}")
            break

    return reviews

def main():
    # Muat URL produk dari file CSV
    product_links = load_product_links()
    if not product_links:
        print("Tidak ada link produk ditemukan. Periksa file CSV Anda.")
        return

    driver = wb.Chrome()  # Pastikan chromedriver sudah tersedia dan sesuai PATH
    driver.implicitly_wait(5)
    all_reviews = []

    for url in tqdm(product_links, desc="Scraping review produk", unit="produk"):
        product_reviews = scrape_product_reviews(driver, url)
        all_reviews.extend(product_reviews)

    driver.quit()

    # Simpan hasil scraping ke file CSV dan Excel
    df = pd.DataFrame(all_reviews)
    now = datetime.datetime.today().strftime('%d-%m-%Y')
    csv_filename = f'Product_Reviews_{now}.csv'
    excel_filename = f'Product_Reviews_{now}.xlsx'
    df.to_csv(csv_filename, index=False)
    df.to_excel(excel_filename, index=False)
    print(f"Scraping selesai. File disimpan sebagai {csv_filename} dan {excel_filename}.")

if __name__ == "__main__":
    main()
