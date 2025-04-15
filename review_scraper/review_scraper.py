import time
import datetime
import pandas as pd
from tqdm import tqdm

from selenium import webdriver as wb
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def scrolling(driver):
    """
    Metode scroll yang membagi tinggi halaman secara bertahap.
    Pendekatan: Menggunakan document.body.scrollHeight dibagi dengan nilai yang meningkat.
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
    Metode scroll tambahan dengan mengirimkan perintah keyboard PAGE_DOWN 
    untuk memastikan halaman bergulir ke bagian bawah.
    """
    body = driver.find_element(By.TAG_NAME, 'body')
    for _ in range(25):
        body.send_keys(Keys.PAGE_DOWN)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)

def full_scrolling(driver):
    """
    Fungsi untuk menggabungkan kedua metode scrolling agar lebih maksimal.
    """
    scrolling(driver)
    reverse_scrolling(driver)

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

def scrape_product_reviews(driver, product_url, timeout=10):
    """
    Melakukan scraping review untuk satu produk berdasarkan URL yang diberikan.
    
    Proses:
      1. Buka URL produk dan tunggu kemunculan elemen review (<article class="css-15m2bcr">).
      2. Lakukan full_scrolling() untuk memastikan semua elemen review termuat.
      3. Ekstrak review dengan selector <p class="css-cvmev1-unf-heading e1qvo2ff8">.
      4. Navigasi ke halaman review berikutnya melalui tombol 
         <button class="css-5p3bh2-unf-pagination-item" aria-label="Laman X">.
    """
    reviews = []
    try:
        driver.get(product_url)
        wait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "article.css-15m2bcr"))
        )
    except TimeoutException:
        print(f"Timeout loading reviews for {product_url}")
        return reviews

    page = 1
    while True:
        # Lakukan gabungan metode scroll untuk memuat seluruh elemen review
        full_scrolling(driver)
        time.sleep(1)  # Beri waktu stabilisasi halaman

        review_elements = driver.find_elements(By.CSS_SELECTOR, "article.css-15m2bcr")
        if not review_elements:
            print("Tidak ditemukan review pada halaman ini.")
            break

        for elem in review_elements:
            try:
                text_elem = elem.find_element(By.CSS_SELECTOR, "p.css-cvmev1-unf-heading.e1qvo2ff8")
                review_text = text_elem.text.strip()
                if review_text:
                    reviews.append({
                        'product_url': product_url,
                        'page': page,
                        'review': review_text
                    })
            except NoSuchElementException:
                continue

        # Coba navigasi ke halaman review selanjutnya jika tombol "Laman X" ada
        try:
            next_page = page + 1
            next_button = driver.find_element(
                By.CSS_SELECTOR,
                f"button.css-5p3bh2-unf-pagination-item[aria-label='Laman {next_page}']"
            )
            if next_button:
                next_button.click()
                page += 1
                time.sleep(2)  # Tunggu konten halaman baru termuat
            else:
                break
        except Exception as e:
            print(f"Pagination review selesai untuk {product_url} pada halaman {page}.")
            break
    return reviews

def main():
    # Memuat URL produk dari CSV
    product_links = load_product_links()
    if not product_links:
        print("No product links found. Please check your CSV file.")
        return

    driver = wb.Chrome()
    driver.implicitly_wait(5)
    all_reviews = []

    # Iterasi tiap URL produk dan lakukan scraping review
    for url in tqdm(product_links, desc="Scraping product reviews", unit="product"):
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