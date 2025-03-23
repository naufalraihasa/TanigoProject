import time
import json
import re
import urllib3
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from tqdm import tqdm

from selenium import webdriver as wb
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Nonaktifkan peringatan urllib3
urllib3.disable_warnings()

# ---------------------------#
#        UTILITY FUNCTIONS   #
# ---------------------------#

def load_keyword(json_file_path):
    """
    Memuat keyword pencarian dari file konfigurasi JSON.
    """
    try:
        with open(json_file_path, 'r') as file:
            data = json.load(file)
            return data.get('keyword')
    except Exception as e:
        print("Error loading keyword from config:", e)
        return None

def load_shop_urls(csv_file='tokopedia_shops.csv'):
    """
    Memuat URL toko dari file CSV.
    """
    try:
        df = pd.read_csv(csv_file)
        return df['url'].dropna().tolist()
    except Exception as e:
        print("Error reading shop URLs:", e)
        return []

def get_shop_name(url):
    """
    Mengekstrak nama toko dari URL.
    """
    parts = url.split('/')
    return parts[3].split('?')[0]

def setup_driver():
    """
    Menginisialisasi Selenium Chrome WebDriver dengan waktu tunggu implisit.
    """
    driver = wb.Chrome()
    driver.implicitly_wait(5)
    return driver

# ---------------------------#
#        SCROLLING           #
# ---------------------------#

def dynamic_scroll(driver, pause_time=1.0, max_iter=20):
    """
    Melakukan scrolling secara dinamis untuk memuat seluruh produk.
    Proses:
      1. Scroll ke bawah penuh.
      2. Scroll ke tengah halaman.
      3. Scroll ke bawah lagi.
    Proses ini diulang hingga tidak ditemukan produk baru.
    """
    last_count = 0
    for _ in range(max_iter):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause_time)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.5);")
        time.sleep(pause_time / 2)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause_time)
        
        product_elements = driver.find_elements(By.CSS_SELECTOR, "div.css-1sn1xa2")
        current_count = len(product_elements)
        if current_count == last_count:
            break
        last_count = current_count
    return last_count

# ---------------------------#
#   PRODUCT DETAIL SCRAPING  #
# ---------------------------#

def get_product_description(driver, product_url):
    """
    Membuka halaman detail produk di tab baru dan mengambil deskripsi produk.
    Mengembalikan teks deskripsi atau None jika tidak ditemukan.
    """
    description = None
    main_window = driver.current_window_handle
    try:
        driver.execute_script(f"window.open('{product_url}', '_blank');")
        driver.switch_to.window(driver.window_handles[-1])
        # Tunggu hingga elemen detail produk termuat
        wait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.css-1wa8o67'))
        )
        try:
            desc_element = driver.find_element(By.CSS_SELECTOR, 'div.css-1wa8o67 span.css-11oczh8.eytdjj00')
            description = desc_element.text
        except (NoSuchElementException, TimeoutException):
            description = None
    except (TimeoutException, NoSuchElementException):
        description = None
    finally:
        driver.close()
        driver.switch_to.window(main_window)
    return description

# ---------------------------#
#   PARSING & SCRAPING       #
# ---------------------------#

def parse_page_source(driver, html, shop_name):
    """
    Memparsing halaman list produk dan mengambil data dari masing-masing produk,
    termasuk mengambil deskripsi dari halaman detail produk.
    Dilengkapi dengan progress bar untuk setiap produk di toko tersebut.
    """
    soup = BeautifulSoup(html, 'html.parser')
    products = []
    
    product_containers = soup.find_all('div', class_='css-1sn1xa2')
    # Progress bar untuk setiap produk pada toko
    for container in tqdm(product_containers, desc=f"---> Scrapping {shop_name} products", leave=False, unit="produk"):
        try:
            # Ambil nama produk dan harga
            name = container.find('div', {'class': 'prd_link-product-name'}).text.strip()
            price_text = container.find('div', {'class': 'prd_link-product-price'}).text.strip()
            price = int(re.sub(r'[^\d]', '', price_text))
            
            # Ambil rating jika tersedia
            rating_elem = container.find('span', {'class': 'prd_rating-average-text'})
            rating = float(rating_elem.text.strip()) if rating_elem else None
            
            # Ambil teks sales langsung (tanpa preprocessing) sesuai permintaan
            sales = container.find('span', {'class': 'prd_label-integrity'}).text
            
            # Ambil URL produk dan URL gambar
            url_elem = container.find('a', {'class': 'pcv3__info-content'})
            product_url = url_elem['href'] if url_elem else None
            img_elem = container.find('img', {'class': 'css-1q90pod'})
            image_url = img_elem['src'] if img_elem else None
            
            # Ambil deskripsi produk dengan mengunjungi halaman detail
            description = get_product_description(driver, product_url) if product_url else None
            
            product = {
                'shop': shop_name,
                'name': name,
                'price': price,
                'rating': rating,
                'sales': sales,
                'url': product_url,
                'image_url': image_url,
                'description': description
            }
            products.append(product)
        except Exception as e:
            print(f"Error processing product in {shop_name}: {e}")
    return products

def scrape_shop(driver, shop_url, query):
    """
    Mengunjungi halaman produk toko, melakukan scrolling, dan mengambil data produk.
    """
    final_url = shop_url + "/product" + query
    driver.get(final_url)
    
    # Lakukan scrolling agar seluruh produk termuat
    dynamic_scroll(driver, pause_time=1.0, max_iter=20)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)
    
    html = driver.page_source
    shop_name = get_shop_name(shop_url)
    products = parse_page_source(driver, html, shop_name)
    print(f"{len(products)} products scraped from {shop_name}")
    return products

# ---------------------------#
#       DATA CLEANING        #
# ---------------------------#

def clean_sales_column(df):
    """
    Membersihkan kolom 'sales' agar tidak perlu preprocessing lagi.
    Proses:
      - Menghapus karakter '+' dan '.' serta kata 'terjual'
      - Mengganti 'rb' dengan '000'
      - Mengkonversi kolom menjadi integer
    """
    df['sales'] = df['sales'].astype(str)\
        .str.replace('+', '', regex=False)\
        .str.replace('.', '', regex=False)\
        .str.replace('terjual', '', regex=False)
    df['sales'] = df['sales'].str.replace('rb', '000', regex=False)
    df['sales'] = pd.to_numeric(df['sales'], errors='coerce').fillna(0).astype(int)
    return df

# ---------------------------#
#           MAIN             #
# ---------------------------#

def main():
    # Memuat konfigurasi dan URL toko
    keyword = load_keyword("config.json")
    shop_urls = load_shop_urls()
    query = f"?q={keyword}"
    
    driver = setup_driver()
    all_products = []
    
    # Progress bar untuk setiap toko
    for shop_url in tqdm(shop_urls, desc="Scraping shops", unit="toko"):
        products = scrape_shop(driver, shop_url, query)
        all_products.extend(products)
    
    driver.quit()
    
    # Membuat DataFrame dan membersihkan kolom 'sales'
    df = pd.DataFrame(all_products)
    df = clean_sales_column(df)
    
    # Menyimpan data ke file Excel (.xlsx)
    file_name = f"tokopedia_products_{keyword}_{datetime.now().strftime('%Y-%m-%d_%H.%M.%S')}.xlsx"
    df.to_excel(file_name, index=False)
    
    # Menampilkan ringkasan hasil scraping
    if df.empty:
        print("No products found.")
    else:
        print(f"Total products found: {len(df)}")
        print(f"Average price: Rp{df['price'].mean():,.2f}")
        print(f"Price range: Rp{df['price'].min():,} - Rp{df['price'].max():,}")
        print(f"Total sales: {df['sales'].sum()}")
        print(f"Average rating: {df['rating'].mean():.2f}")

if __name__ == "__main__":
    main()
