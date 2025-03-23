import time
import random
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

def load_shop_urls(csv_file='shops.csv'):
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
    Menginisialisasi Selenium Chrome WebDriver dengan waktu tunggu implisit dan
    User-Agent yang di-random.
    """
    USER_AGENTS = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.1 Safari/605.1.15',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/119.0',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0'
    ]
    
    options = wb.ChromeOptions()
    options.add_argument(f'user-agent={random.choice(USER_AGENTS)}')
    driver = wb.Chrome(options=options)
    driver.implicitly_wait(5)
    return driver

def dynamic_scroll(driver, pause_time=1.0, max_iter=20):
    """
    Melakukan scrolling dinamis untuk memuat seluruh produk.
    Proses:
      1. Scroll ke bawah penuh.
      2. Scroll ke tengah halaman.
      3. Scroll ke bawah lagi.
    Diulang hingga tidak ada penambahan jumlah produk.
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

def get_product_description(driver, product_url, timeout=5):
    """
    Membuka halaman detail produk di tab baru dan mengambil deskripsi produk.
    Jika halaman tidak termuat dalam waktu 'timeout' detik, maka mengembalikan
    nilai sentinel "TIMEOUT" untuk menandakan bahwa produk ini harus di-skip.
    """
    description = None
    main_window = driver.current_window_handle
    timed_out = False
    try:
        driver.execute_script(f"window.open('{product_url}', '_blank');")
        driver.switch_to.window(driver.window_handles[-1])
        try:
            # Gunakan timeout lebih pendek untuk menghindari RTO
            wait(driver, timeout).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.css-1wa8o67'))
            )
        except TimeoutException:
            timed_out = True

        if not timed_out:
            try:
                desc_element = driver.find_element(
                    By.CSS_SELECTOR, 'div.css-1wa8o67 span.css-11oczh8.eytdjj00'
                )
                description = desc_element.text
            except (NoSuchElementException, TimeoutException):
                description = None
        else:
            description = "TIMEOUT"
    except (TimeoutException, NoSuchElementException):
        description = "TIMEOUT"
    finally:
        driver.close()
        driver.switch_to.window(main_window)
    return description

def parse_page_source(driver, html, shop_name):
    """
    Memparsing halaman list produk dan mengambil data dari masing-masing produk,
    termasuk mengambil deskripsi dari halaman detail produk.
    Produk yang mengalami timeout saat memuat halaman detail akan di-skip.
    Dilengkapi dengan progress bar untuk setiap produk di toko tersebut.
    """
    soup = BeautifulSoup(html, 'html.parser')
    products = []
    
    product_containers = soup.find_all('div', class_='css-1sn1xa2')
    for container in tqdm(
        product_containers, 
        desc=f"---> Scrapping {shop_name} products", 
        leave=False, 
        unit="produk"
    ):
        try:
            # Mengambil nama produk
            name_elem = container.find('div', {'class': 'prd_link-product-name'})
            name = name_elem.text.strip() if name_elem else ""

            # Mengambil harga produk dan mengubahnya menjadi integer
            price_elem = container.find('div', {'class': 'prd_link-product-price'})
            if price_elem:
                price_text = price_elem.text.strip()
                price = int(re.sub(r'[^\d]', '', price_text))
            else:
                price = 0
                
            # Mengambil rating produk jika ada
            rating_elem = container.find('span', {'class': 'prd_rating-average-text'})
            rating = float(rating_elem.text.strip()) if rating_elem else float(0)

            # Mengambil informasi sales
            sales_elem = container.find('span', {'class': 'prd_label-integrity'})
            sales = sales_elem.text if sales_elem else None

            # Mengambil URL produk
            url_elem = container.find('a', {'class': 'pcv3__info-content'})
            product_url = url_elem['href'] if url_elem and url_elem.has_attr('href') else None

            # Mengambil URL gambar produk
            img_elem = container.find('img', {'class': 'css-1q90pod'})
            image_url = img_elem['src'] if img_elem and img_elem.has_attr('src') else None

            
            # Jika ada URL produk, ambil deskripsi dengan timeout
            if product_url:
                description = get_product_description(driver, product_url, timeout=5)
                # Jika terjadi timeout, skip produk ini
                if description == "TIMEOUT":
                    print(f"Skipping product from {shop_name} due to load timeout.")
                    continue
            else:
                description = None
            
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
    Mengunjungi halaman produk suatu toko, melakukan scrolling, dan mengambil data produk.
    """
    final_url = shop_url + "/product" + query
    driver.get(final_url)
    
    # Lakukan scrolling untuk memastikan seluruh produk termuat
    dynamic_scroll(driver, pause_time=1.0, max_iter=20)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)
    
    html = driver.page_source
    shop_name = get_shop_name(shop_url)
    products = parse_page_source(driver, html, shop_name)
    print(f"{len(products)} products scraped from {shop_name}")
    return products

def clean_sales_column(df):
    """
    Membersihkan kolom 'sales' agar siap digunakan untuk analisis.
    Proses:
      - Menghapus karakter '+' dan '.' serta kata 'terjual'
      - Mengganti 'rb' dengan '000'
      - Mengkonversi nilai ke integer
    """
    df['sales'] = df['sales'].astype(str)\
        .str.replace('+', '', regex=False)\
        .str.replace('.', '', regex=False)\
        .str.replace('terjual', '', regex=False)
    df['sales'] = df['sales'].str.replace('rb', '000', regex=False)
    df['sales'] = pd.to_numeric(df['sales'], errors='coerce').fillna(0).astype(int)
    return df

def main():
    # Memuat keyword dan URL toko
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
    
    # Ringkasan hasil scraping
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