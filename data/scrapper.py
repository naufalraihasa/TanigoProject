import datetime
import pandas as pd
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

def scroll_page(driver):
    """Scroll secara bertahap untuk memuat seluruh konten halaman."""
    for fraction in [i/100 for i in range(1, 100)]:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight*%s);" % fraction)

def reverse_scroll(driver, iterations=25):
    """Reverse scrolling untuk memicu pemuatan konten yang mungkin tertunda."""
    body = driver.find_element(By.TAG_NAME, 'body')
    for _ in range(iterations):
        body.send_keys(Keys.PAGE_DOWN)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

def extract_data(driver):
    """Ekstraksi data produk dari halaman saat ini."""
    product_data = []
    wait = WebDriverWait(driver, 30)
    
    driver.refresh()
    scroll_page(driver)
    
    # Mengambil elemen produk
    items = wait.until(EC.presence_of_all_elements_located(
        (By.XPATH, '//div[contains(@class, "css-5wh65g")]')
    ))
    # Jika jumlah item tidak sesuai, refresh ulang dan scroll kembali
    if len(items) != 80:
        driver.refresh()
        scroll_page(driver)
        items = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, '//div[contains(@class, "css-5wh65g")]')
        ))
    
    for item in tqdm(items, desc="Ekstraksi data"):
        element = WebDriverWait(item, 10).until(
            EC.presence_of_element_located((By.XPATH, './/div[@class="bYD8FcVCFyOBiVyITwDj1Q=="]'))
        )
        name = element.find_element(By.XPATH, './/span[@class="_0T8-iGxMpV6NEsYEhwkqEg=="]').text
        
        # Ekstraksi harga
        price_container = element.find_element(By.XPATH, './/div[contains(@class, "XvaCkHiisn2EZFq0THwVug==")]')
        price_elements = price_container.find_elements(By.XPATH, './/div[contains(@class, "_67d6E1xDKIzw+i2D2L0tjw==")]')
        if len(price_elements) == 1:
            price = price_elements[0].text
        else:
            discount_price, original_price = None, None
            for p in price_elements:
                class_attr = p.get_attribute("class")
                if "t4jWW3NandT5hvCFAiotYg==" in class_attr:
                    discount_price = p.text
                else:
                    original_price = p.text
            price = discount_price if discount_price is not None else original_price
        
        store = element.find_element(By.XPATH, './/span[contains(@class, "T0rpy-LEwYNQifsgB-3SQw==")]').text
        
        try:
            sold = element.find_element(By.XPATH, './/span[@class="se8WAnkjbVXZNA8mT+Veuw=="]').text
        except:
            sold = None
        
        details_link = item.find_element(By.XPATH, './/a').get_attribute('href')
        
        product_data.append({
            'name': name,
            'price': price,
            'store': store,
            'sold': sold,
            'details_link': details_link
        })
    
    return product_data

def main():
    driver = webdriver.Chrome()
    driver.get('https://www.tokopedia.com/')
    driver.implicitly_wait(5)
    
    keywords = input("Keywords: ")
    pages = int(input("Pages: "))
    
    search_input = driver.find_element(By.XPATH, '//*[@id="header-main-wrapper"]/div[2]/div[2]/div/div/div/div/input')
    search_input.send_keys(keywords)
    search_input.send_keys(Keys.ENTER)
    
    driver.implicitly_wait(5)
    all_data = []
    
    for _ in range(pages):
        all_data.extend(extract_data(driver))
        try:
            next_btn = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[aria-label="Laman berikutnya"]'))
            )
        except:
            driver.refresh()
            scroll_page(driver)
            reverse_scroll(driver)
            scroll_page(driver)
            next_btn = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[aria-label="Laman berikutnya"]'))
            )
        try:
            next_btn.click()
        except:
            break

    df = pd.DataFrame(all_data)
    now = datetime.datetime.today().strftime('%d-%m-%Y')
    df.to_csv(f'Tokopedia_{now}.csv', index=False)
    df.to_excel(f'Tokopedia_{now}.xlsx', index=False)
    
    driver.quit()

if __name__ == "__main__":
    main()
