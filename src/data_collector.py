import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from src.database_manager import DatabaseManager
import time

class BSHScraper:
    def __init__(self):
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        self.db = DatabaseManager()
        self.db.create_tables()

    def get_product_links(self, category_url):
        links = []
        parsed_uri = urlparse(category_url)
        base_url = f"{parsed_uri.scheme}://{parsed_uri.netloc}"
        
        try:
            response = requests.get(category_url, headers=self.headers)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "html.parser")
            
            for a_tag in soup.find_all("a", href=True):
                href = a_tag["href"]
                if "/product/" in href:
                    full_url = urljoin(base_url, href)
                    if full_url not in links:
                        links.append(full_url)
        except Exception as e:
            print(f"Link toplama hatası: {e}")
        return list(set(links))

    def scrape_product_detail(self, url, brand, category):
        try:
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "html.parser")
            
            title_elem = soup.find(attrs={"data-testid": "buy-area-title-headline"})
            title = title_elem.text.strip() if title_elem else "Bilinmeyen Ürün"
            
            model_code_elem = soup.find(attrs={"data-testid": "product-id-label"})
            model_code = model_code_elem.text.strip() if model_code_elem else "KOD_YOK"
            
            price = 0.0
            price_elem = soup.find(string=lambda text: text and ("₺" in text or "TL" in text))
            
            if price_elem:
                price_str = price_elem.replace("₺", "").replace("TL", "").replace(".", "").replace(",", ".").strip()
                try:
                    price = float(price_str)
                except ValueError:
                    pass
                    
            features = "Standart Özellikler"
            
            if model_code != "KOD_YOK" and price > 0:
                p_id = self.db.insert_or_get_product(brand, category, model_code, title, features)
                self.db.insert_price(p_id, price, "Stokta Var")
                print(f"Kaydedildi: {model_code} - {price} TL")
                
        except Exception as e:
            print(f"Hata ({url}): {e}")

    def run(self, category_url, brand, category):
        print(f"{brand} - {category} linkleri toplanıyor...")
        links = self.get_product_links(category_url)
        print(f"Toplam {len(links)} ürün bağlantısı bulundu. Veriler çekiliyor...")
        
        for link in links:
            self.scrape_product_detail(link, brand, category)
            time.sleep(1)