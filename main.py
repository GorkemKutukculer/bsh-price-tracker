from src.data_collector import BSHScraper
from src.excel_report import ExcelReporter
import time

def main():
    scraper = BSHScraper()
    
    # Bosch için çekilecek tüm kategoriler ve web adresleri
    bosch_categories = [
        {
            "brand": "Bosch",
            "category": "Çamaşır Makinesi",
            "url": "https://www.bosch-home.com.tr/urun-listesi/camasir-ve-kurutma-makineleri/camasir-makineleri"
        },
        {
            "brand": "Bosch",
            "category": "Bulaşık Makinesi",
            "url": "https://www.bosch-home.com.tr/urun-listesi/bulasik-makineleri"
        },
        {
            "brand": "Bosch",
            "category": "Buzdolabı",
            "url": "https://www.bosch-home.com.tr/tr/category/buzdolaplari-derin-dondurucular/buzdolaplari"
        },
        {
            "brand": "Bosch",
            "category": "Klima",
            "url": "https://www.bosch-home.com.tr/urun-listesi/klima-ve-ev-konforu/klimalar"
        }
    ]

    print("--- VERİ TOPLAMA İŞLEMİ BAŞLIYOR ---")
    for item in bosch_categories:
        scraper.run(item["url"], item["brand"], item["category"])
        time.sleep(2) # Kategoriler arası sitenin bizi engellememesi için kısa bir bekleme
        
    print("\n--- RAPORLAMA İŞLEMİ BAŞLIYOR ---")
    reporter = ExcelReporter()
    reporter.generate_report()
    
    print("\nBütün işlemler başarıyla tamamlandı!")

if __name__ == "__main__":
    main()