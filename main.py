from src.data_collector import BSHScraper
from src.excel_report import ExcelReporter
import time

def main():
    scraper = BSHScraper()
    
    # BSH (Bosch & Siemens) için çekilecek tüm kategoriler ve web adresleri
    bsh_categories = [
        # --- BOSCH ---
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
        },
        
        # --- SIEMENS ---
        {
            "brand": "Siemens",
            "category": "Çamaşır Makinesi",
            "url": "https://www.siemens-home.bsh-group.com/tr/urun-listesi/yikama-ve-utuleme-grubu/camasir-makineleri"
        },
        {
            "brand": "Siemens",
            "category": "Bulaşık Makinesi",
            "url": "https://www.siemens-home.bsh-group.com/tr/tr/category/buzdolaplari-ve-derin-dondurucular/ankastre-buzdolaplari-ve-derin-dondurucular"
        },
        {
            "brand": "Siemens",
            "category": "Buzdolabı",
            "url": "https://www.siemens-home.bsh-group.com/tr/urun-listesi/buzdolaplari-ve-derin-dondurucular/alttan-donduruculu-buzdolaplari"
        },
        {
            "brand": "Siemens",
            "category": "Klima",
            "url": "https://www.siemens-home.bsh-group.com/tr/urun-listesi/klima-ve-ev-konforu/klimalar"
        }
    ]

    print("--- BSH VERİ TOPLAMA İŞLEMİ BAŞLIYOR ---")
    for item in bsh_categories:
        scraper.run(item["url"], item["brand"], item["category"])
        time.sleep(2) 
        
    print("\n--- RAPORLAMA İŞLEMİ BAŞLIYOR ---")
    reporter = ExcelReporter()
    reporter.generate_report()
    
    print("\nBütün işlemler başarıyla tamamlandı!")

if __name__ == "__main__":
    main()