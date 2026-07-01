import sqlite3
import pandas as pd
import os

class ExcelReporter:
    def __init__(self, db_name="data/bsh_fiyat_veritabani.db", report_name="reports/BSH_Detayli_Fiyat_Raporu.xlsx"):
        self.db_name = db_name
        self.report_name = report_name
        os.makedirs(os.path.dirname(self.report_name), exist_ok=True)

    def generate_report(self):
        conn = sqlite3.connect(self.db_name)
        
        query = """
        SELECT 
            p.brand AS Marka,
            p.category AS Kategori,
            p.model_code AS Model_Kodu,
            p.title AS Urun_Adi,
            ph.price AS Fiyat,
            ph.stock_status AS Stok_Durumu,
            ph.scrape_date AS Cekilme_Tarihi
        FROM products p
        JOIN price_history ph ON p.id = ph.product_id
        ORDER BY p.brand, p.category, ph.price ASC
        """
        
        df = pd.read_sql_query(query, conn)
        
        if not df.empty:
            df.to_excel(self.report_name, index=False, sheet_name="Fiyat_Listesi")
            print(f"Rapor başarıyla oluşturuldu: {self.report_name}")
        else:
            print("Veritabanında raporlanacak veri bulunamadı.")
            
        conn.close()

if __name__ == "__main__":
    reporter = ExcelReporter()
    reporter.generate_report()