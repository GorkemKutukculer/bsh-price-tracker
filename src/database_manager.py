import sqlite3
import os
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_name="data/bsh_fiyat_veritabani.db"):
        os.makedirs(os.path.dirname(db_name), exist_ok=True)
        self.db_name = db_name
        self.conn = None
        self.cursor = None

    def connect(self):
        self.conn = sqlite3.connect(self.db_name)
        self.cursor = self.conn.cursor()

    def disconnect(self):
        if self.conn:
            self.conn.commit()
            self.conn.close()

    def create_tables(self):
        self.connect()
        
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                brand TEXT NOT NULL,
                category TEXT NOT NULL,
                model_code TEXT UNIQUE NOT NULL,
                title TEXT,
                features TEXT
            )
        """)

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS price_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER,
                price REAL NOT NULL,
                stock_status TEXT,
                scrape_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (product_id) REFERENCES products (id)
            )
        """)
        
        self.disconnect()
        print("Veritabanı ve tablolar başarıyla oluşturuldu.")

    def insert_or_get_product(self, brand, category, model_code, title, features):
        self.connect()
        
        self.cursor.execute("SELECT id FROM products WHERE model_code = ?", (model_code,))
        result = self.cursor.fetchone()
        
        if result:
            product_id = result[0]
        else:
            self.cursor.execute("""
                INSERT INTO products (brand, category, model_code, title, features)
                VALUES (?, ?, ?, ?, ?)
            """, (brand, category, model_code, title, features))
            product_id = self.cursor.lastrowid
            
        self.disconnect()
        return product_id

    def insert_price(self, product_id, price, stock_status):
        self.connect()
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        self.cursor.execute("""
            INSERT INTO price_history (product_id, price, stock_status, scrape_date)
            VALUES (?, ?, ?, ?)
        """, (product_id, price, stock_status, current_time))
        
        self.disconnect()

if __name__ == "__main__":
    db = DatabaseManager()
    db.create_tables()
    
    p_id = db.insert_or_get_product(
        brand="Bosch", 
        category="Çamaşır Makinesi", 
        model_code="WGA254A0TR", 
        title="Bosch 10 kg 1200 Devir Çamaşır Makinesi", 
        features="C Sınıfı, Leke Çıkarma"
    )
    
    db.insert_price(product_id=p_id, price=24500.0, stock_status="Stokta Var")
    print("Örnek veri başarıyla veritabanına işlendi!")