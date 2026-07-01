import sqlite3
import os
import numpy as np
import pandas as pd
from scipy import stats
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

class ExcelReporter:
    CANNIBALIZATION_THRESHOLD = 10.0
    COLOR_BOSCH = "E3170A"
    COLOR_SIEMENS = "00A19A"
    COLOR_HEADER_BG = "1F2937"
    COLOR_HEADER_FONT = "FFFFFF"
    COLOR_RISK = "FBE1E1"
    COLOR_RISK_FONT = "9C1F1F"
    COLOR_SAFE = "E3F3E8"
    COLOR_SAFE_FONT = "1E7A3D"
    COLOR_TITLE_BG = "111827"

    def __init__(self, db_name="data/bsh_fiyat_veritabani.db", report_name="reports/BSH_Kanibalizm_ve_Istatistik_Analizi.xlsx", target_brands=("Bosch", "Siemens")):
        self.db_name = db_name
        self.report_name = report_name
        self.target_brands = list(target_brands)
        os.makedirs(os.path.dirname(self.report_name) or ".", exist_ok=True)

    def _detect_date_column(self, conn):
        cols = pd.read_sql_query("PRAGMA table_info(price_history)", conn)["name"].tolist()
        for candidate in ("scraped_at", "date", "created_at", "timestamp", "recorded_at", "updated_at"):
            if candidate in cols:
                return candidate
        return None

    def _fetch_data(self, conn):
        date_col = self._detect_date_column(conn)
        order_col = f"ph.{date_col}" if date_col else "ph.id"
        query = f"""
            SELECT p.brand         AS Marka,
                   p.category      AS Kategori,
                   p.model_code    AS Model_Kodu,
                   p.title         AS Urun_Adi,
                   ph.price        AS Fiyat,
                   ph.stock_status AS Stok_Durumu
            FROM products p
            JOIN price_history ph ON p.id = ph.product_id
            WHERE ph.id = (
                SELECT ph2.id FROM price_history ph2
                WHERE ph2.product_id = p.id
                ORDER BY {order_col.replace('ph.', 'ph2.')} DESC, ph2.id DESC
                LIMIT 1
            )
        """
        df = pd.read_sql_query(query, conn)
        df["Kategori"] = df["Kategori"].str.strip().str.title()
        return df

    def _build_cannibalization_table(self, df):
        df_bs = df[df["Marka"].isin(self.target_brands)]
        if df_bs.empty:
            return pd.DataFrame()
        stats_df = (
            df_bs.groupby(["Kategori", "Marka"])["Fiyat"]
            .agg(Ortalama_Fiyat="mean", Urun_Sayisi="count", Min_Fiyat="min", Max_Fiyat="max", Standart_Sapma="std")
            .reset_index()
        )
        wide = stats_df.pivot(index="Kategori", columns="Marka", values=["Ortalama_Fiyat", "Urun_Sayisi", "Min_Fiyat", "Max_Fiyat", "Standart_Sapma"])
        wide.columns = [f"{a}_{b}" for a, b in wide.columns]
        wide = wide.reset_index()
        b1, b2 = self.target_brands[0], self.target_brands[1]
        for brand in (b1, b2):
            if f"Ortalama_Fiyat_{brand}" not in wide.columns:
                wide[f"Ortalama_Fiyat_{brand}"] = np.nan
            if f"Urun_Sayisi_{brand}" not in wide.columns:
                wide[f"Urun_Sayisi_{brand}"] = 0
        wide = wide.dropna(subset=[f"Ortalama_Fiyat_{b1}", f"Ortalama_Fiyat_{b2}"]).copy()
        if wide.empty:
            return wide
        wide["Fiyat_Farki_TL"] = wide[f"Ortalama_Fiyat_{b2}"] - wide[f"Ortalama_Fiyat_{b1}"]
        wide[f"{b2}_Prim_Yuzde"] = (wide["Fiyat_Farki_TL"] / wide[f"Ortalama_Fiyat_{b1}"] * 100)
        p_values = []
        sig_results = []
        for cat in wide["Kategori"]:
            b1_prices = df_bs[(df_bs["Kategori"] == cat) & (df_bs["Marka"] == b1)]["Fiyat"].dropna()
            b2_prices = df_bs[(df_bs["Kategori"] == cat) & (df_bs["Marka"] == b2)]["Fiyat"].dropna()
            if len(b1_prices) > 1 and len(b2_prices) > 1:
                t_stat, p_val = stats.ttest_ind(b2_prices, b1_prices, equal_var=False)
                p_values.append(p_val)
                sig_results.append("Anlamlı" if p_val < 0.05 else "Tesadüfi")
            else:
                p_values.append(np.nan)
                sig_results.append("Veri Yetersiz")
        wide["P_Degeri"] = p_values
        wide["Istatistiksel_Fark"] = sig_results
        def label(x):
            if abs(x) < self.CANNIBALIZATION_THRESHOLD:
                return "Kanibalizm Riski"
            return f"{b2} Premium" if x > 0 else f"{b1} Premium"
        wide["Stratejik_Konum"] = wide[f"{b2}_Prim_Yuzde"].apply(label)
        return wide.sort_values(f"{b2}_Prim_Yuzde", ascending=False).reset_index(drop=True)

    def _generate_insight_text(self, wide):
        b1, b2 = self.target_brands[0], self.target_brands[1]
        if wide.empty:
            return "Ortak kategori bulunamadı."
        overall_premium = wide[f"{b2}_Prim_Yuzde"].mean()
        risk_segments = wide[wide["Stratejik_Konum"] == "Kanibalizm Riski"]
        anlamli_segments = wide[wide["Istatistiksel_Fark"] == "Anlamlı"]
        yon = "daha pahalı" if overall_premium >= 0 else "daha ucuz"
        text = (f"Veri bilimi ve istatistiksel T-Testi (Welch's) sonuçlarına göre; ortalama olarak {b2} ürünleri, "
                f"aynı segmentteki {b1} ürünlerine kıyasla %{abs(overall_premium):.1f} {yon} konumlandırılmıştır. ")
        if not anlamli_segments.empty:
            text += f"Bu fiyatlandırma farkı {len(anlamli_segments)} kategoride istatistiksel olarak 'Anlamlı' (p<0.05) bulunmuş olup, bilinçli bir stratejiyi işaret etmektedir. "
        if not risk_segments.empty:
            risky_names = ", ".join(risk_segments["Kategori"].tolist())
            text += (f"Ancak {risky_names} kategorilerinde fiyat farkı hedeflenen "
                     f"%{self.CANNIBALIZATION_THRESHOLD:.0f} barajının altında kaldığı için kanibalizasyon riski tespit edilmiştir.")
        return text

    def _style_header_row(self, ws, row=1, n_cols=None):
        n_cols = n_cols or ws.max_column
        fill = PatternFill("solid", fgColor=self.COLOR_HEADER_BG)
        font = Font(bold=True, color=self.COLOR_HEADER_FONT, size=11)
        for col in range(1, n_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.freeze_panes = ws.cell(row=row + 1, column=1)

    def _autofit_columns(self, ws, min_width=15, max_width=50):
        widths = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                col = cell.column_letter
                widths[col] = max(widths.get(col, 0), len(str(cell.value)))
        for col, width in widths.items():
            header_value = ws[f"{col}1"].value
            if header_value and "Fiyat" in str(header_value):
                ws.column_dimensions[col].width = 22
            elif header_value and "Standart_Sapma" in str(header_value):
                ws.column_dimensions[col].width = 22
            else:
                ws.column_dimensions[col].width = min(max(width + 4, min_width), max_width)

    def _write_title_block(self, ws, title, subtitle, span_cols=7):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span_cols)
        c = ws.cell(row=1, column=1, value=title)
        c.font = Font(bold=True, size=16, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=self.COLOR_TITLE_BG)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 30
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span_cols)
        c2 = ws.cell(row=2, column=1, value=subtitle)
        c2.font = Font(italic=True, size=10, color="6B7280")
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 18

    def generate_report(self):
        conn = sqlite3.connect(self.db_name)
        df = self._fetch_data(conn)
        conn.close()
        if df.empty:
            print("Veritabanında raporlanacak veri bulunamadı.")
            return
        b1, b2 = self.target_brands[0], self.target_brands[1]
        wide = self._build_cannibalization_table(df)
        insight_text = self._generate_insight_text(wide)
        with pd.ExcelWriter(self.report_name, engine="openpyxl") as writer:
            pd.DataFrame().to_excel(writer, sheet_name="Yonetici_Ozeti", index=False)
            wide.to_excel(writer, sheet_name="Kanibalizm_Analizi", index=False, startrow=0)
            df.sort_values(["Kategori", "Marka", "Fiyat"]).to_excel(writer, sheet_name="Detayli_Liste", index=False)
        wb = load_workbook(self.report_name)
        ws_det = wb["Detayli_Liste"]
        self._style_header_row(ws_det)
        self._autofit_columns(ws_det, min_width=15, max_width=60)
        for row in ws_det.iter_rows(min_row=2, min_col=5, max_col=5):
            for cell in row:
                cell.number_format = '#,##0.00 "TL"'
        ws_kan = wb["Kanibalizm_Analizi"]
        self._style_header_row(ws_kan)
        headers = [c.value for c in ws_kan[1]]
        pct_col = headers.index(f"{b2}_Prim_Yuzde") + 1 if f"{b2}_Prim_Yuzde" in headers else None
        konum_col = headers.index("Stratejik_Konum") + 1 if "Stratejik_Konum" in headers else None
        ist_col = headers.index("Istatistiksel_Fark") + 1 if "Istatistiksel_Fark" in headers else None
        price_cols = [i + 1 for i, h in enumerate(headers) if h and ("Fiyat" in h or "Standart_Sapma" in h)]
        for col in price_cols:
            for row in ws_kan.iter_rows(min_row=2, min_col=col, max_col=col):
                for cell in row:
                    cell.number_format = '#,##0.00 "TL"'
        if pct_col:
            for row in ws_kan.iter_rows(min_row=2, min_col=pct_col, max_col=pct_col):
                for cell in row:
                    cell.number_format = '+0.0"%";-0.0"%"'
        if konum_col:
            risk_fill = PatternFill("solid", fgColor=self.COLOR_RISK)
            risk_font = Font(color=self.COLOR_RISK_FONT, bold=True)
            safe_fill = PatternFill("solid", fgColor=self.COLOR_SAFE)
            safe_font = Font(color=self.COLOR_SAFE_FONT)
            for r in range(2, ws_kan.max_row + 1):
                cell = ws_kan.cell(row=r, column=konum_col)
                target_fill = risk_fill if cell.value == "Kanibalizm Riski" else safe_fill
                target_font = risk_font if cell.value == "Kanibalizm Riski" else safe_font
                for c in range(1, ws_kan.max_column + 1):
                    ws_kan.cell(row=r, column=c).fill = target_fill
                ws_kan.cell(row=r, column=konum_col).font = target_font
                ist_cell = ws_kan.cell(row=r, column=ist_col)
                if ist_cell.value == "Anlamlı":
                    ist_cell.font = Font(color="005CBF", bold=True)
                elif ist_cell.value == "Tesadüfi":
                    ist_cell.font = Font(color="856404")
        self._autofit_columns(ws_kan, min_width=16, max_width=35)
        ws_sum = wb["Yonetici_Ozeti"]
        self._write_title_block(ws_sum, "BSH Kurumsal Fiyat Zekası ve Strateji Analizi", f"{b1} vs {b2} | Istatistiksel P-Value Analizi | Oluşturulma: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M')}", span_cols=7)
        ws_sum.merge_cells(start_row=4, start_column=1, end_row=7, end_column=7)
        insight_cell = ws_sum.cell(row=4, column=1, value=insight_text)
        insight_cell.font = Font(size=11, bold=True, color="1F2937")
        insight_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        insight_cell.fill = PatternFill("solid", fgColor="F3F4F6")
        for r in range(4, 8):
            ws_sum.row_dimensions[r].height = 22
        kpi_row = 9
        ws_sum.cell(row=kpi_row, column=1, value="Makro İstatistikler").font = Font(bold=True, size=13)
        kpi_headers = ["Gösterge", "Değer"]
        for i, h in enumerate(kpi_headers):
            cell = ws_sum.cell(row=kpi_row + 1, column=1 + i, value=h)
            cell.font = Font(bold=True, color=self.COLOR_HEADER_FONT)
            cell.fill = PatternFill("solid", fgColor=self.COLOR_HEADER_BG)
        n_total_categories = wide.shape[0]
        n_risk = int((wide["Stratejik_Konum"] == "Kanibalizm Riski").sum()) if not wide.empty else 0
        n_sig = int((wide["Istatistiksel_Fark"] == "Anlamlı").sum()) if not wide.empty else 0
        avg_premium = wide[f"{b2}_Prim_Yuzde"].mean() if not wide.empty else np.nan
        kpis = [
            ("Analiz Edilen Segment Sayısı", n_total_categories),
            ("İstatistiksel Olarak Anlamlı Fark", n_sig),
            ("Kanibalizm Riski Taşıyan Segment", n_risk),
            (f"Ortalama {b2} Primi (%)", None if pd.isna(avg_premium) else round(avg_premium, 1)),
            (f"Toplam {b1} Veri Noktası", int((df["Marka"] == b1).sum())),
            (f"Toplam {b2} Veri Noktası", int((df["Marka"] == b2).sum())),
        ]
        for i, (label, value) in enumerate(kpis, start=1):
            ws_sum.cell(row=kpi_row + 1 + i, column=1, value=label)
            ws_sum.cell(row=kpi_row + 1 + i, column=2, value=value)
        detail_row = kpi_row + len(kpis) + 4
        ws_sum.cell(row=detail_row, column=1, value="Kategori Bazında İstatistiksel Konumlandırma").font = Font(bold=True, size=13)
        mini_headers = ["Kategori", f"{b1} Ort. (TL)", f"{b2} Ort. (TL)", "Fark (%)", "P-Value (T-Test)", "Sonuç", "Konum"]
        for i, h in enumerate(mini_headers):
            cell = ws_sum.cell(row=detail_row + 1, column=1 + i, value=h)
            cell.font = Font(bold=True, color=self.COLOR_HEADER_FONT)
            cell.fill = PatternFill("solid", fgColor=self.COLOR_HEADER_BG)
        for i, r in wide.iterrows():
            row_idx = detail_row + 2 + i
            ws_sum.cell(row=row_idx, column=1, value=r["Kategori"])
            ws_sum.cell(row=row_idx, column=2, value=round(r[f"Ortalama_Fiyat_{b1}"], 2)).number_format = '#,##0 "TL"'
            ws_sum.cell(row=row_idx, column=3, value=round(r[f"Ortalama_Fiyat_{b2}"], 2)).number_format = '#,##0 "TL"'
            ws_sum.cell(row=row_idx, column=4, value=round(r[f"{b2}_Prim_Yuzde"], 1)).number_format = '+0.0"%";-0.0"%"'
            ws_sum.cell(row=row_idx, column=5, value=r["P_Degeri"]).number_format = '0.000'
            ws_sum.cell(row=row_idx, column=6, value=r["Istatistiksel_Fark"])
            konum_cell = ws_sum.cell(row=row_idx, column=7, value=r["Stratejik_Konum"])
            if r["Stratejik_Konum"] == "Kanibalizm Riski":
                fill, font = PatternFill("solid", fgColor=self.COLOR_RISK), Font(color=self.COLOR_RISK_FONT, bold=True)
            else:
                fill, font = PatternFill("solid", fgColor=self.COLOR_SAFE), Font(color=self.COLOR_SAFE_FONT)
            for c in range(1, 8):
                ws_sum.cell(row=row_idx, column=c).fill = fill
            konum_cell.font = font
        self._autofit_columns(ws_sum, min_width=18, max_width=40)
        ws_sum.sheet_view.showGridLines = False
        if not wide.empty:
            chart = BarChart()
            chart.type = "col"
            chart.title = f"Fiyat Konumlandırması: {b1} vs {b2}"
            chart.y_axis.title = "Fiyat (TL)"
            chart.style = 12
            data = Reference(ws_sum, min_col=2, max_col=3, min_row=detail_row + 1, max_row=detail_row + 1 + len(wide))
            cats = Reference(ws_sum, min_col=1, min_row=detail_row + 2, max_row=detail_row + 1 + len(wide))
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.width, chart.height = 18, 9
            ws_sum.add_chart(chart, "J9")
        wb.move_sheet("Yonetici_Ozeti", offset=-len(wb.sheetnames))
        wb.active = 0
        wb.save(self.report_name)
        print(f"Analiz raporu başarıyla oluşturuldu: {self.report_name}")
        print("-" * 50)
        print(insight_text)

if __name__ == "__main__":
    reporter = ExcelReporter()
    reporter.generate_report()