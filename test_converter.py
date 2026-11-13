"""
cPDF Converter - Kapsamlı Fonksiyon Test Paketi
Tüm dönüştürme fonksiyonlarını uçtan uca otomatik test eder.
"""

import io
import sys
import os
import unittest
from PIL import Image
from pypdf import PdfReader, PdfWriter
import docx
from docx import Document
import pandas as pd
from reportlab.pdfgen import canvas

# app.py içerisindeki ana fonksiyonları içeri aktar
from app import (
    convert_pdf_to_docx,
    merge_pdf_files,
    convert_images_to_pdf,
    convert_docx_to_pdf,
    convert_docx_to_excel
)

class MockUploadedFile:
    """Streamlit UploadedFile nesnesini taklit eden yardımcı sınıf."""
    def __init__(self, name, data):
        self.name = name
        self._data = data
    def getvalue(self):
        return self._data
    def read(self):
        return self._data

class TestCPDFConverter(unittest.TestCase):
    
    def setUp(self):
        print("\n" + "="*50)

    def test_01_pdf_to_docx(self):
        print("▶ TEST 1: PDF -> Word (.docx) Dönüştürme Test Ediliyor...")
        # 1. Geçerli bir test PDF'i oluştur
        pdf_buf = io.BytesIO()
        c = canvas.Canvas(pdf_buf)
        c.drawString(100, 750, "cPDF Converter Otomatik Test Belgesi")
        c.drawString(100, 700, "Bu belge PDF'ten Word'e dönüştürme testidir.")
        c.save()
        pdf_bytes = pdf_buf.getvalue()
        
        # 2. Fonksiyonu çalıştır
        docx_bytes = convert_pdf_to_docx(pdf_bytes)
        
        # 3. Doğrula
        self.assertIsNotNone(docx_bytes)
        self.assertGreater(len(docx_bytes), 1000)
        
        # Word belgesi olarak açılabilirliğini doğrula
        doc = Document(io.BytesIO(docx_bytes))
        full_text = " ".join([p.text for p in doc.paragraphs])
        self.assertTrue(len(full_text) > 0)
        print(f"  ✓ BAŞARILI: PDF başarıyla {len(docx_bytes)} baytlık geçerli bir Word belgesine dönüştürüldü.")

    def test_02_merge_pdf(self):
        print("▶ TEST 2: Çoklu PDF Birleştirme (Merge) Test Ediliyor...")
        # 1. İki adet farklı sayfalara sahip PDF üret
        def make_simple_pdf(text):
            buf = io.BytesIO()
            c = canvas.Canvas(buf)
            c.drawString(100, 750, text)
            c.showPage()
            c.save()
            return buf.getvalue()

        pdf1 = MockUploadedFile("belge_1.pdf", make_simple_pdf("Belge 1 İçeriği"))
        pdf2 = MockUploadedFile("belge_2.pdf", make_simple_pdf("Belge 2 İçeriği"))
        pdf3 = MockUploadedFile("belge_3.pdf", make_simple_pdf("Belge 3 İçeriği"))
        
        # 2. Fonksiyonu çalıştır
        merged_bytes = merge_pdf_files([pdf1, pdf2, pdf3])
        
        # 3. Doğrula
        self.assertIsNotNone(merged_bytes)
        reader = PdfReader(io.BytesIO(merged_bytes))
        self.assertEqual(len(reader.pages), 3)
        print(f"  ✓ BAŞARILI: 3 adet bağımsız PDF belgesi birleştirildi. Toplam sayfa sayısı: {len(reader.pages)}")

    def test_03_images_to_pdf_unlimited(self):
        print("▶ TEST 3: Görsellerden PDF'e (Sınırsız PNG/JPG) Test Ediliyor...")
        # 1. Hem RGBA şeffaf PNG hem de RGB JPG oluştur
        imgs = []
        for i in range(5):
            # RGBA şeffaflıklı görsel
            img = Image.new("RGBA", (200, 200), (0, 255, 102, 180))
            b = io.BytesIO()
            img.save(b, format="PNG")
            imgs.append(b)

        # 2. Fonksiyonu çalıştır
        pdf_bytes = convert_images_to_pdf(imgs)
        
        # 3. Doğrula
        self.assertIsNotNone(pdf_bytes)
        reader = PdfReader(io.BytesIO(pdf_bytes))
        self.assertEqual(len(reader.pages), 5)
        print(f"  ✓ BAŞARILI: 5 adet görsel (şeffaf PNG dahil) başarıyla {len(reader.pages)} sayfalık PDF'e dönüştürüldü.")

    def test_04_docx_to_pdf(self):
        print("▶ TEST 4: Word (.docx) -> PDF Dönüştürme Test Ediliyor...")
        # 1. Başlık, paragraf ve tablolu test Word belgesi oluştur
        doc = Document()
        doc.add_heading("Test Raporu: Türkçe Karakterler (ç, ğ, ı, ö, ş, ü)", level=1)
        doc.add_paragraph("Bu paragraf Word'den PDF'e saf Python motorunun testidir.")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Sütun A"
        table.cell(0, 1).text = "Sütun B"
        table.cell(1, 0).text = "Veri 1"
        table.cell(1, 1).text = "Veri 2"

        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_bytes = docx_buf.getvalue()

        # 2. Fonksiyonu çalıştır
        pdf_bytes = convert_docx_to_pdf(docx_bytes)

        # 3. Doğrula
        self.assertIsNotNone(pdf_bytes)
        reader = PdfReader(io.BytesIO(pdf_bytes))
        self.assertGreaterEqual(len(reader.pages), 1)
        print(f"  ✓ BAŞARILI: Word belgesi (başlık, tablo ve metinler) başarıyla PDF'e dönüştürüldü ({len(pdf_bytes)} bayt).")

    def test_05_docx_to_excel(self):
        print("▶ TEST 5: Word (.docx) -> Excel (.xlsx) Dönüştürme Test Ediliyor...")
        # 1. 2 adet tablo ve metin içeren Word belgesi hazırla
        doc = Document()
        doc.add_heading("Satış Raporu", level=1)
        doc.add_paragraph("Aşağıdaki tablolar 2026 yılı satışlarını içermektedir.")
        
        # Tablo 1
        t1 = doc.add_table(rows=3, cols=3)
        t1_data = [
            ["Ürün", "Miktar", "Fiyat"],
            ["Monitör", "5", "12000"],
            ["Klavye", "10", "1500"]
        ]
        for r_i, row in enumerate(t1_data):
            for c_i, val in enumerate(row):
                t1.cell(r_i, c_i).text = val

        # Tablo 2
        t2 = doc.add_table(rows=2, cols=2)
        t2.cell(0, 0).text = "Departman"
        t2.cell(0, 1).text = "Yetkili"
        t2.cell(1, 0).text = "Yazılım"
        t2.cell(1, 1).text = "Ahmet Yılmaz"

        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_bytes = docx_buf.getvalue()

        # 2. Fonksiyonu çalıştır
        excel_bytes, tables_list, df_texts = convert_docx_to_excel(docx_bytes)

        # 3. Doğrula
        self.assertIsNotNone(excel_bytes)
        self.assertEqual(len(tables_list), 2)
        
        # Excel dosyasının sayfalarını ve verilerini test et
        xl = pd.ExcelFile(io.BytesIO(excel_bytes), engine='openpyxl')
        sheet_names = xl.sheet_names
        self.assertIn("Belge_Metinleri", sheet_names)
        self.assertIn("Tablo_1", sheet_names)
        self.assertIn("Tablo_2", sheet_names)
        
        df_tablo1 = pd.read_excel(xl, sheet_name="Tablo_1")
        self.assertEqual(len(df_tablo1), 2)
        self.assertIn("Ürün", df_tablo1.columns)
        print(f"  ✓ BAŞARILI: Word'deki 2 tablo ve metinler başarıyla ayrıştırıldı. Excel sayfaları: {sheet_names}")

if __name__ == "__main__":
    unittest.main()

