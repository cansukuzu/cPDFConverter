"""
cPDF Converter - Eğlenceli, Animasyonlu ve Modern Belge Dönüştürücü
Renk Paleti: Gece Laciverti & Neon Yeşil Detaylar
Özellikler: Canlı Hareketli Arka Plan, Kutlama Balonları (Balloons/XP), TR/EN Çift Dil, %100 Mobil Uyumlu
"""

import io
import os
import tempfile
import docx
from docx import Document
import pandas as pd
from PIL import Image
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import streamlit as st

# ReportLab bileşenleri (Saf Python Word -> PDF motoru)
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------------------------------------------------------
# 1. SAYFA YAPILANDIRMASI
# ---------------------------------------------------------
st.set_page_config(
    page_title="cPDF Converter ⚡",
    page_icon="🎮",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------------------------------------------------
# 2. DİL YÖNETİMİ
# ---------------------------------------------------------
if "lang" not in st.session_state:
    st.session_state.lang = "TR"

TEXTS = {
    "TR": {
        "title_prefix": "cPDF",
        "title_suffix": "CONVERTER",
        "subtitle": "⚡ Yüksek Hızlı, Güvenli ve Eğlenceli Belge İstasyonu",
        "xp_badge": "🎮 LEVEL 99 • +500 XP KAZANMAYA HAZIR",
        "lang_btn": "🇬🇧 English",
        
        # Sekmeler
        "tab_pdf2word": "📄 PDF ➔ Word",
        "tab_merge": "📚 PDF Birleştir",
        "tab_to_pdf": "🖼️ PDF'e Çevir",
        "tab_word2excel": "📊 Word ➔ Excel",
        
        # Sekme 1
        "t1_desc": "PDF belgenizi biçimlendirmesi bozulmadan düzenlenebilir Word (.docx) dosyasına dönüştürün.",
        "t1_drop": "Dönüştürülecek PDF dosyasını buraya yükleyin:",
        "t1_btn": "🚀 WORD'E ATEŞLE (DÖNÜŞTÜR)",
        "t1_converting": "⚡ PDF sayfaları çözümleniyor... Lütfen bekleyin!",
        "t1_success": "🏆 GÖREV TAMAMLANDI! +100 XP • Word dosyanız pırıl pırıl hazırlandı.",
        "t1_download": "🎁 ÖDÜLÜ AL: Word Dosyasını İndir (.docx)",
        
        # Sekme 2
        "t2_desc": "Farklı PDF belgelerini tek bir güçlü master PDF dosyası altında birleştirin.",
        "t2_drop": "Birleştirilecek PDF belgelerini seçin (İstediğiniz kadar ekleyin):",
        "t2_out_name": "Birleşik Belge Adı:",
        "t2_btn": "⚔️ GÜÇLERİ BİRLEŞTİR (PDF MERGE)",
        "t2_warn_min": "⚠️ En az 2 adet PDF dosyası yüklemelisiniz!",
        "t2_merging": "⚡ Füzyon reaktörü devrede! Belgeler birleştiriliyor...",
        "t2_success": "🎉 HARİKA BİRLEŞİM! {count} adet belge tek çatıda toplandı. +150 XP",
        "t2_download": "🛡️ BİRLEŞİK PDF'İ HEMEN İNDİR",
        
        # Sekme 3
        "t3_desc": "Sınırsız sayıda PNG/JPG görselini veya Word belgenizi tek tıkla PDF yapın.",
        "t3_mode_img": "🖼️ Sınırsız Görsellerden PDF (PNG, JPG)",
        "t3_mode_word": "📝 Word Belgesinden PDF'e (.docx ➔ .pdf)",
        "t3_img_drop": "Görselleri buraya yükleyin (Sınırsız PNG & JPG desteği aktif!):",
        "t3_img_info": "👾 Cephanelikte {count} görsel hazır bekliyor!",
        "t3_img_btn": "✨ PİKSELLERİ DERLE (PDF YAP)",
        "t3_img_making": "⚡ Görseller taranıyor ve PDF albümü inşa ediliyor...",
        "t3_img_success": "🎨 PİKSEL ŞÖLENİ! {count} görsel tek bir PDF'e mühürlendi. +120 XP",
        "t3_img_download": "📦 GÖRSEL PDF PAKETİNİ İNDİR",
        "t3_word_drop": "Dönüştürülecek Word (.docx) dosyasını seçin:",
        "t3_word_btn": "🎯 WORD'Ü PDF'E KİLİTLE",
        "t3_word_making": "⚡ Word metinleri ve tabloları PDF matrisine yazılıyor...",
        "t3_word_success": "📜 DÖNÜŞÜM KUSURSUZ! Word belgeniz PDF'e aktarıldı. +100 XP",
        "t3_word_download": "📥 OLUŞTURULAN PDF'İ İNDİR",
        
        # Sekme 4
        "t4_desc": "Word belgenizdeki tüm tabloları ve paragrafları akıllıca ayrıştırıp Excel'e dökün.",
        "t4_drop": "Excel'e aktarılacak Word (.docx) belgesini seçin:",
        "t4_info": "📊 {tables} adet tablo ve {paras} paragraf analiz edildi.",
        "t4_btn": "🎲 EXCEL MATRİSİNİ ÇIKAR (.XLSX)",
        "t4_converting": "⚡ Tablo madenciliği devrede... Hücreler Excel'e aktarılıyor!",
        "t4_success": "💎 VERİ MADENCİLİĞİ BAŞARILI! Tablolar Excel'e aktarıldı. +200 XP",
        "t4_download": "📈 EXCEL ÇALIŞMA KİTABINI İNDİR (.xlsx)",
        "t4_preview_title": "📋 Tablo Önizleme Ekranı",
        
        # Hata & Dipnot
        "err_generic": "⚠️ İşlem sırasında bir aksaklık oldu: {err}",
        "footer": "🎮 cPDF Converter • %100 Güvenli & Yerel İşleme • GG WP!"
    },
    "EN": {
        "title_prefix": "cPDF",
        "title_suffix": "CONVERTER",
        "subtitle": "⚡ High-Speed, Secure & Gamified Document Station",
        "xp_badge": "🎮 LEVEL 99 • READY TO EARN +500 XP",
        "lang_btn": "🇹🇷 Türkçe",
        
        # Tabs
        "tab_pdf2word": "📄 PDF ➔ Word",
        "tab_merge": "📚 Merge PDF",
        "tab_to_pdf": "🖼️ Make PDF",
        "tab_word2excel": "📊 Word ➔ Excel",
        
        # Tab 1
        "t1_desc": "Convert PDF into an editable Word (.docx) file with intact typography & tables.",
        "t1_drop": "Drop your PDF file here:",
        "t1_btn": "🚀 FIRE TO WORD (CONVERT)",
        "t1_converting": "⚡ Laser beams scanning PDF pages... Hold tight!",
        "t1_success": "🏆 MISSION COMPLETE! +100 XP • Your Word document is ready.",
        "t1_download": "🎁 CLAIM LOOT: Download Word (.docx)",
        
        # Tab 2
        "t2_desc": "Merge multiple PDF documents into a single powerhouse PDF file.",
        "t2_drop": "Select PDF files to merge (Add as many as you want):",
        "t2_out_name": "Merged File Name:",
        "t2_btn": "⚔️ COMBINE FORCES (MERGE PDF)",
        "t2_warn_min": "⚠️ You need at least 2 PDF files to merge!",
        "t2_merging": "⚡ Fusion reactor active! Fusing documents together...",
        "t2_success": "🎉 EPIC FUSION! {count} documents consolidated. +150 XP",
        "t2_download": "🛡️ DOWNLOAD MERGED PDF NOW",
        
        # Tab 3
        "t3_desc": "Convert unlimited PNG/JPG images or Word files directly into PDF.",
        "t3_mode_img": "🖼️ Unlimited Images to PDF (PNG, JPG)",
        "t3_mode_word": "📝 Word to PDF (.docx ➔ .pdf)",
        "t3_img_drop": "Upload images here (Unlimited PNG & JPG support!):",
        "t3_img_info": "👾 {count} images loaded in your arsenal!",
        "t3_img_btn": "✨ COMPILE PIXELS (BUILD PDF)",
        "t3_img_making": "⚡ Scanning images and compiling PDF album...",
        "t3_img_success": "🎨 PIXEL MASTERY! {count} images sealed into PDF. +120 XP",
        "t3_img_download": "📦 DOWNLOAD IMAGE PDF BUNDLE",
        "t3_word_drop": "Select the Word (.docx) file to convert:",
        "t3_word_btn": "🎯 LOCK WORD TO PDF",
        "t3_word_making": "⚡ Writing paragraphs and tables into PDF matrix...",
        "t3_word_success": "📜 FLAWLESS CONVERSION! Word file locked to PDF. +100 XP",
        "t3_word_download": "📥 DOWNLOAD GENERATED PDF",
        
        # Tab 4
        "t4_desc": "Extract all tables and paragraphs from Word into clean Excel sheets.",
        "t4_drop": "Select the Word (.docx) document:",
        "t4_info": "📊 {tables} tables and {paras} paragraphs analyzed.",
        "t4_btn": "🎲 EXTRACT EXCEL MATRIX (.XLSX)",
        "t4_converting": "⚡ Table mining active... Weaving cells into Excel!",
        "t4_success": "💎 DATA MINING SUCCESSFUL! Tables mapped to Excel. +200 XP",
        "t4_download": "📈 DOWNLOAD EXCEL WORKBOOK (.xlsx)",
        "t4_preview_title": "📋 Table Preview",
        
        # Error & Footer
        "err_generic": "⚠️ An error occurred during mission: {err}",
        "footer": "🎮 cPDF Converter • 100% In-Memory & Secure • GG WP!"
    }
}

def t(key, **kwargs):
    text = TEXTS[st.session_state.lang].get(key, key)
    if kwargs:
        text = text.format(**kwargs)
    return text


# ---------------------------------------------------------
# 3. CANLI HAREKETLİ ANİMASYONLU ARKA PLAN & GAMING CSS
# ---------------------------------------------------------
ANIMATED_GAMING_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@600;800;900&family=Rajdhani:wght@500;600;700&display=swap');

:root {
    --bg-base: #060a17;
    --neon-green: #00ff66;
    --neon-green-glow: rgba(0, 255, 102, 0.45);
    --neon-cyan: #00f0ff;
    --neon-purple: #9d4edd;
}

/* Streamlit Ana Taşıyıcı ve Hareketli Canlı Arka Plan */
[data-testid="stAppViewContainer"] {
    background: radial-gradient(circle at 15% 20%, rgba(0, 255, 102, 0.14) 0%, transparent 40%),
                radial-gradient(circle at 85% 25%, rgba(0, 240, 255, 0.16) 0%, transparent 45%),
                radial-gradient(circle at 50% 85%, rgba(157, 78, 221, 0.14) 0%, transparent 50%),
                linear-gradient(180deg, #070c1e 0%, #03060f 100%) !important;
    background-size: 180% 180% !important;
    animation: bgMovement 14s ease infinite !important;
    color: #f1f5f9 !important;
    font-family: 'Rajdhani', sans-serif !important;
}

@keyframes bgMovement {
    0% { background-position: 0% 0%; }
    50% { background-position: 100% 100%; }
    100% { background-position: 0% 0%; }
}

/* Sayfa Genişliği Sınırlandırma (Maksimum Konfor) */
.main .block-container {
    max-width: 920px !important;
    padding-top: 2rem !important;
    padding-bottom: 3rem !important;
}

/* Başlık Alanı */
.hero-box {
    text-align: center;
    padding: 10px 10px 18px 10px;
}

.hero-title {
    font-family: 'Orbitron', sans-serif;
    font-size: 2.7rem;
    font-weight: 900;
    color: #ffffff;
    letter-spacing: 2px;
    margin: 0;
    text-shadow: 0 0 15px rgba(0, 255, 102, 0.4), 0 0 35px rgba(0, 255, 102, 0.2);
}

.hero-title span {
    color: var(--neon-green);
    animation: neonPulse 2.5s infinite alternate;
}

@keyframes neonPulse {
    0% { text-shadow: 0 0 10px #00ff66, 0 0 20px #00ff66; }
    100% { text-shadow: 0 0 25px #00ff66, 0 0 50px #00ff66; }
}

.hero-subtitle {
    font-size: 1.15rem;
    font-weight: 600;
    color: var(--neon-cyan);
    margin-top: 6px;
    letter-spacing: 1.2px;
}

.xp-tag {
    display: inline-block;
    margin-top: 10px;
    font-family: 'Orbitron', monospace;
    font-size: 0.78rem;
    font-weight: 800;
    color: #060a17;
    background: linear-gradient(90deg, #00ff66, #00f0ff);
    padding: 4px 16px;
    border-radius: 20px;
    box-shadow: 0 0 18px rgba(0, 255, 102, 0.5);
    letter-spacing: 1px;
}

/* Oyun Temalı Canlı Butonlar */
div.stButton > button:first-child {
    background: linear-gradient(135deg, #0b2f18 0%, #0d1e3d 100%) !important;
    color: #ffffff !important;
    border: 2px solid var(--neon-green) !important;
    border-radius: 12px !important;
    font-family: 'Orbitron', sans-serif !important;
    font-size: 0.95rem !important;
    font-weight: 800 !important;
    letter-spacing: 1.3px !important;
    padding: 14px 24px !important;
    width: 100% !important;
    min-height: 50px;
    box-shadow: 0 0 16px rgba(0, 255, 102, 0.3) !important;
    transition: all 0.2s cubic-bezier(0.175, 0.885, 0.32, 1.275) !important;
    cursor: pointer !important;
}

div.stButton > button:first-child:hover {
    background: linear-gradient(135deg, var(--neon-green) 0%, #00cc52 100%) !important;
    color: #060a17 !important;
    border-color: #ffffff !important;
    box-shadow: 0 0 32px rgba(0, 255, 102, 0.85) !important;
    transform: translateY(-2px) scale(1.01) !important;
}

/* İndirme Butonu */
div.stDownloadButton > button:first-child {
    background: linear-gradient(135deg, #05263d 0%, #034820 100%) !important;
    color: var(--neon-green) !important;
    border: 2px solid var(--neon-green) !important;
    border-radius: 12px !important;
    font-family: 'Orbitron', sans-serif !important;
    font-size: 0.95rem !important;
    font-weight: 800 !important;
    letter-spacing: 1.3px !important;
    padding: 14px 24px !important;
    width: 100% !important;
    min-height: 50px;
    box-shadow: 0 0 18px rgba(0, 255, 102, 0.3) !important;
    transition: all 0.2s ease !important;
}

div.stDownloadButton > button:first-child:hover {
    background: var(--neon-green) !important;
    color: #03060f !important;
    box-shadow: 0 0 35px var(--neon-green) !important;
    transform: translateY(-2px) !important;
}

/* Sekmeler (Tabs) */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: rgba(10, 16, 36, 0.85);
    padding: 8px;
    border-radius: 14px;
    border: 1px solid rgba(0, 255, 102, 0.25);
    margin-bottom: 20px;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 10px;
    color: #94a3b8;
    font-family: 'Rajdhani', sans-serif;
    font-weight: 700;
    font-size: 1.05rem;
    padding: 10px 18px;
    border: none;
    transition: all 0.2s ease;
}

.stTabs [data-baseweb="tab"]:hover {
    color: #ffffff;
    background: rgba(0, 255, 102, 0.1);
}

.stTabs [aria-selected="true"] {
    background: rgba(0, 255, 102, 0.16) !important;
    color: var(--neon-green) !important;
    border: 1px solid var(--neon-green) !important;
    box-shadow: 0 0 16px rgba(0, 255, 102, 0.3) !important;
}

/* Sürükle Bırak Alanı */
[data-testid="stFileUploadDropzone"] {
    background: rgba(10, 18, 38, 0.75) !important;
    border: 2px dashed rgba(0, 255, 102, 0.4) !important;
    border-radius: 14px !important;
    padding: 22px 16px !important;
    transition: all 0.25s ease !important;
}

[data-testid="stFileUploadDropzone"]:hover {
    border-color: var(--neon-green) !important;
    box-shadow: 0 0 20px rgba(0, 255, 102, 0.25) !important;
}

/* Rozet */
.loot-badge {
    display: inline-block;
    background: rgba(0, 255, 102, 0.12);
    color: var(--neon-green);
    border: 1px solid rgba(0, 255, 102, 0.35);
    padding: 6px 14px;
    border-radius: 10px;
    font-size: 0.95rem;
    font-weight: 700;
    margin-bottom: 12px;
}

/* Mobil Görünüm */
@media screen and (max-width: 640px) {
    .hero-title {
        font-size: 1.9rem;
    }
    .hero-subtitle {
        font-size: 0.92rem;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 0.85rem;
        padding: 8px 10px;
        flex: 1 1 45%;
        text-align: center;
    }
}
</style>
"""
st.markdown(ANIMATED_GAMING_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------
# 4. GÜVENİLİR DÖNÜŞTÜRÜCÜ MOTORLARI
# ---------------------------------------------------------
def get_unicode_font():
    """ReportLab için sistemde veya paket içinde bulunan Türkçe/Latin fontu tespit eder."""
    candidates = [
        ("CustomArial", "C:/Windows/Fonts/arial.ttf"),
        ("CustomSegoe", "C:/Windows/Fonts/segoeui.ttf"),
        ("CustomDejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ("CustomDejaVu2", "/usr/share/fonts/dejavu/DejaVuSans.ttf"),
    ]
    try:
        import reportlab
        vera = os.path.join(os.path.dirname(reportlab.__file__), "fonts", "Vera.ttf")
        if os.path.exists(vera):
            candidates.append(("VeraFont", vera))
    except Exception:
        pass

    for font_name, path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, path))
                return font_name
            except Exception:
                continue
    return "Helvetica"


def convert_pdf_to_docx(pdf_bytes):
    """PDF dosyasını Word (.docx) belgesine dönüştürür."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        tmp_pdf.write(pdf_bytes)
        tmp_pdf_path = tmp_pdf.name

    tmp_docx_path = tmp_pdf_path.replace(".pdf", ".docx")
    try:
        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None)
        cv.close()

        with open(tmp_docx_path, "rb") as f:
            docx_bytes = f.read()
        return docx_bytes
    finally:
        if os.path.exists(tmp_pdf_path):
            os.remove(tmp_pdf_path)
        if os.path.exists(tmp_docx_path):
            os.remove(tmp_docx_path)


def merge_pdf_files(pdf_file_list):
    """pypdf'in modern PdfWriter.append metodu ile PDF'leri birleştirir."""
    writer = PdfWriter()
    for item in pdf_file_list:
        writer.append(io.BytesIO(item.getvalue()))
    
    out = io.BytesIO()
    writer.write(out)
    writer.close()
    out.seek(0)
    return out.getvalue()


def convert_images_to_pdf(image_files):
    """Sınırsız sayıda PNG/JPG görselini tek bir PDF albümü yapar."""
    processed = []
    for f in image_files:
        img = Image.open(f)
        if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            bg.paste(img, mask=img.split()[3])
            processed.append(bg)
        else:
            if img.mode != "RGB":
                img = img.convert("RGB")
            processed.append(img)

    if not processed:
        return None

    out = io.BytesIO()
    processed[0].save(
        out,
        format="PDF",
        save_all=True,
        append_images=processed[1:],
        resolution=100.0
    )
    out.seek(0)
    return out.getvalue()


def convert_docx_to_pdf(docx_bytes):
    """Word belgesini saf Python ReportLab motoru ile PDF'e aktarır."""
    doc = Document(io.BytesIO(docx_bytes))
    buf = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)

    font = get_unicode_font()
    styles = getSampleStyleSheet()

    body = ParagraphStyle('FunBody', parent=styles['Normal'], fontName=font, fontSize=10, leading=14, textColor=colors.HexColor('#1e293b'))
    heading = ParagraphStyle('FunHeading', parent=styles['Heading1'], fontName=font, fontSize=13, leading=17, spaceAfter=8, textColor=colors.HexColor('#0f172a'))

    story = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            story.append(Spacer(1, 4))
            continue
        clean = txt.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        if p.style.name.startswith('Heading'):
            story.append(Paragraph(clean, heading))
        else:
            story.append(Paragraph(clean, body))
        story.append(Spacer(1, 4))

    for table in doc.tables:
        t_data = []
        for row in table.rows:
            r_data = []
            for cell in row.cells:
                c_txt = cell.text.strip().replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                r_data.append(Paragraph(c_txt, body))
            t_data.append(r_data)

        if t_data:
            t_obj = Table(t_data)
            t_obj.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e6f9f0')),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#a7f3d0')),
            ]))
            story.append(t_obj)
            story.append(Spacer(1, 10))

    if not story:
        story.append(Paragraph("Belgede metin bulunamadı.", body))

    pdf_doc.build(story)
    buf.seek(0)
    return buf.getvalue()


def convert_docx_to_excel(docx_bytes):
    """Word belgesindeki tabloları ve paragrafları Excel'e dönüştürür."""
    doc = Document(io.BytesIO(docx_bytes))
    tables_found = []
    
    for t_idx, table in enumerate(doc.tables):
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        if rows:
            header = rows[0]
            data = rows[1:] if len(rows) > 1 else []
            clean_hdr = [col if col else f"Sütun_{i+1}" for i, col in enumerate(header)]
            df = pd.DataFrame(data, columns=clean_hdr)
            tables_found.append((f"Tablo_{t_idx+1}", df))

    text_entries = []
    for p_idx, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if txt:
            text_entries.append({"Sıra": p_idx + 1, "Stil": p.style.name, "Metin": txt})
    
    df_texts = pd.DataFrame(text_entries) if text_entries else pd.DataFrame(columns=["Sıra", "Stil", "Metin"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_texts.to_excel(writer, sheet_name="Belge_Metinleri", index=False)
        if tables_found:
            for sname, dft in tables_found:
                dft.to_excel(writer, sheet_name=sname[:31], index=False)
        else:
            pd.DataFrame([{"Not": "Belgede tablo bulunamadı. Metinler listelendi."}]).to_excel(writer, sheet_name="Bilgi", index=False)

    buf.seek(0)
    return buf.getvalue(), tables_found, df_texts


# ---------------------------------------------------------
# 5. EĞLENCELİ BAŞLIK VE HIZLI DİL BUTONU
# ---------------------------------------------------------
col_top_left, col_top_right = st.columns([4, 1])

with col_top_left:
    st.markdown(f"""
    <div class="hero-box">
        <h1 class="hero-title">{t('title_prefix')} <span>{t('title_suffix')}</span></h1>
        <div class="hero-subtitle">{t('subtitle')}</div>
        <div class="xp-tag">{t('xp_badge')}</div>
    </div>
    """, unsafe_allow_html=True)

with col_top_right:
    st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)
    if st.button(t("lang_btn"), key="btn_quick_lang"):
        st.session_state.lang = "EN" if st.session_state.lang == "TR" else "TR"
        st.rerun()


# ---------------------------------------------------------
# 6. DÖNÜŞTÜRÜCÜ SEKMELERİ (TABS)
# ---------------------------------------------------------
tab_pdf2word, tab_merge, tab_to_pdf, tab_word2excel = st.tabs([
    t("tab_pdf2word"),
    t("tab_merge"),
    t("tab_to_pdf"),
    t("tab_word2excel")
])


# =========================================================
# SEKME 1: PDF -> WORD (.DOCX)
# =========================================================
with tab_pdf2word:
    st.markdown(f"<p style='color: #94a3b8; font-size: 1rem; margin-top: 6px;'>{t('t1_desc')}</p>", unsafe_allow_html=True)
    
    up_pdf = st.file_uploader(t("t1_drop"), type=["pdf"], key="up_pdf_word", label_visibility="collapsed")
    
    if up_pdf is not None:
        mb = len(up_pdf.getvalue()) / (1024 * 1024)
        st.markdown(f"<div class='loot-badge'>📄 <b>{up_pdf.name}</b> • {mb:.2f} MB</div>", unsafe_allow_html=True)
        
        if st.button(t("t1_btn"), key="btn_run_pdf2docx"):
            with st.spinner(t("t1_converting")):
                try:
                    docx_data = convert_pdf_to_docx(up_pdf.getvalue())
                    st.balloons()
                    st.success(t("t1_success"))
                    target_name = up_pdf.name.rsplit(".", 1)[0] + ".docx"
                    st.download_button(
                        label=t("t1_download"),
                        data=docx_data,
                        file_name=target_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="dl_docx_fun"
                    )
                except Exception as e:
                    st.error(t("err_generic", err=str(e)))


# =========================================================
# SEKME 2: PDF BİRLEŞTİRİCİ
# =========================================================
with tab_merge:
    st.markdown(f"<p style='color: #94a3b8; font-size: 1rem; margin-top: 6px;'>{t('t2_desc')}</p>", unsafe_allow_html=True)
    
    up_pdfs = st.file_uploader(t("t2_drop"), type=["pdf"], accept_multiple_files=True, key="up_pdfs_merge", label_visibility="collapsed")
    
    if up_pdfs:
        st.markdown(f"<div class='loot-badge'>📚 <b>{len(up_pdfs)} Adet Belge</b> Füzyon Odasında Hazır!</div>", unsafe_allow_html=True)
        out_name = st.text_input(t("t2_out_name"), value="cPDF_Birlestirilmis_Master.pdf" if st.session_state.lang == "TR" else "cPDF_Merged_Master.pdf")
        if not out_name.endswith(".pdf"):
            out_name += ".pdf"
            
        if st.button(t("t2_btn"), key="btn_run_merge"):
            if len(up_pdfs) < 2:
                st.warning(t("t2_warn_min"))
            else:
                with st.spinner(t("t2_merging")):
                    try:
                        merged_data = merge_pdf_files(up_pdfs)
                        st.balloons()
                        st.success(t("t2_success", count=len(up_pdfs)))
                        st.download_button(
                            label=t("t2_download"),
                            data=merged_data,
                            file_name=out_name,
                            mime="application/pdf",
                            key="dl_merged_fun"
                        )
                    except Exception as e:
                        st.error(t("err_generic", err=str(e)))


# =========================================================
# SEKME 3: FARKLI FORMATLARI PDF'E DÖNÜŞTÜR
# (SINIRSIZ PNG/JPG VEYA WORD DOSYALARI)
# =========================================================
with tab_to_pdf:
    st.markdown(f"<p style='color: #94a3b8; font-size: 1rem; margin-top: 6px;'>{t('t3_desc')}</p>", unsafe_allow_html=True)
    
    sub_mode = st.radio("Seçim:", [t("t3_mode_img"), t("t3_mode_word")], horizontal=True, label_visibility="collapsed")
    
    if sub_mode == t("t3_mode_img"):
        up_imgs = st.file_uploader(t("t3_img_drop"), type=["png", "jpg", "jpeg", "webp", "bmp"], accept_multiple_files=True, key="up_imgs_pdf", label_visibility="collapsed")
        
        if up_imgs:
            st.markdown(f"<div class='loot-badge'>{t('t3_img_info', count=len(up_imgs))}</div>", unsafe_allow_html=True)
            
            with st.expander(f"🔍 Cephanelik Önizlemesi ({len(up_imgs)} Görsel)"):
                cols = st.columns(min(len(up_imgs), 4) if up_imgs else 1)
                for i, im in enumerate(up_imgs[:8]):
                    cols[i % 4].image(im, caption=im.name[:14], use_container_width=True)
                if len(up_imgs) > 8:
                    st.caption(f"... ve {len(up_imgs) - 8} adet daha görsel var!")

            if st.button(t("t3_img_btn"), key="btn_run_imgs_pdf"):
                with st.spinner(t("t3_img_making")):
                    try:
                        pdf_data = convert_images_to_pdf(up_imgs)
                        if pdf_data:
                            st.balloons()
                            st.success(t("t3_img_success", count=len(up_imgs)))
                            st.download_button(
                                label=t("t3_img_download"),
                                data=pdf_data,
                                file_name="cPDF_Gorsel_Paketi.pdf" if st.session_state.lang == "TR" else "cPDF_Image_Bundle.pdf",
                                mime="application/pdf",
                                key="dl_imgs_pdf_fun"
                            )
                    except Exception as e:
                        st.error(t("err_generic", err=str(e)))
    else:
        up_docx_pdf = st.file_uploader(t("t3_word_drop"), type=["docx"], key="up_word_pdf", label_visibility="collapsed")
        
        if up_docx_pdf is not None:
            st.markdown(f"<div class='loot-badge'>📝 <b>{up_docx_pdf.name}</b> Algılandı</div>", unsafe_allow_html=True)
            
            if st.button(t("t3_word_btn"), key="btn_run_word_pdf"):
                with st.spinner(t("t3_word_making")):
                    try:
                        word_pdf_data = convert_docx_to_pdf(up_docx_pdf.getvalue())
                        st.balloons()
                        st.success(t("t3_word_success"))
                        out_pdf_name = up_docx_pdf.name.rsplit(".", 1)[0] + ".pdf"
                        st.download_button(
                            label=t("t3_word_download"),
                            data=word_pdf_data,
                            file_name=out_pdf_name,
                            mime="application/pdf",
                            key="dl_word_pdf_fun"
                        )
                    except Exception as e:
                        st.error(t("err_generic", err=str(e)))


# =========================================================
# SEKME 4: WORD'DEN EXCEL'E (.DOCX -> .XLSX)
# =========================================================
with tab_word2excel:
    st.markdown(f"<p style='color: #94a3b8; font-size: 1rem; margin-top: 6px;'>{t('t4_desc')}</p>", unsafe_allow_html=True)
    
    up_docx_xl = st.file_uploader(t("t4_drop"), type=["docx"], key="up_docx_xl", label_visibility="collapsed")
    
    if up_docx_xl is not None:
        try:
            d_obj = Document(io.BytesIO(up_docx_xl.getvalue()))
            t_cnt = len(d_obj.tables)
            p_cnt = len([p for p in d_obj.paragraphs if p.text.strip()])
        except Exception:
            t_cnt, p_cnt = 0, 0
            
        st.markdown(f"<div class='loot-badge'>{t('t4_info', tables=t_cnt, paras=p_cnt)}</div>", unsafe_allow_html=True)
        
        if st.button(t("t4_btn"), key="btn_run_word_xl"):
            with st.spinner(t("t4_converting")):
                try:
                    xl_data, tables_list, _ = convert_docx_to_excel(up_docx_xl.getvalue())
                    st.balloons()
                    st.success(t("t4_success"))
                    
                    if tables_list:
                        st.markdown(f"**{t('t4_preview_title')}**")
                        st.dataframe(tables_list[0][1], use_container_width=True)
                        
                    out_xl_name = up_docx_xl.name.rsplit(".", 1)[0] + ".xlsx"
                    st.download_button(
                        label=t("t4_download"),
                        data=xl_data,
                        file_name=out_xl_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_docx_xl_fun"
                    )
                except Exception as e:
                    st.error(t("err_generic", err=str(e)))


# ---------------------------------------------------------
# 7. EĞLENCELİ ALT BİLGİ (FOOTER)
# ---------------------------------------------------------
st.markdown("<br><hr style='border: none; border-top: 1px solid rgba(0,255,102,0.15); margin-top: 30px;'>", unsafe_allow_html=True)
st.markdown(f"<div style='text-align: center; color: #64748b; font-size: 0.9rem;'>{t('footer')}</div>", unsafe_allow_html=True)
