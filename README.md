# 🎮 cPDF Converter | Gamer Edition

**cPDF Converter**, modern koyu lacivert tonları ve neon yeşil detaylarla tasarlanmış, oyun/arcade estetiğine sahip, **mobil uyumlu (responsive)** ve **Türkçe/İngilizce çift dil destekli** yüksek performanslı bir belge ve format dönüştürme web uygulamasıdır.

Python ve **Streamlit** kullanılarak geliştirilmiş olup, Render veya benzeri bulut platformlarında tek tıkla canlıya alınabilecek şekilde optimize edilmiştir.

---

## 🌐 Çift Dil (TR / EN) ve Mobil Uyumluluk

- **🇹🇷 Türkçe & 🇬🇧 İngilizce Desteği:**
  - Yan menüden (Sidebar) veya mobilde üst kısımdaki hızlı dil butonundan tek tıkla Türkçe ve İngilizce arasında geçiş yapılabilir.
  - Tüm arayüz metinleri, butonlar, hata/başarı bildirimleri ve rehberler seçilen dile dinamik olarak uyum sağlar.
- **📱 %100 Mobil ve Tablet Uyumluluğu:**
  - Özel CSS medya sorguları (`@media screen and (max-width: 768px)`) ile akıllı telefon ve tablet ekranlarına tam uyum.
  - Dokunmatik ekran dostu minimum 44-48px buton yükseklikleri, esnek taşma (wrap) yapan sekmeler ve küçülen neon başlıklar.
  - Mobil cihazlarda sidebar kapalıyken bile üst bardan kolay dil değişimi.

---

## ✨ Temel Özellikler

1. **📄 PDF'ten Word'e (.docx) Dönüştürme:**
   - Arka planda `pdf2docx` motoru kullanılır.
   - Sayfa düzeni, tablolar ve biçimlendirmeleri koruyarak tam düzenlenebilir Word belgesi üretir.
   - Dosya boyutu ve sayfa sayısı gibi ön özet metrikleri sunar.

2. **📚 Çoklu PDF Birleştirici (PDF Merger):**
   - Birden fazla PDF dosyasını aynı anda yükleyin.
   - Yüklenen belgelerin sayfa sayılarını ve boyutlarını inceleyin.
   - Tüm belgeleri tek bir birleşik PDF dosyası haline getirin.

3. **🖼️ Farklı Formatları PDF'e Dönüştürme:**
   - **Sınırsız PNG / JPG Desteği:** Yükleme limiti olmaksızın onlarca veya yüzlerce PNG, JPG, JPEG, WEBP görselini tek bir PDF albümü halinde birleştirin. Şeffaf (RGBA) PNG'ler beyaz arkaplanla temiz RGB'ye dönüştürülür.
   - **Word (.docx) -> PDF:** Microsoft Word belgelerini platform bağımsız saf Python motoru (`reportlab` + `python-docx`) ile PDF formatına dönüştürür. Render/Linux sunucularda Office kurulu olmasa bile kusursuz çalışır.

4. **📊 Word'den Excel'e (.docx -> .xlsx) Dönüştürme:**
   - Word belgesindeki gömülü tabloları otomatik tespit eder.
   - Her tabloyu ayrı bir Excel çalışma sayfasına (`Tablo_1`, `Tablo_2` vb.) aktarır.
   - Belgedeki tüm metin paragraflarını `Belge_Metinleri` sayfasına düzenli olarak yerleştirir.
   - Uygulama içerisinde tablo önizlemesi sunar.

---

## 🎨 Tema & Tasarım

- **Renk Paleti:** Gece laciverti (`#070b19`, `#0d1730`) ve neon yeşil (`#00ff66`, `#39ff14`) detaylar.
- **Oyun Temalı Butonlar:** Neon ışıltılı kenarlıklar, hover efektleri ve oyun içi aksiyon başlıkları (`DÖNÜŞTÜRÜCÜYÜ ATEŞLE`, `GÜÇLERİ BİRLEŞTİR`, `GÖREV ÖDÜLÜNÜ İNDİR`).

---

## 🚀 Yerel Olarak Çalıştırma

Projeyi bilgisayarınızda yerel olarak çalıştırmak için:

1. **Gereksinimleri yükleyin:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Uygulamayı başlatın:**
   ```bash
   streamlit run app.py
   ```

Tarayıcınız otomatik olarak `http://localhost:8501` adresinde açılacaktır.

---

## ☁️ Render'da Canlıya Alma (Deployment Rehberi)

1. Projeyi GitHub veya GitLab üzerinde bir depoya (repository) push edin.
2. [Render Dashboard](https://dashboard.render.com/)'a gidin ve **New +** -> **Web Service** seçeneğine tıklayın.
3. GitHub reponuzu bağlayın.
4. Aşağıdaki ayarları uygulayın:
   - **Name:** `cpdf-converter`
   - **Environment:** `Python 3`
   - **Branch:** `main`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
5. **Create Web Service** butonuna basın. Birkaç dakika içinde uygulamanız canlıya alınacaktır!
