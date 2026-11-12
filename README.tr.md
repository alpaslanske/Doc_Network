# PDF Network: Centralized Bookmarks for PDF Documents
[![License](https://img.shields.io/badge/license-GNU/GPLv3-blue)](LICENSE)
[![Python3](https://img.shields.io/badge/Python-3.x-red)](https://github.com/alpaslanske/Doc_Network)
[![Made In Turkey](https://github.com/alpaslanske/PDF_Network/blob/main/.assets/madeinturkiye.svg)](https://goturkiye.com/)
[![Docs](https://img.shields.io/badge/docs-online-brightgreen)](https://github.com/alpaslanske/Doc_Network/tree/main/docs)
# 📚 Doc Network

[🇬🇧 English](README.md) | [🇹🇷 Türkçe](README.tr.md)

**Doc Network**, PDF, Ofis belgeleri ve e-kitaplar arasında **çapraz bağlantılar (`[[dosya#sayfa]]`)** kurmanızı, hiyerarşik yer imi grupları oluşturmanızı ve tüm kütüphanenizde tam metin arama yapmanızı sağlayan modern, çok platformlu bir belge görüntüleyici ve kişisel bilgi tabanı (Knowledge Base) uygulamasıdır.

---

## ✨ Öne Çıkan Özellikler

* **Geniş Dosya Formatı Desteği:**
  * **PDF ve E-Kitaplar:** `.pdf`, `.epub`, `.mobi`, `.fb2`, `.cbz`
  * **Ofis Belgeleri (Akıllı Önbellek ile):** `.docx`, `.odt`, `.xlsx`, `.ods`, `.pptx`, `.odp`, `.odg`, `.odf`, `.rtf`, `.csv`
* **Çapraz Belge Bağlantıları (`[[dosya#sayfa]]`):**
  * Aldığınız çalışma notlarının içine farklı bir belgenin belirli bir sayfasına giden tıklanabilir bağlantılar ekleyin.
  * Tıkladığınızda hedef belge yeni sekmede açılır ve doğrudan ilgili sayfaya odaklanır.
* **Tarayıcı Tarzı Gezinme Geçmişi (`◀` / `▶`):**
  * Çapraz bağlantılar, yer imleri, içindekiler veya arama sonuçları arasında atlarken bulunduğunuz konum ve sayfa otomatik kaydedilir; tek tıkla geri ve ileri gidebilirsiniz.
* **Çift Modlu Akıllı Arama (Asenkron):**
  * **Açık Belge Modu:** Aktif sekmedeki belgede arka planda arayüzü dondurmadan arama yapar ve eşleşmeleri sayfa üzerinde vurgular.
  * **Tüm Kütüphane & Notlar Modu:** Veritabanındaki tüm yer imi başlıklarında, çalışma notlarında ve kütüphanenizdeki PDF dosyalarının tam metinlerinde arama yapar.
* **PDF Düzenleme ve Veri Çıkarma:**
  * Seçili alanları sarı vurgu (Highlight) ile işaretleme ve yapışkan sayfa notları ekleme.
  * Seçilen tablo alanlarını tek tıkla **Excel (`.xlsx`)** dosyasına aktarma.
  * Taranmış belgelerden veya görsellerden **OCR (Tesseract)** ile metin kopyalama (Türkçe, Lehçe ve İngilizce dil desteği).
  * Yer imlerini ve notları düzenli bir **Markdown (`.md`)** çalışma özeti olarak dışa aktarma.
* **Çoklu Sekme ve Çalışma Panosu (Dashboard):**
  * Son açılan belgeleri sayfa bilgileriyle takip edin ve birden fazla belgeyi sekmeler halinde aynı anda inceleyin.
  * Türkçe (`🇹🇷 TR`) ve İngilizce (`🇬🇧 EN`) arayüz dil desteği.

---

## 🚀 Kurulum ve Çalıştırma

### 1. Hazır Taşınabilir Sürümler (Önerilen)
Projenin **Releases** sayfasından işletim sisteminize uygun paketi indirip kurulum yapmadan doğrudan çalıştırabilirsiniz:
* **Linux:** `Doc_Network_Portable-x86_64.AppImage` (Çalıştırma izni verip başlatın: `chmod +x Doc_Network_Portable-x86_64.AppImage`)
* **Windows:** `Doc_Network_Portable.exe`

### 2. Kaynak Koddan Çalıştırma

1. Depoyu bilgisayarınıza klonlayın:
   ```bash
   git clone https://github.com
   cd Doc_Network
   ```

2. Gerekli Python kütüphanelerini yükleyin:
   ```bash
   pip install -r requirements.txt
   ```

3. Uygulamayı başlatın:
   ```bash
   python3 doc_network.py
   ```

> **Not (İsteğe Bağlı Sistem Araçları):**
> * Ofis belgelerini (`.docx`, `.xlsx`, `.odt` vb.) görüntülemek için sisteminizde **LibreOffice (`soffice`)** kurulu olmalıdır.
> * Görsel üzerinden OCR ile metin tanıma özelliğini kullanmak için sisteminizde **Tesseract OCR** (`tur`, `eng`, `pol` dil paketleriyle) kurulu olmalıdır.

---

## 🖱️ Hızlı Kullanım İpuçları

* **Sağ Tık ile Kaydırma (Pan):** Belge üzerinde sağ fare tuşunu basılı tutarak sayfayı sürükleyebilirsiniz.
* **Sol Tık ile Alan Seçimi:** Metin kopyalamak, yer imi oluşturmak, vurgu eklemek veya Excel'e tablo aktarmak için sol tuşla alan seçip sağ tıklayın.
* **Yakınlaştırma:** `Ctrl + Fare Tekerleği` ile belgeyi yakınlaştırıp uzaklaştırabilirsiniz.
* **Not Düzenleme:** Sol alt paneldeki not alanına çift tıklayarak notunuzu düzenleyebilir ve `[[dosya_adi.pdf#15]]` formatında bağlantılar ekleyebilirsiniz.

---

## 🛠️ Kullanılan Teknolojiler

* **Arayüz:** Python 3, PySide6 (Qt6)
* **Belge İşleme:** PyMuPDF (`fitz`), LibreOffice Headless
* **Veritabanı:** SQLite3
* **OCR ve Veri Aktarımı:** Tesseract OCR, Pillow, Pandas, OpenPyXL
