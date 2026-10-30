# İlçe PDF Birleştirici

> İlçe, kurum ve şube klasörlerindeki PDF dosyalarını sıralı biçimde birleştiren; PDF, Excel ve metin raporları üreten Windows masaüstü uygulaması.

[![Python](https://img.shields.io/badge/Python-3.12%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Windows EXE](https://img.shields.io/badge/Windows-EXE-0078D6?style=flat-square&logo=windows&logoColor=white)](https://github.com/cmtrkr83/pdf-birlestirici/actions)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue?style=flat-square)](LICENSE)

<p align="center">
  <a href="https://www.buymeacoffee.com/cmtrkr">
    <img src="https://img.shields.io/badge/Buy%20Me%20a%20Coffee-FFDD00?style=for-the-badge&logo=buymeacoffee&logoColor=black" alt="Buy Me a Coffee">
  </a>
</p>

## ✨ Özellikler

- 📁 Seçilen ana klasörü ve tüm alt klasörlerini özyinelemeli olarak tarar.
- 📄 Her PDF içeren klasördeki dosyaları doğal isim sırasına göre birleştirir.
- 🏫 Çıktı klasör yapısını koruyarak klasör bazlı PDF’ler üretir.
- 🧩 Oluşturulan PDF’leri tek bir ilçe PDF’sinde birleştirir.
- 📊 İlçe özeti, okul/kurum detayı ve şube detayı içeren `rapor.xlsx` oluşturur.
- 📝 Hızlı paylaşım için `rapor.txt` üretir.
- 🔢 İlk iki sayfayı sınıf listesi, sonraki her iki sayfayı bir öğrenci olarak sayar.
- ⚠️ Okunamayan PDF’leri atlar ve hataları rapora ekler.
- 🖥️ Windows’ta kurulum gerektirmeyen tek dosyalık EXE olarak çalışır.

## 📂 Klasör yapısı

Programın seçilen klasörü aşağıdaki gibi olabilir:

```text
Ilce/
├── Kurum 1/
│   ├── Sube A/
│   │   ├── 1.pdf
│   │   └── 2.pdf
│   └── Sube B/
│       └── sube.pdf
└── Kurum 2/
    └── sube.pdf
```

Program, içinde PDF bulunan her klasörü tarar. Örnek çıktı:

```text
Ilce/
└── Birlestirilmis_PDFler/
    ├── Kurum 1/
    │   ├── Sube A.pdf
    │   └── Sube B.pdf
    ├── Kurum 2.pdf
    ├── Ilce_Birlestirilmis.pdf
    ├── rapor.xlsx
    └── rapor.txt
```

`Birlestirilmis_PDFler` klasörü tekrar girdi olarak taranmaz. Programı yeniden
çalıştırdığınızda mevcut çıktı klasörü, onayınız alındıktan sonra yenilenir.

## 🚀 Kullanım

### Hazır Windows EXE

1. [Actions sayfasından](https://github.com/cmtrkr83/pdf-birlestirici/actions) başarılı bir `Windows EXE` çalışmasının artifact’ini indirin.
2. ZIP dosyasını açın.
3. `Ilce_PDF_Birlestirici.exe` dosyasını çalıştırın.
4. En üstteki ilçe klasörünü seçin ve **Birleştir ve Raporla** düğmesine basın.

### Kaynaktan çalıştırma

Python 3.12 veya üzeri ve gerekli paketler:

```bash
python -m pip install pypdf openpyxl
python pdf_birlestirici.py
```

## 🛠️ Windows EXE derleme

PyInstaller yapılandırması [`Ilce_PDF_Birlestirici.spec`](Ilce_PDF_Birlestirici.spec)
dosyasındadır. GitHub Actions, `master` dalına yapılan her push’ta veya elle
başlatıldığında tek dosyalık Windows EXE üretir.

```bash
pip install pypdf openpyxl pyinstaller
pyinstaller --noconfirm --clean Ilce_PDF_Birlestirici.spec
```

Çıktı `dist/Ilce_PDF_Birlestirici.exe` konumunda oluşur.

## 🔢 Öğrenci sayımı

Varsayılan sayım kuralları:

- İlk 2 sayfa: sınıf listesi
- Sonraki her 2 sayfa: 1 öğrenci
- Eksik kalan tek optik sayfa: 1 öğrenci olarak sayılır

Bu değerler [`pdf_birlestirici.py`](pdf_birlestirici.py) içindeki
`CLASS_LIST_PAGES` ve `STUDENT_PAGES` sabitlerinden değiştirilebilir.

## 🤝 Katkı

Hata bildirimleri ve geliştirme önerileri için
[Issues](https://github.com/cmtrkr83/pdf-birlestirici/issues) bölümünü kullanabilirsiniz.

## ☕ Destek

Bu proje işinize yararsa geliştiriciye bir kahve ısmarlayabilirsiniz:

<p align="center">
  <a href="https://www.buymeacoffee.com/cmtrkr">
    <img src="https://img.shields.io/badge/Buy%20Me%20a%20Coffee-FFDD00?style=for-the-badge&logo=buymeacoffee&logoColor=black" alt="Buy Me a Coffee">
  </a>
</p>

## 📄 Lisans

Bu proje [MIT Lisansı](LICENSE) ile lisanslanmıştır.

© 2026 Cem Türker
