# Preston RPA Üretim Sistemi

Bu depo, Preston V2 web simülatörü üzerinde gerçek dünyadaki bankacılık işlemlerini otomatikleştirmek için geliştirilmiş bir RPA çözümü içerir. Koordinat eşleme, şablon eşleme ve Türkçe OCR kullanarak menüleri tanıyarak işlemleri gerçekleştirir.

## Özellikler
- Streamlit tabanlı web arayüzü
- Excel dosyalarından veri çıkarma ve koordinat formatına dönüştürme
- Tesseract ile Türkçe metin tanıma
- Hibrit element tespitine sahip yeni v3 motoru
- Esnek loglama ve hata yönetimi

## Gereksinimler
- Python 3.10 veya üzeri
- Tesseract OCR ve `tur` dil paketi
- `pip` ile kurulabilen Python bağımlılıkları:
  - `streamlit`
  - `openpyxl`
  - `pyautogui`, `pyperclip`
  - `pytesseract`, `opencv-python`, `numpy`
  - `Pillow`, `pyyaml`

## Kurulum
```bash
git clone <repo>
cd RPA_Expert2
pip install -r requirements.txt
```
Tesseract'ı sisteminize kurduktan sonra `OCR_LANGUAGE` ayarının `tur` olduğundan emin olun.

## Kullanım
```bash
streamlit run main.py
```
Web arayüzü açıldığında bir Excel dosyası yükleyip "Start RPA" düğmesine basarak otomasyonu başlatabilirsiniz.

## Testler
Kod değişikliklerini doğrulamak için:
```bash
pytest
```

## Lisans
Bu proje [MIT Lisansı](LICENSE) ile lisanslanmıştır.
