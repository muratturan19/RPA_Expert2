# Preston RPA Coordinate Map
**Full HD (1920x1080) Tam Ekran Koordinatları**

## Test-Only Steps (Already Active)
| Step | Element | Coordinates | Description |
|------|---------|-------------|-------------|
| 1 | Finans - İzle Menu | (529, 190) | Menu click test |
| 2 | Dropdown: Banka hesap izleme | (529, 210) | Dropdown test |
| 3 | BANKA Toolbar Icon | (18, 219) | Icon test |

## Real Workflow Steps
| Step | Element | Coordinates | Description | Data Source |
|------|---------|-------------|-------------|-------------|
| 4 | Hesap Search "..." | (290, 305) | Account search button | - |
| 5 | Account Selection | (600, 532) | Select: 6293986 | excel: hesap_no |
| 6a | First Tamam | (1163, 820) | First confirmation | - |
| 6b | Second Tamam | (1163, 820) | Final confirmation | - |
| 7a | Date Input | (197, 335) | Start/End date | excel: tarih |
| 7b | Yenile Button | (48, 399) | Refresh data | - |
| 8 | Yeni Button | (223, 448) | New transaction | - |
| 9 | Havale Alma | (848, 566) | Transaction type | - |
| 10 | Yeni Belge | (form) | New document | auto: increment |
| 11a | Belge Tarihi | (form) | Document date | excel: tarih |
| 11b | Valör Tarihi | (form) | Value date | excel: tarih |
| 12 | Banka Kodu | (848, 629) | Bank code | excel: banka_kodu |
| 13 | Cari Kodu | (848, 661) | Client code | excel: cari_kodu |
| 14 | Tutar | (form) | Amount | excel: tutar |
| 15 | Açıklama | (form) | Description | excel: aciklama |
| 16 | Kaydet | (form) | Save | - |
| 17 | Kapat | (form) | Close | - |

## Excel Data Mapping
```python
excel_data = {
    "hesap_no": "6293986",
    "tarih": "25.07.2025", 
    "banka_kodu": "062",
    "cari_kodu": "120.12.001",
    "tutar": "1250.75",
    "aciklama": "POS Tahsilatı"
}
```
