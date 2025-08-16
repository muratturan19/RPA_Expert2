"""Coordinate-based Preston automation implementation."""

import time
from pathlib import Path

import pyautogui


class PrestonRPAV2:
    """Automation workflow that uses predefined screen coordinates."""

    def __init__(self) -> None:
        self.coordinates = self.load_coordinates()

    def load_coordinates(self) -> dict[str, tuple[int, int]]:
        """Load coordinates from mapping."""
        return {
            # Mevcut koordinatlar... (dokunma)
            "hesap_search": (290, 305),
            "hesap_input": (1000, 497),
            "account_item": (990, 532),  # mavi highlight area
            "tamam_button": (1163, 820),
            "date_input": (197, 335),
            "yenile_btn": (48, 399),
            "yeni_btn": (223, 448),
            "havale_alma": (955, 566),  # ✅ Gerçek "Havale alma" koordinatı

            # ✅ YENİ - Direkt input field'lar:
            "banka_input": (629, 277),        # Banka input field (062 yazan yer)
            "cari_input": (629, 309),         # Cari input field (120. yazan yer)
            "belge_tarih_field": (629, 442),  # Belge Tarihi
            "valor_tarih_field": (873, 442),  # Valör Tarihi
            "tutar_input": (629, 514),        # Tutar
            "aciklama_input": (629, 538),     # Açıklama
            "kaydet_btn": (919, 589),         # Kaydet
            "kapat_btn": (840, 589),          # Kapat
        }

    def dismiss_alerts(self) -> None:
        """Attempt to close any blocking JavaScript alerts."""
        try:
            pyautogui.press("esc")
            time.sleep(0.5)
            pyautogui.press("enter")
            time.sleep(0.5)
            pyautogui.press("tab")
            pyautogui.press("enter")
            time.sleep(0.5)
        except Exception:
            pass

    def execute_real_workflow(self, excel_data: dict[str, str]) -> bool:
        """Execute steps 4-17 using values from Excel data."""
        try:
            # Clear potential blocking alerts
            self.dismiss_alerts()

            # Step 4: Click hesap search
            pyautogui.click(*self.coordinates["hesap_search"])
            time.sleep(2)

            # Step 5: Excel'den hesap no yaz
            pyautogui.click(*self.coordinates["hesap_input"])
            pyautogui.typewrite(excel_data["hesap_no"])  # Excel'den tam hesap no
            time.sleep(1.5)  # Filter'ın çalışmasını bekle

            # Step 6: Mavi highlight olan item'a tıkla (seçili yap)
            pyautogui.click(*self.coordinates["account_item"])
            time.sleep(0.5)

            # Step 7a: İlk Tamam
            pyautogui.click(*self.coordinates["tamam_button"])
            time.sleep(0.5)

            # Step 7b: İkinci Tamam (double confirmation)
            pyautogui.click(*self.coordinates["tamam_button"])
            time.sleep(1)

            # Step 8: Fill date
            pyautogui.click(*self.coordinates["date_input"])
            pyautogui.typewrite(excel_data["tarih"])
            time.sleep(0.5)

            # Step 9: Click Yenile
            pyautogui.click(*self.coordinates["yenile_btn"])
            time.sleep(2)

            # Step 10: Click Yeni
            pyautogui.click(*self.coordinates["yeni_btn"])
            time.sleep(1)

            # Step 11: Select Havale alma
            pyautogui.click(*self.coordinates["havale_alma"])
            time.sleep(1)

            # Steps 12-17: Fill form with Excel data
            self.fill_transaction_form(excel_data)

            # Step 18: Save
            self.click_save()

            # Step 19: Close
            self.click_close()

            return True
        except Exception as exc:  # pragma: no cover - runtime safeguard
            print(f"Workflow error: {exc}")
            return False

    def fill_transaction_form(self, data: dict[str, str]) -> None:
        """Fill transaction form fields using Excel data."""

        # Step 13: Banka kodu - direkt input'a yaz
        pyautogui.click(*self.coordinates["banka_input"])
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "a")  # Mevcut text'i seç
        pyautogui.typewrite("062")  # Test banka kodu
        time.sleep(0.5)

        # Step 14: Cari kodu - direkt input'a yaz
        pyautogui.click(*self.coordinates["cari_input"])
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "a")  # Mevcut text'i seç
        pyautogui.typewrite("120.12.001")  # Test cari kodu
        time.sleep(0.5)

        # Step 15: Belge tarihi
        pyautogui.click(*self.coordinates["belge_tarih_field"])
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data["tarih"])
        time.sleep(0.5)

        # Step 16: Valör tarihi
        pyautogui.click(*self.coordinates["valor_tarih_field"])
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data["tarih"])
        time.sleep(0.5)

        # Step 17: Tutar
        pyautogui.click(*self.coordinates["tutar_input"])
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data["tutar"])
        time.sleep(0.5)

        # Step 18: Açıklama
        pyautogui.click(*self.coordinates["aciklama_input"])
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data.get("aciklama", "Test işlemi"))
        time.sleep(0.5)

    def click_save(self) -> None:
        """Trigger save action."""
        pyautogui.click(*self.coordinates["kaydet_btn"])

    def click_close(self) -> None:
        """Close the current window."""
        pyautogui.click(*self.coordinates["kapat_btn"])
