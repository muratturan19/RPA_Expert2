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
            # Real workflow coordinates
            "hesap_search": (290, 305),
            "hesap_input": (1000, 497),
            "account_item": (690, 432),  # mavi highlight 6293986 row
            "tamam_button": (1163, 820),
            "date_input": (197, 335),
            "yenile_btn": (48, 399),
            "yeni_btn": (223, 448),
            "havale_alma": (695, 385),  # EFT'nin biraz üstü (Havale alma)
            "banka_field": (848, 629),
            "cari_field": (848, 661),
            "save_button": (930, 615),
            "close_button": (883, 615),
            "belge_tarih": (611, 462),
            "valor_tarih": (866, 462),
            "tutar_field": (635, 509),
            "aciklama_field": (635, 533),
            "yeni_belge_btn": (760, 383),
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

            # Step 5: Excel'den gelen hesap no'yu input'a yaz
            pyautogui.click(*self.coordinates["hesap_input"])
            pyautogui.typewrite(excel_data["hesap_no"])  # Excel'den tam hesap no
            time.sleep(1.5)  # Filter'ın çalışmasını bekle

            # Step 6: Filtered sonuçtaki tek seçeneği tıkla
            pyautogui.click(*self.coordinates["account_item"])  # Liste itemını seçili yap
            time.sleep(0.5)

            # Step 7a: İlk Tamam
            pyautogui.click(*self.coordinates["tamam_button"])
            time.sleep(0.5)

            # Step 7b: İkinci Tamam
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
        # Step 10: Click Yeni belge
        pyautogui.click(*self.coordinates["yeni_belge_btn"])
        time.sleep(0.5)

        # Step 11a: Belge tarihi
        pyautogui.click(*self.coordinates["belge_tarih"])
        pyautogui.typewrite(data["tarih"])

        # Step 11b: Valör tarihi
        pyautogui.click(*self.coordinates["valor_tarih"])
        pyautogui.typewrite(data["tarih"])

        # Step 12: Banka kodu
        pyautogui.click(*self.coordinates["banka_field"])
        pyautogui.typewrite(data["banka_kodu"])

        # Step 13: Cari kodu
        pyautogui.click(*self.coordinates["cari_field"])
        pyautogui.typewrite(data["cari_kodu"])

        # Step 14: Tutar input
        pyautogui.click(*self.coordinates["tutar_field"])
        pyautogui.typewrite(data["tutar"])

        # Step 15: Açıklama
        pyautogui.click(*self.coordinates["aciklama_field"])
        pyautogui.typewrite(data["aciklama"])

    def click_save(self) -> None:
        """Trigger save action."""
        pyautogui.click(*self.coordinates["save_button"])

    def click_close(self) -> None:
        """Close the current window."""
        pyautogui.click(*self.coordinates["close_button"])
