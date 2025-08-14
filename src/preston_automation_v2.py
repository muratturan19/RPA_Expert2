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
            "account_item": (600, 532),
            "tamam_button": (1163, 820),
            "date_input": (197, 335),
            "yenile_btn": (48, 399),
            "yeni_btn": (223, 448),
            "havale_alma": (848, 566),
            "banka_field": (848, 629),
            "cari_field": (848, 661),
        }

    def execute_real_workflow(self, excel_data: dict[str, str]) -> bool:
        """Execute steps 4-17 using values from Excel data."""
        try:
            # Step 4: Click hesap search
            pyautogui.click(*self.coordinates["hesap_search"])
            time.sleep(1)

            # Step 5: Select account
            pyautogui.click(*self.coordinates["account_item"])
            time.sleep(0.5)

            # Step 6a: First Tamam
            pyautogui.click(*self.coordinates["tamam_button"])
            time.sleep(0.5)

            # Step 6b: Second Tamam
            pyautogui.click(*self.coordinates["tamam_button"])
            time.sleep(1)

            # Step 7a: Fill date
            pyautogui.click(*self.coordinates["date_input"])
            pyautogui.typewrite(excel_data["tarih"])
            time.sleep(0.5)

            # Step 7b: Click Yenile
            pyautogui.click(*self.coordinates["yenile_btn"])
            time.sleep(2)

            # Step 8: Click Yeni
            pyautogui.click(*self.coordinates["yeni_btn"])
            time.sleep(1)

            # Step 9: Select Havale alma
            pyautogui.click(*self.coordinates["havale_alma"])
            time.sleep(1)

            # Steps 10-15: Fill form with Excel data
            self.fill_transaction_form(excel_data)

            # Step 16: Save
            self.click_save()

            # Step 17: Close
            self.click_close()

            return True
        except Exception as exc:  # pragma: no cover - runtime safeguard
            print(f"Workflow error: {exc}")
            return False

    def fill_transaction_form(self, data: dict[str, str]) -> None:
        """Fill transaction form fields using Excel data."""
        # Step 12: Banka kodu
        pyautogui.click(*self.coordinates["banka_field"])
        pyautogui.typewrite(data["banka_kodu"])

        # Step 13: Cari kodu
        pyautogui.click(*self.coordinates["cari_field"])
        pyautogui.typewrite(data["cari_kodu"])

        # Step 14: Tutar (use Tab navigation)
        pyautogui.press("tab")
        pyautogui.typewrite(data["tutar"])

        # Step 15: Açıklama
        pyautogui.press("tab")
        pyautogui.typewrite(data["aciklama"])

    def click_save(self) -> None:
        """Trigger save action."""
        pyautogui.hotkey("ctrl", "s")

    def click_close(self) -> None:
        """Close the current window."""
        pyautogui.hotkey("alt", "f4")
