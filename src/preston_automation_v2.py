"""Coordinate-based Preston automation implementation."""

from __future__ import annotations

import json
import time
from pathlib import Path

import pyautogui

from config import FORM_FILL_DELAY
from logger import get_logger

logger = get_logger(__name__)


class PrestonRPAV2:
    """Automation workflow that uses predefined screen coordinates."""

    def __init__(self, coord_file: str | Path | None = None) -> None:
        self.coordinate_file = (
            Path(coord_file)
            if coord_file
            else Path(__file__).with_name("coordinates.json")
        )
        self.coordinates = self.load_coordinates(self.coordinate_file)

    def load_coordinates(self, path: Path | str | None = None) -> dict[str, tuple[int, int]]:
        """Load coordinates from mapping file.

        Falls back to :meth:`default_coordinates` if file is missing.
        """
        coord_path = Path(path) if path else self.coordinate_file
        try:
            if coord_path.suffix in {".yaml", ".yml"}:
                import yaml  # type: ignore

                data = yaml.safe_load(coord_path.read_text(encoding="utf-8")) or {}
            else:
                data = json.loads(coord_path.read_text(encoding="utf-8"))
            return {k: tuple(v) for k, v in data.items()}
        except FileNotFoundError:
            print(f"Coordinate file '{coord_path}' not found. Using default coordinates.")
            return self.default_coordinates()
        except Exception as exc:  # pragma: no cover - defensive
            print(
                f"Failed to load coordinates from '{coord_path}': {exc}. Using defaults."
            )
            return self.default_coordinates()

    @staticmethod
    def default_coordinates() -> dict[str, tuple[int, int]]:
        """Return built-in fallback coordinates."""
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

    def calibrate_coordinate(self, key: str, path: str | Path | None = None) -> tuple[int, int]:
        """Capture current mouse position and store it under ``key``.

        The captured coordinate is persisted to the mapping file and the in-memory
        coordinates dictionary is updated.
        """

        input(f"Hover over '{key}' and press Enter to capture position...")
        pos = pyautogui.position()
        coord_path = Path(path) if path else self.coordinate_file
        coords = (
            self.load_coordinates(coord_path)
            if coord_path.exists()
            else self.default_coordinates()
        )
        coords[key] = (pos.x, pos.y)
        self._save_coordinates(coords, coord_path)
        self.coordinates = coords
        return pos.x, pos.y

    def _save_coordinates(
        self, coords: dict[str, tuple[int, int]], path: Path | None = None
    ) -> None:
        """Persist coordinates to JSON/YAML file."""

        coord_path = Path(path) if path else self.coordinate_file
        serializable = {k: list(v) for k, v in coords.items()}
        if coord_path.suffix in {".yaml", ".yml"}:
            import yaml  # type: ignore

            coord_path.write_text(
                yaml.safe_dump(serializable), encoding="utf-8"
            )
        else:
            coord_path.write_text(
                json.dumps(serializable, indent=2), encoding="utf-8"
            )

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
            logger.debug("Dismissing alerts")
            self.dismiss_alerts()
            logger.debug("Alerts dismissed")

            logger.debug("Step 4: click hesap search")
            pyautogui.click(*self.coordinates["hesap_search"])
            time.sleep(2)
            logger.debug("Step 4 completed")

            logger.debug("Step 5: enter hesap no")
            pyautogui.click(*self.coordinates["hesap_input"])
            pyautogui.typewrite(excel_data["hesap_no"])  # Excel'den tam hesap no
            time.sleep(1.5)  # Filter'ın çalışmasını bekle
            logger.debug("Step 5 completed")

            logger.debug("Step 6: select account item")
            pyautogui.click(*self.coordinates["account_item"])
            time.sleep(0.5)
            logger.debug("Step 6 completed")

            logger.debug("Step 7a: first Tamam")
            pyautogui.click(*self.coordinates["tamam_button"])
            time.sleep(0.5)
            logger.debug("Step 7a completed")

            logger.debug("Step 7b: second Tamam")
            pyautogui.click(*self.coordinates["tamam_button"])
            time.sleep(1)
            logger.debug("Step 7b completed")

            logger.debug("Step 8: fill date")
            pyautogui.click(*self.coordinates["date_input"])
            pyautogui.typewrite(excel_data["tarih"])
            time.sleep(0.5)
            logger.debug("Step 8 completed")

            logger.debug("Step 9: click Yenile")
            pyautogui.click(*self.coordinates["yenile_btn"])
            time.sleep(2)
            logger.debug("Step 9 completed")

            logger.debug("Step 10: click Yeni")
            pyautogui.click(*self.coordinates["yeni_btn"])
            time.sleep(1)
            logger.debug("Step 10 completed")

            logger.debug("Step 11: select Havale alma")
            pyautogui.click(*self.coordinates["havale_alma"])
            time.sleep(FORM_FILL_DELAY * 2)
            logger.debug("Step 11 completed - modal opened")

            logger.debug("Steps 12-17: fill transaction form")
            self.fill_transaction_form(excel_data)
            logger.debug("Form filling completed")

            logger.debug("Step 18: save")
            self.click_save()
            logger.debug("Step 18 completed")

            logger.debug("Step 19: close window")
            self.click_close()
            logger.debug("Step 19 completed")

            return True
        except Exception as exc:  # pragma: no cover - runtime safeguard
            raise RuntimeError(f"Workflow error: {exc}") from exc

    def fill_transaction_form(self, data: dict[str, str]) -> None:
        """Fill transaction form fields using Excel data."""

        logger.debug("Step 12: enter banka kodu")
        pyautogui.click(*self.coordinates["banka_input"])
        time.sleep(FORM_FILL_DELAY)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data["banka_kodu"])
        time.sleep(FORM_FILL_DELAY)

        logger.debug("Step 13: enter cari kodu")
        pyautogui.click(*self.coordinates["cari_input"])
        time.sleep(FORM_FILL_DELAY)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data["cari_kodu"])
        time.sleep(FORM_FILL_DELAY)

        logger.debug("Step 14: enter belge tarihi")
        pyautogui.click(*self.coordinates["belge_tarih_field"])
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data["tarih"])
        time.sleep(FORM_FILL_DELAY)

        logger.debug("Step 15: enter valör tarihi")
        pyautogui.click(*self.coordinates["valor_tarih_field"])
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data["tarih"])
        time.sleep(FORM_FILL_DELAY)

        logger.debug("Step 16: enter tutar")
        pyautogui.click(*self.coordinates["tutar_input"])
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data["tutar"])
        time.sleep(FORM_FILL_DELAY)

        logger.debug("Step 17: enter açıklama")
        pyautogui.click(*self.coordinates["aciklama_input"])
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(data.get("aciklama", "Test işlemi"))
        time.sleep(FORM_FILL_DELAY)

    def click_save(self) -> None:
        """Trigger save action."""
        logger.debug("Clicking Kaydet button")
        pyautogui.click(*self.coordinates["kaydet_btn"])

    def click_close(self) -> None:
        """Close the current window."""
        logger.debug("Clicking Kapat button")
        pyautogui.click(*self.coordinates["kapat_btn"])
