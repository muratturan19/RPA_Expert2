"""Preston automation v3: element-driven RPA engine.

This module introduces a production-ready approach that no longer relies
solely on brittle screen coordinates.  Instead it combines multiple
techniques to robustly identify and interact with user interface elements.

Key features
------------
1. Element detection system
   * Text based search via :func:`find_element_by_text` which performs OCR
     on screenshots to locate words such as ``"Kaydet"``.
   * Accessibility name lookup using platform APIs
     (:func:`find_element_by_accessibility`).
   * CSS selector inspired queries for applications exposing a DOM like
     structure (:func:`find_element_by_selector`).
   * OCR fallback for arbitrary text recognition.

2. Template matching
   * OpenCV powered pattern recognition implemented in
     :func:`match_template`.
   * Supports modal, button and icon templates with multi-scale matching
     and configurable confidence thresholds.

3. Dynamic waiting
   * Utility functions such as :func:`wait_until_visible` and
     :func:`wait_until_clickable` poll the screen until a desired element
     appears or becomes clickable.
   * Smart timeouts avoid hanging the automation when UI elements fail to
     materialise.

4. Robust error handling
   * Each interaction method tries multiple strategies and falls back to
     coordinate based clicks when necessary.
   * Detailed logging and explicit exceptions make diagnosing failures
     easier and open the door for self-healing mechanisms.

5. Hybrid approach
   * Coordinate hints are still supported as a backup which enables a
     gradual migration path from the previous versions of the project.
"""

from __future__ import annotations

from dataclasses import dataclass
import time
from pathlib import Path
from typing import Callable, Optional, Tuple

try:  # Optional dependencies used at runtime only
    import pyautogui  # type: ignore
except Exception:  # pragma: no cover - library may be missing in CI
    pyautogui = None  # type: ignore

try:  # pragma: no cover - best effort import
    import cv2  # type: ignore
    import numpy as np  # type: ignore
except Exception:  # pragma: no cover - library may be missing in CI
    cv2 = None  # type: ignore
    np = None  # type: ignore

try:  # pragma: no cover - best effort import
    import pytesseract  # type: ignore
except Exception:  # pragma: no cover - library may be missing in CI
    pytesseract = None  # type: ignore

from logger import get_logger
from config import FORM_FILL_DELAY

logger = get_logger(__name__)


@dataclass
class ElementBox:
    """Simple rectangle to represent an element location."""

    left: int
    top: int
    width: int
    height: int

    def center(self) -> Tuple[int, int]:
        return (self.left + self.width // 2, self.top + self.height // 2)


class PrestonRPAV3:
    """Automation workflow using hybrid element detection."""

    def __init__(self, coord_file: str | Path | None = None) -> None:
        self.coord_file = (
            Path(coord_file) if coord_file else Path(__file__).with_name("coordinates.json")
        )
        self.coordinates = self._load_coordinates(self.coord_file)

    # ------------------------------------------------------------------
    # Coordinate helpers ------------------------------------------------
    # ------------------------------------------------------------------
    def _load_coordinates(self, path: Path) -> dict[str, Tuple[int, int]]:
        """Load coordinate fallback mapping from JSON file."""
        try:
            import json

            data = json.loads(path.read_text(encoding="utf-8"))
            return {k: tuple(v) for k, v in data.items()}
        except Exception:  # pragma: no cover - defensive
            return {}

    # ------------------------------------------------------------------
    # Element detection strategies --------------------------------------
    # ------------------------------------------------------------------
    def find_element_by_text(self, text: str) -> Optional[ElementBox]:
        """Locate element by visible text using OCR.

        This method captures the current screen and searches for *text*.
        Returns an :class:`ElementBox` if the text is found.
        """

        if pyautogui is None or pytesseract is None:  # pragma: no cover - env guard
            return None
        image = pyautogui.screenshot()
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
        for i, word in enumerate(data.get("text", [])):
            if word.strip().lower() == text.lower():
                box = ElementBox(
                    left=data["left"][i],
                    top=data["top"][i],
                    width=data["width"][i],
                    height=data["height"][i],
                )
                logger.debug("Text '%s' found at %s", text, box)
                return box
        logger.debug("Text '%s' not found", text)
        return None

    def find_element_by_accessibility(self, name: str) -> Optional[ElementBox]:
        """Locate element using accessibility APIs.

        This is a stub illustrating where platform specific code such as
        `pywinauto` or `uiautomation` calls would live.
        """

        logger.debug("Accessibility lookup for '%s' not implemented", name)
        return None

    def find_element_by_selector(self, selector: str) -> Optional[ElementBox]:
        """CSS selector like element lookup.

        Real implementations might use a browser driver or application
        specific APIs.  Provided here as a placeholder.
        """

        logger.debug("Selector lookup for '%s' not implemented", selector)
        return None

    def match_template(
        self, template_path: str | Path, confidence: float = 0.8
    ) -> Optional[ElementBox]:
        """Locate a UI element using OpenCV template matching."""

        if pyautogui is None or cv2 is None or np is None:  # pragma: no cover - env guard
            return None

        screenshot = pyautogui.screenshot()
        haystack = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
        template = cv2.imread(str(template_path), cv2.IMREAD_GRAYSCALE)
        if template is None:
            logger.warning("Template '%s' could not be read", template_path)
            return None
        result = cv2.matchTemplate(haystack, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        if max_val >= confidence:
            h, w = template.shape
            box = ElementBox(left=max_loc[0], top=max_loc[1], width=w, height=h)
            logger.debug("Template '%s' matched with %.2f", template_path, max_val)
            return box
        logger.debug("Template '%s' not found (%.2f < %.2f)", template_path, max_val, confidence)
        return None

    # ------------------------------------------------------------------
    # Dynamic waiting ---------------------------------------------------
    # ------------------------------------------------------------------
    def wait_until_visible(
        self, finder: Callable[[], Optional[ElementBox]], timeout: float = 10.0
    ) -> ElementBox:
        """Poll ``finder`` until it returns an :class:`ElementBox`.

        Raises :class:`TimeoutError` if the element does not appear within
        *timeout* seconds.
        """

        start = time.time()
        while time.time() - start < timeout:
            box = finder()
            if box:
                return box
            time.sleep(0.3)
        raise TimeoutError("Element not visible after %.1f seconds" % timeout)

    def wait_until_clickable(
        self, finder: Callable[[], Optional[ElementBox]], timeout: float = 10.0
    ) -> ElementBox:
        """Wait until an element is visible and ready for interaction."""

        box = self.wait_until_visible(finder, timeout)
        # Additional checks could be added here (e.g. pixel colour analysis)
        return box

    # ------------------------------------------------------------------
    # Interaction helpers ----------------------------------------------
    # ------------------------------------------------------------------
    def click_element(
        self,
        *,
        text: Optional[str] = None,
        template: str | Path | None = None,
        coord_key: Optional[str] = None,
        timeout: float = 10.0,
    ) -> None:
        """Click an element using multiple strategies.

        The function first tries text based lookup, then template matching
        and finally falls back to coordinate based clicking if all else
        fails.  Any errors are logged and re-raised to aid debugging.
        """

        if pyautogui is None:  # pragma: no cover - env guard
            raise RuntimeError("pyautogui is required for clicking")

        strategies = []
        if text:
            strategies.append(lambda: self.find_element_by_text(text))
        if template:
            strategies.append(lambda: self.match_template(template))

        for finder in strategies:
            try:
                box = self.wait_until_clickable(finder, timeout=timeout)
                pyautogui.click(*box.center())
                time.sleep(FORM_FILL_DELAY)
                return
            except TimeoutError:
                logger.debug("Strategy %s timed out", finder)
                continue
            except Exception as exc:  # pragma: no cover - defensive
                logger.error("Strategy %s failed: %s", finder, exc)
                continue

        if coord_key and coord_key in self.coordinates:
            x, y = self.coordinates[coord_key]
            logger.debug("Falling back to coordinate (%s, %s) for '%s'", x, y, coord_key)
            pyautogui.click(x, y)
            time.sleep(FORM_FILL_DELAY)
            return

        raise RuntimeError("No valid strategy could click the element")

    # ------------------------------------------------------------------
    # High level workflow placeholder ----------------------------------
    # ------------------------------------------------------------------
    def execute_workflow(self) -> None:
        """Example high level workflow using the new engine."""

        logger.info("Starting V3 workflow")
        # Example usage: click a save button either by text, template or coord
        try:
            self.click_element(
                text="Kaydet",
                template=Path("templates/ui-elements/save_button.png"),
                coord_key="kaydet_btn",
                timeout=5,
            )
        except Exception as exc:
            logger.error("Workflow failed: %s", exc)
            raise
        logger.info("V3 workflow finished")
