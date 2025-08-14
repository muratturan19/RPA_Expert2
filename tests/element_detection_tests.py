from src.element_detector import detect_elements

def test_detect_elements_returns_list():
    assert isinstance(detect_elements(), list)
