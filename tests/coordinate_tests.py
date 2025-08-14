from src.coordinate_mapper import map_coordinates

def test_map_coordinates_returns_dict():
    assert isinstance(map_coordinates(), dict)
