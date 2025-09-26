import toir_tra_index_works as sut
from openpyxl import Workbook


def test_extract_reserved_value_digits_and_letters():
    assert sut.extract_reserved_value("II.23.3-01-6M") == "01"
    assert sut.extract_reserved_value("I.7.5-0-1G") == "00"
    assert sut.extract_reserved_value("IV.1.2A-AB-XY") == "AB"
    assert sut.extract_reserved_value("III.4.1") is None


def test_find_suffix_in_tz_file_with_reserved(tmp_path, monkeypatch):
    wb = Workbook()
    ws = wb.active
    ws.title = sut.TZ_SHEET_NAME
    ws.append(["N", "Par.", " ", " ", "", "Podizvoac", "Ukratko", "Reserved", "Comments"])
    ws.append([1, "II.23.3", "-", "-", "6", "Ostral / Senermax", "OST", "01", ""])
    ws.append([2, "II.23.3", "-", "-", "6", "Ostral / Senermax", "SNX", "02", ""])

    test_file = tmp_path / "tz_test.xlsx"
    wb.save(test_file)

    monkeypatch.setattr(sut, "TZ_FILE_PATH", test_file)

    assert sut.find_suffix_in_tz_file("II.23.3", "01") == "OST"
    assert sut.find_suffix_in_tz_file("II.23.3", "02") == "SNX"
    assert sut.find_suffix_in_tz_file("II.23.3", "03") == "OST"
    assert sut.find_suffix_in_tz_file("II.23.3", None) == "OST"
