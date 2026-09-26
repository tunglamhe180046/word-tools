import pytest

from core.cell_addressing import (
    CellAddressError,
    a1_to_rc,
    col_to_index,
    in_range,
    index_to_col,
    is_valid_a1,
    is_valid_locator,
    make_locator,
    parse_locator,
    parse_range,
    range_to_a1,
    rc_to_a1,
)


@pytest.mark.parametrize("ref,rc", [("A1", (1, 1)), ("B3", (3, 2)), ("Z9", (9, 26)), ("AA10", (10, 27)),
                                   ("XFD1048576", (1048576, 16384)), ("b3", (3, 2))])
def test_a1_roundtrip(ref, rc):
    assert a1_to_rc(ref) == rc
    assert rc_to_a1(*rc) == ref.upper()


@pytest.mark.parametrize("bad", ["", "A0", "0A", "A", "1", "$A$1", "A1:B2", "XFE1", "A1048577", "AAAA1", " ", "Sheet1!A1", "A-1"])
def test_a1_invalid(bad):
    assert not is_valid_a1(bad)
    with pytest.raises(CellAddressError):
        a1_to_rc(bad)


def test_column_conversions():
    assert col_to_index("A") == 1 and col_to_index("AZ") == 52 and col_to_index("XFD") == 16384
    assert index_to_col(1) == "A" and index_to_col(26) == "Z" and index_to_col(27) == "AA" and index_to_col(16384) == "XFD"
    for bad in (0, 16385, -1, True):
        with pytest.raises(CellAddressError):
            index_to_col(bad)
    with pytest.raises(CellAddressError):
        col_to_index("XFE")


def test_rc_bounds():
    for row, col in ((0, 1), (1, 0), (1048577, 1), (1, 16385)):
        with pytest.raises(CellAddressError):
            rc_to_a1(row, col)


def test_parse_range_normalizes_and_single_cell():
    assert parse_range("C3:A1") == (1, 1, 3, 3)
    assert parse_range("B2") == (2, 2, 2, 2)
    assert range_to_a1(1, 1, 3, 3) == "A1:C3"
    assert range_to_a1(2, 2, 2, 2) == "B2"
    with pytest.raises(CellAddressError):
        parse_range("A1:")
    with pytest.raises(CellAddressError):
        parse_range("A1:B2:C3")


def test_in_range():
    rect = parse_range("B2:C3")
    assert in_range(2, 2, rect) and in_range(3, 3, rect)
    assert not in_range(1, 2, rect) and not in_range(2, 4, rect)


def test_locator_format_and_roundtrip():
    assert make_locator(1, 3, 2) == "cell_s1_r3_c2"
    assert parse_locator("cell_s2_r10_c27") == (2, 10, 27)
    assert is_valid_locator("cell_s1_r1_c1")


@pytest.mark.parametrize("bad", ["", "cell_s0_r1_c1", "cell_s1_r0_c1", "cell_s1_r1_c0", "cell_s1_r1", "cell_s1_r1_c1_x",
                                 "CELL_s1_r1_c1", "cell_s1_r1048577_c1", "cell_s1_r1_c16385", "cell_s-1_r1_c1", None])
def test_locator_invalid(bad):
    assert not is_valid_locator(bad)
    with pytest.raises(CellAddressError):
        parse_locator(bad)


def test_make_locator_rejects_bad_input():
    with pytest.raises(CellAddressError):
        make_locator(0, 1, 1)
    with pytest.raises(CellAddressError):
        make_locator(1, 0, 1)
