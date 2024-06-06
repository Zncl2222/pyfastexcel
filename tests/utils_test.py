import pytest
from openpyxl_style_writer import CustomStyle

from pyfastexcel.utils import (
    _separate_alpha_numeric,
    column_to_index,
    deprecated_warning,
    index_to_column,
    set_custom_style,
)


@pytest.mark.parametrize(
    'column, expected_index',
    [
        ('A', 1),
        ('Z', 26),
        ('AA', 27),
        ('XFD', 16384),
        ('XFEA', ValueError),
        (92, TypeError),
    ],
)
def test_column_to_index_valid_columns(column, expected_index):
    if not isinstance(expected_index, int):
        with pytest.raises(expected_index):
            column_to_index(column)
    else:
        assert column_to_index(column) == expected_index


@pytest.mark.parametrize(
    'column',
    [
        '0',
        '1A',
        'AA1',
        'XFE',
    ],
)
def test_column_to_index_invalid_columns(column):
    with pytest.raises(ValueError):
        column_to_index(column)


@pytest.mark.parametrize(
    'index, expected_column',
    [
        (1, 'A'),
        (26, 'Z'),
        (27, 'AA'),
        (16384, 'XFD'),
    ],
)
def test_index_to_column_valid_indices(index, expected_column):
    assert index_to_column(index) == expected_column


@pytest.mark.parametrize(
    'index, error_type',
    [
        (0, ValueError),
        (-1, ValueError),
        (16385, ValueError),
        ('A', TypeError),
    ],
)
def test_index_to_column_invalid_indices(index, error_type):
    with pytest.raises(error_type):
        index_to_column(index)


def test_set_custom_style():
    style = CustomStyle(font_size=12, font_bold=True)
    set_custom_style('bold_font', style)


def test_deprecated_warning():
    deprecated_warning('This is a Test')


@pytest.mark.parametrize(
    'index, error_type',
    [
        (';', ValueError),
        ('~~', ValueError),
    ],
)
def test_seperate_alpha_numeric_error(index, error_type):
    with pytest.raises(error_type):
        _separate_alpha_numeric(index)


@pytest.mark.parametrize(
    'index, alpha, num',
    [
        ('A1', 'A', '1'),
        ('XD6', 'XD', '6'),
        ('ZZ999', 'ZZ', '999'),
    ],
)
def test_seperate_alpha_numeric(index, alpha, num):
    a, n = _separate_alpha_numeric(index)
    assert a == alpha
    assert n == num
