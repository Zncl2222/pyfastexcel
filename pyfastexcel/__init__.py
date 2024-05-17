from openpyxl_style_writer import CustomStyle, DefaultStyle

from pyfastexcel.writer import FastWriter, NormalWriter, Workbook

__all__ = [
    'Workbook',
    'FastWriter',
    'NormalWriter',
    # Temporary link the CustomStyle from openpyxl_style_writer for
    # convinent usage.
    'CustomStyle',
    'DefaultStyle',
]
