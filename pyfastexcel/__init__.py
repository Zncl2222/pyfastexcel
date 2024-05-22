from openpyxl_style_writer import CustomStyle, DefaultStyle

from pyfastexcel.writer import StreamWriter, Workbook

__all__ = [
    'Workbook',
    'StreamWriter',
    # Temporary link the CustomStyle from openpyxl_style_writer for
    # convinent usage.
    'CustomStyle',
    'DefaultStyle',
]
