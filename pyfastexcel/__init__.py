from openpyxl_style_writer import CustomStyle, DefaultStyle

from pyfastexcel.workbook import Workbook
from pyfastexcel.writer import StreamWriter

__all__ = [
    'Workbook',
    'StreamWriter',
    # Temporary link the CustomStyle from openpyxl_style_writer for
    # convinent usage.
    'CustomStyle',
    'DefaultStyle',
]
