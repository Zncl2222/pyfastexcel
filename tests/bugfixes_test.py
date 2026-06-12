"""Tests for bug fixes: Bug 1, 3, 4, 5, 6."""

from __future__ import annotations

import pytest

from pyfastexcel import CustomStyle, Workbook
from pyfastexcel.manager import StyleManager
from pyfastexcel.style import DefaultStyle, Protection


class TestBug1WriterEngineTypeHint:
    """Bug 1: _writer_engine type hint had 'NotmalWriter' typo instead of 'NormalWriter'."""

    def test_writer_engine_type_hint_is_correct(self):
        """Verify the source code no longer contains 'NotmalWriter' typo."""
        import inspect

        from pyfastexcel.worksheet import WorkSheet

        source = inspect.getsource(WorkSheet.__init__)
        assert 'NotmalWriter' not in source
        assert 'NormalWriter' in source

    def test_writer_engine_default_value(self):
        from pyfastexcel.worksheet import WorkSheet

        ws = WorkSheet()
        assert ws._writer_engine == 'StreamWriter'

    def test_writer_engine_normal_writer_assignment(self):
        from pyfastexcel.worksheet import WorkSheet

        ws = WorkSheet()
        ws.group_columns('A', 'B')
        assert ws._writer_engine == 'NormalWriter'


class TestBug3SetCellBySlice:
    """Bug 3: _set_cell_by_slice used start_row instead of row for flat list values."""

    def test_flat_list_single_row_slice(self):
        """Flat list assigned to single-row slice should work correctly."""
        wb = Workbook()
        ws = wb['Sheet1']

        ws['A1':'D1'] = [10, 20, 30, 40]

        row0 = ws[0]
        assert row0[0] == (10, 'DEFAULT_STYLE')
        assert row0[1] == (20, 'DEFAULT_STYLE')
        assert row0[2] == (30, 'DEFAULT_STYLE')
        assert row0[3] == (40, 'DEFAULT_STYLE')

    def test_list_of_lists_multi_row_slice(self):
        """List of lists assigned to multi-row slice should write each row correctly.

        Before the fix, start_row was used instead of row, so all rows would
        write to the first row, overwriting previous data.
        """
        wb = Workbook()
        ws = wb['Sheet1']

        ws['A1':'B2'] = [[1, 2], [3, 4]]

        row0 = ws[0]
        row1 = ws[1]
        assert row0[0] == (1, 'DEFAULT_STYLE')
        assert row0[1] == (2, 'DEFAULT_STYLE')
        assert row1[0] == (3, 'DEFAULT_STYLE')
        assert row1[1] == (4, 'DEFAULT_STYLE')

    def test_list_of_lists_three_rows(self):
        """Verify the fix works for 3-row slices too."""
        wb = Workbook()
        ws = wb['Sheet1']

        ws['A1':'C3'] = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]

        assert ws[0][0] == (1, 'DEFAULT_STYLE')
        assert ws[0][2] == (3, 'DEFAULT_STYLE')
        assert ws[1][0] == (4, 'DEFAULT_STYLE')
        assert ws[1][2] == (6, 'DEFAULT_STYLE')
        assert ws[2][0] == (7, 'DEFAULT_STYLE')
        assert ws[2][2] == (9, 'DEFAULT_STYLE')

    def test_flat_list_wrong_length_raises(self):
        """Flat list with wrong length should raise ValueError."""
        wb = Workbook()
        ws = wb['Sheet1']

        with pytest.raises(ValueError):
            ws['A1':'B2'] = [1]  # Need 2 values for 2 cols


class TestBug4CreateTable:
    """Bug 4: Workbook.create_table used self.sheet instead of sheet parameter."""

    def test_create_table_on_non_active_sheet(self):
        """create_table should write to the specified sheet, not the active one."""
        wb = Workbook()
        wb.create_sheet('Sheet2')

        ws2 = wb['Sheet2']
        ws2['A1'] = 'Name'
        ws2['B1'] = 'Age'
        ws2['A2'] = 'Alice'
        ws2['B2'] = 30

        # create_table on Sheet2 (not the active Sheet1)
        wb.create_table('Sheet2', 'A1:B2', 'my_table')

        assert len(wb['Sheet2']._table_list) == 1
        assert len(wb['Sheet1']._table_list) == 0

        table = wb['Sheet2']._table_list[0]
        assert table['range'] == 'A1:B2'
        assert table['name'] == 'my_table'

    def test_create_table_on_active_sheet(self):
        """create_table should still work on the active sheet."""
        wb = Workbook()
        ws = wb['Sheet1']
        ws['A1'] = 'Name'
        ws['B1'] = 'Age'

        wb.create_table('Sheet1', 'A1:B2', 'active_table')

        assert len(wb['Sheet1']._table_list) == 1
        assert wb['Sheet1']._table_list[0]['name'] == 'active_table'

    def test_create_table_nonexistent_sheet_raises(self):
        """create_table on nonexistent sheet should raise KeyError."""
        wb = Workbook()

        with pytest.raises(KeyError):
            wb.create_table('NonExistent', 'A1:B2', 'table')


class TestBug5ProtectionUnpacking:
    """Bug 5: protection used Protection(**self.protection) instead of Protection(**self.protection_params)."""

    def test_default_protection(self):
        """Default CustomStyle should have correct protection."""
        style = CustomStyle()
        assert isinstance(style.protection, Protection)
        assert style.protection.locked is False
        assert style.protection.hidden is False

    def test_custom_protection_via_params(self):
        """protection_params dict should be correctly unpacked."""
        style = CustomStyle(protection_params={'locked': True, 'hidden': True})
        assert isinstance(style.protection, Protection)
        assert style.protection.locked is True
        assert style.protection.hidden is True

    def test_custom_protection_via_individual_params(self):
        """Individual protect/hidden params should work."""
        style = CustomStyle(protect=True, hidden=True)
        assert isinstance(style.protection, Protection)
        assert style.protection.locked is True
        assert style.protection.hidden is True

    def test_clone_and_modify_protection(self):
        """clone_and_modify should preserve protection correctly."""
        style1 = CustomStyle(protect=True, hidden=False)
        style2 = style1.clone_and_modify(protect=False, hidden=True)

        assert style1.protection.locked is True
        assert style1.protection.hidden is False
        assert style2.protection.locked is False
        assert style2.protection.hidden is True

    def test_default_style_protection_class_level(self):
        """DefaultStyle.protection should be a Protection object."""
        assert isinstance(DefaultStyle.protection, Protection)


class TestBug6StyleManagerExceptionSafety:
    """Bug 6: StyleManager global state should be reset even if exceptions occur."""

    def test_style_manager_reset_after_normal_call(self):
        """StyleManager state should be reset after normal read_lib_and_create_excel."""
        wb = Workbook()
        style = CustomStyle(font_bold=True)
        ws = wb['Sheet1']
        ws['A1'] = ('test', style)

        wb.read_lib_and_create_excel()

        assert len(StyleManager.REGISTERED_STYLES) == 1  # Only DEFAULT_STYLE
        assert StyleManager._STYLE_ID == 0
        assert len(StyleManager._style_map) == 0

    def test_style_manager_reset_after_exception(self):
        """StyleManager state should be reset even if an exception occurs during export."""
        wb = Workbook()
        ws = wb['Sheet1']
        ws['A1'] = 'test'

        def failing_transfer():
            raise RuntimeError('Simulated export failure')

        ws._transfer_to_dict = failing_transfer

        with pytest.raises(RuntimeError):
            wb.read_lib_and_create_excel()

        assert len(StyleManager.REGISTERED_STYLES) == 1
        assert StyleManager._STYLE_ID == 0
        assert len(StyleManager._style_map) == 0

    def test_multiple_workbooks_do_not_interfere(self):
        """Successive Workbook instances should not share corrupted style state."""
        wb1 = Workbook()
        style1 = CustomStyle(font_bold=True)
        wb1['Sheet1']['A1'] = ('test', style1)
        wb1.read_lib_and_create_excel()

        assert len(StyleManager.REGISTERED_STYLES) == 1

        wb2 = Workbook()
        style2 = CustomStyle(font_color='FF0000')
        wb2['Sheet1']['A1'] = ('test2', style2)
        wb2.read_lib_and_create_excel()

        assert len(StyleManager.REGISTERED_STYLES) == 1
