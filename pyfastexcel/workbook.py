from __future__ import annotations

from typing import List, Literal, Optional, overload

from pydantic import validate_call as pydantic_validate_call

from pyfastexcel.driver import ExcelDriver, WorkSheet
from pyfastexcel.utils import deprecated_warning

from ._typing import CommentTextStructure, SetPanesSelection
from .chart import (
    Chart,
    ChartAxis,
    ChartDimension,
    ChartLegend,
    ChartPlotArea,
    ChartSeries,
    Fill,
    GraphicOptions,
    Line,
    RichTextRun,
)
from .pivot import PivotTable, PivotTableField
from .utils import CommentText, Selection


class Workbook(ExcelDriver):
    """
    A base class for writing data to Excel files with custom styles.

    This class provides methods to set file properties, cell dimensions,
    merge cells, manipulate sheets, and more.

    Methods:
        remove_sheet(sheet: str) -> None:
            Removes a sheet from the Excel data.
        rename_sheet(self, old_sheet_name: str, new_sheet_name: str) -> None:
            Rename a sheet.
        create_sheet(sheet_name: str) -> None:
            Creates a new sheet.
        switch_sheet(sheet_name: str) -> None:
            Set current self.sheet to a different sheet.
        set_file_props(key: str, value: str) -> None:
            Sets a file property.
        set_cell_width(sheet: str, col: str | int, value: int) -> None:
            Sets the width of a cell.
        set_cell_height(sheet: str, row: int, value: int) -> None:
            Sets the height of a cell.
        set_merge_cell(sheet: str, top_left_cell: str, bottom_right_cell: str) -> None:
            Sets a merge cell range in the specified sheet.
    """

    def remove_sheet(self, sheet: str) -> None:
        """
        Removes a sheet from the Excel data.

        Args:
            sheet (str): The name of the sheet to remove.
        """
        if len(self.workbook) == 1:
            raise ValueError('Cannot remove the only sheet in the workbook.')
        if self.workbook.get(sheet) is None:
            raise IndexError(f'Sheet {sheet} does not exist.')
        self.workbook.pop(sheet)
        self._sheet_list = tuple(self.workbook.keys())
        self.sheet = self._sheet_list[0]

    def rename_sheet(self, old_sheet_name: str, new_sheet_name: str) -> None:
        """
        Renames a sheet in the Excel data.

        Args:
            old_sheet_name (str): The name of the sheet to rename.
            new_sheet_name (str): The new name for the sheet.
        """
        if self.workbook.get(old_sheet_name) is None:
            raise IndexError(f'Sheet {old_sheet_name} does not exist.')
        if self.workbook.get(new_sheet_name) is not None:
            raise ValueError(f'Sheet {new_sheet_name} already exists.')
        self.workbook[new_sheet_name] = self.workbook.pop(old_sheet_name)
        self._sheet_list = tuple(
            [new_sheet_name if x == old_sheet_name else x for x in self._sheet_list]
        )
        self.sheet = new_sheet_name

    def create_sheet(
        self,
        sheet_name: str,
        pre_allocate: dict[str, int] = None,
        plain_data: list[list] = None,
    ) -> None:
        """
        Creates a new sheet, and set it as current self.sheet.

        Args:
            sheet_name (str): The name of the new sheet.
            pre_allocate (dict[str, int], optional): A dictionary containing
                'n_rows' and 'n_cols' keys specifying the dimensions
                for pre-allocating data in new sheet.
            plain_data (list[list[str]], optional): A 2D list of strings
                representing initial data to populate new sheet.
        """
        if self.workbook.get(sheet_name) is not None:
            raise ValueError(f'Sheet {sheet_name} already exists.')
        self.workbook[sheet_name] = WorkSheet(pre_allocate=pre_allocate, plain_data=plain_data)
        self.sheet = sheet_name
        self._sheet_list = tuple([x for x in self._sheet_list] + [sheet_name])

    def switch_sheet(self, sheet_name: str) -> None:
        """
        Set current self.sheet to a different sheet. If sheet does not existed
        then raise error.

        Args:
            sheet_name (str): The name of the sheet to switch to.

        Raises:
            IndexError: If sheet does not exist.
        """
        self._check_if_sheet_exists(sheet_name)
        self.sheet = sheet_name

    @pydantic_validate_call
    def set_file_props(self, key: str, value: str) -> None:
        """
        Sets a file property.

        Args:
            key (str): The property key.
            value (str): The property value.

        Raises:
            ValueError: If the key is invalid.
        """
        if key not in self._FILE_PROPS:
            raise ValueError(f'Invalid file property: {key}')
        self.file_props[key] = value

    @pydantic_validate_call
    def protect_workbook(
        self,
        algorithm: str,
        password: str,
        lock_structure: bool = False,
        lock_windows: bool = False,
    ):
        if algorithm not in self._PROTECT_ALGORITHM:
            raise ValueError(
                f'Invalid algorithm, the options are {self._PROTECT_ALGORITHM}',
            )
        self.protection['algorithm'] = algorithm
        self.protection['password'] = password
        self.protection['lock_structure'] = lock_structure
        self.protection['lock_windows'] = lock_windows

    def set_cell_width(self, sheet: str, col: str | int, value: int) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_cell_width(col, value)

    def set_cell_height(self, sheet: str, row: int, value: int) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_cell_height(row, value)

    @overload
    def set_merge_cell(
        self,
        sheet: str,
        top_lef_cell: Optional[str],
        bottom_right_cell: Optional[str],
    ) -> None: ...

    @overload
    def set_merge_cell(
        self,
        sheet: str,
        cell_range: Optional[str],
    ) -> None: ...

    def set_merge_cell(self, sheet, *args) -> None:
        deprecated_warning(
            "wb.set_merge_cell is going to deprecated in v1.0.0. Please use 'wb.merge_cell' instead",
        )
        self.merge_cell(sheet, *args)

    @overload
    def merge_cell(
        self,
        sheet: str,
        top_lef_cell: Optional[str],
        bottom_right_cell: Optional[str],
    ) -> None: ...

    @overload
    def merge_cell(
        self,
        sheet: str,
        cell_range: Optional[str],
    ) -> None: ...

    def merge_cell(self, sheet: str, *args) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_merge_cell(*args)

    def auto_filter(self, sheet: str, target_range: str) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].auto_filter(target_range)

    def set_panes(
        self,
        sheet: str,
        freeze: bool = False,
        split: bool = False,
        x_split: int = 0,
        y_split: int = 0,
        top_left_cell: str = '',
        active_pane: Literal['bottomLeft', 'bottomRight', 'topLeft', 'topRight', ''] = '',
        selection: Optional[SetPanesSelection | list[Selection] | Selection] = None,
    ) -> None:
        """
        Sets the panes for the worksheet with options for freezing, splitting, and selection.

        Args:
            sheet (str): The name of the sheet.
            freeze (bool): Whether to freeze the panes.
            split (bool): Whether to split the panes.
            x_split (int): The column position to split or freeze.
            y_split (int): The row position to split or freeze.
            top_left_cell (str): The top-left cell in the split or frozen panes.
            active_pane (Literal['bottomLeft', 'bottomRight', 'topLeft', 'topRight', '']):
            The active pane.
            selection (Optional[SetPanesSelection | list[Selection]]): The selection
            details for the panes.

        Raises:
            ValueError: If x_split or y_split is negative, or if active_pane is
                invalid.

        Returns:
            None
        """
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_panes(
            freeze=freeze,
            split=split,
            x_split=x_split,
            y_split=y_split,
            top_left_cell=top_left_cell,
            active_pane=active_pane,
            selection=selection,
        )

    def set_data_validation(
        self,
        sheet: str,
        sq_ref: str = '',
        set_range: Optional[list[int | float]] = None,
        input_msg: Optional[list[str]] = None,
        drop_list: Optional[list[str | int | float] | str] = None,
        error_msg: Optional[list[str]] = None,
    ):
        """
        Set data validation for the specified range.

        Args:
            sheet (str): The name of the sheet.
            sq_ref (str): The range to set the data validation.
            set_range (list[int | float]): The range of values to set the data validation.
            input (list[str]): The input message for the data validation.
            drop_list (list[str] | str): The drop list for the data validation.
            error (list[str]): The error message for the data validation.

        Raises:
            ValueError: If the range is invalid.

        Returns:
            None
        """
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_data_validation(
            sq_ref=sq_ref,
            set_range=set_range,
            input_msg=input_msg,
            drop_list=drop_list,
            error_msg=error_msg,
        )

    def add_comment(
        self,
        sheet: str,
        cell: str,
        author: str,
        text: CommentTextStructure | CommentText | List[CommentText],
    ) -> None:
        """
        Adds a comment to the specified cell.
        Args:
            sheet (str): The name of the sheet.
            cell (str): The cell location to add the comment.
            author (str): The author of the comment.
            text (str | dict[str, str] | list[str | dict[str, str]]): The text of the comment.
        Raises:
            ValueError: If the cell location is invalid.
        Returns:
            None
        """
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].add_comment(cell, author, text)

    def group_columns(
        self,
        sheet: str,
        start_col: str,
        end_col: Optional[str] = None,
        outline_level: int = 1,
        hidden: bool = False,
        engine: Literal['pyfastexcel', 'openpyxl'] = 'pyfastexcel',
    ):
        """
        Groups columns between start_col and end_col with specified outline
        level and visibility.

        Args:
            sheet (str): The name of the sheet.
            start_col (str): The starting column to group.
            end_col (Optional[str]): The ending column to group. If None, only
            start_col will be grouped.
            outline_level (int): The outline level of the group.
            hidden (bool): Whether the grouped columns should be hidden.
            engine (Literal['pyfastexcel', 'openpyxl']): The engine to use for
            grouping.

        Returns:
            None
        """
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].group_columns(
            start_col,
            end_col,
            outline_level,
            hidden,
            engine,
        )

    def group_rows(
        self,
        sheet: str,
        start_row: int,
        end_row: Optional[int] = None,
        outline_level: int = 1,
        hidden: bool = False,
        engine: Literal['pyfastexcel', 'openpyxl'] = 'pyfastexcel',
    ):
        """
        Groups rows between start_row and end_row with specified outline level
        and visibility.

        Args:
            sheet (str): The name of the sheet.
            start_row (int): The starting row to group.
            end_row (Optional[int]): The ending row to group. If None,
            only start_row will be grouped.
            outline_level (int): The outline level of the group.
            hidden (bool): Whether the grouped rows should be hidden.
            engine (Literal['pyfastexcel', 'openpyxl']): The engine to use for
            grouping.

        Returns:
            None
        """
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].group_rows(
            start_row,
            end_row,
            outline_level,
            hidden,
            engine,
        )

    def create_table(
        self,
        sheet: str,
        cell_range: str,
        name: str,
        style_name: str = '',
        show_first_column: bool = True,
        show_last_column: bool = True,
        show_row_stripes: bool = False,
        show_column_stripes: bool = True,
        validate_table: bool = True,
    ):
        """
        Creates a table within the specified cell range with given style and display options.

        Args:
            sheet (str): The name of the sheet.
            cell_range (str): The cell range where the table should be created (e.g., 'A1:D10').
            name (str): The name of the table.
            style_name (str): The style to apply to the table. Defaults to an empty string, which
            applies the default style.
            show_first_column (bool): Whether to emphasize the first column.
            show_last_column (bool): Whether to emphasize the last column.
            show_row_stripes (bool): Whether to show row stripes for alternate row shading.
            show_column_stripes (bool): Whether to show column stripes for alternate column shading.
            validate_table (bool): Whether to validate the table through TableFinalValidation.

        Returns:
            None
        """
        self._check_if_sheet_exists(sheet)
        self.workbook[self.sheet].create_table(
            cell_range,
            name,
            style_name,
            show_first_column,
            show_last_column,
            show_row_stripes,
            show_column_stripes,
            validate_table,
        )

    @overload
    def add_chart(self, sheet: str, cell: str, chart_model: Chart | List[Chart]): ...

    @overload
    def add_chart(
        self,
        sheet: str,
        cell: str,
        chart_type: str,
        series: List[ChartSeries] | ChartSeries,
        graph_format: Optional[GraphicOptions] = None,
        title: Optional[List[RichTextRun]] = None,
        legend: Optional[ChartLegend] = None,
        dimension: Optional[ChartDimension] = None,
        vary_colors: Optional[bool] = None,
        x_axis: Optional[ChartAxis] = None,
        y_axis: Optional[ChartAxis] = None,
        plot_area: Optional[ChartPlotArea] = None,
        fill: Optional[Fill] = None,
        border: Optional[Line] = None,
        show_blanks_as: Optional[str] = None,
        bubble_size: Optional[int] = None,
        hole_size: Optional[int] = None,
        order: Optional[int] = None,
    ): ...

    def add_chart(
        self,
        sheet: str,
        cell: str,
        chart_model: Optional[List[Chart] | Chart] = None,
        chart_type: Optional[str] = None,
        series: Optional[List[ChartSeries] | ChartSeries] = None,
        graph_format: Optional[GraphicOptions] = None,
        title: Optional[List[RichTextRun]] = None,
        legend: Optional[ChartLegend] = None,
        dimension: Optional[ChartDimension] = None,
        vary_colors: Optional[bool] = None,
        x_axis: Optional[ChartAxis] = None,
        y_axis: Optional[ChartAxis] = None,
        plot_area: Optional[ChartPlotArea] = None,
        fill: Optional[Fill] = None,
        border: Optional[Line] = None,
        show_blanks_as: Optional[str] = None,
        bubble_size: Optional[int] = None,
        hole_size: Optional[int] = None,
        order: Optional[int] = None,
    ):
        self._check_if_sheet_exists(sheet)
        if isinstance(chart_model, list):
            self.workbook[sheet].add_chart(cell, chart_model)
        else:
            self.workbook[sheet].add_chart(
                cell,
                chart_model,
                chart_type,
                series,
                graph_format,
                title,
                legend,
                dimension,
                vary_colors,
                x_axis,
                y_axis,
                plot_area,
                fill,
                border,
                show_blanks_as,
                bubble_size,
                hole_size,
                order,
            )

    @overload
    def add_pivot_table(self, sheet: str, pivot_table: PivotTable | list[PivotTable]) -> None: ...

    @overload
    def add_pivot_table(
        self,
        sheet: str,
        data_range: str,
        pivot_table_range: str,
        rows: list[PivotTableField],
        pivot_filter: list[PivotTableField],
        columns: list[PivotTableField],
        data: list[PivotTableField],
        row_grand_totals: Optional[bool],
        column_grand_totals: Optional[bool],
        show_drill: Optional[bool],
        show_row_headers: Optional[bool],
        show_column_headers: Optional[bool],
        show_row_stripes: Optional[bool],
        show_col_stripes: Optional[bool],
        show_last_column: Optional[bool],
        use_auto_formatting: Optional[bool],
        page_over_then_down: Optional[bool],
        merge_item: Optional[bool],
        compact_data: Optional[bool],
        show_error: Optional[bool],
        pivot_table_style_name: Optional[str],
    ) -> None: ...

    def add_pivot_table(
        self,
        sheet: str,
        pivot_table: Optional[PivotTable | list[PivotTable]] = None,
        data_range: Optional[str] = None,
        pivot_table_range: Optional[str] = None,
        rows: Optional[list[PivotTableField]] = None,
        pivot_filter: Optional[list[PivotTableField]] = None,
        columns: Optional[list[PivotTableField]] = None,
        data: Optional[list[PivotTableField]] = None,
        row_grand_totals: Optional[bool] = None,
        column_grand_totals: Optional[bool] = None,
        show_drill: Optional[bool] = None,
        show_row_headers: Optional[bool] = None,
        show_column_headers: Optional[bool] = None,
        show_row_stripes: Optional[bool] = None,
        show_col_stripes: Optional[bool] = None,
        show_last_column: Optional[bool] = None,
        use_auto_formatting: Optional[bool] = None,
        page_over_then_down: Optional[bool] = None,
        merge_item: Optional[bool] = None,
        compact_data: Optional[bool] = None,
        show_error: Optional[bool] = None,
        pivot_table_style_name: Optional[str] = None,
    ) -> None:
        self._check_if_sheet_exists(sheet)
        if isinstance(pivot_table, list) or isinstance(pivot_table, PivotTable):
            self.workbook[sheet].add_pivot_table(pivot_table)
        else:
            self.workbook[sheet].add_pivot_table(
                data_range=data_range,
                pivot_table_range=pivot_table_range,
                rows=rows,
                pivot_filter=pivot_filter,
                columns=columns,
                data=data,
                row_grand_totals=row_grand_totals,
                column_grand_totals=column_grand_totals,
                show_drill=show_drill,
                show_row_headers=show_row_headers,
                show_column_headers=show_column_headers,
                show_row_stripes=show_row_stripes,
                show_col_stripes=show_col_stripes,
                show_last_column=show_last_column,
                use_auto_formatting=use_auto_formatting,
                page_over_then_down=page_over_then_down,
                merge_item=merge_item,
                compact_data=compact_data,
                show_error=show_error,
                pivot_table_style_name=pivot_table_style_name,
            )
