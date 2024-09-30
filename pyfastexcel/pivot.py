from __future__ import annotations

from pydantic import BaseModel, Field, field_validator, field_serializer
from typing import Optional

from .enums import PivotSubTotal


_pivot_table_style_config = ('PivotStyleLight', 'PivotStyleMedium', 'PivotStyleDark')

_pivot_table_style: set[str] = set()
_pivot_table_style.add('')
_pivot_table_style.add(None)

for name in _pivot_table_style_config:
    for i in range(1, 29):
        _pivot_table_style.add(f'{name}{i}')


class PivotTableField(BaseModel):
    """
    Model representing a field in a PivotTable.

    Attributes:
        compact (Optional[bool]): Indicates whether the field is in compact form.
        data (Optional[str]): Represents the data value for the field.
        name (Optional[str]): The name of the field.
        outline (Optional[bool]): Indicates whether the field is in outline form.
        subtotal (Optional[str]): The type of subtotal for the field.
        default_subtotal (Optional[bool]): Indicates whether this field has a default subtotal applied.
    """

    compact: Optional[bool] = Field(None, serialization_alias='Compact')
    data: Optional[str] = Field(None, serialization_alias='Data')
    name: Optional[str] = Field(None, serialization_alias='Name')
    outline: Optional[bool] = Field(None, serialization_alias='Outline')
    subtotal: Optional[str | PivotSubTotal] = Field(None, serialization_alias='Subtotal')
    default_subtotal: Optional[bool] = Field(None, serialization_alias='DefaultSubtotal')

    @field_serializer('subtotal')
    @classmethod
    def subtotal_serializer(cls, subtotal: str | PivotSubTotal | None) -> str:
        if subtotal is None:
            return None
        if isinstance(subtotal, PivotSubTotal):
            return subtotal.value
        return PivotSubTotal.get_enum(subtotal).value


class PivotTable(BaseModel):
    """
    Model representing a PivotTable, with its configuration and field settings.

    Attributes:
        data_range (str): The range of data to be used in the pivot table, e.g., "Sheet1!A1:B2".
        pivot_table_range (str): The range where the pivot table will be positioned, e.g., "Sheet1!C3:D4".
        rows (list[PivotTableField]): List of fields used as rows in the pivot table.
        pivot_filter ([PivotTableField]): List of fields used as filters in the pivot table.
        columns (list[PivotTableField]): List of fields used as columns in the pivot table.
        data (list[PivotTableField]): List of fields used as data fields in the pivot table.
        row_grand_totals (Optional[bool Indicates whether to show row grand totals.
        column_grand_totals (Optional[bool]): Indicates whether to show column grand.
        show_drill (Optional[bool]): Indicates whether to show drill indicators.
        show_row_headers (Optional[bool]): Indicates whether to show row headers.
        show_column_headers (Optional[bool]): Indicates whether to show column headers.
        show_row_stripes (Optional[bool]): Indicates whether to show row stripes.
        show_col_stripes (Optional[bool]): Indicates whether to show column stripes.
        show_last_column (Optional[bool]): Indicates whether to show the last column.
        use_auto_formatting (Optional[bool]): Indicates whether to use automatic formatting.
        page_over_then_down (Optional[bool]): Indicates whether pages should be ordered from top to bottom
            then left to right.
        merge_item (Optional[bool]): Indicates whether to merge items.
        compact_data (Optional[bool]): Indicates whether to show in a compact form.
        show_error (Optional[bool]): Indicates whether to show errors.
        classic_layout (Optional[bool]): Indicates whether to use classic layout.
        pivot_table_style_name (Optional[str]): Specifies the style the pivot table.
    """

    data_range: str = Field(..., serialization_alias='DataRange')
    pivot_table_range: str = Field(..., serialization_alias='PivotTableRange')
    rows: list[PivotTableField] = Field([PivotTableField()], serialization_alias='Rows')
    pivot_filter: list[PivotTableField] = Field([PivotTableField()], serialization_alias='Filter')
    columns: list[PivotTableField] = Field([PivotTableField()], serialization_alias='Columns')
    data: list[PivotTableField] = Field([PivotTableField()], serialization_alias='Data')
    row_grand_totals: Optional[bool] = Field(None, serialization_alias='RowGrandTotals')
    column_grand_totals: Optional[bool] = Field(None, serialization_alias='ColGrandTotals')
    show_drill: Optional[bool] = Field(None, serialization_alias='ShowDrill')
    show_row_headers: Optional[bool] = Field(None, serialization_alias='ShowRowHeaders')
    show_column_headers: Optional[bool] = Field(None, serialization_alias='ShowColHeaders')
    show_row_stripes: Optional[bool] = Field(None, serialization_alias='ShowRowStripes')
    show_col_stripes: Optional[bool] = Field(None, serialization_alias='ShowColStripes')
    show_last_column: Optional[bool] = Field(None, serialization_alias='ShowLastColumn')
    use_auto_formatting: Optional[bool] = Field(None, serialization_alias='UseAutoFormatting')
    page_over_then_down: Optional[bool] = Field(None, serialization_alias='PageOverThenDown')
    merge_item: Optional[bool] = Field(None, serialization_alias='MergeItem')
    compact_data: Optional[bool] = Field(None, serialization_alias='CompactData')
    show_error: Optional[bool] = Field(None, serialization_alias='ShowError')
    classic_layout: Optional[bool] = Field(None, serialization_alias='ClassicLayout')
    pivot_table_style_name: Optional[str] = Field(None, serialization_alias='PivotTableStyleName')

    @field_validator('data_range')
    @classmethod
    def data_range_validator(cls, data_range: str) -> str:
        if '!' not in data_range or ':' not in data_range:
            raise ValueError(
                'Invalid data range. Expected format: Sheet1!A1:B2 or Sheet1!$A$1:$B$2'
            )
        return data_range

    @field_validator('pivot_table_range')
    @classmethod
    def pivot_table_range_validator(cls, pivot_table_range: str) -> str:
        if '!' not in pivot_table_range or ':' not in pivot_table_range:
            raise ValueError(
                'Invalid pivot_table_range. Expected format: Sheet1!A1:B2 or Sheet1!$A$1:$B$2'
            )
        return pivot_table_range

    @field_validator('pivot_table_style_name')
    @classmethod
    def pivot_table_style_name_validator(cls, style: str) -> str:
        if style not in _pivot_table_style:
            raise ValueError(f'Invalid table style name. Expected one of {_pivot_table_style}')
        return style
