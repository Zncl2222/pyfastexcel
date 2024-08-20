import logging

from pydantic import BaseModel, Field, field_validator, model_validator
from typing import Any, Self, Optional, Literal

from ._typing import SetPanesSelection
from .utils import Selection
from .logformatter import formatter
from .utils import cell_reference_to_index, _validate_cell_reference

logger = logging.getLogger(__name__)
style_formatter = logging.StreamHandler()
style_formatter.setFormatter(formatter)

logger.addHandler(style_formatter)
logger.propagate = False

_table_style_config: dict[str, int] = {
    'TableStyleLight': 21,
    'TableStyleMedium': 28,
    'TableStyleDark': 11,
}
_table_style: set[str] = set()
_table_style.add('')

for s, num in _table_style_config.items():
    for i in range(1, num + 1):
        _table_style.add(f'{s}{i}')


class TableValidator(BaseModel):
    cell_range: str
    name: str
    style_name: str = Field(default='', strict=True)
    show_first_column: bool = Field(default=True, strict=True)
    show_last_column: bool = Field(default=True, strict=True)
    show_row_stripes: bool = Field(default=False, strict=True)
    show_column_stripes: bool = Field(default=True, strict=True)
    validate_table: bool = Field(default=True, strict=True)

    @field_validator('cell_range')
    @classmethod
    def validate_cell_range(cls, cell_range: str) -> str:
        if ':' not in cell_range:
            raise ValueError('Invalid cell range. Expected format: A1:B2')
        cell_range_split = cell_range.split(':')
        if len(cell_range_split) != 2:
            raise ValueError('Invalid cell range. Expected format: A1:B2')

        _validate_cell_reference(cell_range_split[0])
        _validate_cell_reference(cell_range_split[1])

        return cell_range

    @field_validator('style_name')
    @classmethod
    def validate_style_name(cls, style_name: str) -> str:
        if style_name not in _table_style:
            raise ValueError(f'Invalid table style name. Expected one of {_table_style}')
        return style_name


class TableFinalValidation(BaseModel):
    data: list
    table_list: list[dict[str, Any]]

    @model_validator(mode='after')
    def validate_table_list(self) -> Self:
        for t in self.table_list:
            if t['validate_table'] is False:
                continue
            t_split = t['range'].split(':')
            start_row, start_col = cell_reference_to_index(t_split[0])
            end_row, end_col = cell_reference_to_index(t_split[1])

            # Check if table range is valid, end_col should +1 because of length comparison
            if end_col + 1 > len(self.data[start_row]):
                raise ValueError(
                    f"Invalid table range for {t['name']}. "
                    'Please write a row for table first row.'
                )

            if len(set(self.data[start_row])) != len(self.data[start_row]):
                raise ValueError(
                    'Invalid table header. ' 'The first row contains duplicate values.'
                )

            end_row = len(self.data) - 1 if end_row >= len(self.data) else end_row
            for col in range(start_col, end_col + 1):
                column_data = [
                    self.data[row][col] for row in range(start_row, end_row + 1) if self.data[row]
                ]
                first_element = column_data[0]
                seen = set(column_data[1:])
                if first_element in seen:
                    raise ValueError('Invalid table data. Column contains duplicate values.')

        return self


class PanesValidator(BaseModel):
    freeze: bool = Field(default=False, strict=True)
    split: bool = Field(default=False, strict=True)
    x_split: int = Field(default=0, strict=True)
    y_split: int = Field(default=0, strict=True)
    top_left_cell: str = Field(default='', strict=True)
    active_pane: Literal['bottomLeft', 'bottomRight', 'topLeft', 'topRight', ''] = ''
    selection: Optional[SetPanesSelection | list[Selection] | Selection] = None

    @model_validator(mode='after')
    def validate_panes(self) -> Self:
        if self.x_split < 0 or self.y_split < 0:
            raise ValueError('Split position should be positive.')
        if self.top_left_cell != '':
            _validate_cell_reference(self.top_left_cell)
        return self


class DataValidationValidator(BaseModel):
    sq_ref: str = Field(default=False, strict=True)
    set_range: Optional[list[int | float]] = None
    input_msg: Optional[list[str]] = None
    drop_list: Optional[list[str | int | float] | str] = None
    error_msg: Optional[list[str]] = None

    @field_validator('sq_ref')
    @classmethod
    def validate_style_name(cls, sq_ref: str) -> str:
        if ':' in sq_ref:
            sq_ref_list = sq_ref.split(':')
            _validate_cell_reference(sq_ref_list[0])
            _validate_cell_reference(sq_ref_list[1])
        else:
            _validate_cell_reference(sq_ref)
        return sq_ref


# Register validators and use them in the validate_call decorator
VALIDATORS = {
    'create_table': TableValidator,
    'set_panes': PanesValidator,
    'set_data_validation': DataValidationValidator,
}


def validate_call(func):
    def wrapper(*args, **kwargs):
        func_name = func.__name__
        if func_name not in VALIDATORS:
            logger.warning(f'No validator found for function {func_name}. Skipping validation.')
            return func(*args, **kwargs)

        validator = VALIDATORS[func_name]
        model_fields = validator.model_fields
        has_self = 'self' in func.__code__.co_varnames

        actual_args = args[1:] if has_self else args

        _kwargs = dict(zip(model_fields, actual_args), **kwargs)
        validator(**_kwargs)

        return func(*args, **kwargs)

    return wrapper
