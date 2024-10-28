# PivotTable

Create a PivotTable in pyfastexcel required the following classes, `PivotTableField` and `PivotTable`. The `PivotTableField` class represents a field within a PivotTable, which includes various settings like name, data, compact display, and subtotal configurations. The `PivotTable` class represents a PivotTable configuration, including data ranges, field settings, and display options.

## Example

```python
import random

from pyfastexcel import Workbook
from pyfastexcel.pivot import PivotTable, PivotTableField


wb = Workbook()
ws = wb['Sheet1']

columns = ['Year', 'Month', 'Market', 'Location', 'Sales']
month = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
]
year = [2024, 2025, 2026, 2027, 2028]
location = ['North', 'South', 'East', 'West']
market = ['A', 'B', 'C', 'D']

ws[0] = columns
for i in range(1, 60):
    ws[f'A{i+1}'] = random.choice(year)
    ws[f'B{i+1}'] = random.choice(month)
    ws[f'C{i+1}'] = random.choice(market)
    ws[f'D{i+1}'] = random.choice(location)
    ws[f'E{i+1}'] = random.randint(-1000, 10000)

pivot = PivotTable(
    data_range='Sheet1!A1:E60',
    pivot_table_range='Sheet1!H2:N60',
    rows=[PivotTableField(data='Year', default_subtotal=True), PivotTableField(data='Month')],
    pivot_filter=[PivotTableField(data='Market')],
    columns=[PivotTableField(data='Location')],
    data=[PivotTableField(data='Sales', name='Summation', subtotal='sum')],
    show_drill=True,
    row_grand_totals=True,
    column_grand_totals=True,
    show_row_headers=True,
    show_column_headers=True,
    show_last_column=True,
    pivot_table_style_name='PivotStyleLight16',
)
ws.add_pivot_table(pivot)

wb.save('PivotTableExample.xlsx')
```

You can also set up the pivot table using the following style

```python
pivot = PivotTable(
    data_range='Sheet1!A1:E60',
    pivot_table_range='Sheet1!H3:N60',
)
pivot.rows[0].data = 'Year'
pivot.rows[0].default_subtotal = True
pivot.rows.append(PivotTableField(data='Month'))
pivot.pivot_filter[0].data = 'Market'
pivot.columns[0].data = 'Location'
pivot.data[0].data = 'Sales'
pivot.data[0].name = 'Summation'
pivot.data[0].subtotal = 'sum'
pivot.show_drill = True
pivot.row_grand_totals = True
pivot.column_grand_totals = True
pivot.show_row_headers = True
pivot.show_column_headers = True
pivot.show_last_column = True
pivot.pivot_table_style_name = 'PivotStyleLight16'
```

<div align='center'>

<img src='../images/pivot_table.png'>

</div>

## PivotTableField

Represents a field within a PivotTable, which includes various settings like name,
data, compact display, and subtotal configurations.

| Attribute         | Type                                     | Description                                                       |
|-------------------|------------------------------------------|-------------------------------------------------------------------|
| `compact`         | `Optional[bool]`                         | Indicates whether the field is displayed in compact form.          |
| `data`            | `Optional[str]`                          | The data value associated with the field.                         |
| `name`            | `Optional[str]`                          | The name of the field.                                             |
| `outline`         | `Optional[bool]`                         | Indicates whether the field is in outline form.                    |
| `subtotal`        | `Optional[str or PivotSubTotal]`          | The subtotal type for the field.                                   |
| `default_subtotal`| `Optional[bool]`                         | Specifies if the field has a default subtotal applied.             |

## PivotTable

Represents a PivotTable configuration, including data ranges, field settings, and display options.

| Attribute             | Type                                    | Description                                                                           |
|-----------------------|-----------------------------------------|---------------------------------------------------------------------------------------|
| `data_range`          | `str`                                   | Range of data used for the pivot table, e.g., "Sheet1!A1:B2".                         |
| `pivot_table_range`   | `str`                                   | Range where the pivot table will be placed, e.g., "Sheet1!C3:D4".                     |
| `rows`                | `list[PivotTableField]`                 | List of fields used as rows in the pivot table.                                       |
| `pivot_filter`        | `list[PivotTableField]`                 | List of fields used as filters in the pivot table.                                    |
| `columns`             | `list[PivotTableField]`                 | List of fields used as columns in the pivot table.                                    |
| `data`                | `list[PivotTableField]`                 | List of fields used as data fields in the pivot table.                                |
| `row_grand_totals`    | `Optional[bool]`                        | Indicates whether to display row grand totals.                                        |
| `column_grand_totals` | `Optional[bool]`                        | Indicates whether to display column grand totals.                                     |
| `show_drill`          | `Optional[bool]`                        | Indicates whether to show drill indicators.                                           |
| `show_row_headers`    | `Optional[bool]`                        | Indicates whether to display row headers.                                             |
| `show_column_headers` | `Optional[bool]`                        | Indicates whether to display column headers.                                          |
| `show_row_stripes`    | `Optional[bool]`                        | Indicates whether to display row stripes.                                             |
| `show_col_stripes`    | `Optional[bool]`                        | Indicates whether to display column stripes.                                          |
| `show_last_column`    | `Optional[bool]`                        | Indicates whether to display the last column.                                         |
| `use_auto_formatting` | `Optional[bool]`                        | Indicates whether to apply automatic formatting.                                      |
| `page_over_then_down` | `Optional[bool]`                        | Indicates whether pages should be ordered top-to-bottom, then left-to-right.          |
| `merge_item`          | `Optional[bool]`                        | Indicates whether to merge items in the pivot table.                                  |
| `compact_data`        | `Optional[bool]`                        | Indicates whether to display data in compact form.                                    |
| `show_error`          | `Optional[bool]`                        | Indicates whether to display errors.                                                  |
| `classic_layout`      | `Optional[bool]`                        | Specifies whether to apply the classic layout style to the pivot table.               |
| `pivot_table_style_name` | `Optional[str]`                      | Specifies the style for the pivot table, chosen from a predefined set of styles.      |
