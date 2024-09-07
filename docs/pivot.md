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
| `filter`              | `list[PivotTableField]`                 | List of fields used as filters in the pivot table.                                    |
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
| `pivot_table_style_name` | `Optional[str]`                      | Specifies the style for the pivot table, chosen from a predefined set of styles.      |
