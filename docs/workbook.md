# Workbook

A workbook contains all the information of the Excel file.
Users can set the title, Excel properties, or create sheets
through the Workbook class.

## Create the Workbook

By default, pyfastexcel creates `Sheet1` as the default sheet when the workbook is created.
You can access the worksheet through an index, as shown in the code snippet below:

| Parameter           | Data Type    | Description                                     |
|---------------------|--------------|-------------------------------------------------|
| `pre_allocate`      | `dict[str, int]` | Pre-allocate the memory space of given row and column numbers |
| `plain_data`        | `list[list]` | Row and Column to write the excel without style |

```python
from pyfastexcel import Workbook


wb = Workbook()
ws = wb['Sheet1']

# Get all the Sheet
sheet_list = wb.sheet_list
```

!!! note WorkSheet
    The `__getitem__` and `__setitem__` methods are also
    modified in the `WorkSheet` object. Therefore, you might need to use
    `#!python wb.sheet_list` to determine how many sheets you have in the
    `Workbook`. Refer to the `WorkSheet` documentation for more information.

## Create and Save Workbook

After writing all the content, pyfastexcel should encode the Python object to a JSON string and pass it to Golang for decoding. The Excel file is then created using Golang code.

To do this, you should call the function `read_lib_and_create_excel()` to obtain the bytes returned and save the workbook.

```python
file_name = 'pyfast_excel.xlsx'
wb.read_lib_and_create_excel()
wb.save(file_name)
```

!!! note="Note"
    `wb.save()` now will call `read_lib_and_create_excel()` automatically.

If you know the dimension of the data you want to write. You can use `pre_allocate`
to pre_allocate the memory space of the pyfastexcel to improve the performance.

```python
from pyfastexcel import Workbook


pre_allocate = {'n_rows': 1000, 'n_cols': 10}
# This will pre_allocate the memory space of the pyfastexcel
wb = Workbook(pre_allocate=pre_allocate)
wb.save('pre_allocate.xlsx')
```

If you don't need any style for the excel. You could also write the excel without any style with the following code **(This is the fastest way to write the excel)**:

```python
from pyfastexcel import Workbook


data = [[1, 2, 3], [4, 5, 6, 7, 8]]
# This will write the data into the default sheet 'Sheet1'
wb = Workbook(plain_data=data)
wb.save('plain_data.xlsx')
```

!!! note="Note"
    "You can only specify either `pre_allocate` or `plain_data` at a time, not both.

## Create the WorkSheet

A worksheet can be created by the function `#!python wb.create_sheet(sheet_name: str)`

```python
wb.create_sheet('New Sheet')
```

## Remove the WorkSheet

Removing the worksheet is achieved with the function `#!python wb.remove_sheet(sheet_name: str)` function

```python
wb.remove_sheet('Sheet1')
```

!!! note "Note"
    The sheet cannot be removed if there is only one sheet in the workbook.

## Switch current sheet

This function switch the instance attributes `#!python self.sheet`. This function is designed
for `StreamWriter` to use.

```python
wb.switch_sheet('Switch Sheet')
```

!!! not "Note"
    If the sheet does not exist, this function will create one and
    switch `self.sheet`.

## Set Excel Properties

Excle properties can be set using the following function

```python
wb.set_file_props('Creator', 'pyfastexcel-example')
```

Here are all the key options for the set_file_props function.

| Key                | type | default value | Description                                   |
| ------------------ | ---- | ------------- | --------------------------------------------- |
| `Category`         | str  |  empty string | Fetches the category of the resource          |
| `ContentStatue`    | str  |  empty string | Updates the content status of the resource     |
| `Created`          | str  |  empty string | Indicates the creation timestamp of the resource |
| `Creator`          | str  |  'pyfastexcel' | Creator of the Excel file                    |
| `Description`      | str  |  empty string | Provides a brief description of the resource  |
| `Identifier`       | str  |  'xlsx'       | Identifies the file format of the resource    |
| `Keywords`         | str  |  'spreadsheet'| Lists keywords associated with the resource   |
| `LastModifiedBy`   | str  |  'pyfastexcel'| Indicates the last modifier of the resource   |
| `Modified`         | str  |  empty string | Indicates the last modification timestamp of the resource |
| `Revision`         | str  |  '0'          | Specifies the revision number of the resource |
| `Subject`          | str  |  empty string | Describes the subject of the resource         |
| `Title`            | str  |  empty string | Provides the title of the resource            |
| `Language`         | str  |  'en-Us'      | Specifies the language of the resource        |
| `Version`          | str  |  empty string | Indicates the version of the resource |

## Set cell width and height

The cell widht can be set with the function

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `sheet`             | str       | The name of the sheet         |
| `col`               | str       | The column number             |
| `value`             | str       | The value of the width        |

```python title="Set Width"
# Set through alphabet
wb.set_cell_width('New Sheet', 'A', 20)
# Set through number
wb.set_cell_width('New Sheet', 2, 23)
```

The cell height can be set with the function

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `sheet`             | str       | The name of the sheet         |
| `row`               | str or int| The row number                |
| `value`             | int       | The value of the height       |

```python title="Set height"
# Set row 15 to height = 20
wb.set_cell_height('New Sheet', 15, 20)
```

## Merge Cell

The cell can be merge through the function

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `sheet`             | str       | The name of the sheet         |
| `top_left_cell`     | str       | The index of the top left cell|
| `bottom_right_cell` | str       | The index of the bottom right cell|

```python title='Merge Cells'
wb.merge_cell('New Sheet', 'A1', 'B2')
```

## AutoFilter

Create an auto filter in a worksheet.

### Parameters

| Parameter      | Data Type | Description                       |
|----------------|-----------|-----------------------------------|
| `sheet`        | str       | The sheet to applied auto filter  |
| `target_range` | str       | The range where the auto filter will be applied. |

### Example

```python title='Auto Filter'
wb.auto_filter("New Sheet", "A1:C1")
```

## WorkBook Protection

Protect a workbook with a password using various encryption algorithms.
The available options for the algorithm are XOR, MD4, MD5, SHA-1,
SHA-256, SHA-384, and SHA-512.

### Parameters

| Parameter      | Data Type | Description                                   |
|----------------|-----------|-----------------------------------------------|
| `algorithm`    | str       | The encryption algorithm to use for protection. |
| `password`     | str       | The password to protect the workbook.         |
| `lock_structure` | bool      | Whether to lock the workbook structure.       |
| `lock_windows` | bool      | Whether to lock the workbook windows.         |

### Example

```python title='WorkBook Protection'
wb.protect_workbook("XOR", "12345", True, False)
```

## Set Panes

Configure the pane settings for a specific sheet in an Excel file using the provided Excelize file.

### Parameters

| Parameter      | Data Type                | Description                                     |
|----------------|--------------------------|-------------------------------------------------|
| `sheet`        | str                      | The sheet to set the panes               |
| `freeze`       | bool                     | Determines if the panes are frozen.              |
| `split`        | bool                     | Determines if the panes are split.               |
| `x_split`      | int                      | The horizontal position where the panes are split, or the column index that should be frozen. |
| `y_split`      | int                      | The vertical position where the panes are split, or the row index that should be frozen.  |
| `top_left_cell`| str                      | The cell at the top left of the visible window.   |
| `active_pane`  | str                      | The active pane.                                 |
| `selection`    | list[dict[str, str]]     | The selection settings for panes.                |

!!! note "Key of selection"
    The key of the selection is `sq_ref`, `active_cell`, and `pane`. `sq_ref` and `active_cell`
    should be a cell reference, and `pane` should be one of the `topLeft`, `topRight`, `bottomLeft`, and `bottomRight`.

!!! note "Options of active_pane"
    The options for `active_pane` and are `topLeft`, `topRight`, `bottomLeft`, and `bottomRight`.

!!! note "Note"
    When `freeze` is set to `true`, `x_split` and `y_split` represent the column or row index where the panes are
    frozen. When `split` is set to `true`, `x_split` and `y_split` represent the pixel position where the panes are split.

### Example

```python title='Panes Configuration'
from pyfastexcel.utils import Selection

# Set panes with Selection instance
wb.set_panes(
    "Sheet1",
    freeze=True,
    split=False,
    x_split=0,
    y_split=0,
    top_left_cell="A1",
    active_pane="topRight",
    selection=[
        Selection(sq_ref="A1", active_cell="A1", pane="topRight")
    ],
)

# Set panes's selection with dict
wb.set_panes(
    "Sheet1",
    freeze=True,
    split=False,
    x_split=0,
    y_split=0,
    top_left_cell="A1",
    active_pane="topRight",
    selection=[
        {
            "sq_ref": "A1",
            "active_cell": "A1",
            "pane": "topRight",
        }
    ],
)
```

## Set Data Validation

Set data validation for a specified range in a worksheet.

### Parameters

| Parameter      | Data Type                | Description                                                             |
|----------------|--------------------------|-------------------------------------------------------------------------|
| `sheet`        | str                      | Sheet name.                                                             |
| `sq_ref`       | str                      | The range to set the data validation.                                   |
| `set_range`    | list[int or float]       | The range of values to set the data validation.                         |
| `input_msg`    | list[str]                | The input message for the data validation. Must be a list with two elements: [Title, Body]. |
| `drop_list`    | list[str] or str         | The drop list for the data validation. Can be a list of strings or a range in the format "A1:B2".  |
| `error_msg`    | list[str]                | The error message for the data validation. Must be a list with two elements: [Title, Body]. |

### Example

```python title='Set Data Validation'
# Example 1: Setting data validation with a specified range, input message, drop-down list, and error message
wb.set_data_validation(
    sheet='Sheet1',
    sq_ref="A1:B2",
    set_range=[1, 10],
    input_msg=["Input Title", "Input Body"],
    drop_list=["Option1", "Option2", "Option3"],
    error_msg=["Error Title", "Error Body"]
)

# Example 2: Setting data validation with a drop-down list based on cell values
wb.set_data_validation(
    sheet='Sheet1',
    sq_ref="A1:B2",
    drop_list="C1:C5",
)
```

## Add Comment

Adds a comment to the specified cell.

### Parameters

| Parameter  | Data Type                                 | Description                                           |
|------------|-------------------------------------------|-------------------------------------------------------|
| `sheet`    | `str`                                     | The name of sheet.                                    |
| `cell`     | `str`                                     | The cell location to add the comment.                 |
| `author`   | `str`                                     | The author of the comment.                            |
| `text`     | `str` or `dict[str, str]` or `list[str or dict[str, str]]` | The text of the comment, and it's font style|

### Example

```python title='Add Comment'
from pyfastexcel.utils import CommentText

# Add a comment to cell A1 with CommentText Instance
comment_text = CommentText(text='Comment', bold=True)
wb.add_comment("Sheet1", "A1", "pyfastexcel", comment_text)

# Add a comment to cell A1 with list of CommentText Instance
comment_text = CommentText(text='Comment', bold=True)
comment_text2 = CommentText(text=' Comment two', color='00ff00')
wb.add_comment("Sheet1", "A1", "pyfastexcel", [comment_text, comment_text2])

# Add a comment to cell A1, and use string as the comment text
wb.add_comment("Sheet1", "A1", "pyfastexcel", "This is a comment.")

# Add a comment to cell A1, and use dictionary as the comment text and set the font style
wb.add_comment("Sheet1", "A1", "pyfastexcel", {"text": "This is a comment.", 'bold': True, 'italic': True})

# Add a comment to cell A1, and use list of dictionary as the comment text and set the font style
# This will create "This is a comment" with bold and italic font style, and "This is another comment" with bold and red color font style.
wb.add_comment(
    "Sheet1",
    "A1",
    "pyfastexcel",
    [
        {
            "text": "This is a comment.",
            'bold': True,
            'italic': True
        },
        {
            "text": "This is another comment.",
            'bold': True,
            'color': 'FF0000'
        }
    ]
)
```

Here is the `key words of the comment when using the dictionary`:

| Key            | Data Type | Description                       |
|----------------|-----------|-----------------------------------|
| `text`         | str       | The text of the comment.          |
| `size`         | int       | The font size of the comment text. |
| `name`         | str       | The font name of the comment text. |
| `bold`         | bool      | Sets the comment text to bold. |
| `italic`       | bool      | Sets the comment text to italic. |
| `underline`    | str       | Sets the underline style of the comment text. |
| `strike`       | bool      | Sets whether the comment text is strike through. |
| `vertAlign`    | str       | Sets the vertical alignment of the comment text. |
| `color`        | str       | Sets the font color of the comment text. |


!!! note "Note"
    The `text` parameter can be a string, a dictionary, or a list of dictionaries. If it is a string, it will be treated as the comment text. If it is a dictionary, it should contain the key `text` with the comment's text as the corresponding value. If it is a list of dictionaries, each dictionary should contain the key `text` with the comment's text as the corresponding value.

## Group Columns

Group columns in a worksheet. This function is currently implemented using `openpyxl`.
It is not recommended to use this function when dealing with large files.

!!! note "Note"
    `Excelize` currently does not support column grouping in Streaming mode.
    As a temporary solution, this function is implemented with `openpyxl`.
    The process requires calling `load_workbook` after `pyfastexcel` writes
    the data, which may cause performance issues for very large Excel files.

### Parameters

| Parameter      | Data Type        | Description                                      |
|----------------|------------------|------------------------------------------------- |
| `sheet`        | str              | The name of sheet.                               |
| `start_col`    | str              | The cell reference where grouping starts.        |
| `end_col`      | Optional[str]    | The cell reference where grouping ends.          |
| `outline_level`| int              | The level of grouping.                           |
| `hidden`       | bool             | Whether to hide the group or not.                |

```python
wb.group_columns('Sheet1', 'A', 'C', 1, False)
```
