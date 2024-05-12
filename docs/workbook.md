## Workbook

A workbook contains all the information of the Excel file.
Users can set the title, Excel properties, or create sheets
through the Workbook class.

### Create the Workbook

By default, pyfastexcel creates `Sheet1` as the default sheet when the workbook is created.
You can access the worksheet through an index, as shown in the code snippet below:

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


### Create and Save Workbook

After writing all the content, pyfastexcel should encode the Python object to a JSON string and pass it to Golang for decoding. The Excel file is then created using Golang code.

To do this, you should call the function `read_lib_and_create_excel()` to obtain the bytes returned and save the workbook.

```python
file_name = 'pyfast_excel.xlsx'
wb.read_lib_and_create_excel()
wb.save(file_name)
```

!!! note "Note"
    `read_lib_and_create_excel()` should be called before saving the file.
    This is the interface between Python and Golang. Without this step,
    the Excel file won't be created.

### Create the WorkSheet

A worksheet can be created by the function `#!python wb.create_sheet(sheet_name: str)`
```python
wb.create_sheet('New Sheet')
```

### Remove the WorkSheet
Removing the worksheet is achieved with the function `#!python wb.remove_sheet(sheet_name: str)` function
```python
wb.remove_sheet('Sheet1')
```

!!! note "Note"
    The sheet cannot be removed if there is only one sheet in the workbook.

### Switch current sheet
This function switch the instance attributes `#!python self.sheet`. This function is designed
for `StreamWriter` to use.
```python
wb.switch_sheet('Switch Sheet')
```

!!! not "Note"
    If the sheet does not exist, this function will create one and
    switch `self.sheet`.

### Set Excel Properties
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
| `Version`          | str  |  empty string | Indicates the version of the resource


### Set cell width and height

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

### Merge Cell
The cell can be merge through the function

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `sheet`             | str       | The name of the sheet         |
| `top_left_cell`     | str       | The index of the top left cell|
| `bottom_right_cell` | str       | The index of the bottom right cell|

```python title='Merge Cells'
wb.set_merge_cell('New Sheet', 'A1', 'B2')
```
