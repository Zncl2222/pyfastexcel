# Overview

`StreamWriter` is a method for writing Excel row by row. `pyfastexcel`.

## StreamWriter

`StreamWriter` provides non-static memory space for data. Users do not need to
allocate the number of rows and columns needed. Instead, they can use `append`
to add data to the last row of `StreamWriter`. Below are steps to demonstrate
how to use `StreamWriter` to create an Excel file.

1. Prepare the data. By default, data should be structured as a list, such as
`#!python list[dict[str, str]]`. However, there is no strict data structure
requirement for `StreamWriter`. You can pass any type of data and implement
your own row-by-row writing method.

    ```python title="Prepare Data"
    def prepare_example_data(rows: int = 1000, cols: int = 10) -> list[dict[str, str]]:
        import random

        random.seed(42)
        headers = [f'Column_{i}' for i in range(cols)]
        data = [[random.random() for _ in range(cols)] for _ in range(rows)]
        records = []
        for row in data:
            record = {}
            for header, value in zip(headers, row):
                record[header] = str(round(value * 100, 2))
            records.append(record)
        return records
    ```

2. Create the style collections

    ```python title="Create Styles"
    from openpyxl_style_writer import CustomStyle
    from openpyxl.styles import Side


    class StyleCollections:
        black_fill_style = CustomStyle(
            font_size='11',
            font_bold=True,
            font_color='F62B00',
            fill_color='000000',
        )
        green_fill_style = CustomStyle(
            font_size='29',
            font_bold=False,
            font_color='000000',
            fill_color='375623',
        )
        test_fill_style = CustomStyle(
            font_params={
                'size': 20,
                'bold': True,
                'italic': True,
                'color': '5e03fc',
            },
            fill_params={
                'patternType': 'solid',
                'fgColor': '375623',
            },
            border_params={
                'left': Side(style='thin', color='e12aeb'),
                'right': Side(style='thick', color='e12aeb'),
                'top': Side(style=None, color='e12aeb'),
                'bottom': Side(style='dashDot', color='e12aeb'),
            },
            ali_params={
                'wrapText': True,
                'shrinkToFit': True,
            },
            number_format='0.00%',
        )
    ```

3. Create a class and inherit `StreamWriter` and `StyleCollections`:

    ```python title="StreamWriter"
    class PyFastExcelStreamExample(StreamWriter, StyleCollections):

        def create_excel(self) -> bytes:
            self._set_header()
            self._create_style()
            self.set_file_props('Creator', 'Hello')
            self._create_single_header()
            self._create_body()
            return self.read_lib_and_create_excel()

        def _set_header(self):
            self.headers = list(self.data[0].keys())

        def _create_single_header(self):
            for h in self.headers:
                self.row_append(h, style='green_fill_style')
            self.create_row()

        def _create_body(self) -> None:
            for row in self.data:
                for h in self.headers:
                    if h[-1] in ('1', '3', '5', '7', '9'):
                        self.row_append(row[h], style='black_fill_style')
                    else:
                        self.row_append(row[h], style='test_fill_style')
                self.create_row()

            self.switch_sheet('Sheet2')
            for row in self.data:
                for h in self.headers:
                    if h[-1] in ('1', '3', '5', '7', '9'):
                        self.row_append(row[h], style=self.green_fill_style)
                    else:
                        self.row_append(row[h], style='black_fill_style')
                self.create_row()

            # Create row by row_append_list
            self.switch_sheet('Sheet3')
            list_data = [1, 2, 3, 4, 5]
            self.row_append_list(list_data, style='black_fill_style')
            self.create_row()

            # Set create_row = True to apply the value to sheet without calling
            # self.create_row()
            self.row_append_list(
                list_data,
                style=self.green_fill_style,
                create_row=True
            )

            # You can also assign the value via index
            self.workbook['Sheet1']['A2'] = ('Hellow World', 'black_fill_style')
            self.workbook['Sheet1']['A3'] = 'I am A3'
            self.workbook['Sheet1']['AB9'] = 'qwer'
    ```

4. Pass the data to initialize the Writer class and create the Excel:

    ```python title="Write Excel"
    data = prepare_example_data(653, 90)
    stream_writer = PyFastExcelStreamExample(data)
    excel_bytes = stream_writer.create_excel()
    file_path = 'pyexample_normal.xlsx'
    stream_writer.save('pyexample_normal.xlsx')
    ```

## Style Modification

The `StreamWriter` provides a method to dynamically modify the style in the
`row_append` function. You can pass keywords from ``CustomStyle` to the
`row_append` function to override the style.

!!! note "Note"
    Currently, if you define a `CustomStyle` through `font_params`, `ali_params`
    or any other `_params` key words. You should also pass `_params` key words
    in `row_append` to override the style. This will be fixed in the future
    version of `openpyxl_style_writer`. For example

    ```python
    style = CustomStyle(
        font_params={
            'size': 20,
            'bold': True,
            'italic': True,
            'color': '5e03fc',
        },
        ali_params={
            'wrapText': True,
            'shrinkToFit': True,
        },
    )

    # This will not work because the original style is using font_params
    sw.row_append('Hello', style=style, font_size=33)

    # This will work
    sw.row_append(
        'Hello',
        style=style,
        font_params={
            'size': 20,
            'bold': True,
            'italic': True,
            'color': '5e03fc',
        }
    )
    ```

!!! warning "Warning"
    By using this method, it might decrease the performance of the writing process.
    It is recommended to use the `StyleCollections` method or call
    `set_custom_style` to create the styles.

```python title="Style Modification"
from pyfastexcel import CustomStyle, StreamWriter
from pyfastexcel.utils import set_custom_style


sw = StreamWriter()

style = CustomStyle()
set_custom_style('normal_style', style)

sw.row_append('Hello', style='normal_style', font_color='00ff00', font_bold=True)
sw.row_append('Hello2', style=style, font_color='ff0000', font_size=33)
sw.create_row()
sw.save('test.xlsx')
```
