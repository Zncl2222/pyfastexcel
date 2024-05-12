## Example

### Write excel via Workbook
The following example show how to create a Workbook
and assign styles and values to cells.

```python title="Workbook"
from pyfastexcel import Workbook
from pyfastexcel.utils import set_custom_style
from openpyxl_style_writer import CustomStyle # (1)


if __name__ == '__main__':
    # Workbook
    wb = Workbook()

    # Set and register CustomStyle
    bold_style = CustomStyle(font_size=15, font_bold=True)
    set_custom_style('bold_style', bold_style)

    ws = wb['Sheet1']
    # Write value with default style
    ws['A1'] = 'A1 value'
    # Write value with custom style
    ws['B1'] = ('B1 value', 'bold_style')

    # Write value in slice with default style
    ws['A2': 'C2'] = [1, 2, 3]
    # Write value in slice with custom style
    ws['A3': 'C3'] = [(1, 'bold_style'), (2, 'bold_style'), (3, 'bold_style')]

    # Write value by row with default style (python index 0 is the index 1 in excel)
    ws[3] = [9, 8, 'go']
    # Write value by row with custom style
    ws[4] = [(9, 'bold_style'), (8, 'bold_style'), ('go', 'bold_style')]

    # Send request to golang lib and create excel
    wb.read_lib_and_create_excel()

    # File path to save
    file_path = 'pyexample_workbook.xlsx'
    wb.save(file_path)

```

1.  CustomStyle will be integrate to the pyfatexcel in next version
    Beside, CustomStyle will be re-implement in future to make it no-longer
    depend on openpyxl_style writer and openpyxl

### Write excel via StreamWriter

!!! warning "Warning"
    Please note that the usage of StreamWriter, such as `FastWriter` and
    `NormalWriter`, may have changed in recent versions (currently v0.0.7).

You can also using the `FastWriter` or `NormalWriter` which was the
subclass of `Workbook` to write excel row by row, see the following steps:

1. Create a class for your `style` registed like `StyleCollections`
in the example.

    !!! Note "Style"
        Every accessible attribute of the `CustomStyle` class will be
        registered once you call the `read_lib_and_create_excel()` function.
        Therefore, you can easily include the `CustomStyle` in your writer
        class, or create a `StyleCollections` class
        and inherit it from your `Writer` class.

2. Create a class for your excel creation implementation and inherit
`NormalWriter` or `FastWriter` and `StyleCollections`.

3. Implement your data writing logic in `def _create_body()` and
`def _create_single_header()`(The latter is not necessary)

!!! example "NormalWriter"

    ```python
    from openpyxl_style_writer import CustomStyle
    from openpyxl.styles import Side
    from pyfastexcel.driver import FastWriter, NormalWriter


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


    class PyExcelizeNormalExample(NormalWriter, StyleCollections):

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

    if __name__ == '__main__':
        data = prepare_example_data(653, 90)
        normal_writer = PyExcelizeFastExample(data)
        excel_normal = normal_writer.create_excel()
        file_path = 'pyexample_normal.xlsx'
        normal_writer.save('pyexample_normal.xlsx')
    ```

The example of FastWriter now supports index assignment. Please see
the last few lines of code in `_create_body()` for reference.

!!! example "FastWriter"
    ```python
    from pyfastexcel.driver import FastWriter


    class PyExcelizeFastExample(FastWriter, StyleCollections):

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

            # Assigning a value with a specific style
            self.workbook['Sheet1']['A2'] = ('Hellow World', 'black_fill_style')

            # Assigning a value without specifying a style (default style used)
            self.workbook['Sheet1']['A3'] = 'I am A3'
            self.workbook['Sheet1']['AB9'] = 'GOGOGO'


    if __name__ == '__main__':
        data = prepare_example_data(653, 90)
        normal_writer = PyExcelizeFastExample(data)
        excel_normal = normal_writer.create_excel()
        file_path = 'pyexample_normal.xlsx'
        normal_writer.save('pyexample_normal.xlsx')

    ```
