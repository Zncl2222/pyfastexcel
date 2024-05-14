# StreamWriter

StreamWriter is a method for writing Excel row by row. `pyfastexcel` currently
provides two StreamWriter options: `NormalWriter` and `FastWriter`.

!!! warning "Warning"
    Please note that `NormalWriter` and `FastWriter` are currently not stable
    for usage. We are planning to refactor them to improve readability and
    usability.

!!! note "Differnce between `NormalWrtier` and `FastWriter`"
    The main difference between `NormalWriter` and `FastWriter` is that
    `NormalWriter` does not pre-allocate memory space for data, while
    `FastWriter` does. When dealing with large datasets to write into
    Excel, `FastWriter` offers faster performance. Another distinction
    is that currently only `FastWriter` supports index assignment, as seen
    in `Workbook`.

## NormalWriter

`NormalWriter` provides non-static memory space for data. Users do not need to
allocate the number of rows and columns needed. Instead, they can use `append`
to add data to the last row of `NormalWriter`. Below are steps to demonstrate
how to use `NormalWriter` to create an Excel file.

1. Prepare the data. By default, data should be structured as a list, such as
`#!python list[dict[str, str]]`. However, there is no strict data structure
requirement for `NormalWriter`. You can pass any type of data and implement
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

    ```python title="NormalWriter"
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

3. Create a class and inherit `NormalWriter` and `StyleCollections`:

    ```python title="NormalWriter"
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
    ```

4. Pass the data to initialize the Writer class and create the Excel:

    ```python title="Write Excel"
    data = prepare_example_data(653, 90)
    normal_writer = PyExcelizeFastExample(data)
    excel_normal = normal_writer.create_excel()
    file_path = 'pyexample_normal.xlsx'
    normal_writer.save('pyexample_normal.xlsx')
    ```

## FastWriter

`FastWriter` should pre-allocate memory space for data. Users need to
allocate the number of rows and columns needed. They have to use `append`
to add data to the current last row of `FastWriter`. Below are steps
to demonstrate how to use `FastWriter` to create an Excel file.

1. Prepare the data. By default, data should be structured as a list, such as
`#!python list[dict[str, str]]`. `FastWriter` utilizes this structured list to
allocate memory space. Therefore, if you wish to input data with other structure
, you need to override the `__init__` method of `FastWriter` and implement
your own memory allocation logic.

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

    ```python title="NormalWriter"
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

3. Create a class and inherit `FastWriter` and `StyleCollections`:

    ```python title="FastWriter"
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
    ```

4. Pass the data to initialize the Writer class and create the Excel:

    ```python title="Write Excel"
    data = prepare_example_data(653, 90)
    normal_writer = PyExcelizeFastExample(data)
    excel_normal = normal_writer.create_excel()
    file_path = 'pyexample_normal.xlsx'
    normal_writer.save('pyexample_normal.xlsx')
    ```
