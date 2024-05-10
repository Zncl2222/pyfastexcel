# Introduction

![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/go.yml?logo=go)
[![Go Report Card](https://goreportcard.com/badge/github.com/Zncl2222/pyfastexcel)](https://goreportcard.com/report/github.com/Zncl2222/pyfastexcel)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/pre-commit.yml?logo=pre-commit&label=pre-commit)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/codeql.yml?logo=github&label=CodeQL)
[![Codacy Badge](https://app.codacy.com/project/badge/Grade/03f42030775045b791586dee20288905)](https://app.codacy.com/gh/Zncl2222/pyfastexcel/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)
[![codecov](https://codecov.io/gh/Zncl2222/pyfastexcel/graph/badge.svg?token=6I03AWUUWL)](https://codecov.io/gh/Zncl2222/pyfastexcel)

---
**Date:**         May 10, 2024

**Version:**      0.0.7

**Project Link:** <https://github.com/Zncl2222/pyfastexcel>

---

This package enables high-performance Excel writing by integrating with the
streaming API from the golang package
[excelize](https://github.com/qax-os/excelize). Users can leverage this
functionality without the need to write any Go code, as the entire process
can be accomplished through Python.

## Features

- Python and Golang Integration: Seamlessly call Golang built shared
libraries from Python.

- No Golang Code Required: Users can solely rely on Python for Excel file
generation, eliminating the need for Golang expertise.

## Installation

### Install via pip (Recommended)

You can easily install the package via pip

```bash
pip install pyfastexcel
```

### Install manually

If you prefer to build the package manually, follow these steps:

1. Clone the repository:

    ```bash
    git clone https://github.com/Zncl2222/pyfastexcel.git
    ```

2. Go to the project root directory:

    ```bash
    cd pyfastexcel
    ```

3. Install the required golang packages:

    ```bash
    go mod download
    ```

4. Build the Golang shared library using the Makefile:

    ```bash
    make
    ```

5. Install the required python packages:

    ```bash
    pip install -r requirements.txt
    ```

    or

    ```bash
    pipenv install
    ```

6. Import the project and start using it!

## How it Works

The core functionality revolves around encoding Excel cell data and styles,
or any other Excel properties, into a JSON string within Python. This JSON
payload is then passed through ctypes to a Golang shared library. In Golang,
the JSON is parsed, and using the streaming writer of
[excelize](https://github.com/qax-os/excelize) to wrtie excel in
high performance.

## Usage

The index assignment is now avaliable in `Workbook` and the `FastWriter`.
Here is the example usage:

```python
from pyfastexcel import Workbook
from pyfastexcel.utils import set_custom_style

# CustomStyle will be integrate to the pyfatexcel in next version
# Beside, CustomStyle will be re-implement in future to make it no-longer
# depend on openpyxl_style writer and openpyxl
from openpyxl_style_writer import CustomStyle


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

You can also using the `FastWriter` or `NormalWriter` which was the
subclass of `Workbook` to write excel row by row, see the following steps:

1. Create a class for your `style` registed like `StyleCollections`
in the example.

2. Create a class for your excel creation implementation and inherit
`NormalWriter` or `FastWriter` and `StyleCollections`.

3. Implement your data writing logic in `def _create_body()` and
`def _create_single_header()`(The latter is not necessary)

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

## Current Limitations & Future Plans

### Problem 1: Dependence on Other Excel Package

Limitations:

This project currently depends on the `CustomStyle` object of
the [openpyxl_style_writer](https://github.com/Zncl2222/openpyxl_style_writer)
package, which is built for openpyxl to write styles in write-only
mode more efficiently without duplicating code.

Future Plans:

This project plans to create its own `Style` object, making it no longer
dependent on the mentioned package.
