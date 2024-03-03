# pyfastexcel

![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/go.yml?logo=go)
[![Go Report Card](https://goreportcard.com/badge/github.com/Zncl2222/pyfastexcel)](https://goreportcard.com/report/github.com/Zncl2222/pyfastexcel)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/pre-commit.yml?logo=pre-commit&label=pre-commit)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/codeql.yml?logo=github&label=CodeQL)
[![Codacy Badge](https://app.codacy.com/project/badge/Grade/03f42030775045b791586dee20288905)](https://app.codacy.com/gh/Zncl2222/pyfastexcel/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)
[![codecov](https://codecov.io/gh/Zncl2222/pyfastexcel/graph/badge.svg?token=6I03AWUUWL)](https://codecov.io/gh/Zncl2222/pyfastexcel)

This package enables high-performance Excel writing by integrating with the
streaming API from the golang package
[excelize](https://github.com/qax-os/excelize). Users can leverage this
functionality without the need to write any Go code, as the entire process
can be accomplished through Python.

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

## Features

- Python and Golang Integration: Seamlessly call Golang built shared
libraries from Python.

- No Golang Code Required: Users can solely rely on Python for Excel file
generation, eliminating the need for Golang expertise.

## How it Works

The core functionality revolves around encoding Excel cell data and styles,
or any other Excel properties, into a JSON string within Python. This JSON
payload is then passed through ctypes to a Golang shared library. In Golang,
the JSON is parsed, and using the streaming writer of
[excelize](https://github.com/qax-os/excelize) to wrtie excel in
high performance.

## Usage

The current version can only be used with this library through example
snippets or the example.py file in the root directory of this repository.
See [limitations](#current-limitations--future-plans) for more details.

The steps are:

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
        return self._read_lib_and_create_excel()

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
    excel_normal = PyExcelizeNormalExample(data).create_excel()
    file_path = 'pyexample_normal.xlsx'
    with open(file_path2, 'wb') as file:
        file.write(excel_normal)
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

### Problem 2: Inflexible Usage

Limitations:

The current version only has the function to write a cell with a style using
the class-based inheritance streaming writer method, similar to the
`Advanced Usage` in
[openpyxl_style_writer](https://github.com/Zncl2222/openpyxl_style_writer).
This means users must inherit the `NormalWriter` or `FastWriter` classes and
also your `StyleCollections` to correctly register and use the Style. In short,
if you want to use this library, you have to create `StyleCollections` and
`Your-Writer-Class` and implement the excel creation in `Your-Writer-Class` as
shown in the code snippet provided.

```python
from openpyxl_style_writer import CustomStyle

class StyleCollections:
    black_fill_style = CustomStyle(
        font_size='11',
        font_bold=True,
        font_color='F62B00',
        fill_color='000000',
    )
    test_fill_style = CustomStyle(
        font_size='19'
    )


class PyExcelizeNormalExample(NormalWriter, StyleCollections):
    headers = ['col1', 'col2', 'col3']

    def create_excel(self) -> None:
        for row in self.data:
            for h in self.headers:
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style='black_fill_style')
                else:
                    self.row_append(row[h], style='test_fill_style')
            self.create_row()

if __name__ == '__main__':
    data = [{'col1': 1, 'col2': 2, 'col3': 3}, {'col1': 4, 'col2': 5, 'col3', 6}]
    excel = PyExcelizeNormalExample(data).create_excel()
    with open('example.xlsx', 'wb') as file:
        file.write(excel)
```

Future Plans:

1. ~~Make the style register in the shared class object on the Python side.
Also, create a function to register the style. By doing so, the style
won't need to depend on the `custom-writer-class`. All the styles registered
through that function will be sent to Golang and registered.~~ (This has been finished
in current version)

2. Add the ability to create a cell and style with the index, similar
to what openpyxl does. (The code snippet provided is only an example,
not the real implementation method planned)

    ```python
    ws['A1'].value = 'test'
    ws['A1'].style = 'black_fill_style'
    ```
