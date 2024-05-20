# Introduction

![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/go.yml?logo=go)
[![Go Report Card](https://goreportcard.com/badge/github.com/Zncl2222/pyfastexcel)](https://goreportcard.com/report/github.com/Zncl2222/pyfastexcel)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/pre-commit.yml?logo=pre-commit&label=pre-commit)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zncl2222/pyfastexcel/codeql.yml?logo=github&label=CodeQL)
[![Codacy Badge](https://app.codacy.com/project/badge/Grade/03f42030775045b791586dee20288905)](https://app.codacy.com/gh/Zncl2222/pyfastexcel/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)
[![codecov](https://codecov.io/gh/Zncl2222/pyfastexcel/graph/badge.svg?token=6I03AWUUWL)](https://codecov.io/gh/Zncl2222/pyfastexcel)
[![Documentation Status](https://readthedocs.org/projects/pyfastexcel/badge/?version=stable)](https://pyfastexcel.readthedocs.io/en/stable/?badge=stable)

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

!!! note "Current Limitations"

    This project currently depends on the `CustomStyle` object of
    the [openpyxl_style_writer](https://github.com/Zncl2222/openpyxl_style_writer)
    package, which is built for openpyxl to write styles in write-only
    mode more efficiently without duplicating code.

!!! info "Future Plans"

    This project plans to create its own `Style` object, making it no longer
    dependent on the mentioned package.

## How it Works

The core functionality revolves around encoding Excel cell data and styles,
or any other Excel properties, into a JSON string within Python. This JSON
payload is then passed through ctypes to a Golang shared library. In Golang,
the JSON is parsed, and using the streaming writer of
[excelize](https://github.com/qax-os/excelize) to wrtie excel in
high performance.

## Dependency

The dependency for python and golang

python:

    openpyxl_style_writer (Depends on openpyxl)
    msgspec (for faster json encoding)

golang:

    excelize (Core functionality)
    marshmallow (for faster json decoding)

!!! Note "Dependency version"

    The dependencies for Go are unlikely to change frequently unless there are
    significant performance improvements or necessary changes.
    The current version of excelize is v2.8.0, and marshmallow is v1.1.5

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


## Benchmark

Comming soon...
