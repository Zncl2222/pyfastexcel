# Benchmark

The following results show the performance comparison between `pyfastexcel` and `openpyxl` when writing a data to an Excel file in different scenario.

## Benchmark Environment

> - OS: Windows 11
> - CPU: Intel(R) Core(TM) i7-12700 CPU
> - RAM: DDR4-3200 32GB
> - Hard Drive: Crucial P5 Plus 1TB Read: 6,600 MB/s Write: 5,000 MB/s
> - Python: 3.11.0
> - openpyxl: 3.1.2
> - pyfastexcel: 0.8.0

## Benchmark Result

### Write 50 rows with 30 columns (Total 1500 cells)

<dev align='center'>
    <img src='../benchmark/50+30_horizontal.png'>
</dev>

### Write 500 rows with 30 columns (Total 15000 cells)

<dev align='center'>
    <img src='../benchmark/500+30_horizontal.png'>
</dev>

### Write 5000 rows with 30 columns (Total 150000 cells)

<dev align='center'>
    <img src='../benchmark/5000+30_horizontal.png'>
</dev>

### Write 50000 rows with 30 columns (Total 1500000 cells)

<dev align='center'>
    <img src='../benchmark/50000+30_horizontal.png'>
</dev>
