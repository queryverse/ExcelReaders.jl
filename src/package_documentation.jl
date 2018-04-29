"""
# ExcelReaders

## Synopsis
ExcelReaders allows you to read data from Excel files.

## Exported functions
* ``openxl``: Open an Excel file for repeated reads.
* ``readxl``: Read data from a specified range in an Excel file.
* ``readxlsheet``: Read data from a specified sheet in an Excel file.

## Exported types
* ``ExcelFile``: A handle to an Excel file that was opened with ``openxl``.
* ``ExcelErrorCell``: Value returned for cells with an Excel error code.

## Further information
The ``ExcelReaders.tutorial`` help topic provides a tutorial to the package.

The package homepage is [https://github.com/davidanthoff/ExcelReaders.jl](https://github.com/davidanthoff/ExcelReaders.jl). Please
report any issues with the package there.
"""
ExcelReaders

"""
# ExcelReaders.tutorial

## Basic usage

The most basic usage is this:

````julia
using ExcelReaders
data = readxl("Filename.xlsx", "Sheet1!A1:C4")
````

This will return an array with all the data in the cell range A1 to
C4 on Sheet1 in the Excel file Filename.xlsx.

If you expect to read multiple ranges from the same Excel file you can get much
better performance by opening the Excel file only once:

````julia
using ExcelReaders
f = openxl("Filename.xlsx")
data1 = readxl(f, "Sheet1!A1:C4")
data2 = readxl(f, "Sheet2!B4:F10")
````

## Reading a whole sheet

The ``readxlsheet`` function reads complete Excel sheets, without a need to
specify precise range information. The most basic usage is

````julia
using ExcelReaders
data = readxlsheet("Filename.xlsx", "Sheet1")
````

This will read all content on Sheet1 in the file Filename.xlsx. Eventual blank
rows and columns at the top and left are skipped. ``readxlsheet`` takes a number
of optional keyword arguments:

- ``skipstartrows`` accepts either ``:blanks`` (default) or a positive integer.
With ``:blank`` any empty initial rows are skipped. An integer skips as many rows
as specified.
- ``skipstartcols`` accepts either ``:blanks`` (default) or a positive integer.
With ``:blank`` any empty initial columns are skipped. An integer skips as many
columns as specified.
- ``nrows`` accepts either ``:all`` (default) or a positive integer. With
``:all``, all rows (except skipped ones) are read. An integer specifies the exact
number of rows to be read.
- ``ncols`` accepts either ``:all`` (default) or a postiive integer. With ``:all``,
all columns (except skipped ones) are read. An integer specifies the exact number of columns to be read.

``readxlsheet`` also accepts an ``ExcelFile`` (as obtained from ``openxl``) as its
first argument.
"""
tutorial = nothing
