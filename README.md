# ExcelReaders

[![Build Status](https://github.com/queryverse/ExcelReaders.jl/actions/workflows/juliaci.yml/badge.svg?branch=main)](https://github.com/queryverse/ExcelReaders.jl/actions/workflows/juliaci.yml)

ExcelReaders is a package that provides functionality to read Excel files.

Both legacy xls files (Excel 97-2003) and modern xlsx files are supported.
The file format is detected from the content of the file, not its extension.
Under the hood legacy files are read via
[LibXLS.jl](https://github.com/queryverse/LibXLS.jl) (a wrapper of the
[libxls](https://github.com/libxls/libxls) C library) and modern files via
[XLSX.jl](https://github.com/JuliaData/XLSX.jl), so the package has no Python
or Java dependency.

## Installation

Use ``Pkg.add("ExcelReaders")`` in Julia to install ExcelReaders and its dependencies.

## Basic usage

The most basic usage is this:

````julia
using ExcelReaders

data = readxl("Filename.xls", "Sheet1!A1:C4")
````

This will return an array with all the data in the cell range A1 to C4 on
Sheet1 in the Excel file Filename.xls.

If you expect to read multiple ranges from the same Excel file you can get much
better performance by opening the Excel file only once:

````julia
using ExcelReaders

f = openxl("Filename.xls")

data1 = readxl(f, "Sheet1!A1:C4")
data2 = readxl(f, "Sheet2!B4:F10")
````

## Reading a whole sheet

The ``readxlsheet`` function reads complete Excel sheets, without a need to specify precise range information. The most basic usage is

````julia
using ExcelReaders

data = readxlsheet("Filename.xls", "Sheet1")
````

This will read all content on Sheet1 in the file Filename.xls. Eventual blank rows and columns at the top and left are skipped. ``readxlsheet`` takes a number of optional keyword arguments:

- ``skipstartrows`` accepts either ``:blanks`` (default) or a positive integer. With ``:blank`` any empty initial rows are skipped. An integer skips as many rows as specified.
- ``skipstartcols`` accepts either ``:blanks`` (default) or a positive integer. With ``:blank`` any empty initial columns are skipped. An integer skips as many columns as specified.
- ``nrows`` accepts either ``:all`` (default) or a positive integer. With ``:all``, all rows (except skipped ones) are read. An integer specifies the exact number of rows to be read.
- ``ncols`` accepts either ``:all`` (default) or a postiive integer. With ``:all``, all columns (except skipped ones) are read. An integer specifies the exact number of columns to be read.

``readxlsheet`` also accepts an ExcelFile (as obtained from ``openxl``) as its first argument.

## Defined names

For xlsx files, workbook-level defined names can be listed and read:

````julia
f = openxl("Filename.xlsx")

readxlnames(f)             # all defined names
readxlrange(f, "MyRange")  # content of the range a name refers to
````

Defined names are not available for legacy xls files, because the underlying
C library does not expose them.

## Closing a file

``close(f)`` releases the resources of an ``ExcelFile`` obtained from
``openxl`` (for xlsx files this is a no-op, for xls files it closes the
underlying C library handle; files also close themselves when garbage
collected).

## Limitations

* Reading only — for writing xlsx files use
  [XLSX.jl](https://github.com/JuliaData/XLSX.jl) or
  [ExcelFiles.jl](https://github.com/queryverse/ExcelFiles.jl).
* Cell formulas are not exposed, only their cached results.
* Reading from a byte buffer or ``IO`` is not supported, only from files.

## Alternatives

[XLSX.jl](https://github.com/JuliaData/XLSX.jl) provides excellent, more
fully-featured support for modern Excel files (including write support), and
is used by this package as its xlsx backend.
