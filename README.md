# ExcelReaders

[![Build Status](https://travis-ci.org/queryverse/ExcelReaders.jl.svg?branch=master)](https://travis-ci.org/queryverse/ExcelReaders.jl)
[![Build status](https://ci.appveyor.com/api/projects/status/v7b60gfrg65qkqt5/branch/master?svg=true)](https://ci.appveyor.com/project/queryverse/excelreaders-jl/branch/master)
[![Coverage Status](https://coveralls.io/repos/queryverse/ExcelReaders.jl/badge.svg)](https://coveralls.io/r/queryverse/ExcelReaders.jl)
[![codecov](https://codecov.io/gh/queryverse/ExcelReaders.jl/branch/master/graph/badge.svg)](https://codecov.io/gh/queryverse/ExcelReaders.jl)
[![ExcelReaders](http://pkg.julialang.org/badges/ExcelReaders_0.3.svg)](http://pkg.julialang.org/?pkg=ExcelReaders&ver=0.3)
[![ExcelReaders](http://pkg.julialang.org/badges/ExcelReaders_0.4.svg)](http://pkg.julialang.org/?pkg=ExcelReaders&ver=0.4)
[![ExcelReaders](http://pkg.julialang.org/badges/ExcelReaders_0.5.svg)](http://pkg.julialang.org/?pkg=ExcelReaders&ver=0.5)
[![ExcelReaders](http://pkg.julialang.org/badges/ExcelReaders_0.6.svg)](http://pkg.julialang.org/?pkg=ExcelReaders&ver=0.5)

ExcelReaders is a package that provides functionality to read Excel files.

**WARNING**: Version v0.9.0 removed all support for [DataFrames.jl](https://github.com/JuliaData/DataFrames.jl)
from this package. The [ExcelFiles.jl](https://github.com/queryverse/ExcelFiles.jl)
package now provides functionality to read data from an Excel file into
a ``DataFrame`` (or any other table type), and users are encouraged to use
that package for tabular data going forward. Version v0.9.0 also no longer
uses [DataArrays.jl](https://github.com/JuliaStats/DataArrays.jl), but instead
is based on [DataValues.jl](https://github.com/queryverse/DataValues.jl).

## Installation

Use ``Pkg.add("ExcelReaders")`` in Julia to install ExcelReaders and its dependencies.

The package uses the Python xlrd library. If either Python or the xlrd package are not installed on your system, the package will use the [conda.jl](https://github.com/Luthaf/Conda.jl) package to install all necessary dependencies automatically.

## Alternatives

The [Taro](https://github.com/aviks/Taro.jl) package also provides Excel file reading functionality. The main difference between the two packages (in terms of Excel functionality) is that ExcelReaders uses the Python package [xlrd](https://github.com/python-excel/xlrd) for its processing, whereas Taro uses the Java packages Apache [Tika](http://tika.apache.org/) and Apache [POI](http://poi.apache.org/).

## Basic usage

The most basic usage is this:

````julia
using ExcelReaders

data = readxl("Filename.xlsx", "Sheet1!A1:C4")
````

This will return an array with all the data in the cell range A1 to C4 on Sheet1 in the Excel file Filename.xlsx.

If you expect to read multiple ranges from the same Excel file you can get much better performance by opening the Excel file only once:

````julia
using ExcelReaders

f = openxl("Filename.xlsx")

data1 = readxl(f, "Sheet1!A1:C4")
data2 = readxl(f, "Sheet2!B4:F10")
````

## Reading a whole sheet

The ``readxlsheet`` function reads complete Excel sheets, without a need to specify precise range information. The most basic usage is

````julia
using ExcelReaders

data = readxlsheet("Filename.xlsx", "Sheet1")
````

This will read all content on Sheet1 in the file Filename.xlsx. Eventual blank rows and columns at the top and left are skipped. ``readxlsheet`` takes a number of optional keyword arguments:

- ``skipstartrows`` accepts either ``:blanks`` (default) or a positive integer. With ``:blank`` any empty initial rows are skipped. An integer skips as many rows as specified.
- ``skipstartcols`` accepts either ``:blanks`` (default) or a positive integer. With ``:blank`` any empty initial columns are skipped. An integer skips as many columns as specified.
- ``nrows`` accepts either ``:all`` (default) or a positive integer. With ``:all``, all rows (except skipped ones) are read. An integer specifies the exact number of rows to be read.
- ``ncols`` accepts either ``:all`` (default) or a postiive integer. With ``:all``, all columns (except skipped ones) are read. An integer specifies the exact number of columns to be read.

``readxlsheet`` also accepts an ExcelFile (as obtained from ``openxl``) as its first argument.
