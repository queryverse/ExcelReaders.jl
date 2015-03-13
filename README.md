# ExcelReaders

[![Build Status](https://travis-ci.org/davidanthoff/ExcelReaders.jl.svg?branch=master)](https://travis-ci.org/davidanthoff/ExcelReaders.jl)
[![Coverage Status](https://coveralls.io/repos/davidanthoff/ExcelReaders.jl/badge.svg)](https://coveralls.io/r/davidanthoff/ExcelReaders.jl)

ExcelReaders is a beta package that provides functionality to read Excel files.

## Installation

You will need to have the Python xlrd library installed on your machine in order to use ExcelReaders.

Once xlrd is installed, then you can just use ``Pkg.clone("ExcelReaders")`` in Julia to install ExcelReaders and its dependencies.

## Basic usage

The most basic usage is this:

````
using ExcelReaders

data = readxl("Filename.xlsx", "Sheet1!A1:C4")
````

This will return a ``DataMatrix{Any}`` with all the data in the cell range A1 to C4 on Sheet1 in the Excel file Filename.xlsx.

If you expect to read multiple ranges from the same Excel file you can get much better performance by opening the Excel file only once:

````
using ExcelReaders

f = openxl("Filename.xlsx")

data1 = readxl(f, "Sheet1!A1:C4")
data2 = readxl(f, "Sheet2!B4:F10")
````

## Reading into a DataFrame

To read into a DataFrame:

````
using ExcelReaders
using DataFrames

df = readxl(DataFrame, "Filename.xlsx", "Sheet1!A1:C4")
````

This code will use the first row in the range A1:C4 as the column names in the DataFrame.

To read in data without a header row use

````
df = readxl(DataFrame, "Filename.xlsx", "Sheet1!A1:C4", header=false)
````

This will auto-generate column names. Alternatively you can specify your own names:

````
df = readxl(DataFrame, "Filename.xlsx", "Sheet1!A1:C4", 
            header=false, colnames=[:name1, :name2, :name3])
````

You can also combine ``header=true`` and a custom ``colnames`` list, in that case the first row in the specified range will just be skipped.
