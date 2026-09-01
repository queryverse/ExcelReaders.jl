module ExcelReaders

using DataValues, Dates

import LibXLS, XLSX

export openxl, readxl, readxlsheet, ExcelErrorCell, ExcelFile, readxlnames, readxlrange

include("package_documentation.jl")

"""
    ExcelFile

A handle to an open Excel file.

You can create an instance of an ``ExcelFile`` by calling ``openxl``.
"""
mutable struct ExcelFile
    workbook::Union{LibXLS.Workbook,XLSX.XLSXFile}
    filename::String
end

"""
    ExcelErrorCell

An Excel cell that has an Excel error.

You cannot create ``ExcelErrorCell`` objects, they are returned if a cell in an
Excel file has an Excel error. ``errorcode`` is the BIFF error code.
"""
mutable struct ExcelErrorCell
    errorcode::Int
end

const ERROR_CODE_TEXTS = Dict{Int,String}(
    0 => "#NULL!",
    7 => "#DIV/0!",
    15 => "#VALUE!",
    23 => "#REF!",
    29 => "#NAME?",
    36 => "#NUM!",
    42 => "#N/A",
    43 => "#GETTING_DATA",
)

const ERROR_TEXT_CODES = Dict{String,Int}(v => k for (k, v) in ERROR_CODE_TEXTS)

function Base.show(io::IO, o::ExcelFile)
    print(io, "ExcelFile <$(o.filename)>")
end

function Base.show(io::IO, o::ExcelErrorCell)
    print(io, get(ERROR_CODE_TEXTS, o.errorcode, "#ERROR($(o.errorcode))"))
end

const OLE2_FILE_HEADER = [0xd0, 0xcf, 0x11, 0xe0] # legacy xls (BIFF)
const ZIP_FILE_HEADER = [0x50, 0x4b, 0x03, 0x04]  # xlsx (OOXML)

"""
    openxl(filename)

Open the Excel file ``filename`` and return an ``ExcelFile`` handle.

The returned ``ExcelFile`` handle can later be passed as the first argument to
``readxl`` or ``readxslsheet`` to read from that file. If you will call either
of those functions more than once, performance will be better if you open the
file only once with ``openxl``.

Both legacy xls files and modern xlsx files are supported; the file format is
detected from the content of the file, not its extension.

# Example
````julia
f = openxl("filename.xls")
data = readxl(f, "Sheet1!A1:C4")
````
"""
function openxl(filename::AbstractString)
    isfile(filename) || error("File $filename not found.")

    header = open(io -> Base.read(io, 4), filename)

    if header == OLE2_FILE_HEADER
        return ExcelFile(LibXLS.openxls(filename), basename(filename))
    elseif header == ZIP_FILE_HEADER
        return ExcelFile(XLSX.readxlsx(filename), basename(filename))
    else
        error("$filename is not a valid Excel file.")
    end
end

function Base.close(file::ExcelFile)
    file.workbook isa LibXLS.Workbook && close(file.workbook)
    return nothing
end

sheetnames(file::ExcelFile) = sheetnames(file.workbook)
sheetnames(wb::LibXLS.Workbook) = LibXLS.sheetnames(wb)
sheetnames(wb::XLSX.XLSXFile) = XLSX.sheetnames(wb)

sheet_handle(file::ExcelFile, sheetname::AbstractString) = sheet_handle(file.workbook, sheetname)

function sheet_handle(wb::LibXLS.Workbook, sheetname::AbstractString)
    LibXLS.is_valid_sheetname(wb, sheetname) || error("Sheet $sheetname not found.")
    return LibXLS.getworksheet(wb, sheetname)
end

function sheet_handle(wb::XLSX.XLSXFile, sheetname::AbstractString)
    XLSX.hassheet(wb, sheetname) || error("Sheet $sheetname not found.")
    return wb[sheetname]
end

sheet_dims(ws::LibXLS.Worksheet) = size(ws)

function sheet_dims(ws::XLSX.Worksheet)
    dim = XLSX.get_dimension(ws)
    dim === nothing && return (0, 0)
    return (dim.stop.row_number, dim.stop.column_number)
end

# Map a backend cell value into the ExcelReaders vocabulary: NA for blank
# cells (and cells holding an empty string, as in previous versions), Float64
# for all numbers, String, Bool, DateTime, Time and ExcelErrorCell.
normalize_value(::Missing) = NA
normalize_value(v::AbstractString) = isempty(v) ? NA : String(v)
normalize_value(v::Bool) = v
normalize_value(v::Real) = Float64(v)
normalize_value(v::Date) = DateTime(v)
normalize_value(v::LibXLS.CellError) = ExcelErrorCell(Int(v.code))
normalize_value(v) = v

function cell_value(ws::LibXLS.Worksheet, row::Integer, col::Integer)
    nrows, ncols = size(ws)
    (1 <= row <= nrows && 1 <= col <= ncols) || return NA
    return normalize_value(ws[row, col])
end

function cell_value(ws::XLSX.Worksheet, row::Integer, col::Integer)
    (1 <= row && 1 <= col) || return NA
    cell = XLSX.getcell(ws, XLSX.CellRef(row, col))
    cell isa XLSX.EmptyCell && return NA
    if XLSX.iserror(cell)
        errortext = XLSX.get_error_string(XLSX.getval(cell))
        return ExcelErrorCell(get(ERROR_TEXT_CODES, errortext, -1))
    end
    return normalize_value(XLSX.getdata(ws, cell))
end

isblank(v) = v isa DataValue && DataValues.isna(v)

"""
    readxlsheet(file, sheet; skipstartrows=:blanks, skipstartcols=:blanks, nrows=:all, ncols=:all)

Read a whole sheet from an Excel file and return its content as a matrix.
`file` is either a filename or an `ExcelFile` from [`openxl`](@ref); `sheet`
is a sheet name or (1-based) index.

Blank rows and columns at the top and left are skipped by default. The
keyword arguments control the range that is read:

- `skipstartrows`/`skipstartcols`: `:blanks` (default) skips empty initial
  rows/columns; an integer skips exactly that many.
- `nrows`/`ncols`: `:all` (default) reads everything after the skipped
  rows/columns; an integer reads exactly that many.
"""
function readxlsheet(filename::AbstractString, sheetindex::Int; args...)
    file = openxl(filename)
    return readxlsheet(file, sheetindex; args...)
end

function readxlsheet(file::ExcelFile, sheetindex::Int; args...)
    return readxlsheet(file, sheetnames(file)[sheetindex]; args...)
end

function readxlsheet(filename::AbstractString, sheetname::AbstractString; args...)
    file = openxl(filename)
    return readxlsheet(file, sheetname; args...)
end

function readxlsheet(file::ExcelFile, sheetname::AbstractString; args...)
    ws = sheet_handle(file, sheetname)
    startrow, startcol, endrow, endcol = convert_args_to_row_col(ws; args...)

    return readxl_internal(ws, startrow, startcol, endrow, endcol)
end

# Function converts "relative" range like skip rows/cols and size of range to "absolute" from row/col to row/col
function convert_args_to_row_col(ws; skipstartrows::Union{Int,Symbol}=:blanks, skipstartcols::Union{Int,Symbol}=:blanks, nrows::Union{Int,Symbol}=:all, ncols::Union{Int,Symbol}=:all)
    isa(skipstartrows, Symbol) && skipstartrows != :blanks && error("Only :blank or an integer is a valid argument for skipstartrows")
    isa(skipstartrows, Int) && skipstartrows < 0 && error("Can't skip a negative number of rows")
    isa(skipstartcols, Symbol) && skipstartcols != :blanks && error("Only :blank or an integer is a valid argument for skipstartcols")
    isa(skipstartcols, Int) && skipstartcols < 0 && error("Can't skip a negative number of columns")
    isa(nrows, Symbol) && nrows != :all && error("Only :all or an integer is a valid argument for nrows")
    isa(nrows, Int) && nrows < 0 && error("nrows should be :all or positive")
    isa(ncols, Symbol) && ncols != :all && error("Only :all or an integer is a valid argument for ncols")
    isa(ncols, Int) && ncols < 0 && error("ncols should be :all or positive")

    sheet_rows, sheet_cols = sheet_dims(ws)

    if skipstartrows == :blanks
        startrow = -1
        for cur_row in 1:sheet_rows, cur_col in 1:sheet_cols
            if !isblank(cell_value(ws, cur_row, cur_col))
                startrow = cur_row
                break
            end
        end
        if startrow == -1
            error("Sheet has no data")
        else
            skipstartrows = startrow - 1
        end
    else
        startrow = 1 + skipstartrows
    end

    if skipstartcols == :blanks
        startcol = -1
        for cur_col in 1:sheet_cols, cur_row in 1:sheet_rows
            if !isblank(cell_value(ws, cur_row, cur_col))
                startcol = cur_col
                break
            end
        end
        if startcol == -1
            error("Sheet has no data")
        else
            skipstartcols = startcol - 1
        end
    else
        startcol = 1 + skipstartcols
    end

    if nrows == :all
        endrow = sheet_rows
    else
        endrow = nrows + skipstartrows
    end

    if ncols == :all
        endcol = sheet_cols
    else
        endcol = ncols + skipstartcols
    end

    return startrow, startcol, endrow, endcol
end

function colnum(col::AbstractString)
    cl = uppercase(col)
    r = 0
    for c in cl
        r = (r * 26) + (c - 'A' + 1)
    end
    return r
end

function convert_ref_to_sheet_row_col(range::AbstractString)
    r = r"('?[^']+'?|[^!]+)!([A-Za-z]*)(\d*)(:([A-Za-z]*)(\d*))?"
    m = match(r, range)
    m === nothing && error("Invalid Excel range specified.")
    sheetname = String(m.captures[1])
    startrow = parse(Int, m.captures[3])
    startcol = colnum(m.captures[2])
    if m.captures[4] === nothing
        endrow = startrow
        endcol = startcol
    else
        endrow = parse(Int, m.captures[6])
        endcol = colnum(m.captures[5])
    end
    if (startrow > endrow) || (startcol > endcol)
        error("Please provide rectangular region from top left to bottom right corner")
    end
    return sheetname, startrow, startcol, endrow, endcol
end

"""
    readxl(file, range)

Read the given range from an Excel file and return its content as a matrix
(or a single value for a single-cell range). `file` is either a filename or
an `ExcelFile` from [`openxl`](@ref); `range` is a full Excel range
specification such as `"Sheet1!A1:C4"`.
"""
function readxl(filename::AbstractString, range::AbstractString)
    excelfile = openxl(filename)

    readxl(excelfile, range)
end

function readxl(file::ExcelFile, range::AbstractString)
    sheetname, startrow, startcol, endrow, endcol = convert_ref_to_sheet_row_col(range)
    ws = sheet_handle(file, sheetname)
    readxl_internal(ws, startrow, startcol, endrow, endcol)
end

function readxl_internal(file::ExcelFile, sheetname::AbstractString, startrow::Integer, startcol::Integer, endrow::Integer, endcol::Integer)
    return readxl_internal(sheet_handle(file, sheetname), startrow, startcol, endrow, endcol)
end

function readxl_internal(ws, startrow::Integer, startcol::Integer, endrow::Integer, endcol::Integer)
    if startrow == endrow && startcol == endcol
        return cell_value(ws, startrow, startcol)
    else
        data = Array{Any}(undef, endrow - startrow + 1, endcol - startcol + 1)

        for row in startrow:endrow
            for col in startcol:endcol
                data[row - startrow + 1, col - startcol + 1] = cell_value(ws, row, col)
            end
        end

        return data
    end
end

"""
    readxlnames(f::ExcelFile)

Return the workbook-level defined names in the Excel file, sorted
alphabetically. Only supported for xlsx files; for legacy xls files an error
is thrown, because the underlying C library does not expose defined names.
"""
function readxlnames(f::ExcelFile)
    f.workbook isa XLSX.XLSXFile || error("Defined names are not supported for legacy xls files.")
    return sort!(collect(keys(f.workbook.workbook.workbook_names)))
end

"""
    readxlrange(f::ExcelFile, name)

Read the range that the defined name `name` refers to and return its content
(a matrix, or a single value for a single-cell name). Only supported for
xlsx files; for legacy xls files an error is thrown.
"""
function readxlrange(f::ExcelFile, range::AbstractString)
    f.workbook isa XLSX.XLSXFile || error("Defined names are not supported for legacy xls files.")
    data = XLSX.getdata(f.workbook, range)
    return data isa AbstractArray ? map(normalize_value, data) : normalize_value(data)
end

end # module
