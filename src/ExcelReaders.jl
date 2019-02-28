module ExcelReaders

using PyCall, DataValues, Dates

export openxl, readxl, readxlsheet, ExcelErrorCell, ExcelFile, readxlnames, readxlrange

const xlrd  = PyNULL()

include("package_documentation.jl")

function __init__()
    copy!(xlrd, pyimport_conda("xlrd", "xlrd"))
end

"""
    ExcelFile

A handle to an open Excel file.

You can create an instance of an ``ExcelFile`` by calling ``openxl``.
"""
mutable struct ExcelFile
    workbook::PyObject
    filename::AbstractString
end

"""
    ExcelErrorCell

An Excel cell that has an Excel error.

You cannot create ``ExcelErrorCell`` objects, they are returned if a cell in an
Excel file has an Excel error.
"""
mutable struct ExcelErrorCell
    errorcode::Int
end

function Base.show(io::IO, o::ExcelFile)
    print(io, "ExcelFile <$(o.filename)>")
end

function Base.show(io::IO, o::ExcelErrorCell)
    print(io, xlrd.error_text_from_code[o.errorcode])
end

"""
    openxl(filename)

Open the Excel file ``filename`` and return an ``ExcelFile`` handle.

The returned ``ExcelFile`` handle can later be passed as the first argument to
``readxl`` or ``readxslsheet`` to read from that file. If you will call either
of those functions more than once, performance will be better if you open the
file only once with ``openxl``.

# Example
````julia
f = openxl("filename.xlsx")
data = readxl(f, "Sheet1!A1:C4")
````
"""
function openxl(filename::AbstractString)
    wb = xlrd.open_workbook(filename)
    return ExcelFile(wb, basename(filename))
end

function readxlsheet(filename::AbstractString, sheetindex::Int; args...)
    file = openxl(filename)
    return readxlsheet(file, sheetindex; args...)
end

function readxlsheet(file::ExcelFile, sheetindex::Int; args...)
    sheetnames = file.workbook.sheet_names()
    return readxlsheet(file, sheetnames[sheetindex]; args...)
end

function readxlsheet(filename::AbstractString, sheetname::AbstractString; args...)
    file = openxl(filename)
    return readxlsheet(file, sheetname; args...)
end

function readxlsheet(file::ExcelFile, sheetname::AbstractString; args...)
    sheet = file.workbook.sheet_by_name(sheetname)
    startrow, startcol, endrow, endcol = convert_args_to_row_col(sheet; args...)

    data = readxl_internal(file, sheetname, startrow, startcol, endrow, endcol)

    return data
end

# Function converts "relative" range like skip rows/cols and size of range to "absolute" from row/col to row/col
function convert_args_to_row_col(sheet;skipstartrows::Union{Int,Symbol}=:blanks, skipstartcols::Union{Int,Symbol}=:blanks, nrows::Union{Int,Symbol}=:all, ncols::Union{Int,Symbol}=:all)
    isa(skipstartrows, Symbol) && skipstartrows!=:blanks && error("Only :blank or an integer is a valid argument for skipstartrows")
    isa(skipstartrows, Int) && skipstartrows<0 && error("Can't skip a negative number of rows")
    isa(skipstartcols, Symbol) && skipstartcols!=:blanks && error("Only :blank or an integer is a valid argument for skipstartcols")
    isa(skipstartcols, Int) && skipstartcols<0 && error("Can't skip a negative number of columns")
    isa(nrows, Symbol) && nrows!=:all && error("Only :all or an integer is a valid argument for nrows")
    isa(nrows, Int) && nrows<0 && error("nrows should be :all or positive")
    isa(ncols, Symbol) && ncols!=:all && error("Only :all or an integer is a valid argument for ncols")
    isa(ncols, Int) && ncols<0 && error("ncols should be :all or positive")
    sheet_rows = sheet.nrows
    sheet_cols = sheet.ncols

    cell_value = sheet.cell_value

    if skipstartrows==:blanks
        startrow = -1
        for cur_row in 1:sheet_rows, cur_col in 1:sheet_cols
            cellval = cell_value(cur_row-1,cur_col-1)
            if cellval!=""
                startrow = cur_row
                break
            end
        end
        if startrow==-1
            error("Sheet has no data")
        else
            skipstartrows = startrow - 1
        end
    else
        startrow = 1 + skipstartrows
    end

    if skipstartcols==:blanks
        startcol = -1
        for cur_col in 1:sheet_cols, cur_row in 1:sheet_rows
            cellval = cell_value(cur_row-1,cur_col-1)
            if cellval!=""
                startcol = cur_col
                break
            end
        end
        if startcol==-1
            error("Sheet has no data")
        else
            skipstartcols = startcol - 1
        end
    else
        startcol = 1 + skipstartcols
    end

    if nrows==:all
        endrow = sheet_rows
    else
        endrow = nrows + skipstartrows
    end

    if ncols==:all
        endcol = sheet_cols
    else
        endcol = ncols + skipstartcols
    end

    return startrow, startcol, endrow, endcol
end

function colnum(col::AbstractString)
    cl=uppercase(col)
    r=0
    for c in cl
        r = (r * 26) + (c - 'A' + 1)
    end
    return r
end

function convert_ref_to_sheet_row_col(range::AbstractString)
    r=r"('?[^']+'?|[^!]+)!([A-Za-z]*)(\d*)(:([A-Za-z]*)(\d*))?"
    m=match(r, range)
    m==nothing && error("Invalid Excel range specified.")
    sheetname=String(m.captures[1])
    startrow=parse(Int,m.captures[3])
    startcol=colnum(m.captures[2])
    if m.captures[4]==nothing
        endrow=startrow
        endcol=startcol
    else
        endrow=parse(Int,m.captures[6])
        endcol=colnum(m.captures[5])
    end
    if (startrow > endrow ) || (startcol>endcol)
        error("Please provide rectangular region from top left to bottom right corner")
    end
    return sheetname, startrow, startcol, endrow, endcol
end

function readxl(filename::AbstractString, range::AbstractString)
    excelfile = openxl(filename)

    readxl(excelfile, range)
end

function readxl(file::ExcelFile, range::AbstractString)
    sheetname, startrow, startcol, endrow, endcol = convert_ref_to_sheet_row_col(range)
    readxl_internal(file, sheetname, startrow, startcol, endrow, endcol)
end

function get_cell_value(ws, row, col, wb)
    cellval = ws.cell_value(row-1,col-1)
    if cellval==""
        return NA
    else
        celltype = ws.cell_type(row-1,col-1)
        if celltype == xlrd.XL_CELL_TEXT
            return convert(String, cellval)
        elseif celltype == xlrd.XL_CELL_NUMBER
            return convert(Float64, cellval)
        elseif celltype == xlrd.XL_CELL_DATE
            date_year,date_month,date_day,date_hour,date_minute,date_sec = xlrd.xldate_as_tuple(cellval, wb.datemode)
            if date_month==0
                return Time(date_hour, date_minute, date_sec)
            else
                return DateTime(date_year, date_month, date_day, date_hour, date_minute, date_sec)
            end
        elseif celltype == xlrd.XL_CELL_BOOLEAN
            return convert(Bool, cellval)
        elseif celltype == xlrd.XL_CELL_ERROR
            return ExcelErrorCell(cellval)
        else
            error("Unknown cell type")
        end
    end
end

function readxl_internal(file::ExcelFile, sheetname::AbstractString, startrow::Integer, startcol::Integer, endrow::Integer, endcol::Integer)
    wb = file.workbook
    ws = wb.sheet_by_name(sheetname)

    if startrow==endrow && startcol==endcol
        return get_cell_value(ws, startrow, startcol, wb)
    else

        data = Array{Any}(undef,endrow-startrow+1,endcol-startcol+1)

        for row in startrow:endrow
            for col in startcol:endcol
                data[row-startrow+1, col-startcol+1] = get_cell_value(ws, row, col, wb)
            end
        end

        return data
    end
end

function readxlnames(f::ExcelFile)
    return [lowercase(i.name) for i in f.workbook.name_obj_list if i.hidden==0]
end

function readxlrange(f::ExcelFile, range::AbstractString)
    name = f.workbook.name_map[lowercase(range)]
    if length(name)!=1
        error("More than one reference per name, this case is not yet handled by ExcelReaders.")
    end

    formula_text = name[1].formula_text
    formula_text = replace(formula_text, "\$"=>"")
    formula_text = replace(formula_text, "'"=>"")

    return readxl(f, formula_text)
end

end # module
