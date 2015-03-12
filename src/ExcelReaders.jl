module ExcelReaders

import Base.show

export openxl, readxl

using PyCall, DataArrays, DataFrames

@pyimport xlrd

type ExcelFile
	workbook::PyObject
	filename::String
end

function show(io::IO, o::ExcelFile)
	print(io, "ExcelFile <$(o.filename)>")
end

function openxl(filename::String)
	wb = xlrd.open_workbook(filename)
	return ExcelFile(wb, basename(filename))
end

function colnum(col::String)
	cl=uppercase(col)
	r=0
	for c in cl
		r = (r * 26) + (c - 'A' + 1)
	end
	return r-1
end

function convert_ref_to_sheet_row_col(range::String)
    r=r"('?[^']+'?|[^!]+)!([A-Za-z]*)(\d*):([A-Za-z]*)(\d*)"
    m=match(r, range)
    sheetname=string(m.captures[1])
    startrow=int(m.captures[3])-1
    startcol=colnum(m.captures[2])
    endrow=int(m.captures[5])-1
    endcol=colnum(m.captures[4])
    if (startrow > endrow ) || (startcol>endcol)
		error("Please provide rectangular region from top left to bottom right corner")
    end

    return sheetname, startrow, startcol, endrow, endcol
end


function readxl(filename::String, range::String)
	excelfile = openxl(filename)

	readxl(excelfile, range)
end

function readxl(file::ExcelFile, range::String)
	sheetname, startrow, startcol, endrow, endcol = convert_ref_to_sheet_row_col(range)

	readxl(file, sheetname, startrow, startcol, endrow, endcol)
end

function readxl(file::ExcelFile, sheetname::String, startrow::Int, startcol::Int, endrow::Int, endcol::Int)
	wb = file.workbook
	ws = wb[:sheet_by_name](sheetname)

	data = DataArray(Any, endrow-startrow+1,endcol-startcol+1)

	for row in startrow:endrow
		for col in startcol:endcol
			cellval = ws[:cell_value](row,col)
			if cellval == ""
				data[row-startrow+1, col-startcol+1] = NA
			else
				celltype = ws[:cell_type](row,col)
				if celltype == xlrd.XL_CELL_TEXT
					data[row-startrow+1, col-startcol+1] = convert(String, cellval)
				elseif celltype == xlrd.XL_CELL_NUMBER
					data[row-startrow+1, col-startcol+1] = convert(Float64, cellval)
				elseif celltype == xlrd.XL_CELL_DATE
					error("Support for date cells not yet implemented")
				elseif celltype == xlrd.XL_CELL_BOOLEAN
					data[row-startrow+1, col-startcol+1] = convert(Bool, cellval)
				elseif celltype == xlrd.XL_CELL_ERROR
					error("Support for error cells not yet implemented")
				else
					error("Unknown cell type")
				end
			end
		end
	end
	
	return data
end

function readxl(::Type{DataFrame}, filename::String, range::String; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
	excelfile = openxl(filename)

	readxl(DataFrame, excelfile, range, header=header, colnames=colnames)
end

function readxl(::Type{DataFrame}, file::ExcelFile, range::String; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
	sheetname, startrow, startcol, endrow, endcol = convert_ref_to_sheet_row_col(range)

	readxl(DataFrame, file, sheetname, startrow, startcol, endrow, endcol, header=header, colnames=colnames)
end

function readxl(::Type{DataFrame}, file::ExcelFile, sheetname::String, startrow::Int, startcol::Int, endrow::Int, endcol::Int; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
	data = readxl(file, sheetname, startrow, startcol, endrow, endcol)

	nrow, ncol = size(data)

	if length(colnames)==0
		if header
			colnames = convert(Array{Symbol},vec(data[1,:]))
		else
			colnames = DataFrames.gennames(ncol)
		end
	elseif length(colnames)!=ncol
		error("Length of colnames must equal number of columns in selected range")
	end

	columns = Array(Any, ncol)

	for i=1:ncol
		if header
			values = data[2:end,i]
		else
			values = data[:,i]
		end

		# Check whether all non-NA values in this column
		# are of the same type
		all_one_type = true
		found_first_type = false
		type_of_el = Any
		for val=values
			if !found_first_type
				if !isna(val)
					type_of_el = typeof(val)
					found_first_type = true
				end
			elseif !isna(val) && (typeof(val)!=type_of_el)
				all_one_type = false
				break
			end
		end
		
		if all_one_type
			columns[i] = convert(DataArray{type_of_el},values)
		else
			columns[i] = values
		end
	end

	df = DataFrame(columns, colnames)

	return df
end

end # module
