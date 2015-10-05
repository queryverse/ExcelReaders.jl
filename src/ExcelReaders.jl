module ExcelReaders

using PyCall, DataArrays, DataFrames

import Base.show

export openxl, readxl, readxlsheet, ExcelErrorCell

@pyimport xlrd

type ExcelFile
	workbook::PyObject
	filename::UTF8String
end

type ExcelErrorCell
	errorcode::Int
end

# TODO Remove this type once there is a Time type in Dates
immutable Time
	hours::Int
	minutes::Int
	seconds::Int
end

function show(io::IO, o::ExcelFile)
	print(io, "ExcelFile <$(o.filename)>")
end

function show(io::IO, o::ExcelErrorCell)
	print(io, xlrd.error_text_from_code[o.errorcode])
end

function openxl(filename::AbstractString)
	wb = xlrd.open_workbook(filename)
	return ExcelFile(wb, basename(filename))
end

function readxlsheet(filename::AbstractString, sheetindex::Int; args...)
	file = openxl(filename)
	return readxlsheet(file, sheetindex; args...)
end

function readxlsheet(file::ExcelFile, sheetindex::Int; args...)
	sheetnames = file.workbook[:sheet_names]()
	return readxlsheet(file, sheetnames[sheetindex]; args...)
end

function readxlsheet(filename::AbstractString, sheetname::AbstractString; args...)
	file = openxl(filename)
	return readxlsheet(file, sheetname; args...)
end

function readxlsheet(file::ExcelFile, sheetname::AbstractString; skipstartrows::Union{Int,Symbol}=:blanks, skipstartcols::Union{Int,Symbol}=:blanks, nrows::Union{Int,Symbol}=:all, ncols::Union{Int,Symbol}=:all)
	isa(skipstartrows, Symbol) && skipstartrows!=:blanks && error("Only :blank or an integer is a valid argument for skipstartrows")
	isa(skipstartrows, Int) && skipstartrows<0 && error("Can't skip a negative number of rows")
	isa(skipstartcols, Symbol) && skipstartcols!=:blanks && error("Only :blank or an integer is a valid argument for skipstartcols")
	isa(skipstartcols, Int) && skipstartcols<0 && error("Can't skip a negative number of columns")
	isa(nrows, Symbol) && nrows!=:all && error("Only :all or an integer is a valid argument for nrows")
	isa(nrows, Int) && nrows<0 && error("nrows should be :all or positive")
	isa(ncols, Symbol) && ncols!=:all && error("Only :all or an integer is a valid argument for ncols")
	isa(ncols, Int) && ncols<0 && error("ncols should be :all or positive")

	sheet = file.workbook[:sheet_by_name](sheetname)
	sheet_rows = sheet[:nrows]
	sheet_cols = sheet[:ncols]

	cell_value = sheet[:cell_value]

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

	data = readxl_internal(file, sheetname, startrow, startcol, endrow, endcol)

	return data
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
    r=r"('?[^']+'?|[^!]+)!([A-Za-z]*)(\d*):([A-Za-z]*)(\d*)"
    m=match(r, range)
    sheetname=string(m.captures[1])
    startrow=parse(Int,m.captures[3])
    startcol=colnum(m.captures[2])
    endrow=parse(Int,m.captures[5])
    endcol=colnum(m.captures[4])
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

function readxl_internal(file::ExcelFile, sheetname::AbstractString, startrow::Int, startcol::Int, endrow::Int, endcol::Int)
	wb = file.workbook
	ws = wb[:sheet_by_name](sheetname)

	data = DataArray(Any, endrow-startrow+1,endcol-startcol+1)

	for row in startrow:endrow
		for col in startcol:endcol
			cellval = ws[:cell_value](row-1,col-1)
			if cellval == ""
				data[row-startrow+1, col-startcol+1] = NA
			else
				celltype = ws[:cell_type](row-1,col-1)
				if celltype == xlrd.XL_CELL_TEXT
					data[row-startrow+1, col-startcol+1] = convert(UTF8String, cellval)
				elseif celltype == xlrd.XL_CELL_NUMBER
					data[row-startrow+1, col-startcol+1] = convert(Float64, cellval)
				elseif celltype == xlrd.XL_CELL_DATE
					date_year,date_month,date_day,date_hour,date_minute,date_sec = xlrd.xldate_as_tuple(cellval, wb[:datemode])
					if date_month==0
						data[row-startrow+1, col-startcol+1] = Time(date_hour, date_minute, date_sec)
					else
						data[row-startrow+1, col-startcol+1] = DateTime(date_year, date_month, date_day, date_hour, date_minute, date_sec)	
					end
				elseif celltype == xlrd.XL_CELL_BOOLEAN
					data[row-startrow+1, col-startcol+1] = convert(Bool, cellval)
				elseif celltype == xlrd.XL_CELL_ERROR
					data[row-startrow+1, col-startcol+1] = ExcelErrorCell(cellval)
				else
					error("Unknown cell type")
				end
			end
		end
	end
	
	return data
end

function readxl(::Type{DataFrame}, filename::AbstractString, range::AbstractString; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
	excelfile = openxl(filename)

	readxl(DataFrame, excelfile, range, header=header, colnames=colnames)
end

function readxl(::Type{DataFrame}, file::ExcelFile, range::AbstractString; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
	sheetname, startrow, startcol, endrow, endcol = convert_ref_to_sheet_row_col(range)

	readxl_internal(DataFrame, file, sheetname, startrow, startcol, endrow, endcol, header=header, colnames=colnames)
end

function readxl_internal(::Type{DataFrame}, file::ExcelFile, sheetname::AbstractString, startrow::Int, startcol::Int, endrow::Int, endcol::Int; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
	data = readxl_internal(file, sheetname, startrow, startcol, endrow, endcol)

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
			vals = data[2:end,i]
		else
			vals = data[:,i]
		end

		# Check whether all non-NA values in this column
		# are of the same type
		all_one_type = true
		found_first_type = false
		type_of_el = Any
		NAs_present = false
		for val=vals
			if !found_first_type
				if !isna(val)
					type_of_el = typeof(val)
					found_first_type = true
				end
			elseif !isna(val) && (typeof(val)!=type_of_el)
				all_one_type = false
				if NAs_present
					break
				end
			end
			if isna(val)
				NAs_present = true
				if all_one_type == false
					break
				end
			end
		end
		
		if all_one_type
			if NAs_present
				# TODO use the following line instead of the shim once upstream
				# bug is fixed
				#columns[i] = convert(DataArray{type_of_el},vals)
				shim_newarray = DataArray(type_of_el, length(vals))
				for l=1:length(vals)
					shim_newarray[l] = vals[l]
				end
				columns[i] = shim_newarray
			else
				# TODO Decide whether this should be converted to Array instead of DataArray
				columns[i] = convert(DataArray{type_of_el},vals)
			end
		else
			columns[i] = vals
		end
	end

	df = DataFrame(columns, colnames)

	return df
end

end # module
