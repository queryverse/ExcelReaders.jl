module ExcelReaders

import Base.show

export openxl, readxl, readxlsheet, ExcelErrorCell

using PyCall, DataArrays, DataFrames, Dates

@pyimport xlrd

type ExcelFile
	workbook::PyObject
	filename::String
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

function openxl(filename::String)
	wb = xlrd.open_workbook(filename)
	return ExcelFile(wb, basename(filename))
end

function readxlsheet(filename::String, sheetindex::Int; args...)
	file = openxl(filename)
	return readxlsheet(file, sheetindex; args...)
end

function readxlsheet(file::ExcelFile, sheetindex::Int; args...)
	sheetnames = file.workbook[:sheet_names]()
	return readxlsheet(file, sheetnames[sheetindex]; args...)
end

function readxlsheet(filename::String, sheetname::String; args...)
	file = openxl(filename)
	return readxlsheet(file, sheetname; args...)
end

function readxlsheet(file::ExcelFile, sheetname::String; skipstartrows::Int=0, skipstartcols::Int=0, nrows::Int=-1, ncols::Int=-1, skipblanks::Symbol=:notset, skipblankrows::Symbol=:notset, skipblankcols::Symbol=:notset)
	if skipblanks != :notset
		if !(skipblanks==:none || skipblanks==:start || skipblanks==:all)
			error("Only :none, :start or :all are valid skipblanks arguments.")
		end

		if skipblankrows!=:notset || skipblankcols!=:notset
			error("skipblanks cannot be used simultaniously with either skipblankrows or skipblankcols.")
		end

		skipblankrows = skipblanks
		skipblankcols = skipblanks
	else
		if !(skipblankrows==:notset || skipblankrows==:none || skipblankrows==:start || skipblankrows==:all)
			error("Only :none, :start or :all are valid skipblankrows arguments.")
		end

		if !(skipblankcols==:notset || skipblankcols==:none || skipblankcols==:start || skipblankcols==:all)
			error("Only :none, :start or :all are valid skipblankcols arguments.")
		end

		if skipblankrows==:notset
			if skipstartrows==0 && nrows==-1
				skipblankrows = :start
			else
				skipblankrows = :none
			end
		end

		if skipblankcols==:notset
			if skipstartcols==0 && ncols==-1
				skipblankcols = :start
			else
				skipblankcols = :none
			end
		end
	end

	if skipblankrows!=:none && skipstartrows>0
		error("When skipstartrows is used, no option to skip blank rows can be used.")
	end

	if skipblankcols!=:none && skipstartcols>0
		error("When skipstartcols is used, no option to skip blank cols can be used.")
	end

	if skipblankrows!=:none && nrows!=-1
		error("When nrows is used, no option to skip blank rows can be used.")
	end

	if skipblankcols!=:none && ncols!=-1
		error("When ncols is used, no option to skip blank cols can be used.")
	end

	sheet = file.workbook[:sheet_by_name](sheetname)

	startrow = 1 + skipstartrows
	startcol = 1 + skipstartcols

	if nrows==-1
		endrow = sheet[:nrows] - skipstartrows
	else
		endrow = nrows + skipstartrows
	end

	if ncols==-1
		endcol = sheet[:ncols] - skipstartcols
	else
		endcol = ncols + skipstartcols
	end

	datawithblanks = readxl_internal(file, sheetname, startrow, startcol, endrow, endcol)

	if skipblankrows!=:none || skipblankcols!=:none
		nrows, ncols = size(datawithblanks)

		rows_to_keep = Array(Int,0)
		found_row_with_data = false
		for i=1:nrows
			if skipblankrows==:none || (skipblankrows==:start && found_row_with_data)
				push!(rows_to_keep, i)
			else
				for l=1:ncols
					if !isna(datawithblanks[i,l])
						push!(rows_to_keep, i)
						found_row_with_data = true
						break
					end
				end
			end
		end

		cols_to_keep = Array(Int,0)
		found_col_with_data = false
		for i=1:ncols
			if skipblankcols==:none || (skipblankcols==:start && found_col_with_data)
				push!(cols_to_keep,i)
			else
				for l=1:nrows
					if !isna(datawithblanks[l,i])
						push!(cols_to_keep, i)
						found_col_with_data = true
						break
					end
				end
			end
		end

		target_rows = length(rows_to_keep)
		target_cols = length(cols_to_keep)

		data = DataArray(Any, target_rows, target_cols)
		for i=1:target_rows
			for l=1:target_cols
				data[i,l] = datawithblanks[rows_to_keep[i], cols_to_keep[l]]
			end
		end

		return data
	else
		return datawithblanks
	end
end

function colnum(col::String)
	cl=uppercase(col)
	r=0
	for c in cl
		r = (r * 26) + (c - 'A' + 1)
	end
	return r
end

function convert_ref_to_sheet_row_col(range::String)
    r=r"('?[^']+'?|[^!]+)!([A-Za-z]*)(\d*):([A-Za-z]*)(\d*)"
    m=match(r, range)
    sheetname=string(m.captures[1])
    startrow=int(m.captures[3])
    startcol=colnum(m.captures[2])
    endrow=int(m.captures[5])
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

	readxl_internal(file, sheetname, startrow, startcol, endrow, endcol)
end

function readxl_internal(file::ExcelFile, sheetname::String, startrow::Int, startcol::Int, endrow::Int, endcol::Int)
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
					data[row-startrow+1, col-startcol+1] = convert(String, cellval)
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

function readxl(::Type{DataFrame}, filename::String, range::String; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
	excelfile = openxl(filename)

	readxl(DataFrame, excelfile, range, header=header, colnames=colnames)
end

function readxl(::Type{DataFrame}, file::ExcelFile, range::String; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
	sheetname, startrow, startcol, endrow, endcol = convert_ref_to_sheet_row_col(range)

	readxl_internal(DataFrame, file, sheetname, startrow, startcol, endrow, endcol, header=header, colnames=colnames)
end

function readxl_internal(::Type{DataFrame}, file::ExcelFile, sheetname::String, startrow::Int, startcol::Int, endrow::Int, endcol::Int; header::Bool=true, colnames::Vector{Symbol}=Symbol[])
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
