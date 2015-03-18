using ExcelReaders
using Base.Test
using PyCall
using DataArrays
using DataFrames
using Dates

# TODO Throw julia specific exceptions for these errors
@test_throws PyCall.PyError openxl("FileThatDoesNotExist.xlsx")
@test_throws PyCall.PyError openxl("runtests.jl")

filename = normpath(Pkg.dir("ExcelReaders"),"test", "TestData.xlsx")
file = openxl(filename)
@test file.filename == "TestData.xlsx"

buffer = IOBuffer()
show(buffer, file)
@test takebuf_string(buffer) == "ExcelFile <TestData.xlsx>"

for (k,v) in [0=>"#NULL!",7=>"#DIV/0!",23 => "#REF!",42=>"#N/A",29=>"#NAME?",36=>"#NUM!",15=>"#VALUE!"]
	errorcell = ExcelErrorCell(k)
	buffer = IOBuffer()
	show(buffer, errorcell)
	@test takebuf_string(buffer) == v
end


# Read into DataArray
for f in [file, filename]
	@test_throws ErrorException readxl(f, "Sheet1!C4:G3")
	@test_throws ErrorException readxl(f, "Sheet1!G2:B5")
	@test_throws ErrorException readxl(f, "Sheet1!G5:B2")

	data = readxl(f, "Sheet1!C3:N7")
	@test size(data) == (5,12)
	@test data[4,1] == 2.0
	@test data[2,2] == "A"
	@test data[2,3] == true
	@test isna(data[4,5])
	@test data[2,9] == Date(2015,3,3)
	@test data[3,9] == DateTime(2015,2,4,10,14)
	@test data[4,9] == DateTime(1988,4,9,0,0)
	@test data[5,9] == ExcelReaders.Time(15,2,0)
	@test data[3,10] == DateTime(1950,8,9,18,40)
	@test isna(data[5,10])
	@test isa(data[2,11], ExcelErrorCell)
	@test isa(data[3,11], ExcelErrorCell)
	@test isa(data[4,12], ExcelErrorCell)
	@test isna(data[5,12])

	df = readxl(DataFrame, f, "Sheet1!C3:N7")
	@test ncol(df) == 12
	@test nrow(df) == 4
	@test isa(df[symbol("Some Float64s")], DataVector{Float64})
	@test isa(df[symbol("Some Strings")], DataVector{UTF8String})
	@test isa(df[symbol("Some Bools")], DataVector{Bool})
	@test isa(df[symbol("Mixed column")], DataVector{Any})
	@test isa(df[symbol("Mixed with NA")], DataVector{Any})
	@test isa(df[symbol("Some dates")], DataVector{Any})
	@test isa(df[symbol("Dates with NA")], DataVector{Any})
	@test df[4,symbol("Some Float64s")] == 2.5
	@test df[4,symbol("Some Strings")] == "DDDD"
	@test df[4,symbol("Some Bools")] == true
	@test df[1,symbol("Mixed column")] == 2.0
	@test df[2,symbol("Mixed column")] == "EEEEE"
	@test df[3,symbol("Mixed column")] == false
	@test isna(df[3,symbol("Mixed with NA")])
	@test df[1,symbol("Float64 with NA")] == 3.
	@test isna(df[2,symbol("Float64 with NA")])
	@test df[1,symbol("String with NA")] == "FF"
	@test isna(df[2,symbol("String with NA")])
	@test df[2,symbol("Bool with NA")] == true
	@test isna(df[1,symbol("Bool with NA")])
	@test df[1,symbol("Dates with NA")] == Date(1965,4,3)
	@test df[2,symbol("Some dates")] == DateTime(2015,2,4,10,14)
	@test df[4,symbol("Some dates")] == ExcelReaders.Time(15,2,0)
	@test isna(df[4,symbol("Dates with NA")])
	# TODO Add a test that checks the error code, not just type
	@test isa(df[1,symbol("Some errors")], ExcelErrorCell)
	@test isna(df[4,symbol("Errors with NA")])


	df = readxl(DataFrame, f, "Sheet1!C4:N7", header=false)
	@test ncol(df) == 12
	@test nrow(df) == 4
	@test isa(df[1], DataVector{Float64})
	@test isa(df[2], DataVector{UTF8String})
	@test isa(df[3], DataVector{Bool})
	@test isa(df[4], DataVector{Any})
	@test isa(df[5], DataVector{Any})
	@test isa(df[9], DataVector{Any})
	@test isa(df[10], DataVector{Any})
	@test df[4,1] == 2.5
	@test df[4,2] == "DDDD"
	@test df[4,3] == true
	@test df[1,4] == 2.0
	@test df[2,4] == "EEEEE"
	@test df[3,4] == false
	@test isna(df[3,5])
	@test df[1,6] == 3.
	@test isna(df[2,6])
	@test df[1,7] == "FF"
	@test isna(df[2,7])
	@test df[2,8] == true
	@test isna(df[1,8])
	@test df[1,10] == Date(1965,4,3)
	@test df[2,9] == DateTime(2015,2,4,10,14)
	@test df[4,9] == ExcelReaders.Time(15,2,0)
	@test isna(df[4,10])
	# TODO Add a test that checks the error code, not just type
	@test isa(df[1,11], ExcelErrorCell)
	@test isna(df[4,12])

	df = readxl(DataFrame, f, "Sheet1!C4:N7", header=false, colnames=[:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8, :c9, :c10, :c11, :c12])
	@test ncol(df) == 12
	@test nrow(df) == 4
	@test isa(df[:c1], DataVector{Float64})
	@test isa(df[:c2], DataVector{UTF8String})
	@test isa(df[:c3], DataVector{Bool})
	@test isa(df[:c4], DataVector{Any})
	@test isa(df[:c5], DataVector{Any})
	@test isa(df[:c9], DataVector{Any})
	@test isa(df[:c10], DataVector{Any})
	@test df[4,:c1] == 2.5
	@test df[4,:c2] == "DDDD"
	@test df[4,:c3] == true
	@test df[1,:c4] == 2.0
	@test df[2,:c4] == "EEEEE"
	@test df[3,:c4] == false
	@test isna(df[3,:c5])
	@test df[1,:c6] == 3.
	@test isna(df[2,:c6])
	@test df[1,:c7] == "FF"
	@test isna(df[2,:c7])
	@test df[2,:c8] == true
	@test isna(df[1,:c8])
	@test df[1,:c10] == Date(1965,4,3)
	@test df[2,:c9] == DateTime(2015,2,4,10,14)
	@test df[4,:c9] == ExcelReaders.Time(15,2,0)
	@test isna(df[4,:c10])
	# TODO Add a test that checks the error code, not just type
	@test isa(df[1,:c11], ExcelErrorCell)
	@test isna(df[4,:c12])

	df = readxl(DataFrame, f, "Sheet1!C3:N7", header=true, colnames=[:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8, :c9, :c10, :c11, :c12])
	@test ncol(df) == 12
	@test nrow(df) == 4
	@test isa(df[:c1], DataVector{Float64})
	@test isa(df[:c2], DataVector{UTF8String})
	@test isa(df[:c3], DataVector{Bool})
	@test isa(df[:c4], DataVector{Any})
	@test isa(df[:c5], DataVector{Any})
	@test isa(df[:c9], DataVector{Any})
	@test isa(df[:c10], DataVector{Any})
	@test df[4,:c1] == 2.5
	@test df[4,:c2] == "DDDD"
	@test df[4,:c3] == true
	@test df[1,:c4] == 2.0
	@test df[2,:c4] == "EEEEE"
	@test df[3,:c4] == false
	@test isna(df[3,:c5])
	@test df[1,:c6] == 3.
	@test isna(df[2,:c6])
	@test df[1,:c7] == "FF"
	@test isna(df[2,:c7])
	@test df[2,:c8] == true
	@test isna(df[1,:c8])
	@test df[1,:c10] == Date(1965,4,3)
	@test df[2,:c9] == DateTime(2015,2,4,10,14)
	@test df[4,:c9] == ExcelReaders.Time(15,2,0)
	@test isna(df[4,:c10])
	@test isa(df[1,:c11], ExcelErrorCell)
	@test isna(df[4,:c12])

	# Too few colnames
	@test_throws ErrorException df = readxl(DataFrame, f, "Sheet1!C3:N7", header=true, colnames=[:c1, :c2, :c3, :c4])

	# Test readxlsheet function
	@test_throws ErrorException readxlsheet(f, "Empty Sheet")
	for sheetinfo=["Second Sheet", 2]
		@test_throws ErrorException readxlsheet(f, sheetinfo, skipstartrows=-1)
		@test_throws ErrorException readxlsheet(f, sheetinfo, skipstartrows=:nonsense)

		@test_throws ErrorException readxlsheet(f, sheetinfo, skipstartcols=-1)
		@test_throws ErrorException readxlsheet(f, sheetinfo, skipstartcols=:nonsense)

		@test_throws ErrorException readxlsheet(f, sheetinfo, nrows=-1)
		@test_throws ErrorException readxlsheet(f, sheetinfo, nrows=:nonsense)

		@test_throws ErrorException readxlsheet(f, sheetinfo, ncols=-1)
		@test_throws ErrorException readxlsheet(f, sheetinfo, ncols=:nonsense)

		data = readxlsheet(f, sheetinfo)
		@test size(data) == (6, 6)
		@test data[2,1] == 1.
		@test data[5,2] == "CCC"
		@test data[3,3] == false
		@test data[6,6] == ExcelReaders.Time(15,2,00)
		@test isna(data[4,3])
		@test isna(data[4,6])

		data = readxlsheet(f, sheetinfo, skipstartrows=:blanks, skipstartcols=:blanks)
		@test size(data) == (6, 6)
		@test data[2,1] == 1.
		@test data[5,2] == "CCC"
		@test data[3,3] == false
		@test data[6,6] == ExcelReaders.Time(15,2,00)
		@test isna(data[4,3])
		@test isna(data[4,6])

		data = readxlsheet(f, sheetinfo, skipstartrows=0, skipstartcols=0)
		@test size(data) == (6+7, 6+3)
		@test data[2+7,1+3] == 1.
		@test data[5+7,2+3] == "CCC"
		@test data[3+7,3+3] == false
		@test data[6+7,6+3] == ExcelReaders.Time(15,2,00)
		@test isna(data[4+7,3+3])
		@test isna(data[4+7,6+3])

		data = readxlsheet(f, sheetinfo, skipstartrows=0, )
		@test size(data) == (6+7, 6)
		@test data[2+7,1] == 1.
		@test data[5+7,2] == "CCC"
		@test data[3+7,3] == false
		@test data[6+7,6] == ExcelReaders.Time(15,2,00)
		@test isna(data[4+7,3])
		@test isna(data[4+7,6])

		data = readxlsheet(f, sheetinfo, skipstartcols=0)
		@test size(data) == (6, 6+3)
		@test data[2,1+3] == 1.
		@test data[5,2+3] == "CCC"
		@test data[3,3+3] == false
		@test data[6,6+3] == ExcelReaders.Time(15,2,00)
		@test isna(data[4,3+3])
		@test isna(data[4,6+3])

		data = readxlsheet(f, sheetinfo, skipstartrows=1, skipstartcols=1, nrows=11, ncols=7)
		@test size(data) == (11, 7)
		@test data[2+6,1+2] == 1.
		@test data[5+6,2+2] == "CCC"
		@test data[3+6,3+2] == false
		@test_throws BoundsError data[6+6,6+2] == ExcelReaders.Time(15,2,00)
		@test isna(data[4+6,2+2])
	end
end
