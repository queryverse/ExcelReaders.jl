using ExcelReaders
using Base.Test
using PyCall
using DataArrays
using DataFrames

# TODO Throw julia specific exceptions for these errors
@test_throws PyCall.PyError openxl("FileThatDoesNotExist.xlsx")
@test_throws PyCall.PyError openxl("runtests.jl")

filename = normpath(Pkg.dir("ExcelReaders"),"test", "TestData.xlsx")
file = openxl(filename)
@test file.filename == "TestData.xlsx"

buffer = IOBuffer()
show(buffer, file)
@test takebuf_string(buffer) == "ExcelFile <TestData.xlsx>"

# Read into DataArray
for f in [file, filename]
	@test_throws ErrorException readxl(f, "Sheet1!C4:G3")
	@test_throws ErrorException readxl(f, "Sheet1!G2:B5")
	@test_throws ErrorException readxl(f, "Sheet1!G5:B2")

	data = readxl(f, "Sheet1!C3:J7")
	@test size(data) == (5,8)
	@test data[4,1] == 2.0
	@test data[2,2] == "A"
	@test data[2,3] == true
	@test isna(data[4,5])

	# TODO Read in C3:J7 once a bug in DataArrays is fixed
	df = readxl(DataFrame, f, "Sheet1!C3:G7")
	@test ncol(df) == 5
	@test nrow(df) == 4
	@test isa(df[symbol("Some Float64s")], DataVector{Float64})
	@test isa(df[symbol("Some Strings")], DataVector{UTF8String})
	@test isa(df[symbol("Some Bools")], DataVector{Bool})
	@test isa(df[symbol("Mixed column")], DataVector{Any})
	@test isa(df[symbol("Mixed with NA")], DataVector{Any})
	@test df[4,symbol("Some Float64s")] == 2.5
	@test df[4,symbol("Some Strings")] == "DDDD"
	@test df[4,symbol("Some Bools")] == true
	@test df[1,symbol("Mixed column")] == 2.0
	@test df[2,symbol("Mixed column")] == "EEEEE"
	@test df[3,symbol("Mixed column")] == false
	@test isna(df[3,symbol("Mixed with NA")])

	# TODO Read in C3:J7 once a bug in DataArrays is fixed
	df = readxl(DataFrame, f, "Sheet1!C4:G7", header=false)
	@test ncol(df) == 5
	@test nrow(df) == 4
	@test isa(df[1], DataVector{Float64})
	@test isa(df[2], DataVector{UTF8String})
	@test isa(df[3], DataVector{Bool})
	@test isa(df[4], DataVector{Any})
	@test isa(df[5], DataVector{Any})
	@test df[4,1] == 2.5
	@test df[4,2] == "DDDD"
	@test df[4,3] == true
	@test df[1,4] == 2.0
	@test df[2,4] == "EEEEE"
	@test df[3,4] == false
	@test isna(df[3,5])

	# TODO Read in C3:J7 once a bug in DataArrays is fixed
	df = readxl(DataFrame, f, "Sheet1!C4:G7", header=false, colnames=[:c1, :c2, :c3, :c4, :c5])
	@test ncol(df) == 5
	@test nrow(df) == 4
	@test isa(df[:c1], DataVector{Float64})
	@test isa(df[:c2], DataVector{UTF8String})
	@test isa(df[:c3], DataVector{Bool})
	@test isa(df[:c4], DataVector{Any})
	@test isa(df[:c5], DataVector{Any})
	@test df[4,:c1] == 2.5
	@test df[4,:c2] == "DDDD"
	@test df[4,:c3] == true
	@test df[1,:c4] == 2.0
	@test df[2,:c4] == "EEEEE"
	@test df[3,:c4] == false
	@test isna(df[3,:c5])

	# TODO Read in C3:J7 once a bug in DataArrays is fixed
	df = readxl(DataFrame, f, "Sheet1!C3:G7", header=true, colnames=[:c1, :c2, :c3, :c4, :c5])
	@test ncol(df) == 5
	@test nrow(df) == 4
	@test isa(df[:c1], DataVector{Float64})
	@test isa(df[:c2], DataVector{UTF8String})
	@test isa(df[:c3], DataVector{Bool})
	@test isa(df[:c4], DataVector{Any})
	@test isa(df[:c5], DataVector{Any})
	@test df[4,:c1] == 2.5
	@test df[4,:c2] == "DDDD"
	@test df[4,:c3] == true
	@test df[1,:c4] == 2.0
	@test df[2,:c4] == "EEEEE"
	@test df[3,:c4] == false
	@test isna(df[3,:c5])

	# Too few colnames
	@test_throws ErrorException df = readxl(DataFrame, f, "Sheet1!C3:G7", header=true, colnames=[:c1, :c2, :c3, :c4])
end
