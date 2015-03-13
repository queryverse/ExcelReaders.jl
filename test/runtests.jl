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

	df = readxl(DataFrame, f, "Sheet1!C3:J7")
	@test ncol(df) == 8
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
	@test df[1,symbol("Float64 with NA")] == 3.
	@test isna(df[2,symbol("Float64 with NA")])
	@test df[1,symbol("String with NA")] == "FF"
	@test isna(df[2,symbol("String with NA")])
	@test df[2,symbol("Bool with NA")] == true
	@test isna(df[1,symbol("Bool with NA")])

	df = readxl(DataFrame, f, "Sheet1!C4:J7", header=false)
	@test ncol(df) == 8
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
	@test df[1,6] == 3.
	@test isna(df[2,6])
	@test df[1,7] == "FF"
	@test isna(df[2,7])
	@test df[2,8] == true
	@test isna(df[1,8])

	df = readxl(DataFrame, f, "Sheet1!C4:J7", header=false, colnames=[:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8])
	@test ncol(df) == 8
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
	@test df[1,:c6] == 3.
	@test isna(df[2,:c6])
	@test df[1,:c7] == "FF"
	@test isna(df[2,:c7])
	@test df[2,:c8] == true
	@test isna(df[1,:c8])

	df = readxl(DataFrame, f, "Sheet1!C3:J7", header=true, colnames=[:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8])
	@test ncol(df) == 8
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
	@test df[1,:c6] == 3.
	@test isna(df[2,:c6])
	@test df[1,:c7] == "FF"
	@test isna(df[2,:c7])
	@test df[2,:c8] == true
	@test isna(df[1,:c8])

	# Too few colnames
	@test_throws ErrorException df = readxl(DataFrame, f, "Sheet1!C3:J7", header=true, colnames=[:c1, :c2, :c3, :c4])
end
