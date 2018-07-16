using ExcelReaders
using Dates
using PyCall
using DataValues
using Test

@testset "ExcelReaders" begin

# TODO Throw julia specific exceptions for these errors
@test_throws PyCall.PyError openxl("FileThatDoesNotExist.xlsx")
@test_throws PyCall.PyError openxl("runtests.jl")

filename = normpath(@__DIR__, "TestData.xlsx")
file = openxl(filename)
@test file.filename == "TestData.xlsx"

buffer = IOBuffer()
show(buffer, file)
@test String(take!(buffer)) == "ExcelFile <TestData.xlsx>"

for (k,v) in Dict(0=>"#NULL!",7=>"#DIV/0!",23 => "#REF!",42=>"#N/A",29=>"#NAME?",36=>"#NUM!",15=>"#VALUE!")
    errorcell = ExcelErrorCell(k)
    buffer = IOBuffer()
    show(buffer, errorcell)
    @test String(take!(buffer)) == v
end

# Read into DataValueArray
for f in [file, filename]
    @test_throws ErrorException readxl(f, "Sheet1!C4:G3")
    @test_throws ErrorException readxl(f, "Sheet1!G2:B5")
    @test_throws ErrorException readxl(f, "Sheet1!G5:B2")

    data = readxl(f, "Sheet1!C3:N7")
    @test size(data) == (5,12)
    @test data[4,1] == 2.0
    @test data[2,2] == "A"
    @test data[2,3] == true
    @test DataValues.isna(data[4,5])
    @test data[2,9] == Date(2015,3,3)
    @test data[3,9] == DateTime(2015,2,4,10,14)
    @test data[4,9] == DateTime(1988,4,9,0,0)
    @test data[5,9] == Time(15,2,0)
    @test data[3,10] == DateTime(1950,8,9,18,40)
    @test DataValues.isna(data[5,10])
    @test isa(data[2,11], ExcelErrorCell)
    @test isa(data[3,11], ExcelErrorCell)
    @test isa(data[4,12], ExcelErrorCell)
    @test DataValues.isna(data[5,12])

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
        @test data[6,6] == Time(15,2,00)
        @test DataValues.isna(data[4,3])
        @test DataValues.isna(data[4,6])

        data = readxlsheet(f, sheetinfo, skipstartrows=:blanks, skipstartcols=:blanks)
        @test size(data) == (6, 6)
        @test data[2,1] == 1.
        @test data[5,2] == "CCC"
        @test data[3,3] == false
        @test data[6,6] == Time(15,2,00)
        @test DataValues.isna(data[4,3])
        @test DataValues.isna(data[4,6])

        data = readxlsheet(f, sheetinfo, skipstartrows=0, skipstartcols=0)
        @test size(data) == (6+7, 6+3)
        @test data[2+7,1+3] == 1.
        @test data[5+7,2+3] == "CCC"
        @test data[3+7,3+3] == false
        @test data[6+7,6+3] == Time(15,2,00)
        @test DataValues.isna(data[4+7,3+3])
        @test DataValues.isna(data[4+7,6+3])

        data = readxlsheet(f, sheetinfo, skipstartrows=0, )
        @test size(data) == (6+7, 6)
        @test data[2+7,1] == 1.
        @test data[5+7,2] == "CCC"
        @test data[3+7,3] == false
        @test data[6+7,6] == Time(15,2,00)
        @test DataValues.isna(data[4+7,3])
        @test DataValues.isna(data[4+7,6])

        data = readxlsheet(f, sheetinfo, skipstartcols=0)
        @test size(data) == (6, 6+3)
        @test data[2,1+3] == 1.
        @test data[5,2+3] == "CCC"
        @test data[3,3+3] == false
        @test data[6,6+3] == Time(15,2,00)
        @test DataValues.isna(data[4,3+3])
        @test DataValues.isna(data[4,6+3])

        data = readxlsheet(f, sheetinfo, skipstartrows=1, skipstartcols=1, nrows=11, ncols=7)
        @test size(data) == (11, 7)
        @test data[2+6,1+2] == 1.
        @test data[5+6,2+2] == "CCC"
        @test data[3+6,3+2] == false
        @test_throws BoundsError data[6+6,6+2] == Time(15,2,00)
        @test DataValues.isna(data[4+6,2+2])
    end
end

end
