@testitem "ExcelReaders" begin
    using Dates, DataValues

    @test_throws ErrorException openxl("FileThatDoesNotExist.xls")
    @test_throws ErrorException openxl(normpath(@__DIR__, "runtests.jl"))

    filename = normpath(@__DIR__, "TestData.xls")
    file = openxl(filename)
    @test file.filename == "TestData.xls"

    @test sprint(show, file) == "ExcelFile <TestData.xls>"

    for (k, v) in Dict(0 => "#NULL!", 7 => "#DIV/0!", 23 => "#REF!", 42 => "#N/A", 29 => "#NAME?", 36 => "#NUM!", 15 => "#VALUE!")
        errorcell = ExcelErrorCell(k)
        @test sprint(show, errorcell) == v
    end

# Read into DataValueArray
    for f in [file, filename]
        @test_throws ErrorException readxl(f, "Sheet1!C4:G3")
        @test_throws ErrorException readxl(f, "Sheet1!G2:B5")
        @test_throws ErrorException readxl(f, "Sheet1!G5:B2")

        data = readxl(f, "Sheet1!C3:N7")
        @test size(data) == (5, 12)
        @test data[4,1] == 2.0
        @test data[2,2] == "A"
        @test data[2,3] == true
        @test DataValues.isna(data[4,5])
        @test data[2,9] == Date(2015, 3, 3)
        @test data[3,9] == DateTime(2015, 2, 4, 10, 14)
        @test data[4,9] == DateTime(1988, 4, 9, 0, 0)
        @test data[5,9] == Time(15, 2, 0)
        @test data[3,10] == DateTime(1950, 8, 9, 18, 40)
        @test DataValues.isna(data[5,10])
        @test isa(data[2,11], ExcelErrorCell)
        @test isa(data[3,11], ExcelErrorCell)
        @test isa(data[4,12], ExcelErrorCell)
        @test DataValues.isna(data[5,12])

    # Test readxlsheet function
        @test_throws ErrorException readxlsheet(f, "Empty Sheet")
        for sheetinfo = ["Second Sheet", 2]
            @test_throws ErrorException readxlsheet(f, sheetinfo, skipstartrows = -1)
            @test_throws ErrorException readxlsheet(f, sheetinfo, skipstartrows = :nonsense)

            @test_throws ErrorException readxlsheet(f, sheetinfo, skipstartcols = -1)
            @test_throws ErrorException readxlsheet(f, sheetinfo, skipstartcols = :nonsense)

            @test_throws ErrorException readxlsheet(f, sheetinfo, nrows = -1)
            @test_throws ErrorException readxlsheet(f, sheetinfo, nrows = :nonsense)

            @test_throws ErrorException readxlsheet(f, sheetinfo, ncols = -1)
            @test_throws ErrorException readxlsheet(f, sheetinfo, ncols = :nonsense)

            data = readxlsheet(f, sheetinfo)
            @test size(data) == (6, 6)
            @test data[2,1] == 1.
            @test data[5,2] == "CCC"
            @test data[3,3] == false
            @test data[6,6] == Time(15, 2, 00)
            @test DataValues.isna(data[4,3])
            @test DataValues.isna(data[4,6])

            data = readxlsheet(f, sheetinfo, skipstartrows = :blanks, skipstartcols = :blanks)
            @test size(data) == (6, 6)
            @test data[2,1] == 1.
            @test data[5,2] == "CCC"
            @test data[3,3] == false
            @test data[6,6] == Time(15, 2, 00)
            @test DataValues.isna(data[4,3])
            @test DataValues.isna(data[4,6])

            data = readxlsheet(f, sheetinfo, skipstartrows = 0, skipstartcols = 0)
            @test size(data) == (6 + 7, 6 + 3)
            @test data[2 + 7,1 + 3] == 1.
            @test data[5 + 7,2 + 3] == "CCC"
            @test data[3 + 7,3 + 3] == false
            @test data[6 + 7,6 + 3] == Time(15, 2, 00)
            @test DataValues.isna(data[4 + 7,3 + 3])
            @test DataValues.isna(data[4 + 7,6 + 3])

            data = readxlsheet(f, sheetinfo, skipstartrows = 0, )
            @test size(data) == (6 + 7, 6)
            @test data[2 + 7,1] == 1.
            @test data[5 + 7,2] == "CCC"
            @test data[3 + 7,3] == false
            @test data[6 + 7,6] == Time(15, 2, 00)
            @test DataValues.isna(data[4 + 7,3])
            @test DataValues.isna(data[4 + 7,6])

            data = readxlsheet(f, sheetinfo, skipstartcols = 0)
            @test size(data) == (6, 6 + 3)
            @test data[2,1 + 3] == 1.
            @test data[5,2 + 3] == "CCC"
            @test data[3,3 + 3] == false
            @test data[6,6 + 3] == Time(15, 2, 00)
            @test DataValues.isna(data[4,3 + 3])
            @test DataValues.isna(data[4,6 + 3])

            data = readxlsheet(f, sheetinfo, skipstartrows = 1, skipstartcols = 1, nrows = 11, ncols = 7)
            @test size(data) == (11, 7)
            @test data[2 + 6,1 + 2] == 1.
            @test data[5 + 6,2 + 2] == "CCC"
            @test data[3 + 6,3 + 2] == false
            @test_throws BoundsError data[6 + 6,6 + 2] == Time(15, 2, 00)
            @test DataValues.isna(data[4 + 6,2 + 2])
        end
    end

end

@testitem "ExcelReaders xlsx backend" begin
    using Dates, DataValues

    filename = normpath(@__DIR__, "TestData.xlsx")

    file = openxl(filename)
    @test file.filename == "TestData.xlsx"
    @test sprint(show, file) == "ExcelFile <TestData.xlsx>"

    # TestData.xlsx holds the same content as TestData.xls, so the very same
    # assertions must hold for both backends.
    for f in [file, filename]
        @test_throws ErrorException readxl(f, "Sheet1!C4:G3")
        @test_throws ErrorException readxl(f, "Sheet1!G2:B5")
        @test_throws ErrorException readxl(f, "Sheet1!G5:B2")

        data = readxl(f, "Sheet1!C3:N7")
        @test size(data) == (5, 12)
        @test data[4,1] == 2.0
        @test data[2,2] == "A"
        @test data[2,3] == true
        @test DataValues.isna(data[4,5])
        @test data[2,9] == Date(2015, 3, 3)
        @test data[2,9] isa DateTime            # backend-independent type vocabulary
        @test data[3,9] == DateTime(2015, 2, 4, 10, 14)
        @test data[4,9] == DateTime(1988, 4, 9, 0, 0)
        @test data[5,9] == Time(15, 2, 0)
        @test data[3,10] == DateTime(1950, 8, 9, 18, 40)
        @test DataValues.isna(data[5,10])
        @test isa(data[2,11], ExcelErrorCell)
        @test isa(data[3,11], ExcelErrorCell)
        @test isa(data[4,12], ExcelErrorCell)
        @test DataValues.isna(data[5,12])

        # single cell read
        @test readxl(f, "Sheet1!C4") == 1.0

        @test_throws ErrorException readxlsheet(f, "Empty Sheet")

        data = readxlsheet(f, "Second Sheet")
        @test size(data) == (6, 6)
        @test data[2,1] == 1.
        @test data[5,2] == "CCC"
        @test data[3,3] == false
        @test data[6,6] == Time(15, 2, 00)
        @test DataValues.isna(data[4,3])
        @test DataValues.isna(data[4,6])
    end
end

@testitem "Cross-backend parity" begin
    using Dates, DataValues

    # The two test files hold the same content in the two file formats, so
    # every cell in the shared range must come back identical from both
    # backends.
    xls = openxl(normpath(@__DIR__, "TestData.xls"))
    xlsx = openxl(normpath(@__DIR__, "TestData.xlsx"))

    function compare_cells(data_xls, data_xlsx)
        @test size(data_xls) == size(data_xlsx)
        for i in eachindex(data_xls)
            a, b = data_xls[i], data_xlsx[i]
            if a isa ExcelErrorCell
                @test b isa ExcelErrorCell
                @test a.errorcode == b.errorcode
            elseif a isa DataValue
                @test b isa DataValue && DataValues.isna(b)
            else
                @test typeof(a) == typeof(b)
                @test a == b
            end
        end
    end

    compare_cells(readxl(xls, "Sheet1!C3:N7"), readxl(xlsx, "Sheet1!C3:N7"))

    for sheet in ["Second Sheet", 2]
        compare_cells(readxlsheet(xls, sheet), readxlsheet(xlsx, sheet))
    end

    # the internal entry point ExcelFiles relies on
    @test ExcelReaders.readxl_internal(xls, "Sheet1", 4, 3, 4, 3) == 1.0
    @test ExcelReaders.readxl_internal(xlsx, "Sheet1", 4, 3, 4, 3) == 1.0

    # close works for both backends (a no-op for xlsx)
    close(xls)
    close(xlsx)
    @test_throws ErrorException readxl(xls, "Sheet1!C4")
end

@testitem "Defined names" begin
    using Dates, DataValues
    import ExcelReaders.XLSX

    filename = joinpath(mktempdir(), "named.xlsx")
    XLSX.openxlsx(filename, mode="w") do xf
        sh = xf[1]
        XLSX.rename!(sh, "Sheet1")
        sh["B2"] = 1.0
        sh["C2"] = 2.0
        sh["B3"] = 3.0
        sh["C3"] = 4.0
        sh["E1"] = "hello"
        XLSX.addDefinedName(xf, "block", "Sheet1!B2:C3")
        XLSX.addDefinedName(xf, "single", "Sheet1!E1")
    end

    f = openxl(filename)
    @test readxlnames(f) == ["block", "single"]

    data = readxlrange(f, "block")
    @test size(data) == (2, 2)
    @test data[1, 1] == 1.0
    @test data[2, 2] == 4.0

    @test readxlrange(f, "single") == "hello"

    # Defined names are not available for legacy xls files, because the
    # underlying C library does not expose them.
    xls = openxl(normpath(@__DIR__, "TestData.xls"))
    @test_throws ErrorException readxlnames(xls)
    @test_throws ErrorException readxlrange(xls, "block")
end

@testitem "Open-ended ranges" begin
    using Dates, DataValues

    for name in ["TestData.xls", "TestData.xlsx"]
        f = openxl(normpath(@__DIR__, name))

        data = readxl(f, "Sheet1!C4:C")
        @test size(data, 2) == 1
        @test data[1, 1] == 1.0
        @test data[2, 1] == 1.5
        @test data[3, 1] == 2.0

        # a fully open range spans all rows of the sheet
        data2 = readxl(f, "Sheet1!C:C")
        @test size(data2, 1) == size(data, 1) + 3
        @test DataValues.isna(data2[1, 1])
        @test data2[4, 1] == 1.0

        data3 = readxl(f, "Sheet1!C3:D")
        @test data3[1, 1] == "Some Float64s"
        @test data3[2, 2] == "A"

        @test_throws ErrorException readxl(f, "Sheet1!C4:G3")
        @test_throws ErrorException readxl(f, "Sheet1!:C")
        @test_throws ErrorException readxl(f, "Sheet1!C4:")
    end
end

@testitem "Ranges beyond the data" begin
    using DataValues

    # Ranges that extend beyond the used part of the sheet come back as
    # NA-filled cells rather than erroring (issue #49).
    for name in ["TestData.xls", "TestData.xlsx"]
        f = openxl(normpath(@__DIR__, name))

        data = readxl(f, "Sheet1!A1:AB60")
        @test size(data) == (60, 28)
        @test data[4, 3] == 1.0
        @test DataValues.isna(data[1, 1])
        @test DataValues.isna(data[60, 28])

        @test DataValues.isna(readxl(f, "Sheet1!AB60"))
    end
end
