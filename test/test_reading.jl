
@testset "Reading" begin
    @testset "Case №1: Simple reading" begin
        xlsx_book_bytes = parse_xlsx(read("data/simple_book.xlsx"))

        @test xlsx_book_bytes isa ExcelBook
        @test xlsx_book_bytes.sheet_names == ["Лист1"]
        @test xlsx_book_bytes.sheets |> length == 1

        xlsx_book_path = parse_xlsx("data/simple_book.xlsx")
        @test xlsx_book_path isa ExcelBook
    end

    @testset "Case №2: Invalid format file reading" begin
        @test_throws "WrongExtension: supported only .xlsx format" begin
            xlsx_book = parse_xlsx(read("data/invalid.xls"))
        end
    end

    @testset "Case №3: File without xl/workbook.xml" begin
        @test_throws "ParseError: xl/workbook.xml not found" begin
            xlsx_book = parse_xlsx(read("data/without_workbook.xlsx"))
        end
    end
end
