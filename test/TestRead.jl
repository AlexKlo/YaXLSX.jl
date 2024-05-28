
import YaXLSX: ExcelBook

@testset "Read" begin
    @testset "Case №1: Simple reading" begin
        xlsx_book_bytes = parse_xlsx(read("data/simple_book.xlsx"))

        @test xlsx_book_bytes isa ExcelBook
        @test xlsx_book_bytes.sheet_names == ["Лист1"]
        @test xlsx_book_bytes.sheets |> length == 1

        xlsx_book_path = parse_xlsx("data/simple_book.xlsx")
        @test xlsx_book_path isa ExcelBook
    end

    @testset "Case №2: Invalid format file reading" begin
        @test_throws ErrorException("WrongExtension: supported only .xlsx format") begin
            xlsx_book = parse_xlsx(read("data/invalid.xls"))
        end
    end

    @testset "Case №3: Check ExcelSheet data" begin
        xlsx_book_bytes = parse_xlsx(read("data/simple_book.xlsx"))

        sheet = xlsx_book_bytes.sheets[1]

        @test sheet.name == "Лист1"
        @test sheet.data == Dict(
            "A1" => "Numbers", "B1" => "Names",
            "A2" => 1.0, "B2" => "a",
            "A3" => 2.0, "B3" => "b",
            "A4" => 3.0, "B4" => "c",
            "A5" => 4.0, "B5" => "d",
            "A6" => 5.0, "B6" => "e",
        )
        @test sheet.dimension == (6, 2)
    end

    @testset "Case №5: Reading book with two sheets" begin
        xlsx_book_bytes = parse_xlsx(read("data/two_sheets.xlsx"))

        @test xlsx_book_bytes.sheet_names == ["Лист1", "Лист2"]
        @test xlsx_book_bytes.sheets |> length == 2

        sheet1 = xlsx_book_bytes.sheets[1]
        @test sheet1.name == "Лист1"
        @test sheet1.data == Dict(
            "A1" => "Numbers", "B1" => "Names",
            "A2" => 1.0, "B2" => "a",
            "A3" => 2.0, "B3" => "b",
            "A4" => 3.0, "B4" => "c",
            "A5" => 4.0, "B5" => "d",
            "A6" => 5.0, "B6" => "e",
        )
        @test sheet1.dimension == (6, 2)

        sheet2 = xlsx_book_bytes.sheets[2]
        @test sheet2.name == "Лист2"
        @test sheet2.data == Dict(
            "A1" => "Numbers", "B1" => "Names",
            "A2" => 1.0, "B2" => "a",
            "A3" => 2.0, "B3" => "b",
            "A4" => 3.0, "B4" => "c",
            "A5" => 4.0, "B5" => "d", "C5" => false,
            "A6" => 5.0, "B6" => "e", "C6" => "D8",
            "C7" => "D8/0",
            "D8" => 100.0,
        )
        @test sheet2.dimension == (8, 4)
    end
end
