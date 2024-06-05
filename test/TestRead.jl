
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
        @test sheet.data == [
            "Numbers" "Names";
            1.0       "a";
            2.0       "b";
            3.0       "c";
            4.0       "d";
            5.0       "e";
        ]
    end

    @testset "Case №4: Reading book with two sheets" begin
        xlsx_book_bytes = parse_xlsx(read("data/two_sheets.xlsx"))

        @test xlsx_book_bytes.sheet_names == ["Лист1", "Лист2"]
        @test xlsx_book_bytes.sheets |> length == 2

        sheet1 = xlsx_book_bytes.sheets[1]
        @test sheet1.name == "Лист1"
        @test sheet1.data == [
            "Numbers" "Names";
            1.0       "a";
            2.0       "b";
            3.0       "c";
            4.0       "d";
            5.0       "e";
        ]

        sheet2 = xlsx_book_bytes.sheets[2]
        @test sheet2.name == "Лист2"
        @test sheet2.data[1:6, 1:2] == [
            "Numbers" "Names";
            1.0       "a";
            2.0       "b";
            3.0       "c";
            4.0       "d";
            5.0       "e";
        ]
        @test sheet2.data[6, 3] == "D8"
        @test sheet2.data[7, 3] == "D8/0"
        @test sheet2.data[8, 4] == 100.0;
    end

    @testset "Case №4: Reading same edge cases" begin
        @test try parse_xlsx(read("data/blank.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/book_sparse.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/book_sparse_2.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/inlinestr.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/style_strings.xlsx")); true catch; false end
    end
end
