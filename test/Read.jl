@testset "Read" begin
    @testset "Case №1: Simple reading" begin

        xlsx_book_bytes = parse_xlsx(read("data/simple_table.xlsx"))

        @test xlsx_book_bytes isa XLSX
        @test xlsx_book_bytes.sheets |> length == 1

        xlsx_book_path = parse_xlsx("data/simple_table.xlsx")
        @test xlsx_book_path isa XLSX
    end

    @testset "Case №2: Invalid format file reading" begin
        @test_throws ErrorException("WrongExtension: supported only .xlsx format") begin
            xlsx_book = parse_xlsx(read("data/old.xls"))
        end
    end

    @testset "Case №3: Reading different cases" begin
        @test try parse_xlsx(read("data/two_sheets.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/blank.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/book_sparse.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/book_sparse_2.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/inlinestr.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/style_strings.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/file_example_XLSX_5000.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/missing.xlsx")); true catch; false end
        @test try parse_xlsx(read("data/general_tables.xlsx")); true catch; false end
    end
end