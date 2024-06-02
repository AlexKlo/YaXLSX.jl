@testset "SheetInterfaces" begin
    @testset "Case №1: get table rows" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        table_rows = xl_rowtable(xl_sheet, "A1:B6")

        @test table_rows isa Tables.RowIterator
        @test table_rows.A == ["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0]
        @test table_rows.B == ["Names", "a", "b", "c", "d", "e"]
    end

    @testset "Case №2: get table not from the beginning" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        table_rows = xl_rowtable(xl_sheet, "B2:C3")

        @test table_rows.B == ["a", "b"]
        @test all(ismissing, table_rows.C)
    end

    @testset "Case №3: get table columns" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        table_cols = xl_columntable(xl_sheet, "A1:B6")

        @test table_cols isa NamedTuple
        @test table_cols.A == ["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0]
        @test table_cols.B == ["Names", "a", "b", "c", "d", "e"]
    end

    @testset "Case №4: get table columns with specified headers" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        table_cols = xl_columntable(xl_sheet, "A1:B6"; headers = ["Numbers", "Names"])

        @test table_cols isa NamedTuple
        @test table_cols.var"Numbers" == ["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0]
        @test table_cols.var"Names" == ["Names", "a", "b", "c", "d", "e"]
    end

    @testset "Case №5: get all table data" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")

        @test xl_rowtable(xl_sheet) isa Tables.RowIterator
        @test xl_columntable(xl_sheet) isa NamedTuple
    end

    @testset "Case №6: get data by column names only" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")

        @test xl_rowtable(xl_sheet, "A:B") isa Tables.RowIterator
        @test xl_columntable(xl_sheet, "A:B") isa NamedTuple
    end

    @testset "Case №7: get data by column indices only" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")

        @test xl_rowtable(xl_sheet, "1:2") isa Tables.RowIterator
        @test xl_columntable(xl_sheet, "1:2") isa NamedTuple
    end

    @testset "Case №8: invalid cell range exceptions" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")

        @test_throws ErrorException("KeyError: invalid cell range `AB`") begin
            xl_rowtable(xl_sheet, "AB")
        end
        @test_throws ErrorException("KeyError: invalid cell range `A1:B`") begin
            xl_rowtable(xl_sheet, "A1:B")
        end
        @test_throws ErrorException("KeyError: invalid cell range `1:B6`") begin
            xl_rowtable(xl_sheet, "1:B6")
        end
        @test_throws ErrorException("KeyError: invalid cell range `3:2`") begin
            xl_rowtable(xl_sheet, "3:1")
        end
        @test_throws ErrorException("KeyError: invalid cell range `a:b`") begin
            xl_rowtable(xl_sheet, "a:b")
        end
        @test_throws ErrorException("KeyError: invalid cell range `A2:B1`") begin
            xl_rowtable(xl_sheet, "A2:B1")
        end
    end

    @testset "Case №9: invalind headers length" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        
        exp_msg = "KeyError: headers=[\"Numbers\"] length must be equal \
            to cell range length: 2"
        @test_throws ErrorException(exp_msg) begin
            xl_columntable(xl_sheet, "A1:B6"; headers = ["Numbers"])
        end
    end
end