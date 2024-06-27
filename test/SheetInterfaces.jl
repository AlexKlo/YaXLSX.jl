@testset "SheetInterfaces" begin

    sheet1_rowtable = NamedTuple[
        (A = "Numbers", B = "Names"),
        (A = 1.0, B = "a"),
        (A = 2.0, B = "b"),
        (A = 3.0, B = "c"),
        (A = 4.0, B = "d"),
        (A = 5.0, B = "e"),
    ]
    sheet1_columntable = (
        A = Any["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0], 
        B = Any["Names", "a", "b", "c", "d", "e"]
    )

    @testset "Case №1: get table rows" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")
        table_rows = xl_rowtable(xl_sheet, "A1:B6")

        @test table_rows isa RowTable
        @test table_rows == sheet1_rowtable
    end

    @testset "Case №2: get table not from the beginning" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")
        table_rows = xl_rowtable(xl_sheet, "B2:C3")

        @test table_rows == NamedTuple[
            (B = "a", C = nothing),
            (B = "b", C = nothing)
        ]
    end

    @testset "Case №3: get table columns" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")
        table_cols = xl_columntable(xl_sheet, "A1:B6")

        @test table_cols isa ColumnTable
        @test table_cols == sheet1_columntable
    end

    @testset "Case №4: get table columns with specified headers" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")
        table_cols = xl_columntable(xl_sheet, "A1:B6"; column_labels = ["Numbers", "Names"])

        @test table_cols isa ColumnTable
        @test table_cols == (
            Numbers = Any["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0], 
            Names = Any["Names", "a", "b", "c", "d", "e"]
        )
    end

    @testset "Case №5: get all table data" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")

        @test xl_rowtable(xl_sheet) == sheet1_rowtable
        @test xl_columntable(xl_sheet) == sheet1_columntable
    end

    @testset "Case №6: get data by column names only" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")

        @test xl_rowtable(xl_sheet, "A:B") == sheet1_rowtable
        @test xl_columntable(xl_sheet, "A:B") == sheet1_columntable
    end

    @testset "Case №7: get data by column indices only" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")

        @test xl_rowtable(xl_sheet, "1:2") == sheet1_rowtable
        @test xl_columntable(xl_sheet, "1:2") == sheet1_columntable
    end

    @testset "Case №8: invalid cell range exceptions" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")

        @test_throws ErrorException("KeyError: invalid range `AB`") begin
            xl_rowtable(xl_sheet, "AB")
        end
        @test_throws ErrorException("KeyError: invalid range `A1:B`") begin
            xl_rowtable(xl_sheet, "A1:B")
        end
        @test_throws ErrorException("KeyError: invalid range `1:B6`") begin
            xl_rowtable(xl_sheet, "1:B6")
        end
        @test_throws ErrorException("KeyError: invalid range `3:1`") begin
            xl_rowtable(xl_sheet, "3:1")
        end
        @test_throws ErrorException("KeyError: invalid range `a:b`") begin
            xl_columntable(xl_sheet, "a:b")
        end
        @test_throws ErrorException("KeyError: invalid range `A2:B1`") begin
            xl_columntable(xl_sheet, "A2:B1")
        end
        @test_throws ErrorException("KeyError: invalid range `16380:16385`") begin
            xl_columntable(xl_sheet, "16380:16385")
        end
        @test_throws ErrorException("KeyError: invalid range `A1:A1048577`") begin
            xl_columntable(xl_sheet, "A1:A1048577")
        end
    end

    @testset "Case №9: invalind headers length" begin
        xl_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Sheet1")
        
        exp_msg = "column_labels=[\"Nums\"] length must be equal to table columns length: 2"
        @test_throws ErrorException(exp_msg) begin
            xl_columntable(xl_sheet, "A1:B6"; column_labels = ["Nums"])
        end
    end
end