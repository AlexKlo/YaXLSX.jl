@testset "SheetInterfaces" begin
    @testset "Case №1: get table rows" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        table_rows = xl_rowtable(xl_sheet, "A1:B6")

        exp_df_rows = DataFrame(
            ["Numbers" "Names"; 1.0 "a"; 2.0 "b"; 3.0 "c"; 4.0 "d"; 5.0 "e"],
            ["A","B"]
        ) |> eachrow

        @test table_rows == exp_df_rows
    end

    @testset "Case №2: get table not from the beginning" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        table_rows = xl_rowtable(xl_sheet, "B2:C3")

        @test table_rows.B == ["a", "b"]
        @test all(ismissing, table_rows.C)
    end

    @testset "Case №3: invalid cell range" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        
        @test_throws ErrorException("KeyError: invalid cell range `B6:A1`") begin
            xl_rowtable(xl_sheet, "B6:A1")
        end
    end

    @testset "Case №4: get table columns" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        table_cols = xl_columntable(xl_sheet, "A1:B6")

        exp_df_cols = DataFrame(
            ["Numbers" "Names"; 1.0 "a"; 2.0 "b"; 3.0 "c"; 4.0 "d"; 5.0 "e"],
            ["A","B"]
        ) |> eachcol

        @test table_cols == exp_df_cols
    end

    @testset "Case №5: get table columns with specified headers" begin
        xl_book = parse_xlsx(read("data/simple_book.xlsx"))
        xl_sheet = xl_sheets(xl_book, "Лист1")
        table_cols = xl_columntable(xl_sheet, "A1:B6"; headers = ["Numbers", "Names"])

        exp_df_cols = DataFrame(
            ["Numbers" "Names"; 1.0 "a"; 2.0 "b"; 3.0 "c"; 4.0 "d"; 5.0 "e"],
            ["Numbers", "Names"]
        ) |> eachcol

        @test table_cols == exp_df_cols
    end
end