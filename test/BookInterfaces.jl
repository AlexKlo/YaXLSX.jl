@testset "BookInterfaces" begin

    sheet1_table = Any[
        "Numbers"  "Names";
        1.0         "a";
        2.0         "b";
        3.0         "c";
        4.0         "d";
        5.0         "e";
    ]

    sheet2_table = Any[
        "Numbers"  "Names"       nothing     nothing;
        1.0         "a"          nothing     nothing;
        2.0         "b"          nothing     nothing;
        3.0         "c"          nothing     nothing;
        4.0         "d"          false       nothing;
        5.0         "e"          "D8"        nothing;
        nothing    nothing       "D8/0"      nothing;
        nothing    nothing       nothing     100.0;
    ]

    @testset "Case №1: get sheets names" begin
        xlsx_book = parse_xlsx(read("data/simple_table.xlsx"))
        @test xl_sheetnames(xlsx_book) == ["Sheet1"]

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheetnames(xlsx_book) == ["Sheet1", "Sheet2"]
    end

    @testset "Case №2: get all sheets" begin
        xlsx_book = parse_xlsx(read("data/simple_table.xlsx"))
        @test xl_sheets(xlsx_book) |> length == 1

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheets(xlsx_book) |> length == 2
    end

    @testset "Case №3: get sheets by name" begin
        xlsx_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet1 = xl_sheets(xlsx_book, "Sheet1")
        @test xl_sheet1.table == sheet1_table

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        xl_sheet2 = xl_sheets(xlsx_book, "Sheet2")
        @test xl_sheet2.table == sheet2_table
    end

    @testset "Case №4: get sheets by index" begin
        xlsx_book = parse_xlsx(read("data/simple_table.xlsx"))
        xl_sheet1 = xl_sheets(xlsx_book, 1)
        @test xl_sheet1.table == sheet1_table

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        xl_sheet2 = xl_sheets(xlsx_book, 2)
        @test xl_sheet2.table == sheet2_table
    end

    @testset "Case №5: get sheets by list" begin
        xlsx_book = parse_xlsx(read("data/simple_table.xlsx"))
        @test xl_sheets(xlsx_book, ["Sheet1"]) |> length == 1

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheets(xlsx_book, ["Sheet1", "Sheet2"]) |> length == 2
    end

    @testset "Case №6: get sheets by range" begin
        xlsx_book = parse_xlsx(read("data/simple_table.xlsx"))
        @test xl_sheets(xlsx_book, 1:1) |> length == 1

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheets(xlsx_book, 1:2) |> length == 2
    end

    @testset "Case №7: invalid sheet keys" begin
        xlsx_book = parse_xlsx("data/two_sheets.xlsx")

        @test_throws ErrorException("KeyError: no sheet name `Sheet3`") begin
            xl_sheets(xlsx_book, "Sheet3")
        end

        @test_throws ErrorException("KeyError: no sheet index `3`") begin
            xl_sheets(xlsx_book, 3)
        end

        exp_msg = "KeyError: invalid sheet names list `[\"Sheet1\", \"Sheet3\"]`"
        @test_throws ErrorException(exp_msg) begin
            xl_sheets(xlsx_book, ["Sheet1", "Sheet3"])
        end

        @test_throws ErrorException("KeyError: invalid sheet names range `2:3`") begin
            xl_sheets(xlsx_book, 2:3)
        end
    end
end