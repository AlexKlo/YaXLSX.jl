import YaXLSX: ExcelSheet

@testset "BookInterfaces" begin
    @testset "Case №1: get sheets names" begin
        xlsx_book = parse_xlsx(read("data/simple_book.xlsx"))
        @test xl_sheetnames(xlsx_book) == ["Лист1"]

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheetnames(xlsx_book) == ["Лист1", "Лист2"]
    end

    @testset "Case №2: get all sheets" begin
        xlsx_book = parse_xlsx(read("data/simple_book.xlsx"))
        @test xl_sheets(xlsx_book) |> length == 1

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheets(xlsx_book) |> length == 2
    end

    @testset "Case №3: get sheets by name" begin
        xlsx_book = parse_xlsx(read("data/simple_book.xlsx"))
        @test xl_sheets(xlsx_book, "Лист1") isa ExcelSheet

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheets(xlsx_book, "Лист2") isa ExcelSheet
    end

    @testset "Case №4: get sheets by index" begin
        xlsx_book = parse_xlsx(read("data/simple_book.xlsx"))
        @test xl_sheets(xlsx_book, 1) isa ExcelSheet

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheets(xlsx_book, 2) isa ExcelSheet
    end

    @testset "Case №5: get sheets by list" begin
        xlsx_book = parse_xlsx(read("data/simple_book.xlsx"))
        @test xl_sheets(xlsx_book, ["Лист1"]) |> length == 1

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheets(xlsx_book, ["Лист1", "Лист2"]) |> length == 2
    end

    @testset "Case №6: get sheets by range" begin
        xlsx_book = parse_xlsx(read("data/simple_book.xlsx"))
        @test xl_sheets(xlsx_book, 1:1) |> length == 1

        xlsx_book = parse_xlsx("data/two_sheets.xlsx")
        @test xl_sheets(xlsx_book, 1:2) |> length == 2
    end

    @testset "Case №7: invalid sheet keys" begin
        xlsx_book = parse_xlsx("data/two_sheets.xlsx")

        @test_throws ErrorException("KeyError: no sheet name `Лист3`") begin
            xl_sheets(xlsx_book, "Лист3")
        end

        @test_throws ErrorException("KeyError: no sheet index `3`") begin
            xl_sheets(xlsx_book, 3)
        end

        exp_msg = "KeyError: invalid sheet names list `[\"Лист1\", \"Лист3\"]`"
        @test_throws ErrorException(exp_msg) begin
            xl_sheets(xlsx_book, ["Лист1", "Лист3"])
        end

        @test_throws ErrorException("KeyError: invalid sheet names range `2:3`") begin
            xl_sheets(xlsx_book, 2:3)
        end
    end
end