
@testset "Reading" begin
    @testset "Case №1: Simple reading" begin
        xlsx_book = parse_xlsx(read("data/simple_book.xlsx"))
        @test xlsx_book isa IOBuffer
    end

    @testset "Case №2: Invalid format file reading" begin
        @test_throws "WrongExtension: " begin
            xlsx_book = parse_xlsx(read("data/invalid.xls"))
        end
    end
end
