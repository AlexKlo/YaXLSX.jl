# SampleData

export xlsx_simple_table, xlsx_two_sheets_table

function xlsx_simple_table()
    return read(joinpath(@__DIR__, "../test/data/simple_table.xlsx"))
end

function xlsx_two_sheets_table()
    return read(joinpath(@__DIR__, "../test/data/two_sheets.xlsx"))
end