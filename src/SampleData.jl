# SampleData

export xlsx_simple_table

function xlsx_simple_table()
    return read(joinpath(@__DIR__, "../test/data/simple_table.xlsx"))
end