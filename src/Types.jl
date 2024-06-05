"""
    ExcelSheet

Structure with the contents of the book sheets contained in ExcelBook structure.

## Fields
- `name::String`: sheet name.
- `data::Matrix`: sheet data.
"""
mutable struct ExcelSheet
    name::String
    data::Matrix
end

function Base.show(io::IO, es::ExcelSheet)
    print(io, "ExcelSheet($(es.name))")
end

"""
    ExcelBook

Structure with the contents of the .xlsx file returned by the [`parse_xlsx`](@ref) function.

## Fields
- `sheets::Vector{ExcelSheet}`: vector book sheets.
- `sheet_names::Vector{String}`: vector book sheet names.
"""
mutable struct ExcelBook
    sheets::Vector{ExcelSheet}
    sheet_names::Vector{String}
end

function Base.show(io::IO, eb::ExcelBook)
    print(io, "ExcelBook")
end