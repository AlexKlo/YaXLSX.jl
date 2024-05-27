mutable struct ExcelSheet
    name::String
    data::Dict
    dimension::Tuple{Int64, Int64}
end

mutable struct ExcelBook
    io::IOBuffer
    sheets::Vector{ExcelSheet}
    sheet_names::Vector{String}
end