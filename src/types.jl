mutable struct ExcelSheet
    
end

mutable struct ExcelBook
    io::IOBuffer
    sheets::Vector{ExcelSheet}
    sheet_names::Vector{String}
end