# Types

#__ Workbook

struct WSheet
    sheetId::Int64
    name::String
    id::String
    state::Union{String, Nothing}
end

struct WSheets
    sheet::Vector{WSheet}
end

function Serde.deser(::Type{WSheets}, ::Type{Vector{W}}, x::AbstractDict) where {W<:WSheet}
    return W[Serde.deser(W, x)]
end

struct Workbook
    sheets::WSheets
end

#__ Sheet

struct CellValue
    _::Union{Nothing,String}
end

struct CellText
    _::String
end

function Serde.deser(::Type{CellText}, ::Type{String}, x::Nothing)
    return ""
end

struct RichText
    t::Union{Nothing, CellText}
end

struct IS
    t::Union{Nothing,CellText}
    r::Vector{RichText}
end

function Serde.deser(::Type{IS}, ::Type{Vector{R}}, x::AbstractDict) where {R<:RichText}
    return R[Serde.deser(R, x)]
end

function Serde.deser(::Type{IS}, ::Type{Vector{R}}, x::Nothing) where {R<:RichText}
    return R[]
end

struct Cell
    v::Union{Nothing,CellValue}
    t::Union{Nothing,String}
    r::String
    s::Union{Nothing,Int64}
    is::Union{Nothing, IS}
end

struct Row
    r::Int64
    c::Vector{Cell}
end

function Serde.deser(::Type{Row}, ::Type{Vector{C}}, x::AbstractDict) where {C<:Cell}
    return C[Serde.deser(C, x)]
end

function Serde.deser(::Type{Row}, ::Type{Vector{C}}, x::Nothing) where {C<:Cell}
    return C[]
end

struct Rows
    row::Vector{Row}
end

function Serde.deser(::Type{Rows}, ::Type{Vector{R}}, x::AbstractDict) where {R<:Row}
    return R[Serde.deser(R, x)]
end

function Serde.deser(::Type{Rows}, ::Type{Vector{R}}, x::Nothing) where {R<:Row}
    return R[]
end

mutable struct Sheet
    sheetData::Rows
    table::Union{Nothing, Matrix}
    name::Union{Nothing, String}
end

Base.isempty(sheet::Sheet) = isempty(sheet.sheetData.row)

function Base.show(io::IO, x::Sheet)
    return print(io, "Sheet(\"$(x.name)\")")
end

#__ sharedStrings

struct SI
    t::Union{Nothing,CellText}
    r::Vector{RichText}
end

function Serde.deser(::Type{SI}, ::Type{Vector{R}}, x::AbstractDict) where {R<:RichText}
    return R[Serde.deser(R, x)]
end

function Serde.deser(::Type{SI}, ::Type{Vector{R}}, x::Nothing) where {R<:RichText}
    return R[]
end

struct sharedStrings
    uniqueCount::Int64
    si::Union{Nothing,Vector{SI}}
    count::Union{Nothing,Int64}
end

function Serde.deser(::Type{sharedStrings}, ::Type{Vector{S}}, x::AbstractDict) where {S<:SI}
    return S[Serde.deser(S, x)]
end

#__ XLSX

struct XLSX
    workbook::Workbook
    sheets::Vector{Sheet}
    shared_strings::Union{Nothing,sharedStrings}
end

function Base.show(io::IO, x::XLSX)
    len = length(x.sheets)
    return print(io, string("XLSX with $len sheet", len==1 ? "" : "s"))
end
