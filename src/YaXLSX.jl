module YaXLSX

export parse_xlsx, xl_sheets

using Serde, ZipFile

#__ Workbook

struct WSheet
    sheetId::Int64
    name::String
    id::String
    state::Union{String, Nothing}
end

struct WSheets
    sheet::Vector{WSheet} # здесь проблема, если в Workbook 1 лист.
end

struct Workbook
    sheets::WSheets
end

#__ Sheet

struct CellValue
    _::String
end

struct Cell
    v::Union{Nothing,CellValue}
    t::Union{Nothing,String}
    r::String
    s::Int64
end

struct Row
    r::Int64
    c::Vector{Cell}
end

struct Rows
    row::Vector{Row}
end

struct Sheet
    sheetData::Rows
    #__
end

#__ sharedStrings

struct sharedValue
    _::String
end

struct Si
    t::sharedValue
end

struct sharedStrings
    uniqueCount::Int64
    si::Vector{Si}
    count::Int64
end

#__

struct XLSX
    sheets::Vector{Sheet}
    shared_strings::sharedStrings
end

function Base.show(io::IO, x::XLSX)
    return print(io, "XLSX with $(length(x.sheets)) sheets")
end

#__

function unzip(zip_bytes::Vector{UInt8})
    files = Dict{String,Vector{UInt8}}()
    zip_archive = ZipFile.Reader(IOBuffer(zip_bytes))
    for entry in zip_archive.files
        files[entry.name] = read(entry)
    end
    return files
end

function sheet_names(x::Workbook)
    return map(x -> string("sheet", x.sheetId), x.sheets.sheet)
end

function parse_xlsx(x::AbstractString)
    return parse_xlsx(Vector{UInt8}(x))
end

function parse_xlsx(x::Vector{UInt8})
    files = unzip(x)
    workbook = deser_xml(Workbook, files["xl/workbook.xml"])
    sheets = map(
        x -> deser_xml(Sheet, files["xl/worksheets/$(x).xml"]),
        sheet_names(workbook),
    )
    shared_strings = deser_xml(sharedStrings, files["xl/sharedStrings.xml"])
    return XLSX(sheets, shared_strings)
end

#__ mvp

function xl_sheets(xlsx::XLSX, n::Int64)

    max_rows = length(xlsx.sheets[n].sheetData.row)
    max_cols = maximum(x -> length(x.c), xlsx.sheets[n].sheetData.row)

    tbls = Matrix{Any}(nothing, max_rows, max_cols)

    for row in xlsx.sheets[n].sheetData.row
        for (c_index, col) in enumerate(row.c)
            value = if col.v == nothing
                nothing
            elseif col.t == "s"
                xlsx.shared_strings.si[parse(Int64, col.v._)+1].t._
            elseif col.t == "b"
                col.v._ == "1"
            else
                col.v._
            end
            tbls[row.r, c_index] = value
        end
    end

    return tbls
end

end
