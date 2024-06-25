# Read

const ZIP_FILE_HEADER = UInt8[ 0x50, 0x4b, 0x03, 0x04 ] # valid Zip file header

function check_file_format(zip_bytes::Vector{UInt8})
    header = @view(zip_bytes[1:4])
    header == ZIP_FILE_HEADER || error("WrongExtension: supported only .xlsx format")
    return nothing
end

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

function xl_sheetnames(x::Workbook)
    return map(x -> x.name, x.sheets.sheet)
end


"""
    parse_xlsx(x::AbstractString) -> XLSX
    parse_xlsx(x::Vector{UInt8}) -> XLSX

Parse a XLSX file into a structure of type [`XLSX`](@ref).

## Examples

```julia-repl
julia> xl_book = parse_xlsx(xlsx_simple_table())
XLSX with 1 sheet
```
"""
function parse_xlsx end

function parse_xlsx(x::AbstractString)
    return parse_xlsx(read(x))
end

function parse_xlsx(x::Vector{UInt8})
    check_file_format(x)
    files = unzip(x)
    workbook = deser_xml(Workbook, files["xl/workbook.xml"])
    sheets = map(
        x -> deser_xml(Sheet, files["xl/worksheets/$(x).xml"]),
        sheet_names(workbook),
    )
    shared_strings = if haskey(files, "xl/sharedStrings.xml") 
        deser_xml(sharedStrings, files["xl/sharedStrings.xml"])
    else
        nothing
    end
    map(x -> setfield!(x, :table, data2table(x, shared_strings)), sheets)
    map((x, y) -> setfield!(x, :name, y), sheets, xl_sheetnames(workbook))

    return XLSX(workbook, sheets, shared_strings)
end

#__ Table

@inline function text_string(s::Union{SI, IS})
    str_content = ""
    if !isnothing(s.t)
        str_content *= s.t._
    else
        for r in s.r
            str_content *= r.t._
        end
    end
    return str_content
end

@inline function ref2col(ref::String)
    return letter2num(match(r"\D+", ref).match)
end

function max_row_number(sheet)
    return maximum(x -> x.r, sheet.sheetData.row)
end

function max_col_number(sheet)
    return maximum(ref2col(last(x.c).r) for x in sheet.sheetData.row if !isempty(x.c))
end

function data2table(sheet::Sheet, shared_strings::Union{Nothing, sharedStrings})
    isempty(sheet) && return Matrix{Any}(nothing, 0, 0)

    max_rows = max_row_number(sheet)
    max_cols = max_col_number(sheet)

    matrix = Matrix{Any}(nothing, max_rows, max_cols)

    for row in sheet.sheetData.row
        for col in row.c
            value = if col.t == "inlineStr"
                text_string(col.is)
            elseif !isnothing(col.f)
                isnothing(col.v) ? col.f._ : col.v._
            elseif isnothing(col.v)
                nothing 
            elseif col.t == "s"
                text_string(shared_strings.si[parse(Int64, col.v._)+1])
            elseif col.t == "b"
                col.v._ == "1"
            elseif col.t == "str" || col.t == "e"
                col.v._
            else
                parse(Float64, col.v._)
            end
            c_index = ref2col(col.r)
            matrix[row.r, c_index] = value
        end
    end

    return matrix
end
