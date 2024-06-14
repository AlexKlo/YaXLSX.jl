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

function parse_xlsx(x::AbstractString)
    return parse_xlsx(read(x))
end

function parse_xlsx(zip_bytes::Vector{UInt8})
    check_file_format(zip_bytes)
    files = unzip(zip_bytes)
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
    tables = map(s -> get_table(s, shared_strings), sheets)

    return XLSX(workbook, sheets, tables, shared_strings)
end

#__ Table

@inline function letter2num(letter::AbstractString)::Int64
    @assert !isempty(letter)
    id = 0
    n = length(letter)
    for i in eachindex(letter)
        j = Int64(letter[n-i+1]) - 65 + 1
        id = id + j + 26
    end
    return id - 26 - n + 1
end

@inline function get_text_string(s::Union{SI, IS})
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

function get_table(sheet::Sheet, shared_strings::Union{Nothing, sharedStrings})
    isempty(sheet) && return Matrix{Any}(nothing, 0, 0)

    max_rows = max_row_number(sheet)
    max_cols = max_col_number(sheet)

    matrix = Matrix{Any}(nothing, max_rows, max_cols)

    for row in sheet.sheetData.row
        for col in row.c
            value = if col.t == "inlineStr"
                get_text_string(col.is)
            elseif isnothing(col.v)
                nothing 
            elseif col.t == "s"
                get_text_string(shared_strings.si[parse(Int64, col.v._)+1])
            elseif col.t == "b"
                col.v._ == "1"
            elseif col.t in ["str", "e"]
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