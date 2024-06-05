const ZIP_FILE_HEADER = UInt8[ 0x50, 0x4b, 0x03, 0x04 ] # valid Zip file header
const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

function _check_file_format(byte_array::Vector{UInt8})
    header = @view(byte_array[1:4])
    header == ZIP_FILE_HEADER || error("WrongExtension: supported only .xlsx format")
    return nothing
end

function _read_xml_by_name(zip::ZipFile.Reader, name::String)
    xml_doc = nothing
    for f in zip.files
        f.name == name || continue
        xml_doc = EzXML.parsexml(read(f))
        break
    end
    return xml_doc
end

function _get_shared_strings(ss_doc::Union{Nothing, EzXML.Document})
    shared_strings = String[]
    isnothing(ss_doc) && return shared_strings

    for si in EzXML.findall("//x:si", ss_doc.root, ["x"=>ns])
        str_content = ""

        for parts in EzXML.findall(".//x:t", si, ["x" => ns])
            str_content *= EzXML.nodecontent(parts)
        end
        push!(shared_strings, str_content)
    end

    return shared_strings
end

function _parse_cell(cell::EzXML.Node, shared_strings::Vector{String})
    t = haskey(cell, "t") ? cell["t"] : ""

    for child_elem in EzXML.eachelement(cell)
        child_name = EzXML.nodename(child_elem)

        if t == "inlineStr"

            if child_name == "is"
                t_node = EzXML.findfirst(".//x:t", child_elem, ["x"=>ns])
                isnothing(t_node) || return EzXML.nodecontent(t_node)
            end
        else
            if child_name == "v"
                value = EzXML.nodecontent(child_elem)
                isempty(value) && return value
                t == "s" && return shared_strings[parse(Int, value) + 1]
                t == "b" && return value == "TRUE"
                return parse(Float64, value)
            elseif child_name == "f"
                return EzXML.nodecontent(child_elem)
            end
        end
    end
    
    return nothing
end

function _parse_sheet(ws_doc::EzXML.Document, shared_strings::Vector{String}, sheet_name::String)
    rows = findall("/x:worksheet/x:sheetData/x:row", ws_doc.root, ["x"=>ns])
    isempty(rows) && return ExcelSheet(sheet_name, Matrix{Any}(missing, 0, 0))

    data = Tuple[]
    max_row = 0
    max_col = 0
    for row in rows
        cells = EzXML.findall("x:c", row, ["x"=>ns])

        for cell in cells
            cell_ref = cell["r"]

            cell_content = _parse_cell(cell, shared_strings)
            isnothing(cell_content) && continue
            r, c = _cell_to_indices(cell_ref)
            push!(data, (r, c, cell_content))

            r > max_row && (max_row = r)
            c > max_col && (max_col = c)
        end
    end

    matrix = Matrix{Any}(missing, max_row, max_col)
    for (r, c, value) in data
        matrix[r, c] = value
    end

    return ExcelSheet(sheet_name, matrix)
end

"""
    parse_xlsx(byte_array::Vector{UInt8}) -> ExcelBook

Passing a byte array from an XLSX file.

## Examples

```julia-repl
julia> parse_xlsx(read("path_to/file.xlsx"))
```
"""
function parse_xlsx(byte_array::Vector{UInt8})
    _check_file_format(byte_array)
    io = IOBuffer(byte_array)
    zip = ZipFile.Reader(io)

    wb_doc = _read_xml_by_name(zip, "xl/workbook.xml")
    isnothing(wb_doc) && error("ParseError: xl/workbook.xml not found")

    sheet_names = String[ 
        sheet["name"] for sheet in EzXML.findall("//x:sheet", wb_doc.root, ["x"=>ns]) 
    ]

    ss_doc = _read_xml_by_name(zip, "xl/sharedStrings.xml")
    shared_strings = _get_shared_strings(ss_doc)

    sheets = Vector{ExcelSheet}(undef, length(sheet_names))
    for (i, name) in enumerate(sheet_names)
        ws_doc = _read_xml_by_name(zip, "xl/worksheets/sheet$(i).xml")
        isnothing(ws_doc) && error("ParseError: xl/worksheets/sheet$(i).xml not found")
        sheets[i] = _parse_sheet(ws_doc, shared_strings, name)
    end

    return ExcelBook(sheets, sheet_names)
end

"""
    parse_xlsx(byte_array::Vector{UInt8}) -> ExcelBook

Passing an XLSX file from path.

## Examples

```julia-repl
julia> parse_xlsx("path_to/file.xlsx")
```
"""
function parse_xlsx(path::AbstractString)
    return parse_xlsx(read(path))
end