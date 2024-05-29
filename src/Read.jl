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

function _parse_cell(cell::EzXML.Node, shared_strings::Vector{String})
    t = haskey(cell, "t") ? cell["t"] : ""

    for child_elem in EzXML.eachelement(cell)
        child_name = EzXML.nodename(child_elem)
        if child_name == "v"
            value = EzXML.nodecontent(child_elem)
            t == "s" && return shared_strings[parse(Int, value) + 1]
            t == "b" && return value == "TRUE"
            return parse(Float64, value)
        elseif child_name == "f"
            return EzXML.nodecontent(child_elem)
        end
    end
    
    return nothing
end

function _parse_sheet(ws_doc::EzXML.Document, shared_strings::Vector{String}, sheet_name::String)
    rows = findall("/x:worksheet/x:sheetData/x:row", ws_doc.root, ["x"=>ns])
    data = Dict{String, Any}()
    last_cell_ref = ""
    for row in rows
        cells = EzXML.findall("x:c", row, ["x"=>ns])

        for cell in cells
            cell_ref = cell["r"]

            cell_content = _parse_cell(cell, shared_strings)
            isnothing(cell_content) && continue
            data[cell_ref] = cell_content

            last_cell_ref = cell_ref
        end
    end

    return ExcelSheet(sheet_name, data)
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

    shared_strings = []
    ss_doc = _read_xml_by_name(zip, "xl/sharedStrings.xml")
    isnothing(ss_doc) && error("ParseError: xl/sharedStrings.xml not found")
    shared_strings = [
        EzXML.nodecontent(EzXML.findfirst("x:t", si, ["x"=>ns])) 
        for si in EzXML.findall("//x:si", ss_doc.root, ["x"=>ns])
    ]

    sheets = Vector{ExcelSheet}(undef, length(sheet_names))
    for (i, name) in enumerate(sheet_names)
        ws_doc = _read_xml_by_name(zip, "xl/worksheets/sheet$(i).xml")
        isnothing(ws_doc) && error("ParseError: xl/worksheets/sheet$(i).xml not found")
        sheet = _parse_sheet(ws_doc, shared_strings, name)
        sheets[i] = sheet
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