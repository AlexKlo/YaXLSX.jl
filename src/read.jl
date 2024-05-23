const ZIP_FILE_HEADER = UInt8[ 0x50, 0x4b, 0x03, 0x04 ] # valid Zip file header

function _check_file_format(byte_array::Vector{UInt8})
    header = @view(byte_array[1:4])

    if header == ZIP_FILE_HEADER 
        return nothing
    else
        error("WrongExtension: supported only .xlsx format")
    end
    return nothing
end

"""
    parse_xlsx(byte_array::Vector{UInt8})

Passing a byte array from an XLSX file
"""
function parse_xlsx(byte_array::Vector{UInt8})
    _check_file_format(byte_array)
    io = IOBuffer(byte_array)

    zip = ZipFile.Reader(io)
    wb_xml = nothing
    for f in zip.files
        f.name == "xl/workbook.xml" || continue
        wb_content = read(f, String)
        wb_xml = EzXML.root(EzXML.readxml(IOBuffer(wb_content)))
        break
    end
    @assert !isnothing(wb_xml) "ParseError: xl/workbook.xml not found"
    sheets = ExcelSheet[]
    sheet_names = String[]
    for node in EzXML.eachelement(wb_xml)
        EzXML.nodename(node) == "sheets" || continue

        for sheet_node in EzXML.eachelement(node)
            EzXML.nodename(sheet_node) == "sheet" || continue
            push!(sheets, ExcelSheet())
            push!(sheet_names, sheet_node["name"])
        end

        break
    end

    return ExcelBook(io, sheets, sheet_names)
end

"""
    parse_xlsx(byte_array::Vector{UInt8})

Passing an XLSX file from path
"""
function parse_xlsx(path::AbstractString)
    return parse_xlsx(read(path))
end