const ZIP_FILE_HEADER = UInt8[ 0x50, 0x4b, 0x03, 0x04 ] # valid Zip file header
const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

function _check_file_format(byte_array::Vector{UInt8})
    header = @view(byte_array[1:4])

    if header == ZIP_FILE_HEADER 
        return nothing
    else
        error("WrongExtension: supported only .xlsx format")
    end
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


function _parse_sheet(ws_doc::EzXML.Document, shared_strings::Vector{String}, sheet_name::String)
    rows = findall("/x:worksheet/x:sheetData/x:row", ws_doc.root, ["x"=>ns])

    col_lengths = zeros(length(rows))
    for (j, row) in enumerate(rows)
        col_lengths[j] =  EzXML.findall("x:c", row, ["x"=>ns]) |> length
    end
    max_cols = Int(maximum(col_lengths))
    data = Matrix(undef, length(rows), max_cols)
    fill!(data, missing)
    col_names = String[]
    
    for row in rows
        cells = EzXML.findall("x:c", row, ["x"=>ns])

        for (col_index, cell) in enumerate(cells)
            col_name = replace(cell["r"], r"[1-9]"=>"")
            col_name in col_names || push!(col_names, col_name)
            row_index = parse(Int, match(r"\d+", cell["r"]).match)

            if EzXML.findfirst("x:v", cell, ["x"=>ns]) !== nothing
                value = EzXML.nodecontent(EzXML.findfirst("x:v", cell, ["x"=>ns]))

                if EzXML.haskey(cell, "t") && EzXML.getindex(cell, "t") == "s"
                    data[row_index, col_index] = shared_strings[parse(Int, value) + 1]
                else
                    data[row_index, col_index] = parse(Float64, value)
                end
            end
        end
    end
    return ExcelSheet(sheet_name, DataFrame(data, col_names))
end

"""
    parse_xlsx(byte_array::Vector{UInt8})

Passing a byte array from an XLSX file
"""
function parse_xlsx(byte_array::Vector{UInt8})
    _check_file_format(byte_array)
    io = IOBuffer(byte_array)
    zip = ZipFile.Reader(io)

    wb_doc = _read_xml_by_name(zip, "xl/workbook.xml")
    @assert !isnothing(wb_doc) "ParseError: xl/workbook.xml not found"

    sheet_names = String[ 
        sheet["name"] for sheet in EzXML.findall("//x:sheet", wb_doc.root, ["x"=>ns]) 
    ]

    # Чтение sharedStrings.xml
    shared_strings = []
    ss_doc = _read_xml_by_name(zip, "xl/sharedStrings.xml")
    shared_strings = [
        EzXML.nodecontent(EzXML.findfirst("x:t", si, ["x"=>ns])) 
        for si in EzXML.findall("//x:si", ss_doc.root, ["x"=>ns])
    ]

    sheets = Vector{ExcelSheet}(undef, length(sheet_names))
    for (i, name) in enumerate(sheet_names)
        ws_doc = _read_xml_by_name(zip, "xl/worksheets/sheet$(i).xml")
        sheet = _parse_sheet(ws_doc, shared_strings, name)
        sheets[1] = sheet
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