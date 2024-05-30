var documenterSearchIndex = {"docs":
[{"location":"#YaXLSX-Documentation","page":"Home","title":"YaXLSX Documentation","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"Welcome to the documentation for YaXLSX.","category":"page"},{"location":"#Parsing","page":"Home","title":"Parsing","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"YaXLSX.parse_xlsx","category":"page"},{"location":"#YaXLSX.parse_xlsx","page":"Home","title":"YaXLSX.parse_xlsx","text":"parse_xlsx(byte_array::Vector{UInt8}) -> ExcelBook\n\nPassing a byte array from an XLSX file.\n\nExamples\n\njulia> parse_xlsx(read(\"path_to/file.xlsx\"))\n\n\n\n\n\nparse_xlsx(byte_array::Vector{UInt8}) -> ExcelBook\n\nPassing an XLSX file from path.\n\nExamples\n\njulia> parse_xlsx(\"path_to/file.xlsx\")\n\n\n\n\n\n","category":"function"},{"location":"#Book-interfaces","page":"Home","title":"Book interfaces","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"YaXLSX.xl_sheetnames\nYaXLSX.xl_sheets","category":"page"},{"location":"#YaXLSX.xl_sheetnames","page":"Home","title":"YaXLSX.xl_sheetnames","text":"xl_sheetnames(xl_book::ExcelBook)\n\nGetting a list of all sheet names in a book.\n\n\n\n\n\n","category":"function"},{"location":"#YaXLSX.xl_sheets","page":"Home","title":"YaXLSX.xl_sheets","text":"xl_sheets(xl_book::ExcelBook)\n\nGetting a list of all ExcelSheet in a book.\n\n\n\n\n\nxl_sheets(xl_book::ExcelBook, key::Union{String, Int64})\n\nGetting a ExcelSheet by index or name.\n\n\n\n\n\nxl_sheets(xl_book::ExcelBook, keys_set::Union{Vector{String}, UnitRange{Int64}})\n\nGetting a list of ExcelSheet by names or range.\n\n\n\n\n\n","category":"function"},{"location":"#Sheet-interfaces","page":"Home","title":"Sheet interfaces","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"YaXLSX.xl_rowtable\nYaXLSX.xl_columntable","category":"page"},{"location":"#YaXLSX.xl_rowtable","page":"Home","title":"YaXLSX.xl_rowtable","text":"xl_rowtable(sheet::ExcelSheet, cell_range::AbstractString)\n\nGetting the rows table from the specified range.\n\nExamples\n\njulia> xl_book = parse_xlsx(read(\"path_to/file.xlsx\"))\n\njulia> xl_sheet = xl_sheets(xl_book, \"Sheet1\")\n\njulia> xl_rowtable(xl_sheet, \"A1:B2\")\n\n\n\n\n\n","category":"function"},{"location":"#YaXLSX.xl_columntable","page":"Home","title":"YaXLSX.xl_columntable","text":"xl_columntable(\n    sheet::ExcelSheet, \n    cell_range::AbstractString; \n    headers::Union{Nothing, Vector{String}}=nothing\n)\n\nGetting the columns table from the specified range.\n\nExamples\n\njulia> xl_book = parse_xlsx(read(\"path_to/file.xlsx\"))\n\njulia> xl_sheet = xl_sheets(xl_book, \"Sheet1\")\n\njulia> xl_columntable(xl_sheet, \"A1:B2\")\n\njulia> xl_columntable(xl_sheet, \"A1:B2\"; headers=[\"col1\", \"col2\"])\n\n\n\n\n\n","category":"function"}]
}
