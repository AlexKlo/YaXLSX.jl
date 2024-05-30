# YaXLSX

[![Build Status](https://github.com/AlexKlo/YaXLSX.jl/actions/workflows/CI.yml/badge.svg?branch=main)](https://github.com/AlexKlo/YaXLSX.jl/actions/workflows/CI.yml?query=branch%3Amain)
[![Documentation](https://img.shields.io/badge/docs-latest-blue.svg)](https://AlexKlo.github.io/YaXLSX.jl/)

YaXLSX is a Julia package for parsing data from .xlsx files.

# Usage

The following is an example of how you can parse data from .xlsx file.

It is possible to parse a file from a vector of bytes or along the path

```julia-repl
julia> using YaXLSX

julia> xl_book = parse_xlsx(read("path_to/file.xlsx"))
ExcelBook

julia> xl_book = parse_xlsx("path_to/file.xlsx")
ExcelBook
```

There are interfaces for obtaining data from the book: 

`xl_sheetnames` to get a list of all sheet names in a book
```julia-repl
julia> xl_sheetnames(xl_book)
2-element Vector{String}:
 "Sheet1"
 "Sheet2"
```

`xl_sheets` to sheets from a book

```julia-repl
julia> xl_sheets(xl_book)
2-element Vector{YaXLSX.ExcelSheet}:
 ExcelSheet(Sheet1)
 ExcelSheet(Sheet2)

julia> xl_sheets(xl_book, "Sheet1")
ExcelSheet(Sheet1)

julia> xl_sheets(xl_book, 2)
ExcelSheet(Sheet2)

julia> xl_sheets(xl_book, ["Sheet1", "Sheet2"])
2-element Vector{YaXLSX.ExcelSheet}:
 ExcelSheet(Sheet1)
 ExcelSheet(Sheet2)

 julia> xl_sheets(xl_book, 1:2)
2-element Vector{YaXLSX.ExcelSheet}:
 ExcelSheet(Sheet1)
 ExcelSheet(Sheet2)
```

There are also interfaces for obtaining tabular data from sheets `xl_rowtable`, `xl_columntable`

```julia-repl
julia> xl_sheet = xl_sheets(xl_book, 2)
ExcelSheet(Sheet2)

julia> xl_rowtable(xl_sheet, "A1:B6")
6×2 DataFrameRows
 Row │ A        B
     │ Any      Any
─────┼────────────────
   1 │ Numbers  Names
   2 │ 1.0      a
   3 │ 2.0      b
   4 │ 3.0      c
   5 │ 4.0      d
   6 │ 5.0      e

julia> xl_columntable(xl_sheet, "A1:C3"; headers=["h1", "h2", "h3"])
3×3 DataFrameColumns
 Row │ h1       h2     h3
     │ Any      Any    Any
─────┼─────────────────────────
   1 │ Numbers  Names  missing
   2 │ 1.0      a      missing
   3 │ 2.0      b      missing
```