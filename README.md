# YaXLSX

[![Build Status](https://github.com/AlexKlo/YaXLSX.jl/actions/workflows/CI.yml/badge.svg?branch=main)](https://github.com/AlexKlo/YaXLSX.jl/actions/workflows/CI.yml?query=branch%3Amain)
[![Documentation](https://img.shields.io/badge/docs-latest-blue.svg)](https://AlexKlo.github.io/YaXLSX.jl/)

YaXLSX is a Julia package for parsing data from .xlsx files.

# Usage

The following is an example of how you can parse data from .xlsx file.

It is possible to parse a file from a vector of bytes or along the path

```julia
julia> using YaXLSX

julia> xl_book = parse_xlsx(read("path_to/file.xlsx"))
ExcelBook

julia> xl_book = parse_xlsx("path_to/file.xlsx")
ExcelBook
```

There are interfaces for obtaining data from the book: 

`xl_sheetnames` to get a list of all sheet names in a book
```julia
julia> xl_sheetnames(xl_book)
2-element Vector{String}:
 "Sheet1"
 "Sheet2"
```

`xl_sheets` to get sheets from a book

```julia
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

```julia
julia> using DataFrames

julia> xl_sheet = xl_sheets(xl_book, 2)
ExcelSheet(Sheet2)

julia> xl_rowtable(xl_sheet, "A1:B6") |> DataFrame
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

julia> xl_columntable(xl_sheet, "A1:C3"; headers=["h1", "h2", "h3"]) |> DataFrame
3×3 DataFrameColumns
 Row │ h1       h2     h3
     │ Any      Any    Any
─────┼─────────────────────────
   1 │ Numbers  Names  missing
   2 │ 1.0      a      missing
   3 │ 2.0      b      missing
```

If you want to see more getting data options, then take a look at the corresponding [section](https://alexklo.github.io/YaXLSX.jl/#Sheet-interfaces) of the documentation.

# Collaboration with Serde

Collaboration with the (de)serialization package [Serde.jl](https://github.com/bhftbootcamp/Serde.jl) is also possible:

```julia
julia> using Serde, YaXLSX

julia> struct Sheet
           A::AbstractArray
           B::AbstractArray
           C::AbstractArray
       end

julia> xl_book=parse_xlsx("path_to/file.xlsx")
ExcelBook

julia> xl_sheet = xl_sheets(xl_book, 1)
ExcelSheet(Sheet1)

julia> column_data = xl_columntable(xl_sheet, "A:C")
(A = Any["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0], B = Any["Names", "a", "b", "c", "d", "e"], C = Any[missing, missing, missing, missing, missing, missing])

julia> Serde.deser(Sheet, column_data)
Sheet(Any["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0], Any["Names", "a", "b", "c", "d", "e"], Any[missing, missing, missing, missing, missing, missing])

julia> struct SheetCustomHeaders
           Numbers::AbstractArray
           Names::AbstractArray
           empty::AbstractArray
       end

julia> column_data = xl_columntable(xl_sheet, "A:C"; headers=["Numbers", "Names", "empty"])
(Numbers = Any["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0], Names = Any["Names", "a", "b", "c", "d", "e"], empty = Any[missing, missing, missing, missing, missing, missing])

julia> s = Serde.deser(SheetCustomHeaders, column_data)
SheetCustomHeaders(Any["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0], Any["Names", "a", "b", "c", "d", "e"], Any[missing, missing, missing, missing, missing, missing])

julia> s.Names
6-element view(::Matrix{Any}, :, 2) with eltype Any:
 "Names"
 "a"
 "b"
 "c"
 "d"
 "e"
```