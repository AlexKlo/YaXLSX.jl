# YaXLSX.jl

YaXLSX package helps to read .xlsx files and construct tables for data processing.

## Installation

To install YaXLSX, simply use the Julia package manager:

```julia
] add YaXLSX
```

## Usage

The following is an example of how you can parse data from .xlsx file, retrieve sheet information, and create tabular data.

```julia
using YaXLSX

julia> xl_book = parse_xlsx(xlsx_simple_table())
XLSX with 1 sheet

julia> xl_sheetnames(xl_book)
1-element Vector{String}:
 "Sheet1"

julia> xl_sheet = xl_sheets(xl_book, "Sheet1")
Sheet("Sheet1")

julia> xl_rowtable(xl_sheet, "A1:B6")
6-element Vector{NamedTuple{(:A, :B)}}:
 (A = "Numbers", B = "Names")
 (A = 1.0, B = "a")
 (A = 2.0, B = "b")
 (A = 3.0, B = "c")
 (A = 4.0, B = "d")
 (A = 5.0, B = "e")

julia> xl_columntable(xl_sheet, "1:2"; column_labels=["column1", "column2"])
(
    column1 = Any["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0], 
    column2 = Any["Names", "a", "b", "c", "d", "e"]
)
```

It is also possible to convert tabular data into a `DataFrame` using [DataFrames.jl](https://dataframes.juliadata.org/stable/)

```julia
using DataFrames

using YaXLSX

julia> xl_rowtable(xl_sheets(parse_xlsx(xlsx_simple_table()), 1)) |> DataFrame
6×2 DataFrame
 Row │ A        B
     │ Any      String
─────┼─────────────────
   1 │ Numbers  Names
   2 │ 1.0      a
   3 │ 2.0      b
   4 │ 3.0      c
   5 │ 4.0      d
   6 │ 5.0      e
```

Another feature is that you can use [Serde.jl](https://bhftbootcamp.github.io/Serde.jl/stable/) to serialize data, for example to XML format

```julia
using Serde

using YaXLSX

xl_sheet = xl_sheets(parse_xlsx(xlsx_simple_table()), 1)

julia> to_xml(xl_sheet.sheetData) |> print
<xml>
  <row r="1">
    <c t="s" r="A1" s="1">
      <v>0</v>
    </c>
    <c t="s" r="B1" s="2">
      <v>1</v>
    </c>
  </row>
  <row r="2">
    <c r="A2" s="3">
      <v>1</v>
    </c>
    <c t="s" r="B2" s="2">
      <v>2</v>
    </c>
  </row>
    ...
</xml>
```