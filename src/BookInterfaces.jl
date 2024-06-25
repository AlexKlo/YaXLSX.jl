# BookInterfaces

"""
    xl_sheetnames(xl_book::XLSX)

Get a list of all [`Sheet`](@ref) names in a book.

```julia-repl
julia> xl_book = parse_xlsx(xlsx_two_sheets_table())
xXLSX with 2 sheets

julia> xl_sheetnames(xl_book)
2-element Vector{String}:
 "Sheet1"
 "Sheet2"
```
"""
function xl_sheetnames(xl_book::XLSX)
    return xl_sheetnames(xl_book.workbook)
end

"""
    xl_sheets(xl_book::XLSX) -> Vector{Sheet}
    xl_sheets(xl_book::XLSX, keys_set::Vector{String}) -> Vector{Sheet}
    xl_sheets(xl_book::XLSX, keys_set::UnitRange{Int64}) -> Vector{Sheet}
    xl_sheets(xl_book::XLSX, key::String) -> Sheet
    xl_sheets(xl_book::XLSX, key::Int64) -> Sheet

Get a [`Sheet`](@ref) or Vector{Sheet} by key or set of keys.

```julia-repl
julia> xl_book = parse_xlsx(xlsx_two_sheets_table())
xXLSX with 2 sheets

julia> xl_sheets(xl_book)
2-element Vector{Sheet}:
 Sheet("Sheet1")
 Sheet("Sheet2")

julia> xl_sheets(xl_book, ["Sheet1", "Sheet2"])
2-element Vector{Sheet}:
 Sheet("Sheet1")
 Sheet("Sheet2")

julia> xl_sheets(xl_book, 1:2)
2-element Vector{Sheet}:
 Sheet("Sheet1")
 Sheet("Sheet2")

julia> xl_sheets(xl_book, "Sheet1")
Sheet("Sheet1")

julia> xl_sheets(xl_book, 1)
Sheet("Sheet1")
```
"""
function xl_sheets end

function xl_sheets(xl_book::XLSX)
    return xl_book.sheets
end

function xl_sheets(xl_book::XLSX, key::String)
    key in xl_sheetnames(xl_book) || error("KeyError: no sheet name `$key`")

    index = findfirst(s -> s.name == key, xl_book.sheets)
    return xl_book.sheets[index]
end

function xl_sheets(xl_book::XLSX, key::Int64)
    key <= length(xl_book.sheets) || error("KeyError: no sheet index `$key`")

    return xl_book.sheets[key]
end

function xl_sheets(xl_book::XLSX, keys_set::Vector{String})
    all(key -> key in xl_sheetnames(xl_book), keys_set) || 
        error("KeyError: invalid sheet names list `$keys_set`")

    indices = findall(s -> s.name in keys_set, xl_book.sheets)
    return xl_book.sheets[indices]
end

function xl_sheets(xl_book::XLSX, keys_set::UnitRange{Int64})
    keys_set ⊆ 1:length(xl_book.sheets) || 
        error("KeyError: invalid sheet names range `$keys_set`")
    
    return xl_book.sheets[keys_set]
end