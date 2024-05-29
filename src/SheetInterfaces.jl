@inline function _number_to_xl_column(n::Int)
    column_name = ""
    while n > 0
        n -= 1
        column_name = string(Char(n % 26 + 65)) * column_name
        n = div(n, 26)
    end
    return column_name
end

@inline function _cell_to_indices(cell_ref::AbstractString)
    row_idx = parse(Int64, match(r"\d+", cell_ref).match)

    col_letter = match(r"[A-Z]+", cell_ref).match
    col_idx = 0
    length_col = length(col_letter)
    for (i, char) in enumerate(col_letter)
        col_idx += (Int(char) - Int('A') + 1) * 26^(length_col - i)
    end

    return (row_idx, col_idx)
end

function _get_table(sheet::ExcelSheet, cell_range::AbstractString)
    left_top, right_bottom = split(cell_range, ":")

    top, left = _cell_to_indices(left_top)
    bottom, right = _cell_to_indices(right_bottom)

    left <= right && top <= bottom || error("KeyError: invalid cell range `$cell_range`")

    data = Array{Any, 2}(undef, bottom - top + 1, right - left + 1)
    fill!(data, missing)

    for (cell_ref, value) in sheet.data
        row, col = _cell_to_indices(cell_ref)
        top <= row <= bottom || continue
        left <= col <= right  || continue
        rel_row = row - top + 1
        rel_col = col - left + 1
        data[rel_row, rel_col] = value
    end
    names = [_number_to_xl_column(i) for i in left:right]

    return DataFrame(data, names)
end

"""
    xl_rowtable(sheet::ExcelSheet, cell_range::AbstractString)

Getting the rows table from the specified range.

## Examples

```julia-repl
julia> xl_book = parse_xlsx(read("path_to/file.xlsx"))

julia> xl_sheet = xl_sheets(xl_book, "Sheet1")

julia> xl_rowtable(xl_sheet, "A1:B2")
```
"""
function xl_rowtable(sheet::ExcelSheet, cell_range::AbstractString)
    df = _get_table(sheet, cell_range)
    return df |> eachrow
end


"""
    xl_columntable(
        sheet::ExcelSheet, 
        cell_range::AbstractString; 
        headers::Union{Nothing, Vector{String}}=nothing
    )

Getting the columns table from the specified range.

## Examples

```julia-repl
julia> xl_book = parse_xlsx(read("path_to/file.xlsx"))

julia> xl_sheet = xl_sheets(xl_book, "Sheet1")

julia> xl_columntable(xl_sheet, "A1:B2")

julia> xl_columntable(xl_sheet, "A1:B2"; headers=["col1", "col2"])
```
"""
function xl_columntable(
    sheet::ExcelSheet, 
    cell_range::AbstractString; 
    headers::Union{Nothing, Vector{String}}=nothing
)
    df = _get_table(sheet, cell_range)
    isnothing(headers) || rename!(df, headers)
    return df |> eachcol
end