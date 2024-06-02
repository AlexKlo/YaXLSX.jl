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

function _get_table(sheet::ExcelSheet, top_left::Tuple, bottom_right::Tuple)
    top, left = top_left
    bottom, right = bottom_right

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

function _cell_range_to_indices(cell_range::UnitRange{Int64}, n_rows::Int64)
    top = 1
    left = cell_range.start
    bottom = n_rows
    right = cell_range.stop

    left<=right || error("KeyError: invalid cell range `$cell_range`")

    return (top, left), (bottom, right)
end

function _cell_range_to_indices(cell_range::AbstractString, n_rows::Int64)
    range = split(cell_range, ":")

    length(range) == 2 || error("KeyError: invalid cell range `$cell_range`")

    top_left, bottom_right = range[1], range[2]

    r1 = r"^\D*$"
    if occursin(r1, top_left)
        if occursin(r1, bottom_right)
            top_left *= "1"
            bottom_right *= "$n_rows"
        else
            error("KeyError: invalid cell range `$cell_range`")
        end
    end

    r2 = r"^\d+$"
    if occursin(r2, top_left)
        if occursin(r2, bottom_right)
            l = parse(Int64, top_left)
            r = parse(Int64, bottom_right)
            return _cell_range_to_indices(l:r, n_rows)
        else
            error("KeyError: invalid cell range `$cell_range`")
        end
    end
    
    r3 = r"^[A-Z]+[0-9]+$"
    occursin(r3, top_left) && occursin(r3, bottom_right) || 
        error("KeyError: invalid cell range `$cell_range`")

    top, left = _cell_to_indices(top_left)
    bottom, right = _cell_to_indices(bottom_right)

    if iszero(n_rows)
        left<=right || error("KeyError: invalid cell range `$cell_range`")
    else
        left<=right && top<=bottom || error("KeyError: invalid cell range `$cell_range`")
    end

    return (top, left), (bottom, right)
end

"""
    xl_rowtable(sheet::ExcelSheet, cell_range::AbstractString)

Getting the rows table from the specified range.

## Examples

```julia-repl
julia> xl_book = parse_xlsx(read("path_to/file.xlsx"))

julia> xl_sheet = xl_sheets(xl_book, "Sheet1")

julia> xl_rowtable(xl_sheet, "A1:B2")

julia> xl_rowtable(xl_sheet, "A:B")

julia> xl_rowtable(xl_sheet, "1:2")

julia> xl_rowtable(xl_sheet, 1:2)

julia> xl_rowtable(xl_sheet)
```
"""
function xl_rowtable(sheet::ExcelSheet, cell_range::Union{AbstractString, UnitRange})
    top_left, bottom_right = _cell_range_to_indices(cell_range, sheet.dim.n_rows)
    df = _get_table(sheet, top_left, bottom_right)
    return df |> eachrow
end

function xl_rowtable(sheet::ExcelSheet)
    df = _get_table(sheet, (1, 1), (sheet.dim.n_rows, sheet.dim.n_cols))
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

julia> xl_columntable(xl_sheet, "A:B")

julia> xl_columntable(xl_sheet, "1:2")

julia> xl_columntable(xl_sheet, 1:2)

julia> xl_columntable(xl_sheet)
```
"""
function xl_columntable(
    sheet::ExcelSheet, 
    cell_range::Union{AbstractString, UnitRange}; 
    headers::Union{Nothing, Vector{String}}=nothing
)
    top_left, bottom_right = _cell_range_to_indices(cell_range, sheet.dim.n_rows)
    df = _get_table(sheet, top_left, bottom_right)
    isnothing(headers) || rename!(df, headers)
    return df |> eachcol
end

function xl_columntable(sheet::ExcelSheet; headers::Union{Nothing, Vector{String}}=nothing)
    df = _get_table(sheet, (1, 1), (sheet.dim.n_rows, sheet.dim.n_cols))
    isnothing(headers) || rename!(df, headers)
    return df |> eachcol
end