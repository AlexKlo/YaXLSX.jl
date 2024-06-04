@inline function _number_to_xl_column(n::Int)
    column_name = ""
    while n > 0
        n -= 1
        column_name = string(Char(n % 26 + 65)) * column_name
        n = div(n, 26)
    end
    return Symbol(column_name)
end

@inline function _cell_to_indices(cell_ref::AbstractString)
    row_idx = parse(Int64, match(r"\d+", cell_ref).match)

    col_letter = match(r"[A-Z]+", cell_ref).match
    col_idx = 0
    length_col = length(col_letter)
    for (i, char) in enumerate(col_letter)
        col_idx = (Int(char) - Int('A') + 1) + col_idx * 26
    end

    return (row_idx, col_idx)
end

function _get_table(
    sheet::ExcelSheet, top_left::Tuple,
    bottom_right::Tuple;
    headers::Union{Nothing, Vector{String}}=nothing
)
    top, left = top_left
    bottom, right = bottom_right

    width = right - left + 1
    height = bottom - top + 1
    names = if isnothing(headers)
        Tuple(_number_to_xl_column(i) for i in left:right)
    else
        length(headers) == width || error(
            "KeyError: headers=$headers length must be equal \
            to cell range length: $width"
        )
        Tuple(Symbol(h) for h in headers)
    end

    data = Array{Any, 2}(missing, height, width)

    for (cell_ref, value) in sheet.data
        row, col = _cell_to_indices(cell_ref)
        top <= row <= bottom || continue
        left <= col <= right  || continue
        rel_row = row - top + 1
        rel_col = col - left + 1
        data[rel_row, rel_col] = value
    end

    return NamedTuple{names}(eachcol(data))
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

    r = r"([A-Z]*)(\d*):([A-Z]*)(\d*)"
    m = match(r, cell_range)
    if all(!isempty, m.captures)

        top, left = _cell_to_indices(top_left)
        bottom, right = _cell_to_indices(bottom_right)

    elseif all(!isempty, m.captures[1:2:end]) && all(isempty, m.captures[2:2:end])

        top, left = _cell_to_indices(top_left*"1")
        bottom, right = _cell_to_indices(bottom_right*"$n_rows")

    elseif all(isempty, m.captures[1:2:end]) && all(!isempty, m.captures[2:2:end])

        l = parse(Int64, top_left)
        r = parse(Int64, bottom_right)
        return _cell_range_to_indices(l:r, n_rows)        

    else
        error("KeyError: invalid cell range `$cell_range`")
    end

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
    ntp = _get_table(sheet, top_left, bottom_right)
    return ntp |> Tables.rows
end

function xl_rowtable(sheet::ExcelSheet)
    ntp = _get_table(sheet, (1, 1), (sheet.dim.n_rows, sheet.dim.n_cols))
    return ntp |> Tables.rows
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
    ntp = _get_table(sheet, top_left, bottom_right; headers=headers)
    return ntp |> Tables.columns
end

function xl_columntable(sheet::ExcelSheet; headers::Union{Nothing, Vector{String}}=nothing)
    ntp = _get_table(sheet, (1, 1), (sheet.dim.n_rows, sheet.dim.n_cols); headers=headers)
    return ntp |> Tables.columns
end