const MAX_COLUMN_NUMBER = 16384
const MAX_ROW_NUMBER = 1048576

_sheet_dims(sheet::ExcelSheet) = NamedTuple{(:n_rows,:n_cols)}(size(sheet.data))

@inline function _number_to_xl_column(n::Int)
    n, r1 = divrem(n - 1, 26)
    iszero(n) && return string(Char(r1 + 65))

    n, r2 = divrem(n - 1, 26)
    iszero(n) && return string(Char(r2 + 65), Char(r1 + 65))

    _, r3 = divrem(n - 1, 26)
    return string(Char(r3 + 65), Char(r2 + 65), Char(r1 + 65))
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
        Tuple(Symbol(_number_to_xl_column(i)) for i in left:right)
    else
        length(headers) == width || error(
            "KeyError: headers=$headers length must be equal \
            to cell range length: $width"
        )
        Tuple(Symbol(h) for h in headers)
    end

    data = Array{Any, 2}(undef, height, width)

    for col in left:right
        for row in top:bottom
            rel_row = row - top + 1
            rel_col = col - left + 1
            data[rel_row, rel_col] = get(sheet.data, (row, col), missing)
        end
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

        top, left = 1, parse(Int64, top_left)
        bottom, right = n_rows, parse(Int64, bottom_right)      

    else
        error("KeyError: invalid cell range `$cell_range`")
    end

    if iszero(n_rows)
        left<=right || error("KeyError: invalid cell range `$cell_range`")
    else
        left<=right && top<=bottom || error("KeyError: invalid cell range `$cell_range`")
    end

    (left <= MAX_COLUMN_NUMBER && right <= MAX_COLUMN_NUMBER ) || 
        error("KeyError: column index is too large. Maximum 16384 or XFD")
    (top <= MAX_ROW_NUMBER && bottom <= MAX_ROW_NUMBER ) || 
        error("KeyError: row index is too large. Maximum 1048576")

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
    top_left, bottom_right = _cell_range_to_indices(cell_range, _sheet_dims(sheet).n_rows)
    ntp = _get_table(sheet, top_left, bottom_right)
    return ntp |> Tables.rows
end

function xl_rowtable(sheet::ExcelSheet)
    ntp = _get_table(sheet, (1, 1), Tuple(_sheet_dims(sheet)))
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
    top_left, bottom_right = _cell_range_to_indices(cell_range, _sheet_dims(sheet).n_rows)
    ntp = _get_table(sheet, top_left, bottom_right; headers=headers)
    return ntp |> Tables.columns
end

function xl_columntable(sheet::ExcelSheet; headers::Union{Nothing, Vector{String}}=nothing)
    ntp = _get_table(sheet, (1, 1), Tuple(_sheet_dims(sheet)); headers=headers)
    return ntp |> Tables.columns
end