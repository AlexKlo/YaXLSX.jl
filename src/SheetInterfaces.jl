# SheetInterfaces

function is_valid_index_range(r::UnitRange{Int64})
    return r.start <= r.stop
end

function is_valid_index_range(m::Union{Nothing, RegexMatch})
    isnothing(m) && return false
    all(isempty, @view(m.captures[1:2:end])) && 
        all(!isempty, @view(m.captures[2:2:end])) || return false
    
    l = letter2num(m[2])
    r = letter2num(m[4])

    is_valid_index_range(l:r) || return false

    return true
end

function is_valid_cell_range(m::Union{Nothing, RegexMatch})
    isnothing(m) && return false
    any(isempty, m.captures) && return false
    
    l = letter2num(m[1])
    t = parse(Int64, m[2])
    r = letter2num(m[3])
    b = parse(Int64, m[4])

    l <= r && t <= b || return false

    return true
end

function is_valid_column_name_range(m::Union{Nothing, RegexMatch})
    isnothing(m) && return false
    all(!isempty, @view(m.captures[1:2:end])) && 
        all(isempty, @view(m.captures[2:2:end])) || return false
    
    l = letter2num(m[1])
    r = letter2num(m[3])

    l <= r || return false

    return true
end

function column_range(m::RegexMatch)::Tuple{Int64, Int64}
    l = parse(Int64, m[2])
    r = parse(Int64, m[4])

    return l, r
end

function cell_range(m::RegexMatch)::Tuple{Tuple{Int64, Int64}, Tuple{Int64, Int64}}
    l = letter2num(m[1])
    t = parse(Int64, m[2])
    r = letter2num(m[3])
    b = parse(Int64, m[4])

    return (t, b), (l, r)
end

function column_name_range(m::RegexMatch)::Tuple{Int64, Int64}
    l = letter2num(m[1])
    r = letter2num(m[3])

    return l, r
end

function table_bounds(sheet::Sheet, s::AbstractString)
    m = match(r"([A-Z]*)(\d*):([A-Z]*)(\d*)", s)

    if is_valid_cell_range(m)
        return cell_range(m)
    elseif is_valid_column_name_range(m)
        return (1, size(sheet.table, 1)), column_name_range(m)
    elseif is_valid_index_range(m)
        return (1, size(sheet.table, 1)), column_range(m)
    else
        error("KeyError: invalid range `$s`")
    end
end

function table_bounds(sheet::Sheet, r::UnitRange{Int64})
    if is_valid_index_range(r)
        return (1, size(sheet.table, 1)), (r.start, r.stop)
    else
        error("KeyError: invalid range `$r`")
    end
end

function construct_table(sheet::Sheet, top_bottom::Tuple, left_right::Tuple)
    t, b = top_bottom
    l, r = left_right

    result_table = Matrix{Any}(nothing, b - t + 1, r - l + 1)

    for col in l:r
        for row in t:b
            rel_row = row - t + 1
            rel_col = col - l + 1
            result_table[rel_row, rel_col] = get(sheet.table, (row, col), nothing)
        end
    end

    return result_table
end

function table_labels(l_r::Tuple{Int64, Int64}, column_labels::Nothing)::Tuple
    return Tuple(Symbol(num2letter(i)) for i in UnitRange(l_r...))
end

function table_labels(l_r::Tuple{Int64, Int64}, column_labels::Vector{String})::Tuple
    width = l_r[2] - l_r[1] + 1
    width == length(column_labels) || error(
        "column_labels=$column_labels length must be equal \
        to table columns length: $width"
    )
    return Tuple(Symbol(label) for label in column_labels)
end

"""
    xl_rowtable(sheet::Sheet, s::Union{UnitRange, AbstractString}; column_labels=nothing) -> RowTable
    xl_rowtable(sheet::Sheet; column_labels=nothing) -> RowTable

Get a [`RowTable`](@ref) from [`Sheet`](@ref) by cells or columns range.

```julia-repl
xl_book = parse_xlsx(xlsx_two_sheets_table())

xl_sheet = xl_sheets(xl_book, 2)

julia> xl_rowtable(xl_sheet, "A1:C3")
3-element Vector{NamedTuple}:
 (A = "Numbers", B = "Names", C = nothing)
 (A = 1.0, B = "a", C = nothing)
 (A = 2.0, B = "b", C = nothing)

julia> xl_rowtable(xl_sheet, "B:C"; column_labels=["c1", "c2"])
8-element Vector{NamedTuple}:
 (c1 = "Names", c2 = nothing)
 (c1 = "a", c2 = nothing)
 (c1 = "b", c2 = nothing)
 (c1 = "c", c2 = nothing)
 (c1 = "d", c2 = false)
 (c1 = "e", c2 = "D8")
 (c1 = nothing, c2 = "D8/0")
 (c1 = nothing, c2 = nothing)
```
"""
function xl_rowtable end

function xl_rowtable(
    sheet::Sheet, s::Union{UnitRange, AbstractString}; column_labels=nothing
)
    t_b, l_r = table_bounds(sheet, s)
    result_table = construct_table(sheet, t_b, l_r)
    labels = table_labels(l_r, column_labels)

    return RowTable{NamedTuple}(map(row -> NamedTuple{labels}(row), eachrow(result_table)))
end

function xl_rowtable(sheet::Sheet; column_labels=nothing)
    cols_n = size(sheet.table, 2)
    labels = table_labels((1, cols_n), column_labels)

    return RowTable{NamedTuple}(map(row -> NamedTuple{labels}(row), eachrow(sheet.table)))
end

"""
    xl_columntable(sheet::Sheet, s::Union{UnitRange, AbstractString}; column_labels=nothing) -> ColumnTable
    xl_columntable(sheet::Sheet; column_labels=nothing) -> ColumnTable

Get a [`ColumnTable`](@ref) from [`Sheet`](@ref) by cells or columns range.

```julia-repl
xl_book = parse_xlsx(xlsx_two_sheets_table())

xl_sheet = xl_sheets(xl_book, 2)

julia> xl_columntable(xl_sheet, "1:2")
(
    A = Any["Numbers", 1.0, 2.0, 3.0, 4.0, 5.0, nothing, nothing], 
    B = Any["Names", "a", "b", "c", "d", "e", nothing, nothing]
)

julia> xl_columntable(xl_sheet, 3:4; column_labels=["C_col", "D_col"])
(
    C_col = Any[nothing, nothing, nothing, nothing, false, "D8", "D8/0", nothing], 
    D_col = Any[nothing, nothing, nothing, nothing, nothing, nothing, nothing, 100.0]
)
```
"""
function xl_columntable end

function xl_columntable(
    sheet::Sheet, s::Union{UnitRange, AbstractString}; column_labels=nothing
)
    t_b, l_r = table_bounds(sheet, s)
    result_table = construct_table(sheet, t_b, l_r)
    labels = table_labels(l_r, column_labels)

    return NamedTuple{labels}(eachcol(result_table))::ColumnTable
end

function xl_columntable(sheet::Sheet; column_labels=nothing)
    cols_n = size(sheet.table, 2)
    labels = table_labels((1, cols_n), column_labels)

    return NamedTuple{labels}(eachcol(sheet.table))::ColumnTable
end