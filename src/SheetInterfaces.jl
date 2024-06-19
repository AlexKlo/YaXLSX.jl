# SheetInterfaces

function is_valid_index_range(r::UnitRange{Int64})
    return r.start <= r.stop
end

function is_valid_index_range(m::Union{Nothing, RegexMatch})
    isnothing(m) && return false
    all(isempty, m.captures[1:2:end]) && all(!isempty, m.captures[2:2:end]) || return false
    
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
    all(!isempty, m.captures[1:2:end]) && all(isempty, m.captures[2:2:end]) || return false
    
    l = letter2num(m[1])
    r = letter2num(m[3])

    l <= r || return false

    return true
end

function column_range(m::RegexMatch)::UnitRange{Int64}
    l = parse(Int64, m[2])
    r = parse(Int64, m[4])

    return l:r
end

function cell_range(m::RegexMatch)::Tuple{Tuple{Int64, Int64}, Tuple{Int64, Int64}}
    l = letter2num(m[1])
    t = parse(Int64, m[2])
    r = letter2num(m[3])
    b = parse(Int64, m[4])

    return (t, b), (l, r)
end

function column_name_range(m::RegexMatch)::UnitRange{Int64}
    l = letter2num(m[1])
    r = letter2num(m[3])

    return l:r
end

function construct_table(sheet::Sheet, s::AbstractString)
    m = match(r"([A-Z]*)(\d*):([A-Z]*)(\d*)", s)

    if is_valid_cell_range(m)
        return construct_table(sheet.table, cell_range(m)...)
    elseif is_valid_column_name_range(m)
        return construct_table(sheet, column_name_range(m))
    elseif is_valid_index_range(m)
        return construct_table(sheet, column_range(m))
    else
        error("KeyError: invalid range `$s`")
    end
end

function construct_table(sheet::Sheet, r::UnitRange{Int64})
    if is_valid_index_range(r)
        return construct_table(sheet.table, (1, size(sheet.table, 1)), (r.start, r.stop))
    else
        error("KeyError: invalid range `$r`")
    end
end

function construct_table(table::Matrix, top_bottom::Tuple, left_right::Tuple)
    t, b = top_bottom
    l, r = left_right

    result_table = Matrix{Any}(nothing, b - t + 1, r - l + 1)

    for col in l:r
        for row in t:b
            rel_row = row - t + 1
            rel_col = col - l + 1
            result_table[rel_row, rel_col] = get(table, (row, col), nothing)
        end
    end

    names = Tuple(Symbol(num2letter(i)) for i in l:r)

    return NamedTuple{names}(eachcol(result_table))
end

function construct_table(sheet::Sheet)
    rows_n, cols_n = size(sheet.table)
    return construct_table(sheet.table, (1, rows_n), (1, cols_n))
end