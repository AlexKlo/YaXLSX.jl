function xl_sheetnames(xl_book::ExcelBook)
    return xl_book.sheet_names
end

function xl_sheets(xl_book::ExcelBook)
    return xl_book.sheets
end

function xl_sheets(xl_book::ExcelBook, key::Union{String, Int64})
    index = if key isa String 
        key in xl_sheetnames(xl_book) || error("KeyError: no sheet name `$key`")

        findfirst(s -> s.name == key, xl_book.sheets)
    else
        key <= length(xl_book.sheets) || error("KeyError: no sheet index `$key`")

        key
    end

    return xl_book.sheets[index]
end

function xl_sheets(xl_book::ExcelBook, keys_set::Union{Vector{String}, UnitRange{Int64}})
    indices = if keys_set isa Vector{String} 
        all(key -> key in xl_sheetnames(xl_book), keys_set) ||
            error("KeyError: invalid sheet names list `$keys_set`")

        findall(s -> s.name in keys_set, xl_book.sheets)
    else
        keys_set ⊆ 1:length(xl_book.sheets) || 
            error("KeyError: invalid sheet names range `$keys_set`")

        keys_set
    end
    
    return xl_book.sheets[indices]
end