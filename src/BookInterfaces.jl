"""
    xl_sheetnames(xl_book::ExcelBook)

Getting a list of all sheet names in a book.
"""
function xl_sheetnames(xl_book::ExcelBook)
    return xl_book.sheet_names
end

"""
    xl_sheets(xl_book::ExcelBook)

Getting a list of all ExcelSheet in a book.
"""
function xl_sheets(xl_book::ExcelBook)
    return xl_book.sheets
end

"""
    xl_sheets(xl_book::ExcelBook, key::String)

Getting a ExcelSheet by name.
"""
function xl_sheets(xl_book::ExcelBook, key::String)
    key in xl_sheetnames(xl_book) || error("KeyError: no sheet name `$key`")

    index = findfirst(s -> s.name == key, xl_book.sheets)
    return xl_book.sheets[index]
end

"""
    xl_sheets(xl_book::ExcelBook, key::Int64)

Getting a ExcelSheet by index.
"""
function xl_sheets(xl_book::ExcelBook, key::Int64)
    key <= length(xl_book.sheets) || error("KeyError: no sheet index `$key`")

    return xl_book.sheets[key]
end

"""
    xl_sheets(xl_book::ExcelBook, keys_set::Vector{String})

Getting a list of ExcelSheet by names.
"""
function xl_sheets(xl_book::ExcelBook, keys_set::Vector{String})
    all(key -> key in xl_sheetnames(xl_book), keys_set) || 
        error("KeyError: invalid sheet names list `$keys_set`")

    indices = findall(s -> s.name in keys_set, xl_book.sheets)
    return xl_book.sheets[indices]
end

"""
    xl_sheets(xl_book::ExcelBook, keys_set::UnitRange{Int64})

Getting a list of ExcelSheet by names or range.
"""
function xl_sheets(xl_book::ExcelBook, keys_set::UnitRange{Int64})
    keys_set ⊆ 1:length(xl_book.sheets) || 
        error("KeyError: invalid sheet names range `$keys_set`")
    
    return xl_book.sheets[keys_set]
end