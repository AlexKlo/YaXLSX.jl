# Utils

@inline function letter2num(letter::AbstractString)::Int64
    id = 0
    n = length(letter)
    for i in eachindex(letter)
        j = Int64(letter[n-i+1]) - 65 + 1
        id = id + j + 26
    end
    return id - 26 - n + 1
end

@inline function num2letter(n::Int)::String
    n, r1 = divrem(n - 1, 26)
    iszero(n) && return string(Char(r1 + 65))

    n, r2 = divrem(n - 1, 26)
    iszero(n) && return string(Char(r2 + 65), Char(r1 + 65))

    _, r3 = divrem(n - 1, 26)
    return string(Char(r3 + 65), Char(r2 + 65), Char(r1 + 65))
end