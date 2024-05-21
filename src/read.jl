const ZIP_FILE_HEADER = [ 0x50, 0x4b, 0x03, 0x04 ] # valid Zip file header

function _check_file_format(byte_array::Vector{UInt8})
    header = byte_array[1:4]

    if header == ZIP_FILE_HEADER 
        return nothing
    else
        error("WrongExtension: ")
    end

    return nothing
end

function parse_xlsx(byte_array::Vector{UInt8})
    _check_file_format(byte_array)
    io = IOBuffer(byte_array)
    return io
end