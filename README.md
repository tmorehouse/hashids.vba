# hashids.vba
Hashids, ported for VBA (http://www.hashids.org)

# NAME

Hashids - generate short hashes from numbers

# SYNOPSIS

    Sub testHashids()
        Dim hid As Hashids
        Dim hash As String, hexStr As String
        Dim number As Long, numbers As Variant, i As Long
    
        Set hid = New Hashids
        hid.Params("this is my salt")
    
        ' encode a single number
        hash = hid.Encode(123)          
        Debug.Print "Encode 123:", hash   ' "YDx"
    
        number = hid.Decode("YDx")      
        Debug.Print "Decode YDx:", number ' 123
    
        ' or a list
        hash = hid.Encode(1, 2, 3)        ' "laHquq"
        Debug.Print "Encode list 1,2,3:", hash
    
        numbers = hid.Decode("laHquq")    ' (1, 2, 3)
        Debug.Print "Decode laHquq:"
        For i = LBound(numbers) To UBound(numbers)
            Debug.Print numbers(i)
        Next i
    
        ' or an Array
        hash = hid.Encode(Array(1,2,3))   ' "laHquq"
        Debug.Print "Encode list 1,2,3:", hash

        numbers = hid.Decode("laHquq")    ' (1, 2, 3)
        Debug.Print "Decode laHquq:"
        For i = LBound(numbers) To UBound(numbers)
            Debug.Print numbers(i)
        Next i
        
        ' or a hex string
        hash = hid.EncodeHex("deadbeef")   ' "kRNrpKlJ"
        Debug.Print "EncodeHex deadbeef:", hash

        hexStr = hid.DecodeHex("kRNrpKlJ") ' "deadbeef"
        Debug.Print "Decode kRNrpKlJ:", hexStr
        
    End Sub

# DESCRIPTION

This is a port of the Hashids JavaScript library for VBA.

Hashids was designed for use in URL shortening, tracking stuff,
validating accounts or making pages private (through abstraction.)
Instead of showing items as `1`, `2`, or `3`, you could show them as
`b9iLXiAa`, `EATedTBy`, and `Aaco9cy5`.  Hashes depend on your salt
value.

**IMPORTANT**: This implementation follows the v1.0.0 API release of
hashids.js.

This implementation is also compatible with the v0.3.x hashids.js API.

# METHODS

- `set hid = New Hashids`

    Make a new Hashids object.  This constructor does not accept any options

- `hid.Params(salt,minHashLength,alphabet)`

    - `salt = "this is my salt"`

        Salt string, this should be unique per Hashid object. Defaults to ""

    - `minHashLength = 5`

        Minimum hash length.  Use this to control how long the generated hash
        string should be. Defaults to 0

    - `alphabet = 'abcdefghijklmnop'`

        Alphabet set to use.  This is optional as Hashids comes with a default
        set suitable for URL shortening.  Should you choose to supply a custom
        alphabet, make sure that it is at least 16 characters long, has no
        spaces, and only has unique characters.

- `hash = hid.Encode(x, [y, z, ...])`

    Encode a single number (or a list/array of numbers) into a hash
    string. If encoding an array of numbers, the array must be the only 
    parameter passed to this function.

    _hid.Encrypt()_ is an alias for this method, for compatibility with v0.3.x
    hashids.js API.

- `hash = hid.EncodeHex("deadbeef")`

    Encode a hex string into a hash string.

- `number = hid.Decode(hash)`

    Decode a hash string into its number (or numbers.)  Returns either a
    a single number, an array of numbers if it decrypted a set, or Null if 
    given bad input. You should use a variable of type Variant for receiving the 
    return value(s).

    _decrypt()_ is an alias for this method, for compatibility with v0.3.x
    hashids.js API.

- `hexStr = hid.DecodeHex(hash)`

    Opposite of _hid.EncodeHex()_.  Unlike _hid.Decode()_, this will always
    return a string, including the empty string if the hash is invalid.

# SEE ALSO

[Hashids](http://www.hashids.org)

# CHANGE LOG

***1.0.2***
- Now supports numbers up to the value of `9,007,199,254,740,991`
- Improved Hex encoding and decoding

***1.0.1***
- Implemented decoding routines
- Allows array of numbers to be passed to `Encode`

***1.0.0***
- Initial release
- Limited to max number value of `2,147,483,647` (signed 32bit integer)

# LICENSE

MIT License. See the `LICENSE` file. You can use Hashids in open source projects and commercial products.

Batteries not included, some assembly required, not 
suitable for children under the age of 6, slippery when wet, bridge freezes
before road, do not use while consuming alcohol, use at your own risk.
