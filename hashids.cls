VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hashids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Make sure arrays start at 0
Option Base 0
'
' Range of numbers supported: 0 to 9007199254740991
' Decoded numbers are returned a Decimal (stored in variant)
'
'===================================================================
' Constants
'===================================================================
Private Const c_version = "1.0.2"
Private Const c_sepDiv As Double = 3.5
Private Const c_guardDiv As Double = 12#
Private Const c_minAlphaLen As Integer = 16
Private Const c_minHashLength As Integer = 0
Private Const c_seps As String = "cfhistuCFHISTU"
Private Const c_alphabet As String = _
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
Private Const c_hex As String = "0123456789abcdefABCDEF"

'===================================================================
' Private variables
'===================================================================
Private p_alphabet As String
Private p_salt As String
Private p_seps As String
Private p_guards As String
Private p_minHashLength As Integer

'===================================================================
' Error Codes
'===================================================================
Public Enum HashidsErr
    hidAlphaShort = vbObjectError + 1000
    hidAlphaSpace = vbObjectError + 1001
    hidNonInteger = vbObjectError + 1002
End Enum

'===================================================================
' Properties
'===================================================================

'-------------------------------------------------------------------
' Our Version
Public Property Get Version() As String
    Version = c_version
End Property

'-------------------------------------------------------------------
' MinHashLength
Public Property Get MinHashLength() As Integer
    MinHashLength = p_minHashLength
End Property

Private Property Let MinHashLength(ByVal l As Integer)
    p_minHashLength = l
End Property

'-------------------------------------------------------------------
' Salt property
Public Property Get Salt() As String
    Salt = p_salt
End Property

Private Property Let Salt(ByVal s As String)
    p_salt = P_Nz(s, "")
End Property

'-------------------------------------------------------------------
' Seps property
Public Property Get Seps() As String
    Seps = p_seps
End Property

Private Property Let Seps(ByVal s As String)
    p_seps = s
End Property

'-------------------------------------------------------------------
' Guards Property
Public Property Get Guards() As String
    Guards = p_guards
End Property

Private Property Let Guards(ByVal g As String)
    p_guards = g
End Property

'-------------------------------------------------------------------
' Alphabet Property
Public Property Get Alphabet() As String
    Alphabet = p_alphabet
End Property

Private Property Let Alphabet(ByVal alpha As String)
    p_alphabet = alpha
End Property

'===================================================================
' Allow initialization with optional salt, minhashlen and alphabet
'===================================================================
Public Sub Params( _
        Optional saltStr As Variant = "", _
        Optional minHashLen As Variant = 0, _
        Optional alphabetStr As Variant = c_alphabet _
    )
    ' Now compute the guards and seps
    P_Init P_Nz(saltStr, ""), P_Nz(minHashLen, 0), P_Nz(alphabetStr, c_alphabet)
End Sub


'===================================================================
' obj.Encode(nums)
'  Parameters:
'   nums (variant) single number or list of numbers, or a
'   variant array of numbers
' Returns
'  hashid (string)
'===================================================================
Public Function Encode(ParamArray vars() As Variant) As String
    Dim ret As String
    Dim i As Integer, j As Integer, l As Integer
    Dim num As Variant, nums As Variant, tmpNum As Variant
    
    ' Handle passing of an array
    If IsArray(vars(0)) Then
        nums = vars(0)
    Else
        nums = vars
    End If
    
    i = 0
    For Each num In nums
        tmpNum = CDec(num)
        If tmpNum <> Int(tmpNum / CDec(1)) Or tmpNum < CDec(0) Then
            Err.Raise hidNonInteger, "Hashids", "Non positive/integer value in Encode list"
        End If
        i = i + 1
    Next num
    
    If i = 0 Then
        ' empty array
        Encode = ""
    Else
        Encode = P_Encode(nums)
    End If
    
End Function

'===================================================================
' obj.EncodeHex(nums)
'  Parameters:
'   hex value as a string
' Returns
'  hashid (string)
'===================================================================
Public Function EncodeHex(ByVal hexStr As Variant) As String
    Dim i As Integer, nums() As Variant, tmp As String

    hexStr = CStr(hexStr)
    
    ' Ensure only hex characters
    For i = 1 To Len(hexStr) Step 1
        If InStr(1, c_hex, Mid(hexStr, i, 1), vbBinaryCompare) < 1 Then
            EncodeHex = ""
            Exit Function
        End If
    Next
        
    ' Break the input string into groups of 12 hex chars max
    ' and convert to decimal
    i = 0
    While Len(hexStr) > 0
        ReDim Preserve nums(i)
        nums(i) = P_Hex2Dec("1" & Left(CStr(hexStr), 12))
        hexStr = Mid(hexStr, 13)
        i = i + 1
    Wend
    
    ' Encode the numbers
    If i = 0 Then
        ' empty strig
        EncodeHex = ""
    Else
        EncodeHex = P_Encode(nums)
    End If
    
End Function

'===================================================================
' obj.Decode(hashStr)
' Returns:
'   Decimal(variant) if single number, or arary if multiple
'===================================================================
Public Function Decode( _
        ByVal hashStr As String, _
        Optional forceArray As Boolean = False _
    ) As Variant
    
    If P_Nz(hashStr, "") = "" Then
        Decode = Null
        Exit Function
    End If
    
    Decode = P_Decode(hashStr, Me.Alphabet, forceArray)

End Function

'===================================================================
' obj.DecodeHex(hashStr)
' Returns:
'   hex String
'===================================================================
Public Function DecodeHex( _
        ByVal hashStr As String _
    ) As Variant
    
    Dim nums As Variant, num As Variant, a As Long, b As Long
    Dim hexStr As String
    
    If P_Nz(hashStr, "") = "" Then
        DecodeHex = Null
        Exit Function
    End If
    
    nums = P_Decode(hashStr, Me.Alphabet, True)
    hexStr = ""
    For Each num In nums
         hexStr = hexStr & Mid(P_Dec2Hex(num),2)
    Next num
    
    DecodeHex = hexStr

End Function

'===================================================================
' obj.Encrypt  Alias of Encode
'===================================================================
Public Function Encrypt( _
        ParamArray nums() As Variant _
    ) As String
    
    Encrypt = Encode(nums)
    
End Function

'===================================================================
' obj.Decrypt  Alias of Decode
'===================================================================
Public Function Decrypt( _
        hashStr As String, _
        Optional forceArray As Boolean = False _
    ) As Variant
    
    Decrypt = Decode(hashStr)
    
End Function

'===================================================================
' p_Encode(nums) - Private encode
'  Parameters:
'   num (long) array of numbers to hash
' Returns
'  hash (string) hash string
'===================================================================
Private Function P_Encode(ByVal nums As Variant) As String

    Dim lottery As String, buffer As String, alpha As String
    Dim ret As String, guard As String, lst As String
    Dim num As Variant
    Dim i As Integer, sepsIdx As Integer, guardIdx As Integer
    Dim half As Integer, excess As Integer, numSize As Integer
    Dim numHashInt As Integer
  
    alpha = Me.Alphabet
    numSize = UBound(nums) - LBound(nums) + 1
    numHashInt = 0
    
    i = 0
    For Each num In nums
        numHashInt = numHashInt + CInt(P_Mod(num, i + 100))
        i = i + 1
    Next num
    
    lottery = Mid(alpha, (numHashInt Mod Len(alpha)) + 1, 1)
    ret = lottery
    
    i = 0
    For Each num In nums
        buffer = lottery & Me.Salt & alpha
        alpha = P_ConsistentShuffle(alpha, Mid(buffer, 1, Len(alpha)))
        lst = P_Hash(num, alpha)
        ret = ret & lst
        If (i + 1) < numSize Then
            num = P_Mod(num, Asc(lst) + i)
            sepsIdx = CInt(P_Mod(num, Len(Me.Seps)))
            ret = ret & Mid(Me.Seps, sepsIdx + 1, 1)
        End If
        i = i + 1
    Next num
    
    If Len(ret) < Me.MinHashLength Then
        guardIdx = (numHashInt + Asc(Mid(ret, 1, 1))) Mod Len(Me.Guards)
        guard = Mid(Me.Guards, guardIdx + 1, 1)
        ret = guard & ret
        If Len(ret) < Me.MinHashLength Then
            guardIdx = (numHashInt + Asc(Mid(ret, 3, 1))) Mod Len(Me.Guards)
            guard = Mid(Me.Guards, guardIdx + 1, 1)
            ret = ret & guard
        End If
    End If
        
    half = Int(Len(alpha) / 2)
    While Len(ret) < Me.MinHashLength
        alpha = P_ConsistentShuffle(alpha, alpha)
        ret = Mid(alpha, half + 1) & ret & Mid(alpha, 1, half)
        excess = Len(ret) - Me.MinHashLength
        If excess > 0 Then
            ret = Mid(ret, (excess \ 2) + 1, Me.MinHashLength)
        End If
    Wend
    
    P_Encode = ret
End Function

'===================================================================
' P_Decode(hash, alpha) - Private Decode
'  Parameters:
'   hash (string) hashid to decode
'   alpha (string) alphabet to use for decoding
' Returns
'  single number or array of numbers
'===================================================================
Private Function P_Decode( _
        ByVal hashStr As String, _
        ByVal alphaStr As String, _
        Optional forceArray As Boolean = False _
    ) As Variant
    
    Dim nums() As Variant, hashArray As Variant
    Dim i As Integer, l As Integer
    Dim lottery As String, subHash As Variant, buffer As String
    Dim c As String

    ' Strip out guards
    For i = 1 To Len(Me.Guards)
        c = Mid(Me.Guards, i, 1)
        hashStr = Replace(hashStr, c, " ", 1, -1, vbBinaryCompare)
    Next
    hashArray = Split(hashStr, " ", -1, vbBinaryCompare)
    i = 0
    l = UBound(hashArray) - LBound(hashArray) + 1
    If l = 3 Or l = 2 Then
        i = 1
    End If
    hashStr = P_Nz(hashArray(LBound(hashArray) + i), "")
    If hashStr = "" Then
        P_Decode = Null
        Exit Function
    End If
    
    lottery = Left(hashStr, 1)
    hashStr = Mid(hashStr, 2)
    
    ' Break into array of number hashes
    For i = 1 To Len(Me.Seps)
        c = Mid(Me.Seps, i, 1)
        hashStr = Replace(hashStr, c, " ", 1, -1, vbBinaryCompare)
    Next
    hashArray = Split(hashStr, " ", -1, vbBinaryCompare)
    
    ' Decode each number in the hash array
    i = 0
    For Each subHash In hashArray
        ReDim Preserve nums(i)
        buffer = Left(lottery & Me.Salt & alphaStr, Len(alphaStr))
        alphaStr = P_ConsistentShuffle(alphaStr, buffer)
        nums(i) = P_UnHash(subHash, alphaStr)
        i = i + 1
    Next subHash

    If i = 1 And Not forceArray Then
        P_Decode = nums(0)
    Else
        P_Decode = nums
    End If

End Function
        

'===================================================================
' P_Hash(num)
'  Parameters:
'   num (long) number to hash
' Returns
'  hash (string) hash string
'===================================================================
Private Function P_Hash( _
        ByVal num As Variant, _
        ByVal alphaStr As String _
    ) As String
        
    Dim hash As String
    Dim la As Integer, pos As Integer
    
    num = CDec(num)
    hash = ""
    la = CDec(Len(alphaStr))
    
    Do
        pos = CInt(P_Mod(num, la))
        hash = Mid(alphaStr, pos + 1, 1) & hash
        num = CDec(Int(num / la))
    Loop While CDec(num) > CDec(0)
    
    P_Hash = hash
    
End Function

'===================================================================
' P_UnHash(hash, alpha)
'  Parameters:
'   hash (string) single value hash string
'   alpha (string) alphabet to use for decoding
' Returns
'  number
'===================================================================
Private Function P_UnHash( _
        ByVal hashStr As String, _
        ByVal alphaStr As String _
    ) As Variant
    
    Dim i As Integer, pos As Integer, lh As Integer, la As Integer
    Dim num As Variant

    num = CDec(0)
    lh = CDec(Len(hashStr))
    la = CDec(Len(alphaStr))
    
    For i = 1 To lh Step 1
        pos = CDec(InStr(1, alphaStr, Mid(hashStr, i, 1), vbBinaryCompare) - 1)
        num = num + (CDec(pos) * (la ^ (lh - CDec(i))))
    Next
    
    P_UnHash = num
    
End Function
        
'===================================================================
' consisentShuffle(alphabet, salt)
' Parameters:
'  alphaStr (string) Alphabet to Shuffle
'  salt (string) Salt to see teh shuffle
' Returns:
'  alphabet (string)
'===================================================================
Private Function P_ConsistentShuffle( _
        ByVal alphaStr As String, _
        Optional ByVal saltStr As String = "") As String
        
    Dim tmpa As String, tmpb As String
    Dim n As Integer, i As Integer, j As Integer, v As Integer, p As Integer
    
    If Len(saltStr) = 0 Then
        P_ConsistentShuffle = alphaStr
    Else
        i = Len(alphaStr) - 1
        While i > 0
            v = v Mod Len(saltStr)
            n = Asc(Mid(saltStr, v + 1, 1))
            p = p + n
            j = (n + v + p) Mod i
            tmpa = Mid(alphaStr, i + 1, 1)
            tmpb = Mid(alphaStr, j + 1, 1)
            alphaStr = Mid(alphaStr, 1, i) & tmpb & Mid(alphaStr, i + 2)
            alphaStr = Mid(alphaStr, 1, j) & tmpa & Mid(alphaStr, j + 2)
            i = i - 1
            v = v + 1
        Wend
        P_ConsistentShuffle = alphaStr
    End If
    
End Function

'===================================================================
' Private P_Ceil Function (VBA has none)
'===================================================================
Private Function P_Ceil(ByVal num As Double) As Double
    If Abs(num) = Int(Abs(num)) Then
        P_Ceil = num
    Else
        P_Ceil = CDbl(Sgn(num) * (Int(Abs(num) + 1)))
    End If
End Function

'===================================================================
' Private Nz function
'  to compensate for lackof Nz() in non MS Access apps
'===================================================================
Private Function P_Nz( _
        ByVal data As Variant, _
        Optional nullValue As Variant = "") As Variant

    If IsNull(data) Or IsEmpty(data) Then
        P_Nz = nullValue
    Else
        P_Nz = data
    End If

End Function

'===================================================================
' Private MOD function
'  to compensate for VBAs MOD function not working with large values
'===================================================================
Private Function P_Mod( _
        ByVal number As Variant, _
        ByVal divisor As Variant _
    ) As Variant

    number = CDec(number)
    divisor = CDec(divisor)
    
    P_Mod = number - (divisor * Int(number / divisor))

End Function

'===================================================================
' Private Dec2Hex function
'  to compensate for VBAs limit of 8 characters max
'===================================================================
Private Function P_Dec2Hex( _
        ByVal number As Variant _
    ) As String

    Dim hexStr As String
    
    number = CDec(number)
    hexStr = ""
    
    While number > CDec(0)
        hexStr = Hex(P_Mod(number, 16)) & hexStr
        number = CDec(Int(number / CDec(16)))
    Wend

    P_Dec2Hex = LCase(hexStr)
    
End Function

'===================================================================
' Private HEX2DEC function
'  to compensate for VBAs limited hex handling
'===================================================================
Private Function P_Hex2Dec( _
        ByVal hexStr As Variant _
    ) As Variant
    
    Dim number As Variant

    number = CDec(0)
    hexStr = CStr(hexStr)
    
    While Len(hexStr) > 0
        number = (number * CDec(16)) + CDec("&h" & Left(hexStr, 1))
        hexStr = Mid(hexStr, 2)
    Wend

    P_Hex2Dec = CDec(number)
    
End Function


'===================================================================
' Private Initialization routine to compute seps, guards, alphabet
'===================================================================
Private Sub P_Init(saltStr As String, minHashLen As Integer, alphaStr As String)
    Dim uniqueStr As String, c As String, g As String
    Dim i As Integer, l As Integer, d As Integer

    Seps = c_seps
    Guards = ""
    Salt = saltStr
    MinHashLength = minHashLen
    Alphabet = alphaStr
    
    ' Ensure alphabet only has uniqueStr characters
    uniqueStr = ""
    For i = 1 To Len(Me.Alphabet)
        If InStr(1, uniqueStr, Mid(Alphabet, i, 1), vbBinaryCompare) = 0 Then
            uniqueStr = uniqueStr & Mid(Alphabet, i, 1)
        End If
    Next
    Alphabet = uniqueStr
    
    If InStr(1, Alphabet, " ", vbBinaryCompare) Then
        ' Alphabet cannot contain spaces
        Err.Raise hidAlphaSpace, "Hashids", "Alphabet cannot contain spaces"
    ElseIf Len(Alphabet) < 16 Then
        ' Alphabet must be at least 16 characters
        Err.Raise hidAlphaShort, "Hashids", "Alphabet must contain at least " & c_minAlphaLen & " unique characters"
    End If
    
    ' Seps must be from alphabet and alphabet cannot contain seps
    Seps = ""
    For i = 1 To Len(c_seps) Step 1
        c = Mid(c_seps, i, 1)
        If InStr(1, Alphabet, c, vbBinaryCompare) > 0 Then
            Seps = Seps & c
            Alphabet = Replace(Alphabet, c, "", 1, -1, vbBinaryCompare)
        End If
    Next
    Seps = P_ConsistentShuffle(Seps, Salt)

    If Len(Seps) = 0 Or (Len(Alphabet) / Len(Seps)) > c_sepDiv Then
        l = P_Ceil(Len(Alphabet) / c_sepDiv)
        If l = 1 Then
            l = 2
        End If
        If l > Len(Me.Seps) Then
            d = l - Len(Me.Seps)
            Seps = Seps & Mid(Alphabet, 1, d)
            Alphabet = Mid(Alphabet, d + 1)
        Else
            Seps = Mid(Seps, 1, l)
        End If
    End If
    
    Alphabet = P_ConsistentShuffle(Alphabet, Salt)
    g = P_Ceil(Len(Alphabet) / c_guardDiv)
    If Len(Alphabet) < 3 Then
        Guards = Mid(Seps, 1, g)
        Seps = Mid(Seps, g + 1)
    Else
        Guards = Mid(Alphabet, 1, g)
        Alphabet = Mid(Alphabet, g + 1)
    End If
End Sub

'===================================================================
' Class Initialization
'===================================================================
Private Sub Class_Initialize()
    ' Initialize with defaults
    P_Init "", 0, c_alphabet
End Sub
