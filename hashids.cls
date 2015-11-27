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

'===================================================================
' Constants
'===================================================================
Private Const c_version = "1.0.0"
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
' Allow initialization with optional salt. min hash len and alphabet
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
Public Function Encode(ParamArray nums() As Variant)
    Dim ret As String
    Dim i As Integer, j As Integer, l As Integer
    Dim num As Variant
    
    i = 0
    For Each num In nums
        If num Mod 1 <> 0 Or num < 0 Then
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
Public Function EncodeHex(ByVal hexStr As String)
    Dim tmpHex As String
    Dim i As Integer
    Dim nums As Variant

    For i = 1 To Len(hexStr) Step 1
        c = Mid(hexStr, i, 1)
        If InStr(1, c_hex, Mid(hexStr, i, 1), vbBinaryCompare) < 1 Then
            EncodeHex = ""
            Exit Function
        End If
    Next
        
    i = 0
    While Len(hexStr) > 0
        ReDim Preserve nums(i)
        'VBA only handles 8 hex chars max, so we break it into two parts
        tmpHex = Left("00000000000" & Left(hexStr, 11), 11)
        nums(i) = (CLng("&H" & Left(tmpHex, 7)) * 65536) + CLng("&H" & Right(tmpHex, 4))
        hexStr = Mid(hexStr, 12)
        i = i + 1
    Wend
    
    If i = 0 Then
        ' empty strig
        EncodeHex = ""
    Else
        EncodeHex = P_Encode(nums)
    End If
    
End Function

'===================================================================
' obj.Decode(hashStr)
' Params:
'   hashStr (string)
'   alphaStr (string)
' Returns
'   Decoded number (if only 1 num) or array of numbers
'===================================================================
Public Function Decode(ByVal hashStr As String) As Variant
    
    If P_Nz(hashStr, "") = "" Then
        Decode = vbNull
    End If

    Decode = P_Decode(hashStr, Me.Alphabet)
    
End Function

'===================================================================
' obj.DecodeHex(hashStr)
' Returns:
'   hex String
'===================================================================
Public Function DecodeHex(ByVal hashStr As String) As Variant
    Dim nums As Variant
    Dim num As Long, a As Long, b As Long
    Dim hexStr As String
    
    If P_Nz(hashStr, "") = "" Then
        Decode = vbNull
    End If
    
    nums = P_Decode(hashStr, Me.Alphabet)
    hexStr = ""
    For Each num In nums
        'VBA only handles 8 hex chars max
        b = Hex(Int(num / 65536))
        a = Hex(num - (b * 65536))
        hexStr = hexStr & Mid(b, 2) & a
    Next num
    
    DecodeHex = hexStr

End Function

'===================================================================
' obj.Encrypt  Alias of Encode
'===================================================================
Public Function Encrypt(ParamArray nums() As Variant) As String
    Encrypt = Encode(nums)
End Function

'===================================================================
' obj.Decrypt  Alias of Decode
'===================================================================
Public Function Decrypt(hashStr As String) As Variant
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
    Dim half As Integer, excess As Integer, numSize As Integer, numHashInt As Integer
  
    alpha = Me.Alphabet
    numSize = UBound(nums) - LBound(nums) + 1
    numHashInt = 0
    
    i = 0
    For Each num In nums
        numHashInt = numHashInt + (num Mod (i + 100))
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
            num = num Mod (Asc(lst) + i)
            sepsIdx = num Mod Len(Me.Seps)
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
            ret = Mid(ret, Int(excess / 2) + 1, Me.MinHashLength)
        End If
    Wend
    
    P_Encode = ret
End Function

'===================================================================
' P_Decode
' Params:
'   hashStr (string)
'   alphaStr (string)
' Returns
'   Decoded number (if only 1 num) or array of numbers
'===================================================================
Private Function P_Decode( _
        ByVal hashStr As String, _
        ByVal alphaStr As String _
    ) As Variant

    Dim nums As Variant, hashArray As Variant
    Dim i As Integer, l As Integer
    Dim lottery As String, subHash As String, buffer As String
    Dim c As String

    For i = 1 To leng(Me.Guards)
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
        P_Decode = vbNull
        Exit Function
    End If
    
    lottery = Left(hashStr, 1)
    hashStr = Mid(hashStr, 2)
    
    For i = 1 To leng(Me.Seps)
        c = Mid(Me.Seps, i, 1)
        hashStr = Replace(hashStr, c, " ", 1, -1, vbBinaryCompare)
    Next
    hashArray = Split(hashStr, " ", -1, vbBinaryCompare)
    
    i = 0
    For Each subHash In hashArray
        ReDim Preserve nums(i)
        buffer = Left(lottery & Me.Salt & alphaStr, Len(alphaStr))
        alphaStr = P_ConsistentShuffle(alphaStr, buffer)
        nums(i) = P_UnHash(subHash, alphaStr)
        i = i + 1
    Next subHash

    If i = 1 Then
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
        ByVal num As Long, _
        ByVal alphaStr As String) As String
        
    Dim h As String
    Dim l As Integer, p As Integer
    h = ""
    l = Len(alphaStr)
    While num > 0
        p = num Mod l
        h = Mid(alphaStr, p + 1, 1) & h
        num = Int(num / l)
    Wend
    
    P_Hash = h
End Function

'===================================================================
' P_UnHash
'===================================================================
Private Function P_UnHash( _
        ByVal hashStr As String, _
        ByVal alphaStr As String) As Long
    Dim i As Integer, pos As intger, l As Integer
    Dim num As Long

    num = 0
    l = Len(hashStr)
    For i = 1 To l Step 1
        pos = InStr(1, alphaStr, Mid(hashStr, i, 1), vbBinaryCompare)
        num = num + ((pos - 1) * (Len(alphaStr) ^ (l - i - 2)))
    Next
    
    P_UnHash = num
    
End Function
        
'===================================================================
' P_ConsisentShuffle(alphabet, salt)
' Parameters:
'  alphaStr (string) Alphabet to Shuffle
'  salt (string) Salt to see teh shuffle
' Returns:
'  alphabet (string)
'===================================================================
Private Function P_ConsistentShuffle( _
        ByVal alphaStr As String, _
        Optional ByVal saltStr As String = "" _
    ) As String
        
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
    
    ' Seps must be from alphabet and alphabet cant contain seps
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
' Private Nz function
'  to compensate for lack of Nz() in non MS Access apps
'===================================================================
Private Function P_Nz( _
        data As Variant, _
        Optional nullValue As Variant = "" _
    ) As Variant

    If IsNull(data) Then
        P_Nz = nullValue
    Else
        P_Nz = data
    End If

End Function

'===================================================================
' Class Initialization
'===================================================================
Private Sub Class_Initialize()
    ' Initialize with defaults
    P_Init "", 0, c_alphabet
End Sub
