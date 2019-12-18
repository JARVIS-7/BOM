VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRC4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic RC4 Implementation
' By: David Midkiff (mznull@earthlink.net)
'
' Standard RC4 implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. RC4 is an extremely fast
' and very secure stream cipher from RSA Data Security, Inc. I recommend
' this for high risk low resource environments. It's speed is very very
' attractive. Patents no longer apply for commercial use.
'
' Information on the algorithm can be found at:
' http://www.rsasecurity.com/rsalabs/faq/3-6-3.html
'
'
Private Const mbytEQUALSIGN As Byte = 61
Private Const mbytMASK1 As Byte = 3
Private Const mbytMASK2 As Byte = 15
Private Const mbytMASK3 As Byte = 63
'Private Const k_bytMask4 As Byte = 192
'Private Const k_bytMask5 As Byte = 240
'Private Const k_bytMask6 As Byte = 252
Private Const mbytSHIFT2 As Byte = 4
Private Const mbytSHIFT4 As Byte = 16
Private Const mbytSHIFT6 As Byte = 64
Private Const mlMAXBYTESPERLINE As Long = 152&
'...
Private msKey As String
Private miarrBox() As Integer
Private mbytarrIndex() As Byte
Private mbytarrReverseIndex() As Byte
'
Event Progress(lPercent As Long)
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Initialize64()
    mbytarrIndex(0) = 65 'Asc("A")
    mbytarrIndex(1) = 66 'Asc("B")
    mbytarrIndex(2) = 67 'Asc("C")
    mbytarrIndex(3) = 68 'Asc("D")
    mbytarrIndex(4) = 69 'Asc("E")
    mbytarrIndex(5) = 70 'Asc("F")
    mbytarrIndex(6) = 71 'Asc("G")
    mbytarrIndex(7) = 72 'Asc("H")
    mbytarrIndex(8) = 73 'Asc("I")
    mbytarrIndex(9) = 74 'Asc("J")
    mbytarrIndex(10) = 75 'Asc("K")
    mbytarrIndex(11) = 76 'Asc("L")
    mbytarrIndex(12) = 77 'Asc("M")
    mbytarrIndex(13) = 78 'Asc("N")
    mbytarrIndex(14) = 79 'Asc("O")
    mbytarrIndex(15) = 80 'Asc("P")
    mbytarrIndex(16) = 81 'Asc("Q")
    mbytarrIndex(17) = 82 'Asc("R")
    mbytarrIndex(18) = 83 'Asc("S")
    mbytarrIndex(19) = 84 'Asc("T")
    mbytarrIndex(20) = 85 'Asc("U")
    mbytarrIndex(21) = 86 'Asc("V")
    mbytarrIndex(22) = 87 'Asc("W")
    mbytarrIndex(23) = 88 'Asc("X")
    mbytarrIndex(24) = 89 'Asc("Y")
    mbytarrIndex(25) = 90 'Asc("Z")
    mbytarrIndex(26) = 97 'Asc("a")
    mbytarrIndex(27) = 98 'Asc("b")
    mbytarrIndex(28) = 99 'Asc("c")
    mbytarrIndex(29) = 100 'Asc("d")
    mbytarrIndex(30) = 101 'Asc("e")
    mbytarrIndex(31) = 102 'Asc("f")
    mbytarrIndex(32) = 103 'Asc("g")
    mbytarrIndex(33) = 104 'Asc("h")
    mbytarrIndex(34) = 105 'Asc("i")
    mbytarrIndex(35) = 106 'Asc("j")
    mbytarrIndex(36) = 107 'Asc("k")
    mbytarrIndex(37) = 108 'Asc("l")
    mbytarrIndex(38) = 109 'Asc("m")
    mbytarrIndex(39) = 110 'Asc("n")
    mbytarrIndex(40) = 111 'Asc("o")
    mbytarrIndex(41) = 112 'Asc("p")
    mbytarrIndex(42) = 113 'Asc("q")
    mbytarrIndex(43) = 114 'Asc("r")
    mbytarrIndex(44) = 115 'Asc("s")
    mbytarrIndex(45) = 116 'Asc("t")
    mbytarrIndex(46) = 117 'Asc("u")
    mbytarrIndex(47) = 118 'Asc("v")
    mbytarrIndex(48) = 119 'Asc("w")
    mbytarrIndex(49) = 120 'Asc("x")
    mbytarrIndex(50) = 121 'Asc("y")
    mbytarrIndex(51) = 122 'Asc("z")
    mbytarrIndex(52) = 48 'Asc("0")
    mbytarrIndex(53) = 49 'Asc("1")
    mbytarrIndex(54) = 50 'Asc("2")
    mbytarrIndex(55) = 51 'Asc("3")
    mbytarrIndex(56) = 52 'Asc("4")
    mbytarrIndex(57) = 53 'Asc("5")
    mbytarrIndex(58) = 54 'Asc("6")
    mbytarrIndex(59) = 55 'Asc("7")
    mbytarrIndex(60) = 56 'Asc("8")
    mbytarrIndex(61) = 57 'Asc("9")
    mbytarrIndex(62) = 43 'Asc("+")
    mbytarrIndex(63) = 47 'Asc("/")
    mbytarrReverseIndex(65) = 0 'Asc("A")
    mbytarrReverseIndex(66) = 1 'Asc("B")
    mbytarrReverseIndex(67) = 2 'Asc("C")
    mbytarrReverseIndex(68) = 3 'Asc("D")
    mbytarrReverseIndex(69) = 4 'Asc("E")
    mbytarrReverseIndex(70) = 5 'Asc("F")
    mbytarrReverseIndex(71) = 6 'Asc("G")
    mbytarrReverseIndex(72) = 7 'Asc("H")
    mbytarrReverseIndex(73) = 8 'Asc("I")
    mbytarrReverseIndex(74) = 9 'Asc("J")
    mbytarrReverseIndex(75) = 10 'Asc("K")
    mbytarrReverseIndex(76) = 11 'Asc("L")
    mbytarrReverseIndex(77) = 12 'Asc("M")
    mbytarrReverseIndex(78) = 13 'Asc("N")
    mbytarrReverseIndex(79) = 14 'Asc("O")
    mbytarrReverseIndex(80) = 15 'Asc("P")
    mbytarrReverseIndex(81) = 16 'Asc("Q")
    mbytarrReverseIndex(82) = 17 'Asc("R")
    mbytarrReverseIndex(83) = 18 'Asc("S")
    mbytarrReverseIndex(84) = 19 'Asc("T")
    mbytarrReverseIndex(85) = 20 'Asc("U")
    mbytarrReverseIndex(86) = 21 'Asc("V")
    mbytarrReverseIndex(87) = 22 'Asc("W")
    mbytarrReverseIndex(88) = 23 'Asc("X")
    mbytarrReverseIndex(89) = 24 'Asc("Y")
    mbytarrReverseIndex(90) = 25 'Asc("Z")
    mbytarrReverseIndex(97) = 26 'Asc("a")
    mbytarrReverseIndex(98) = 27 'Asc("b")
    mbytarrReverseIndex(99) = 28 'Asc("c")
    mbytarrReverseIndex(100) = 29 'Asc("d")
    mbytarrReverseIndex(101) = 30 'Asc("e")
    mbytarrReverseIndex(102) = 31 'Asc("f")
    mbytarrReverseIndex(103) = 32 'Asc("g")
    mbytarrReverseIndex(104) = 33 'Asc("h")
    mbytarrReverseIndex(105) = 34 'Asc("i")
    mbytarrReverseIndex(106) = 35 'Asc("j")
    mbytarrReverseIndex(107) = 36 'Asc("k")
    mbytarrReverseIndex(108) = 37 'Asc("l")
    mbytarrReverseIndex(109) = 38 'Asc("m")
    mbytarrReverseIndex(110) = 39 'Asc("n")
    mbytarrReverseIndex(111) = 40 'Asc("o")
    mbytarrReverseIndex(112) = 41 'Asc("p")
    mbytarrReverseIndex(113) = 42 'Asc("q")
    mbytarrReverseIndex(114) = 43 'Asc("r")
    mbytarrReverseIndex(115) = 44 'Asc("s")
    mbytarrReverseIndex(116) = 45 'Asc("t")
    mbytarrReverseIndex(117) = 46 'Asc("u")
    mbytarrReverseIndex(118) = 47 'Asc("v")
    mbytarrReverseIndex(119) = 48 'Asc("w")
    mbytarrReverseIndex(120) = 49 'Asc("x")
    mbytarrReverseIndex(121) = 50 'Asc("y")
    mbytarrReverseIndex(122) = 51 'Asc("z")
    mbytarrReverseIndex(48) = 52 'Asc("0")
    mbytarrReverseIndex(49) = 53 'Asc("1")
    mbytarrReverseIndex(50) = 54 'Asc("2")
    mbytarrReverseIndex(51) = 55 'Asc("3")
    mbytarrReverseIndex(52) = 56 'Asc("4")
    mbytarrReverseIndex(53) = 57 'Asc("5")
    mbytarrReverseIndex(54) = 58 'Asc("6")
    mbytarrReverseIndex(55) = 59 'Asc("7")
    mbytarrReverseIndex(56) = 60 'Asc("8")
    mbytarrReverseIndex(57) = 61 'Asc("9")
    mbytarrReverseIndex(43) = 62 'Asc("+")
    mbytarrReverseIndex(47) = 63 'Asc("/")
End Sub

Private Function Decode64(sInput As String) As String
    
If LenB(sInput) Then
    Decode64 = StrConv(DecodeArray64(sInput), vbUnicode)
End If

End Function

Private Function DecodeArray64(sInput As String) As Byte()
    
    Dim bytInput() As Byte
    Dim bytWorkspace() As Byte
    Dim bytResult() As Byte
    Dim lInputCounter As Long
    Dim lWorkspaceCounter As Long
    
    
    If mbytarrReverseIndex(47) <> 63 Then Call Initialize64
    
    bytInput = Replace(Replace(sInput, vbCrLf, ""), "=", "")
    ReDim bytWorkspace(LBound(bytInput()) To (UBound(bytInput()) * 2)) As Byte
    
    lWorkspaceCounter = LBound(bytWorkspace())
    For lInputCounter = LBound(bytInput()) To UBound(bytInput())
        bytInput(lInputCounter) = mbytarrReverseIndex(bytInput(lInputCounter))
    Next 'lInputCounter
    
    For lInputCounter = LBound(bytInput()) To (UBound(bytInput()) - ((UBound(bytInput()) Mod 8) + 8)) Step 8
        bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * mbytSHIFT2) + (bytInput(lInputCounter + 2) \ mbytSHIFT4)
        bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And mbytMASK2) * mbytSHIFT4) + (bytInput(lInputCounter + 4) \ mbytSHIFT2)
        bytWorkspace(lWorkspaceCounter + 2) = ((bytInput(lInputCounter + 4) And mbytMASK1) * mbytSHIFT6) + bytInput(lInputCounter + 6)
        lWorkspaceCounter = lWorkspaceCounter + 3
    Next 'lInputCounter
    
    Select Case (UBound(bytInput()) Mod 8)
        Case 3
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * mbytSHIFT2) + (bytInput(lInputCounter + 2) \ mbytSHIFT4)
        Case 5
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * mbytSHIFT2) + (bytInput(lInputCounter + 2) \ mbytSHIFT4)
            bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And mbytMASK2) * mbytSHIFT4) + (bytInput(lInputCounter + 4) \ mbytSHIFT2)
            lWorkspaceCounter = lWorkspaceCounter + 1
        Case 7
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * mbytSHIFT2) + (bytInput(lInputCounter + 2) \ mbytSHIFT4)
            bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And mbytMASK2) * mbytSHIFT4) + (bytInput(lInputCounter + 4) \ mbytSHIFT2)
            bytWorkspace(lWorkspaceCounter + 2) = ((bytInput(lInputCounter + 4) And mbytMASK1) * mbytSHIFT6) + bytInput(lInputCounter + 6)
            lWorkspaceCounter = lWorkspaceCounter + 2
        Case Else
    End Select
    
    ReDim bytResult(LBound(bytWorkspace()) To lWorkspaceCounter) As Byte
    If LBound(bytWorkspace()) = 0 Then lWorkspaceCounter = lWorkspaceCounter + 1
    Call CopyMemory(VarPtr(bytResult(LBound(bytResult()))), VarPtr(bytWorkspace(LBound(bytWorkspace()))), lWorkspaceCounter)
    
    DecodeArray64 = bytResult()
End Function

Private Function Encode64(ByRef sInput As String) As String
    
    Dim bytTemp() As Byte
    
    If LenB(sInput) Then
        bytTemp() = StrConv(sInput, vbFromUnicode)
        Encode64 = EncodeArray64(bytTemp())
    End If
    
End Function

Private Function EncodeArray64(ByRef bytInput() As Byte) As String

    On Error GoTo ErrorHandler
    
    
    Dim bytWorkspace() As Byte, bytResult() As Byte
    Dim bytCrLf(0 To 3) As Byte, lCounter As Long
    Dim lWorkspaceCounter As Long, lLineCounter As Long
    Dim lCompleteLines As Long, lBytesRemaining As Long
    Dim lpWorkSpace As Long, lpResult As Long
    Dim lpCrLf As Long

    If mbytarrReverseIndex(47) <> 63 Then Call Initialize64
    
    
    If UBound(bytInput()) < 1024 Then
        ReDim bytWorkspace(LBound(bytInput()) To (LBound(bytInput()) + 4096)) As Byte
    Else
        ReDim bytWorkspace(LBound(bytInput()) To (UBound(bytInput()) * 4)) As Byte
    End If

    lWorkspaceCounter = LBound(bytWorkspace())
    
    For lCounter = LBound(bytInput()) To (UBound(bytInput()) - ((UBound(bytInput()) Mod 3) + 3)) Step 3
        bytWorkspace(lWorkspaceCounter) = mbytarrIndex((bytInput(lCounter) \ mbytSHIFT2))
        bytWorkspace(lWorkspaceCounter + 2) = mbytarrIndex(((bytInput(lCounter) And mbytMASK1) * mbytSHIFT4) + ((bytInput(lCounter + 1)) \ mbytSHIFT4))
        bytWorkspace(lWorkspaceCounter + 4) = mbytarrIndex(((bytInput(lCounter + 1) And mbytMASK2) * mbytSHIFT2) + (bytInput(lCounter + 2) \ mbytSHIFT6))
        bytWorkspace(lWorkspaceCounter + 6) = mbytarrIndex(bytInput(lCounter + 2) And mbytMASK3)
        lWorkspaceCounter = lWorkspaceCounter + 8
    Next 'lCounter
    
    Select Case (UBound(bytInput()) Mod 3):
        Case 0
            bytWorkspace(lWorkspaceCounter) = mbytarrIndex((bytInput(lCounter) \ mbytSHIFT2))
            bytWorkspace(lWorkspaceCounter + 2) = mbytarrIndex((bytInput(lCounter) And mbytMASK1) * mbytSHIFT4)
            bytWorkspace(lWorkspaceCounter + 4) = mbytEQUALSIGN
            bytWorkspace(lWorkspaceCounter + 6) = mbytEQUALSIGN
        Case 1
            bytWorkspace(lWorkspaceCounter) = mbytarrIndex((bytInput(lCounter) \ mbytSHIFT2))
            bytWorkspace(lWorkspaceCounter + 2) = mbytarrIndex(((bytInput(lCounter) And mbytMASK1) * mbytSHIFT4) + ((bytInput(lCounter + 1)) \ mbytSHIFT4))
            bytWorkspace(lWorkspaceCounter + 4) = mbytarrIndex((bytInput(lCounter + 1) And mbytMASK2) * mbytSHIFT2)
            bytWorkspace(lWorkspaceCounter + 6) = mbytEQUALSIGN
        Case 2:
            bytWorkspace(lWorkspaceCounter) = mbytarrIndex((bytInput(lCounter) \ mbytSHIFT2))
            bytWorkspace(lWorkspaceCounter + 2) = mbytarrIndex(((bytInput(lCounter) And mbytMASK1) * mbytSHIFT4) + ((bytInput(lCounter + 1)) \ mbytSHIFT4))
            bytWorkspace(lWorkspaceCounter + 4) = mbytarrIndex(((bytInput(lCounter + 1) And mbytMASK2) * mbytSHIFT2) + ((bytInput(lCounter + 2)) \ mbytSHIFT6))
            bytWorkspace(lWorkspaceCounter + 6) = mbytarrIndex(bytInput(lCounter + 2) And mbytMASK3)
        Case Else
    End Select
    
    lWorkspaceCounter = lWorkspaceCounter + 8
    
    If lWorkspaceCounter <= mlMAXBYTESPERLINE Then
        EncodeArray64 = Mid$(bytWorkspace, 1, InStr(1, bytWorkspace, Chr$(0)) - 1)
    Else
        bytCrLf(0) = 13
        bytCrLf(1) = 0
        bytCrLf(2) = 10
        bytCrLf(3) = 0
        ReDim bytResult(LBound(bytWorkspace()) To UBound(bytWorkspace()))
        lpWorkSpace = VarPtr(bytWorkspace(LBound(bytWorkspace())))
        lpResult = VarPtr(bytResult(LBound(bytResult())))
        lpCrLf = VarPtr(bytCrLf(LBound(bytCrLf())))
        lCompleteLines = Fix(lWorkspaceCounter / mlMAXBYTESPERLINE)
        
        For lLineCounter = 0 To lCompleteLines
            Call CopyMemory(lpResult, lpWorkSpace, mlMAXBYTESPERLINE)
            lpWorkSpace = lpWorkSpace + mlMAXBYTESPERLINE
            lpResult = lpResult + mlMAXBYTESPERLINE
            Call CopyMemory(lpResult, lpCrLf, 4&)
            lpResult = lpResult + 4&
        Next 'lLineCounter
        
        lBytesRemaining = lWorkspaceCounter - (lCompleteLines * mlMAXBYTESPERLINE)
        If lBytesRemaining > 0 Then Call CopyMemory(lpResult, lpWorkSpace, lBytesRemaining)
        EncodeArray64 = Mid$(bytResult, 1, InStr(1, bytResult, Chr$(0)) - 1)
    End If
    
    Exit Function

ErrorHandler:
    
    Erase bytResult()
    EncodeArray64 = bytResult()
End Function

Private Sub DecryptByte(byteArray() As Byte, Optional sKey As String)
Call EncryptByte(byteArray(), sKey)
End Sub
Public Function EncryptString(ByVal sTxt As String, Optional ByVal sKey As String, Optional bOutputIn64 As Boolean) As String
    
    Dim byteArray() As Byte
    
    byteArray() = StrConv(sTxt, vbFromUnicode)
    Call EncryptByte(byteArray(), sKey)
    EncryptString = StrConv(byteArray(), vbUnicode)
    If bOutputIn64 = True Then EncryptString = Encode64(EncryptString)
    
End Function
Public Function DecryptString(ByVal sTxt As String, Optional ByVal sKey As String, Optional bIsTextIn64 As Boolean) As String
    
    Dim byteArray() As Byte
    
    If bIsTextIn64 = True Then sTxt = Decode64(sTxt)
    
    byteArray() = StrConv(sTxt, vbFromUnicode)
    Call DecryptByte(byteArray(), sKey)
    DecryptString = StrConv(byteArray(), vbUnicode)
    
End Function
Private Sub EncryptByte(byteArray() As Byte, Optional sKey As String)

Dim i As Long, j As Long, bytTmp As Byte, lOffset As Long, lOrigLen As Long
Dim lCipherLen As Long, lCurrPercent As Long, lNextPercent As Long
Dim iarrBox() As Integer

ReDim iarrBox(0 To 255) As Integer

If (Len(sKey) > 0) Then Me.Key = sKey
Call CopyMem(iarrBox(0), miarrBox(0), 512)
lOrigLen = UBound(byteArray) + 1
lCipherLen = lOrigLen

For lOffset = 0 To (lOrigLen - 1)
    i = (i + 1) Mod 256
    j = (j + iarrBox(i)) Mod 256
    bytTmp = iarrBox(i)
    iarrBox(i) = iarrBox(j)
    iarrBox(j) = bytTmp
    byteArray(lOffset) = byteArray(lOffset) Xor (iarrBox((iarrBox(i) + iarrBox(j)) Mod 256))
    If (lOffset >= lNextPercent) Then
        lCurrPercent = Int((lOffset / lCipherLen) * 100)
        lNextPercent = (lCipherLen * ((lCurrPercent + 1) / 100)) + 1
        RaiseEvent Progress(lCurrPercent)
    End If
Next

If (lCurrPercent <> 100) Then RaiseEvent Progress(100)
Erase iarrBox()

End Sub
Public Property Let Key(sNewValue As String)
    
    Dim i As Long, j As Long, bytTmp As Byte, bytarrKey() As Byte, lKeyLen As Long
    
    If (msKey <> sNewValue) Then
            
        msKey = sNewValue
        bytarrKey() = StrConv(msKey, vbFromUnicode)
        lKeyLen = Len(msKey)
        For i = 0 To 255
            miarrBox(i) = i
        Next 'i
        
        For i = 0 To 255
            j = (j + miarrBox(i) + bytarrKey(i Mod lKeyLen)) Mod 256
            bytTmp = miarrBox(i)
            miarrBox(i) = miarrBox(j)
            miarrBox(j) = bytTmp
        Next
        
    End If
    
End Property

Private Sub Class_Initialize()

ReDim miarrBox(0 To 255) As Integer
ReDim mbytarrIndex(0 To 63) As Byte
ReDim mbytarrReverseIndex(0 To 255) As Byte

End Sub


Private Sub Class_Terminate()

ReDim miarrBox(0 To 0) As Integer
ReDim mbytarrIndex(0 To 0) As Byte
ReDim mbytarrReverseIndex(0 To 0) As Byte

Erase miarrBox(), mbytarrIndex(), mbytarrReverseIndex()
End Sub
