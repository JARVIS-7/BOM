VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msFileName As String
Private mcolEntries As Collection

Private Sub ResetIni()
    
    Do While mcolEntries.Count > 0
        Do While mcolEntries(1).Count > 0
            mcolEntries(1).Remove 1
        Loop
        mcolEntries.Remove 1
    Loop
    
End Sub

Public Function ReadIni(Optional sFileName As String, Optional sSection As String = "") As Boolean
    
    Dim i As Integer
    Dim sTmp As String
    Dim sSec As String
    Dim sEnt As String
    Dim sValue As String
    Dim sArr() As String
    
    Call ResetIni
    
    If LenB(sFileName) Then msFileName = sFileName
    
    If LenB(msFileName) Then
        sTmp = ReadFromFile(msFileName)
        If LenB(sTmp) Then
        
            sArr() = Split(sTmp & vbCrLf, vbCrLf)
            
            For i = LBound(sArr()) To UBound(sArr())
                
                If Trim$(sArr(i)) Like "[[]*[]]" Then
                    sSec = Trim$(Mid$(sArr(i), 2, Len(sArr(i)) - 2))
                Else
                    If InStr(1, sArr(i), "=") > 0 And Mid$(sArr(i), 1, 1) Like "[a-zA-Z]" Then
                        sEnt = Trim$(Mid$(sArr(i), 1, InStr(1, sArr(i), "=") - 1))
                        sValue = Trim$(Mid$(sArr(i), InStr(1, sArr(i), "=") + 1))
                        If Mid$(sValue, 1, 1) = """" And Right$(sValue, 1) = """" And Len(sValue) > 1 Then sValue = Mid$(sValue, 2, Len(sValue) - 2)
                    Else
                        sEnt = ""
                        sValue = sArr(i)
                    End If
                    
                    If UCase$(sSec) = UCase$(sSection) Or LenB(sSection) = 0 Then Call SetValue(sSec, sEnt, sValue)
                          
                End If
            
            Next 'i
        End If 'LenB(sTmp)
    End If 'LenB(msFileName)
    
    ReadIni = True
    
    Erase sArr()
    
End Function

Public Sub SetValue(ByVal sSection As String, ByVal sEntry As String, ByVal vValue As Variant)
    
    Dim sSectionTmp As String
    Dim sEntryTmp As String
    
    If VarType(vValue) = vbBoolean Then
      If vValue Then
        vValue = "1"
      Else
        vValue = "0"
      End If
    End If
    
    sSectionTmp = UCase$(sSection)
    sEntryTmp = UCase$(sEntry)
    
    If Not ExistCollectionKey(mcolEntries, sSectionTmp) Then
        mcolEntries.Add New Collection, sSectionTmp
    End If
    
    If ExistCollectionKey(mcolEntries(sSectionTmp), sEntryTmp) Then
        With mcolEntries(sSectionTmp)
            Call .Add("dummy", "dummy", sEntryTmp)
            Call .Remove(sEntryTmp)
            Call .Add(Array(sSection, sEntry, vValue), sEntryTmp, "dummy")
            Call .Remove("dummy")
        End With
    Else
        mcolEntries(sSectionTmp).Add Array(sSection, sEntry, vValue), sEntryTmp
    End If
    
End Sub

Public Function GetValue(ByVal sSection As String, ByVal sEntry As String, Optional ByVal sDefaultValue As String = "") As String
    
    Dim sSectionTmp As String
    Dim sEntryTmp As String
    
    sSectionTmp = UCase(sSection)
    sEntryTmp = UCase(sEntry)
    
    GetValue = sDefaultValue
    If ExistCollectionKey(mcolEntries, sSectionTmp) Then
        If ExistCollectionKey(mcolEntries(sSectionTmp), sEntryTmp) Then
            GetValue = mcolEntries(sSectionTmp)(sEntryTmp)(2)
        End If
    End If
    
End Function

Private Sub DeleteSection(ByVal sSection As String)
    
    'MD-Marker , Sub wird nicht aufgerufen
    
'    Dim uSection As String
'
'    uSection = UCase(Section)
'    If ExistCollectionKey(mcolEntries, uSection) Then
'        Do While mcolEntries(uSection).Count > 0
'            mcolEntries(uSection).Remove 1
'        Loop
'        mcolEntries.Remove uSection
'    End If
    
End Sub

Public Function WriteIni(Optional sFileName As String) As Boolean
    
    On Error GoTo ERR_HANDLER
    Dim i As Integer
    Dim j As Integer
    Dim sTmp As String
    Dim sValue As String
    Dim sArr() As String
    Dim c As Integer
    Dim iArrSize As Integer
    Dim sFileTmp As String


    
    If sFileName > "" Then msFileName = sFileName
    
    If LenB(msFileName) Then
            
        
        ReDim sArr(1 To 100) As String
        iArrSize = 100
        
        For i = 1 To mcolEntries.Count
            If mcolEntries(i).Count > 0 Then
                
                c = c + 1
                If c > iArrSize Then
                    iArrSize = iArrSize + 100
                    ReDim Preserve sArr(1 To iArrSize) As String
                End If
                sArr(c) = "[" & Trim(mcolEntries(i)(1)(0)) & "]"
                
                For j = 1 To mcolEntries(i).Count
                    sValue = mcolEntries(i)(j)(2)
                    If mcolEntries(i)(j)(1) > "" Then
                        If sValue <> """" And (Left(sValue, 1) = " " Or Left(sValue, 1) = """" Or Right(sValue, 1) = " " Or Right(sValue, 1) = """") Then
                            sValue = """" & sValue & """"
                        End If
                    End If
                    
                    c = c + 1
                    If c > iArrSize Then
                        iArrSize = iArrSize + 100
                        ReDim Preserve sArr(1 To iArrSize) As String
                    End If
                    If mcolEntries(i)(j)(1) > "" Then
                        sArr(c) = Trim(mcolEntries(i)(j)(1)) & "=" & sValue
                    Else
                        sArr(c) = sValue
                    End If
                    
                Next j
            End If
        Next i
                
                
        If c > 0 Then
            ReDim Preserve sArr(1 To c) As String
            sTmp = Join(sArr, vbCrLf)
        End If
        
        sFileTmp = MakeTempFile()
        If sFileTmp > "" Then
            Call SaveToFile(sTmp, sFileTmp)
            If Dir(sFileTmp) > "" Then
                If FileLen(sFileTmp) > 0 Then
                    If Dir(msFileName) > "" Then Call Kill(msFileName)
                    Call FileCopy(sFileTmp, msFileName)
                    Call Kill(sFileTmp)
                    WriteIni = True
                End If
            End If
        End If
    End If 'LenB(msFileName)
    
    Erase sArr()
    
    Exit Function
    
ERR_HANDLER:
    
End Function

Private Sub Class_Initialize()
Set mcolEntries = New Collection
End Sub

Private Sub Class_Terminate()
    
    Call ResetIni
    
    Set mcolEntries = Nothing
    
End Sub
