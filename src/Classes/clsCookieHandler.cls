VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCookieHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolUserCookies As Collection

Private Sub Class_Initialize()

  Set mcolUserCookies = New Collection

End Sub

Public Function GetCookieHeader(strUrl As String, strUser As String) As String
    
    On Error Resume Next
    Dim colCookies As Collection
    mcolUserCookies.Add New Collection, LCase(strUser)
    Set colCookies = mcolUserCookies.Item(LCase(strUser))
    On Error GoTo 0
    
    Dim strUrlDomain As String
    Dim strUrlPath As String
    Dim intUrlPort As Integer
    Dim blnUrlSecure As Boolean
    Dim strName As String
    Dim strValue As String
    Dim strDomain As String
    Dim strPath As String
    Dim intPort As Integer
    Dim blnSecure As Boolean
    Dim strCookies As String
    Dim blnAddCookie As Boolean
    Dim i As Integer
    
    
    strUrlDomain = GetDomainFromUrl(strUrl)
    strUrlPath = GetPathFromUrl(strUrl)
    intUrlPort = GetPortFromUrl(strUrl)
    blnUrlSecure = IIf(strUrl Like "https://*", 1, 0)
    
    If intUrlPort = 0 Then
        If strUrl Like "https://*" Then
            intUrlPort = 443
        Else
            intUrlPort = 80
        End If
    End If
    
    For i = 1 To colCookies.Count
        strName = ""
        strValue = ""
        strDomain = ""
        strPath = ""
        intPort = 0
        blnSecure = False
        
        On Error Resume Next
        strName = colCookies(i).Item("Name")
        strValue = colCookies(i).Item("Value")
        strDomain = colCookies(i).Item("Domain")
        strPath = colCookies(i).Item("Path")
        intPort = colCookies(i).Item("Port")
        blnSecure = colCookies(i).Item("Secure")
        On Error GoTo 0
        
        blnAddCookie = True
        
        If strValue = "" Then blnAddCookie = False
        If Not strUrlDomain Like "*" & strDomain Then blnAddCookie = False
        If Not strUrlPath Like strPath & "*" Then blnAddCookie = False
        If intPort > 0 And intPort <> intUrlPort Then blnAddCookie = False
        If blnSecure And Not blnUrlSecure Then blnAddCookie = False
        
        If blnAddCookie Then
            If strCookies > "" Then strCookies = strCookies & "; "
            strCookies = strCookies & strName & "=" & strValue
        End If
        
    Next i
    
    GetCookieHeader = strCookies
    
End Function

Public Sub ExtractCookies(ByVal strHeader As String, strUser As String)
    
    Dim iPos1 As Integer
    Dim iPos2 As Integer
    
    iPos1 = InStr(1, strHeader, vbCrLf & "Set-Cookie:", vbTextCompare)
    Do While (iPos1 > 0)
        iPos1 = iPos1 + 13
        iPos2 = InStr(iPos1, strHeader, vbCrLf)
        If iPos2 > 0 Then
            'Debug.Print "Set-Cookie:" & Mid(strHeader, iPos1, iPos2 - iPos1)
            Call ExtractCookie(Trim(Mid(strHeader, iPos1, iPos2 - iPos1)), strUser)
            iPos1 = InStr(iPos2, strHeader, vbCrLf & "Set-Cookie:", vbTextCompare)
        End If
    Loop
    
End Sub

Private Sub ExtractCookie(strCookie As String, strUser As String)
    
    
    Dim i As Integer
    Dim strDomain As String
    Dim strCookieKey As String
    Dim tmpArray As Variant
    Dim tmpArray2 As Variant
    Dim colCookies As Collection
    Dim colNewCookie As Collection
    
    On Error Resume Next
    mcolUserCookies.Add New Collection, LCase(strUser)
    Set colCookies = mcolUserCookies.Item(LCase(strUser))
    On Error GoTo 0
    
    tmpArray = Split(strCookie, ";")
    
    If UBound(tmpArray) >= LBound(tmpArray) Then
        
        Set colNewCookie = New Collection
        
        For i = LBound(tmpArray) To UBound(tmpArray)
            
            tmpArray2 = Split(tmpArray(i), "=")
            
            If UBound(tmpArray2) >= LBound(tmpArray2) Then
                If i = LBound(tmpArray) Then
                    Call colNewCookie.Add(Trim(tmpArray2(LBound(tmpArray2))), "Name")
                    Call colNewCookie.Add(Trim(tmpArray2(LBound(tmpArray2) + 1)), "Value")
                    'Debug.Print "Name: " & Trim(tmpArray2(LBound(tmpArray2)))
                    'Debug.Print "Value: " & Trim(tmpArray2(LBound(tmpArray2) + 1))
                Else
                    If UBound(tmpArray2) > LBound(tmpArray2) Then
                        Call colNewCookie.Add(Trim(tmpArray2(LBound(tmpArray2) + 1)), Trim(tmpArray2(LBound(tmpArray2))))
                        'Debug.Print Trim(tmpArray2(LBound(tmpArray2))) & ": " & Trim(tmpArray2(LBound(tmpArray2) + 1))
                    Else
                        colNewCookie.Add True, Trim(tmpArray2(LBound(tmpArray2)))
                        'Debug.Print Trim(tmpArray2(LBound(tmpArray2)))
                    End If
                End If
            End If
        Next i

        On Error Resume Next
        strDomain = colNewCookie.Item("Domain")
        strCookieKey = colNewCookie.Item("Name") & "/" & strDomain

        'ein evtl. vorhandenes Cookie leermachen und entfernen
        For i = 1 To colCookies.Item(strCookieKey).Count
           colCookies.Item(strCookieKey).Remove 1
        Next i
        Call colCookies.Remove(strCookieKey)
        On Error GoTo 0
        
        'dann das neue Cookie hinzufügen
        Call colCookies.Add(colNewCookie, strCookieKey)
        
    End If

End Sub

Private Sub Class_Terminate()
Set mcolUserCookies = Nothing
End Sub


