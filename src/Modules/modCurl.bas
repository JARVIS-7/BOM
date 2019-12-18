Attribute VB_Name = "modCurl"
Option Explicit
Private mcolCookieFiles As New Collection
Private mcolParams As New Collection

Public Function Curl(sUrl As String, sPostData As String, sReferer As String, sEBayUser As String, Optional bWait As Boolean = True) As String
          
10        On Error GoTo ERROR_HANDLER
          
          Dim i As Integer
          Dim F As Long
          Dim t As String
          Dim p1 As Integer
          Dim p2 As Integer
          Dim c As String
          Dim sCookieFile As String
          Dim sOutFile As String
          Dim myParams  As String
          Dim lProcessID As Long
          Dim lTimestamp As Long
          Dim sKey As String
          Dim a As Variant
          Dim v As Variant
          Dim colCurlConfig As Collection
          Dim colCurlOptions As Collection
          Dim stmpPass As String
          Dim myTime As Long
          
20        Set colCurlConfig = New Collection
30        Set colCurlOptions = New Collection

40        If sEBayUser = "" Then
50            sEBayUser = "default"
60        Else
              ' es gibt einen User, dann auch ein Passwort
70            stmpPass = gsPass
80        End If
          
          ' für den output:
90        sOutFile = MakeTempFile()
          
      '    a = Array("FILE_HEADER", "FILE_STDOUT", "FILE_STDERR", "FILE_TRACE", "FILE_CONFIG")
      '
      '
      '    If ExistCollectionKey(mcolCookieFiles, sEBayUser) Then
      '        sCookieFile = mcolCookieFiles(sEBayUser)
      '    Else
      '        sCookieFile = MakeTempFile()
      '        mcolCookieFiles.Add sCookieFile, sEBayUser
      '    End If
          
      '    With colCurlConfig
      '        .Add sCookieFile, "FILE_COOKIE"
      '
      '        For Each v In a
      '            Call .Add(MakeTempFile(), CStr(v))
      '            Call SaveToFile("", colCurlConfig(CStr(v)))
      '        Next
      '
      '        Call .Add("Accept-Language: " & gsBrowserLanguage, "HEADER")
      '        Call .Add(gsBrowserIdString, "USER_AGENT")
      '        Call .Add(Int(glHttpTimeOut / 1000), "CONNECT_TIMEOUT")
      '        Call .Add(Int(glHttpTimeOut / 1000), "TRANSFER_TIMEOUT")
      '        Call .Add(10, "MAX_REDIRECT")
      '        Call .Add(3, "MAX_RETRY")
      '        Call .Add(gsProxyName & IIf(giProxyPort > 0, ":" & CStr(giProxyPort), ""), "PROXY_SERVER")
      '        Call .Add(gsProxyUser & IIf(gsProxyPass > "", ":" & gsProxyPass, ""), "PROXY_CREDENTIALS")
      '        Call .Add(sUrl, "URL")
      '        Call .Add(sReferer, "REFERER")
      '        Call .Add(sPostData, "DATA")
      '    End With 'colCurlConfig
          
      '    With colCurlOptions
      '        Call .Add("--dump-header {FILE_HEADER}")
      '        Call .Add("--compressed")
      ''       Call .Add("--fail")
      '        Call .Add("--silent")
      '        Call .Add("--show-error")
      '        Call .Add("--user-agent {USER_AGENT}")
      '        Call .Add("--connect-timeout {CONNECT_TIMEOUT}")
      '        Call .Add("--location")
      '        Call .Add("--max-time {TRANSFER_TIMEOUT}")
      '        Call .Add("--max-redirs {MAX_REDIRECT}")
      '        Call .Add("--retry {MAX_RETRY}")
      '        Call .Add("--output {FILE_STDOUT}")
      '        Call .Add("--stderr {FILE_STDERR}")
      '        Call .Add("--header {HEADER}")
      '        Call .Add("--url {URL}")
      '        Call .Add("--ebayuser {eBayUser}")
      '        Call .Add("--ebaypass {eBayPass}")
      '
      '        If sReferer > "" Then Call .Add("--referer {REFERER}")
      '        If sPostData > "" Then Call .Add("--data {DATA}")
      '
      '        If sEBayUser <> "anonymous" Then
      '          Call .Add("--cookie {FILE_COOKIE}")
      '          Call .Add("--cookie-jar {FILE_COOKIE}")
      '        End If
      '
      '        If gbUseProxy Then
      '            Call .Add("--proxy {PROXY_SERVER}")
      '            If gsProxyUser > "" Then
      '                Call .Add("--proxy-user {PROXY_USER}:{PROXY_PASS}")
      '                Call .Add("--proxy-anyauth")
      '            End If
      '        End If
      '
      '        If Dir(App.Path & "\curl-ca-bundle.crt") = "" Then
      '            Call .Add("--insecure")
      '        End If
      '
      '        Call .Add("#--trace-ascii {FILE_TRACE}")
      '        Call .Add("#--trace-time")
      '        Call .Add("#--limit-rate 7168")
      '    End With 'colCurlOptions
          
      '    F = FreeFile()
      '    Open colCurlConfig("FILE_CONFIG") For Output As F
      '        For i = 1 To colCurlOptions.Count
      '            t = colCurlOptions.Item(i)
      '            p1 = InStr(1, t, "{")
      '            p2 = InStr(p1 + 1, t, "}")
      '            Do While p1 > 0 And p2 > 0
      '                c = Mid(t, p1 + 1, p2 - p1 - 1)
      '                c = Replace(colCurlConfig.Item(c), IIf(Left(c, 5) = "FILE_", "\", "/"), "/")
      '                t = Replace(t, Mid(t, p1, p2 - p1 + 1), """" & c & """")
      '                p1 = InStr(1, t, "{")
      '                p2 = InStr(p1 + 1, t, "}")
      '            Loop
      '            Print #F, t
      '        Next i
      '    Close F
          
100       myParams = "url=""" & sUrl & """ u=""" & sEBayUser & """ p=""" & stmpPass & """ of=""" & sOutFile & """ d=""" & sPostData & """ ua=""" & gsBrowserIdString & """"
110       Debug.Print myParams
          
120       myTime = GetTickCount
          
130       lProcessID = ShellStart("""" & App.Path & "\JARVIS-7.exe """ & myParams & """", vbHide)
         'lProcessID = ShellStart("""" & App.Path & "\curl.exe"" --config """ & colCurlConfig("FILE_CONFIG") & """", vbHide)
          
140       If lProcessID <= 0 Then DebugPrint "no curl process id, url = " & sUrl
          
150       lTimestamp = Timer * 100
160       sKey = lProcessID & "/" & lTimestamp
          
          'mcolParams.Add a, "a" & sKey
170       mcolParams.Add sOutFile, "c" & sKey
180       mcolParams.Add sUrl, "u" & sKey
190       mcolParams.Add sKey, "k" & sKey
          
200       If bWait Then
210           Do
220               Call Sleep(10)
230               DoEvents
240           Loop While ShellStillRunning(lProcessID)
250           Curl = CurlGetData(sKey)
260       Else
270           mcolParams.Add lProcessID, sKey
280           Curl = lProcessID
290       End If
          
300       Debug.Print "duration: " & GetTickCount - myTime
          
310       Set colCurlConfig = Nothing
320       Set colCurlOptions = Nothing
          
330   Exit Function
ERROR_HANDLER:
340       DebugPrint "error in function curl: " & Err.Description & " Line: " & Erl
        Debug.Print "error in function curl: " & Err.Description & " Line: " & Erl
350       Err.Clear
              
End Function

Private Function CurlGetData(sKey As String, Optional ByRef sUrlReturn As String) As String
          
          Dim a As Variant
          Dim v As Variant
          'Dim vntCurlConfig As Variant
          Dim sUrl As String
          Dim sOutFile As String
          Dim bOk As Boolean
          Dim b() As Byte
          'Dim sServerHeader As String
          Dim sError As String

10        If ExistCollectionKey(mcolParams, "c" & sKey) Then
          
              'a = mcolParams("a" & sKey)
20            sOutFile = mcolParams("c" & sKey)
30            sUrl = mcolParams("u" & sKey)
              
      '        sUrlReturn = sUrl
              
40            If FileLen(sOutFile) > 0 Then
50                b() = ReadFromFile(sOutFile, True)
60                If UBound(b()) >= LBound(b()) Then
70                    bOk = True
80                End If
90            End If
              
100           If bOk Then
                  
      '            sServerHeader = ReadFromFile(vntCurlConfig("FILE_HEADER"))
      '            'MsgBox "SiteEncoding: " & gsSiteEncoding & vbCrLf & "ServerHeader: " & vbCrLf & sServerHeader
      '            If InStr(1, sServerHeader, "charset=utf-8", vbTextCompare) > 0 Then
110                   CurlGetData = ByteArray2String(Decode_UTF8(b))
                      'CurlGetData = Replace(CurlGetData, "charset=utf-8", "charset=ISO-8859-1", , , vbTextCompare)
      '            Else
      '                CurlGetData = StrConv(b, vbUnicode)
      '            End If
120           Else
130               sError = "unknown error"
      '            sError = ReadFromFile(vntCurlConfig("FILE_STDERR"))
      '            Do While Right(sError, 1) = vbCr Or Right(sError, 1) = vbLf: sError = Left(sError, Len(sError) - 1): Loop
140               frmHaupt.SetStatus sError, True
150               DoEvents
160               Call DebugPrint(sError & " (" & sUrl & ")")
170           End If
              
180           On Error Resume Next
      '        For Each v In a
      '            Call Kill(vntCurlConfig.Item(v))
      '        Next
190           Call Kill(sOutFile)

200           On Error GoTo ERROR_HANDLER
              
      '        Do While vntCurlConfig.Count > 0
      '            vntCurlConfig.Remove 1
      '        Loop
              
210           'mcolParams.Remove "a" & sKey
220           mcolParams.Remove "c" & sKey
230           mcolParams.Remove "u" & sKey
240           mcolParams.Remove "k" & sKey
250           If ExistCollectionKey(mcolParams, sKey) Then mcolParams.Remove sKey
260       End If 'ExistCollectionKey(mcolParams, "c" & sKey)
          
270       Exit Function
          
ERROR_HANDLER:
280           DebugPrint "error in function CurlGetData: " & Err.Description & " Line: " & Erl
290           Err.Clear
         
End Function

Public Function PollPendingCurls(sUrlReturn As String, sDataReturn As String) As Boolean
    
    Dim i As Integer
    Dim sKey As String
    
    For i = 1 To mcolParams.Count
        If TypeName(mcolParams(i)) = "String" Then
            sKey = mcolParams(i)
            If ExistCollectionKey(mcolParams, "" & sKey) And ExistCollectionKey(mcolParams, "k" & sKey) Then
                If Not ShellStillRunning(mcolParams(sKey)) Then
                    sDataReturn = CurlGetData(sKey, sUrlReturn)
                    PollPendingCurls = CBool(sDataReturn > "")
                    Exit For
                End If
            End If
        End If
    Next i
    
End Function

Public Sub RemoveCookies()
    
    On Error Resume Next
    Dim i As Integer
    
    For i = 1 To mcolCookieFiles.Count
        Call Kill(mcolCookieFiles(i))
    Next
    On Error GoTo 0
    
End Sub

Public Function TestForCurl() As Boolean
    
    If Dir(App.Path & "\JARVIS-7.exe") > "" Then TestForCurl = True
    
End Function
