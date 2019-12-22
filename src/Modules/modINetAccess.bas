Attribute VB_Name = "modINetAccess"
'******************************************************
'
' BOM default header
'
' this file is OpenSource.
' license model: GPL
' please respect the limitations
'
' main language: german
' compiled under VB6 SP5 german
'
' $author: internet :-)$
' $id: V 2.0.4 date 090404 ingo$
' $version: 2.0.4$
' $file: $
'
' last modified:
' &date: 090404$
'
' contact: visit http://de.groups.yahoo.com/group/BOMInfo
'
'*******************************************************
Option Explicit
'
' gesammelte Werke für den INET- und POP- Zugriff
'

'local Consts

Private Const msEOM As String = vbCrLf & "." & vbCrLf

'Mode constants for WaitFor
Private Const mlSTATUS As Long = 1&
Private Const mlDOT As Long = 2&
Private Const mlCLOSED As Long = 4&
Private Const mlCONNECTED As Long = 3&
Private Const mlSMTPDATA As Long = 6&
Private Const mlNTPDATA As Long = 8&

Private mlGlobHandle As Long
         
  Private Declare Function URLDownloadToFile Lib "urlmon.dll" Alias "URLDownloadToFileA" _
                  (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
                  ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
         
' Für Modem- Betrieb

  Private Declare Function InternetDial Lib "wininet.dll" _
                 (ByVal hwndParent As Long, ByVal lpszConiID _
                 As String, ByVal dwFlags As Long, ByRef hCon _
                 As Long, ByVal dwReserved As Long) As Long

  Private Declare Function InternetHangUp Lib "wininet.dll" _
                 (ByVal hCon As Long, ByVal dwReserved _
                 As Long) As Long
                 
  'Public Declare Function InternetGetConnectedState Lib "wininet.dll" _
                 (ByRef dwFlags As Long, _
                  ByVal dwReserved As Long) As Long

  Private Declare Function RasEnumEntries Lib "RasApi32.dll" _
                 Alias "RasEnumEntriesA" (ByVal Reserved$, ByVal _
                 lpszPhonebook$, lprasentryname As Any, lpcb As Long, _
                 lpcEntries As Long) As Long
                 
  Private Declare Function RasEnumConnections Lib "RasApi32.dll" _
                 Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As _
                 Long, lpcConnections As Long) As Long

  'Zugriff auf die Registry raus

'Stati aus GetConnectState
'Public Const INTERNET_CONNECTION_MODEM As Long = &H1
'Public Const INTERNET_CONNECTION_LAN As Long = &H2
'Public Const INTERNET_CONNECTION_PROXY As Long = &H4
'Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
'Public Const INTERNET_RAS_INSTALLED As Long = &H10
'Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
'Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40

'RAS Dial Const
'Const DIAL_UNATTENDED = &H8000
'Const DIAL_FORCE_ONLINE = 1
Private Const DIAL_FORCE_UNATTENDED As Long = 2&
Private Const RAS95_MaxEntryName As Long = 256&

Private Type RASENTRYNAME95
           dwSize As Long
           szEntryName(RAS95_MaxEntryName) As Byte
End Type

Private Type RASCONN
  dwSize As Long
  hRasConn As Long
  szEntryName(256) As Byte
  szDeviceType(16) As Byte
  szDeviceName(128) As Byte
End Type

Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2&
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
  Bias As Long                  ' Basis-Zeitverschiebung in Minuten
  StandardName(1 To 64) As Byte ' Name der Sommerzeit-Zeitzone
  StandardDate As SYSTEMTIME    ' Beginn der Standardzeit
  StandardBias As Long          ' Zusätzliche Zeitverschiebung der Standardzeit
  DaylightName(1 To 64) As Byte ' Name der Sommerzeit-Zeitzone
  DaylightDate As SYSTEMTIME    ' Beginn der Sommerzeit
  DaylightBias As Long          ' Zusätzliche Zeitverschiebung der Sommerzeit
End Type

'ACHTUNG SYSTEMTIME IST UTC
Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub GetPageToFile(strUrl As String, sFileName As String)

  Call DownloadFromWeb(strUrl, sFileName)

'Dim b() As Byte
'
'On Error Resume Next
'
'If Not RequestSema(INET1_Sema) Then Exit Sub
'
''Cancel any operations
'frmBrowser.Inet1.Cancel
''Set protocol HTTP und wenn ich mal Zeit hab auch https
'frmBrowser.Inet1.protocol = icHTTP
''Set the URL
'frmBrowser.Inet1.URL = strURL
''Retrieve
'b() = frmBrowser.Inet1.OpenUrl(, icByteArray)
''Create file
'Open filename For Binary Access Write As #1
''und ab
'Put #1, , b()
'Close #1
'
'ReleaseSema INET1_Sema

End Sub
Public Function ShortPost(ByVal strUrl As String, Optional ByVal sPostData As String = "", Optional ByVal sReferer As String = "", Optional ByVal sEBayUser As String = "", Optional bWait As Boolean = True, Optional SkipCurl As Boolean = False) As String
Dim bOk As Boolean
Dim iFileNo As Integer
Dim sFollowToUrl As String
Dim iRetries As Integer

On Error Resume Next

If sEBayUser = "" Then sEBayUser = gsUser ' Fallback auf Standarduser

Dim sUrlOrg As String
Dim strTimeStart As String
Dim strTimeEnd As String
Dim dblTimeStart As Double
Dim dblTimeEnd As Double

'no more semas needed ..

If gbUsesModem Then
    bOk = CheckInternetConnection
    If Not bOk Then
        bOk = ModemConnect(mlGlobHandle)
    End If
End If

If Len(sPostData) > 0 Then
    If Right(strUrl, 1) = "?" Then strUrl = Left(strUrl, Len(strUrl) - 1)
End If

sUrlOrg = strUrl & IIf(sPostData <> "", "?" & sPostData, "")
DebugPrint "HttpRequest: " & sUrlOrg, 3
strTimeStart = GetDateTimeString()

Redirect:

strUrl = Trim(strUrl)

dblTimeStart = Timer
If gbUseCurl And SkipCurl = False Then
  ShortPost = Curl(strUrl, sPostData, sReferer, sEBayUser, bWait)
Else
  ShortPost = WinInetReadEx(strUrl, sPostData, sReferer, sEBayUser)
End If
dblTimeEnd = Timer

If ShortPost Like "Redirect:*" Then
  sReferer = strUrl
  strUrl = Mid(ShortPost, 10)
  sPostData = ""
  DebugPrint "HttpRedirect: " & strUrl, 3
  GoTo Redirect
End If

strTimeEnd = GetDateTimeString()

If Len(ShortPost) > 100 Then
  If dblTimeEnd < dblTimeStart Then dblTimeEnd = dblTimeEnd + (24 * 60 * 60)
  If (dblTimeEnd - dblTimeStart) > 0.1 Then
    gfLzMittel = (gfLzMittel + (dblTimeEnd - dblTimeStart)) / 2
    DebugPrint "Mittlere Laufzeit: " & Round(gfLzMittel, 1), 4
  End If
End If

iRetries = 2
Do While iRetries > 0 And (ShortPost = "" Or (InStr(1, ShortPost, "<HTML", vbTextCompare) > 0 And InStr(1, ShortPost, "</HTML", vbTextCompare) <= 0))
  'Wenn Ergebnis leer oder ein <HTML aber kein </HTML vorhanden, dann nochmal versuchen, lg 03.04.2004
  If gbUseCurl Then
    ShortPost = Curl(strUrl, sPostData, sReferer, sEBayUser, bWait)
  Else
    ShortPost = WinInetReadEx(strUrl, sPostData, sReferer, sEBayUser)
  End If
  iRetries = iRetries - 1
Loop

If gbLogHtml Then
  OpenLogfile
  iFileNo = FreeFile()
  Open gsAppDataPath & "\HtmlRequest.log" For Append As #iFileNo
    Print #iFileNo, strTimeStart & " - " & sUrlOrg
  Close #iFileNo
  iFileNo = FreeFile()
  Open gsAppDataPath & "\HtmlResponse.log" For Append As #iFileNo
    Print #iFileNo, strTimeStart & vbCrLf & ShortPost & vbCrLf & strTimeEnd & vbCrLf & "-----------------------------------------------"
  Close #iFileNo
  CloseLogfile
End If

sFollowToUrl = GetMetaHttpEquivRefresh(ShortPost)
If Len(sFollowToUrl) > 0 Then
  ShortPost = ShortPost(sFollowToUrl, "", strUrl, sEBayUser)
End If

End Function

Public Function GetPop() As Variant
'ein POP- SMTP- Protokoll
On Error Resume Next
Dim lMessages As Long
Dim lStart As Long
Dim lFinish As Long
Dim sTmp As String
Dim lPos As Long
Dim bWasEncrypted As Boolean
Dim sSubjDecrypted As String
Dim c As Long   ' counter, long is better ;-)
Dim vntSuba As Variant
Dim sSubj As String, sSender As String, sRecv As String
Dim col As Collection
Dim oRC4 As clsRC4
Dim oSSLWrapper As clsSSLWrapper

Set col = New Collection
Set GetPop = col

If Not RequestSema(gtTcpSema) Then Exit Function

gbFatalError = False

Call ClearTests
    
Set oSSLWrapper = New clsSSLWrapper
    
If gbPopUseSSL Then
    If LenB(gsPopCmdSSL) Then
        Call oSSLWrapper.StartSSLWrapper(gsPopCmdSSL, gbHideSSLWindow, gsPopServer, giPopPort, glSSLStartupDelay)
    End If
End If

If Not frmHaupt.tcpIn.State = sckClosed Then
    frmHaupt.tcpIn.Close
    Call WaitFor(mlCLOSED, gsOutText)
End If

Call DebugPrint("Connecting to POP server...", 2)

If gbPopUseSSL And gsPopCmdSSL > "" Then
    frmHaupt.tcpIn.Connect "127.0.0.1", giPopPort
Else
    frmHaupt.tcpIn.Connect gsPopServer, giPopPort
End If
    
If Err Then
    gbFatalError = True
    Err = 0
    Call ReleaseSema(gtTcpSema)
    Exit Function
End If
    
'On Error GoTo 0
    
Call WaitFor(mlSTATUS, gsOutText)

If gbFatalError Then
    Call ReleaseSema(gtTcpSema)
    Exit Function
End If
    
gbSessionClosed = False

If gsOutText = "+OK" Then
    Call Send("USER " & gsPopUser)  ' Usernamen senden
Else
    GoTo errExit
End If

Call WaitFor(mlSTATUS, gsOutText)
If gsOutText = "+OK" Then
    Call Send("PASS " & gsPopPass)  ' Passwort senden
Else
    GoTo errExit
End If

Call WaitFor(mlSTATUS, gsOutText)
If gsOutText = "+OK" Then ' Login ok
    Call Send("STAT")          ' Postfachstatistik abfragen
    Call WaitFor(mlSTATUS, gsOutText)
    If gsOutText <> "+OK" Then
        GoTo errExit
    End If

    lStart = InStr(gsWholeThing, " ") + 1
    lFinish = InStr(lStart, gsWholeThing, " ")
    lMessages = Val(Mid$(gsWholeThing, lStart, lFinish - lStart))
    
    If lMessages = 0 Then    'keine Nachrichten im Postfach
        GoTo errExit
    End If
    
    'Get size and header
    
    For c = 1 To lMessages ' Jede vorhandene Mail abarbeiten

        Call Send("TOP " & c & " 1")  ' Nachrichteninfo abrufen
    
        Call WaitFor(mlDOT, gsOutText)
            
        'nur das Subj interessiert!
        'Steuersequenz aus Mailer, alles im Subject
            
        lStart = InStr(LCase$(gsWholeThing), vbCrLf & "subject:") + 10
        lFinish = InStr(lStart, gsWholeThing, vbCrLf)
        Do While (Mid(gsWholeThing, lFinish + 2, 1) = " ")
            lFinish = InStr(lFinish + 2, gsWholeThing, vbCrLf)
        Loop
        sSubj = Mid$(gsWholeThing, lStart, lFinish - lStart)
            
        vntSuba = Split(sSubj, vbCrLf)
        sSubj = ""
        For lPos = LBound(vntSuba) To UBound(vntSuba)
            sSubj = sSubj & frmHaupt.SMTP_1.DecodeHeader(Mid(vntSuba(lPos), 2))
        Next
            
        bWasEncrypted = False
            
        lStart = InStr(LCase$(sSubj), "biet-o-matic:")
        If lStart = 0 Then
            lStart = InStr(LCase$(sSubj), "b-o-m:")
        End If
            
        If lStart = 0 Then ' nochmal mit Entschlüsseln probieren
            Set oRC4 = New clsRC4
            sSubjDecrypted = oRC4.DecryptString(sSubj, gsPass, True)
            Set oRC4 = Nothing
            
            lStart = InStr(LCase$(sSubjDecrypted), "biet-o-matic:")
            If lStart = 0 Then
                lStart = InStr(LCase$(sSubjDecrypted), "b-o-m:")
            End If
              
              If lStart > 0 Then ' okay, jetzt in das normale Subj übertragen
                  sSubj = sSubjDecrypted
                  bWasEncrypted = True
              End If
        End If
            
        If gbPopEncryptedOnly And Not bWasEncrypted Then
            sSubj = "b-o-m: encryption_needed"
        End If
                    
        'Ist der Befehl biet-o-matic: oder b-o-m: im Subjekt
        '
        If lStart > 0 Then ' Wenn ja, dann abarbeiten
            
            sSubj = sSubj & "  " 'blank dran ist wichtig!
                
            'Sender lesen, für Erlaubnis und Quittungsmail
            lStart = InStr(LCase$(gsWholeThing), vbLf & "from:") + 6
            lFinish = InStr(lStart, gsWholeThing, vbCrLf)
            sSender = Trim(Mid$(gsWholeThing, lStart, lFinish - lStart))
                
            If TesteAbsender(sSender) Then 'ok, erlaubte Adresse :-)
                
                'mal sehen, ob es ein "delivered" ist ..
                lStart = InStr(LCase$(gsWholeThing), vbLf & "delivered-to:")
                If lStart = 0 Then
                    lStart = InStr(LCase$(gsWholeThing), vbLf & "to:") + 4
                End If
                'Empfänger auslesen
                lFinish = InStr(lStart, gsWholeThing, vbCrLf)
                sRecv = Trim(Mid$(gsWholeThing, lStart, lFinish - lStart))
                    
                'mal sehen, ob noch irgendwelche Infos im recv stehen ..
                lPos = 1
                lFinish = InStr(1, sRecv, "@")
                lStart = InStr(1, sRecv, " ")
                While lStart > 0 And lStart < lFinish
                    lPos = lStart + 1
                    lStart = InStr(lPos, sRecv, " ")
                Wend
                sRecv = Trim(Mid$(sRecv, lPos))
                'mal sehen, ob wir den Body lesen müssen ..
                    
                'vorher stand da artikel?
                If InStr(LCase$(sSubj), "readcsv") Then 'Steuerkommand readcsv
                    lPos = Len(gsWholeThing) - 5
                        
                    Call Send("RETR " & c)  ' ganze Nachricht abrufen
                    Call WaitFor(mlDOT, gsOutText)
                    gsWholeThing = frmHaupt.SMTP_1.Decode_qp(gsWholeThing)
                    lPos = InStr(1, gsWholeThing, vbCrLf & vbCrLf)
                    lFinish = InStr(lPos, gsWholeThing, "-- " & vbCrLf)
                        
                    If lFinish = 0 Or lFinish <= lPos Then ' Nachrichtentext an die Subjektzeile anhängen
                        sSubj = sSubj & vbCrLf & Mid$(gsWholeThing, lPos)
                    Else
                        sSubj = sSubj & vbCrLf & Mid$(gsWholeThing, lPos, lFinish - lPos)
                    End If
                End If
                'und Message rauswerfen, um sie nicht doppelt zu haben
                Call Send("DELE " & c)
                Call WaitFor(mlSTATUS, gsOutText)
                col.Add Array(sSubj, sSender, sRecv)
            End If
        Else
            '
            'Steuersequenz aus "Artikel an Freund"
            '
            sTmp = frmHaupt.SMTP_1.DecodeHeader(CStr(gsWholeThing))
            sTmp = Replace(sTmp, vbCrLf, "")
            sTmp = Replace(sTmp, vbCr, "")
            sTmp = Replace(sTmp, vbLf, "")
            sTmp = Replace(sTmp, "_", "")
            sTmp = Replace(sTmp, " ", "")
                
            lStart = InStr(1, sTmp, gsAnsMailToFriend)
            If lStart > 0 Then
                'Sender lesen, für Erlaubnis und Quittungsmail
                lStart = InStr(1, LCase$(gsWholeThing), vbLf & "from:") + 6
                lFinish = InStr(lStart, gsWholeThing, vbCrLf)
                sSender = Trim(Mid$(gsWholeThing, lStart, lFinish - lStart))
                
                If sSender Like "*" & gsAnsMailToFriendAddressStart & "*" & gsAnsMailToFriendAddressEnd & "*" Then
                    sSender = Trim(Mid(sSender, 1 + InStr(1, sSender, gsAnsMailToFriendAddressStart) + 1, Len(sSender) - 1 - InStr(1, sSender, gsAnsMailToFriendAddressStart) - Len(sSender) + InStr(1, sSender, gsAnsMailToFriendAddressEnd) - 1))
                End If
                
                If TesteAbsender(sSender) Then 'ok, erlaubte Adresse :-)
                        
                    lStart = InStr(LCase$(gsWholeThing), vbLf & "to:") + 4 'mit vbLf prüfen wegen Zeilenanfang (gmx hat z.B. ein Delivered-To: GMX delivery to hans@gmx.de)
                    lFinish = InStr(lStart, gsWholeThing, vbCrLf)
                    'Empfänger auslesen
                    sRecv = Trim(Mid$(gsWholeThing, lStart, lFinish - lStart))
                        
                    Call DebugPrint("Send RETR, OldLen= " & Len(gsWholeThing), 2)

                    Call Send("RETR " & c)  ' ganze Nachricht abrufen
                    Call WaitFor(mlDOT, gsOutText)
                    gsWholeThing = frmHaupt.SMTP_1.Decode_qp(gsWholeThing)
                        
                    Call DebugPrint("RETR done, Quit=" & gsOutText & " AktLen= " & Len(gsWholeThing), 2)
                        
                    lPos = lFinish ' Finish merken
                    'den Subject- string müssen wir uns selbst zusammenbasteln ..
                    'If InStr(1, sSubj, "=?ISO", vbTextCompare) > 0 Then sSubj = frmHaupt.SMTP_1.DecodeHeader(Replace(sSubj, " ", ""))
                    'lStart = InStr(1, sSubj, gsAnsMailToFriendItemStart) + Len(gsAnsMailToFriendItemStart)
                    lStart = InStr(1, gsWholeThing, gsAnsMailToFriendItemStart) + Len(gsAnsMailToFriendItemStart)
                    If lStart > Len(gsAnsMailToFriendItemStart) Then
                        'lFinish = InStr(lStart, sSubj, gsAnsMailToFriendItemEnd)
                        'sTmp = Mid$(sSubj, lStart, lFinish - lStart)
                        lFinish = InStr(lStart, gsWholeThing, gsAnsMailToFriendItemEnd)
                        sTmp = Val(Mid$(gsWholeThing, lStart, lFinish - lStart))
                        sSubj = "art" & sTmp & " "
                    End If
                        
                    lFinish = lPos ' Finish wiederherstellen
                        
                    'Keyword Biet-O-Matic suchen
                    lStart = InStr(lFinish, LCase$(gsWholeThing), "biet-o-matic:") + 13
                    If lStart = 13 Then
                        lStart = InStr(lFinish, LCase$(gsWholeThing), "b-o-m:") + 6
                        If lStart = 6 Then lStart = 1
                    End If
                        
                    lFinish = InStr(lStart, gsWholeThing, vbCrLf)
                    If lFinish = 0 Then lFinish = Len(gsWholeThing)
                    If InStr(lStart, gsWholeThing, "<", vbTextCompare) < lFinish Then
                        lFinish = InStr(lStart, gsWholeThing, "<", vbTextCompare)
                    End If
                        
                    Call DebugPrint("Artikel= " & sSubj & " Keyword b-o-m?" & lStart - 1, 2)
                        
                    If lStart > 1 Then
                        'es ist für mich :-)
                        sTmp = Mid$(gsWholeThing, lStart, lFinish - lStart)
                        sSubj = sSubj & sTmp & "  " 'blank dran ist wichtig!
                        If gbPopEncryptedOnly Then sSubj = "b-o-m: encryption_needed"
                            
                        'und Message rauswerfen, um sie nicht doppelt zu haben
                        Call Send("DELE " & c)
                        Call WaitFor(mlSTATUS, gsOutText)
                        col.Add Array(sSubj, sSender, sRecv)
                    End If
                End If
            End If
        End If 'Steuersequenz aus Artikel
    Next
End If
    
errExit:

Call PopDisconnect
Call ReleaseSema(gtTcpSema)

End Function
Private Sub WaitFor(lMode As Long, sTestFor As String)

'Schweinisches Quasi- Warten auf die quittung

sTestFor = "-ER"

'ohne Timeouthandling kann es fürchterlich knallen, wenn eine Msg verlorengeht!
gbTimeOutOccurs = False
giPopTimeOutCount = 0
giTimeOutTimerTimeOut = giPopTimeOut
frmHaupt.TimeoutTimer.Enabled = True

    Do
        Select Case lMode
            Case mlSTATUS
                If gsResponseState = "+OK" Or gsResponseState = "-ER" Then
                    gsStatusTxt = gsThisChunk 'show status
                    sTestFor = gsResponseState
                    Exit Do
                End If
            Case mlDOT
                If gsResponseState = "-ER" Then  'if got here by mistake
                    gsStatusTxt = gsThisChunk 'show status
                    sTestFor = gsResponseState
                    Exit Do
                End If
                If gsDotLine = msEOM Then
                    Exit Do
                End If
            Case mlCONNECTED
                If frmHaupt.tcpIn.State = sckConnected Then
                    sTestFor = "+OK"
                    Exit Do

                End If
            Case mlNTPDATA
                If Len(gsWholeThing) >= 4 Then
                    gsStatusTxt = gsThisChunk 'show status
                    sTestFor = gsWholeThing
                    Exit Do
                End If
            
            Case mlSMTPDATA
                If giSmtpResponse > 0 Then
                    gsStatusTxt = gsThisChunk 'show status
                    sTestFor = gsResponseState
                    Exit Do
                End If
            
            Case mlCLOSED
                If gbSessionClosed Then
                    sTestFor = "+OK"
                    Exit Do
                End If
        End Select

        If gbFatalError Then
            frmHaupt.tcpIn.Close
            Exit Do
        End If

        DoEvents
        Call Sleep(10)
        If gbTimeOutOccurs Then
            gsStatusTxt = gsStatusTxt & "Timeout bei WaitFor"
            Exit Do
        End If
    Loop
    
        frmHaupt.TimeoutTimer.Enabled = False

End Sub

Private Sub ClearTests()
    
    gsResponseState = ""
    giSmtpResponse = 0
    gsThisChunk = ""
    gsWholeThing = ""
    gsDotLine = ""
    Err.Clear
    
End Sub

Private Sub Send(sTextOut As String)
    
    On Error Resume Next
    
    Call ClearTests
    
    frmHaupt.tcpIn.SendData sTextOut & vbCrLf
    On Error GoTo 0
    
End Sub

Public Function PopTest() As Boolean
    
    On Error Resume Next
    Dim sTmp As String
    Dim oSSL As clsSSLWrapper
    
    PopTest = False
    
    If Not RequestSema(gtTcpSema) Then Exit Function
    
    gbFatalError = False
    
    Call ClearTests
    
    Set oSSL = New clsSSLWrapper
    
    If gbPopUseSSL Then
        If LenB(gsPopCmdSSL) Then
            Call oSSL.StartSSLWrapper(gsPopCmdSSL, gbHideSSLWindow, gsPopServer, giPopPort, glSSLStartupDelay)
        End If
    End If
    
    If Not frmHaupt.tcpIn.State = sckClosed Then
        frmHaupt.tcpIn.Close
        Call WaitFor(mlCLOSED, gsOutText)
    End If
    
    If gbPopUseSSL And gsPopCmdSSL > "" Then
        frmHaupt.tcpIn.Connect "127.0.0.1", giPopPort
    Else
        frmHaupt.tcpIn.Connect gsPopServer, giPopPort
    End If
    
    If Err Then
        gbFatalError = True
        Err = 0
        GoTo errhdl
    End If
    gbSessionClosed = False
    
    On Error GoTo 0
    Call WaitFor(mlSTATUS, gsOutText)
    
    sTmp = gsStatusTxt
    
    If gbFatalError Then GoTo errhdl
    
    If gsOutText = "+OK" Then
        Call Send("USER " & gsPopUser)
    Else
        GoTo errhdl
    End If
    
    Call WaitFor(mlSTATUS, gsOutText)
    sTmp = sTmp & gsStatusTxt
    
    If gsOutText = "+OK" Then
        Call Send("PASS " & gsPopPass)
    Else
        GoTo errhdl
    End If
    
    Call WaitFor(mlSTATUS, gsOutText)
    sTmp = sTmp & gsStatusTxt
    
    If gsOutText = "+OK" Then PopTest = True

    
errhdl:

    Call PopDisconnect
    gsStatusTxt = sTmp
    
    Call ReleaseSema(gtTcpSema)
    
End Function
Private Sub PopDisconnect()
    
    Call Send("QUIT")
    Call WaitFor(mlCLOSED, gsOutText)
    
End Sub
' Testonly erweitert, lg 29.05.03
Public Function SendSMTP(ByVal sSender As String, ByVal sRecv As String, ByVal sMsg As String, Optional bDebugMode As Boolean = False, Optional bTestOnly As Boolean = False) As Boolean
Dim iEnde As Integer
Dim iPos As Integer
Dim sSenderName As String
Dim sRecvName As String
Dim oSSLWrapper As clsSSLWrapper
Dim vaRecvs As Variant
Dim vRecv As Variant

SendSMTP = False

Set oSSLWrapper = New clsSSLWrapper

If gbSmtpUseSSL Then
    If LenB(gsSmtpCmdSSL) Then
        Call oSSLWrapper.StartSSLWrapper(gsSmtpCmdSSL, gbHideSSLWindow, gsSmtpServer, giSmtpPort, glSSLStartupDelay)
    End If
End If

'ggf. Sender und Empfänger aufdröseln
iPos = InStr(1, sSender, "<")
If iPos > 0 Then
    iEnde = InStr(iPos, sSender, ">") - 1
    sSenderName = Trim(Left(sSender, iPos - 1))
    sSender = Mid(sSender, iPos + 1, iEnde - iPos)
    If Left(sSenderName, 1) = """" And Right(sSenderName, 1) = """" Then sSenderName = Trim(Mid(sSenderName, 2, Len(sSenderName) - 2))
    If Left(sSenderName, 1) = "'" And Right(sSenderName, 1) = "'" Then sSenderName = Trim(Mid(sSenderName, 2, Len(sSenderName) - 2))
End If

vaRecvs = Split(sRecv, ";")
For Each vRecv In vaRecvs

    sRecv = vRecv
    sRecvName = ""

    iPos = InStr(1, sRecv, "<")
    If iPos > 0 Then
        iEnde = InStr(iPos, sRecv, ">") - 1
        sRecvName = Trim(Left(sRecv, iPos - 1))
        sRecv = Mid(sRecv, iPos + 1, iEnde - iPos)
        If Left(sRecvName, 1) = """" And Right(sRecvName, 1) = """" Then sRecvName = Trim(Mid(sRecvName, 2, Len(sRecvName) - 2))
        If Left(sRecvName, 1) = "'" And Right(sRecvName, 1) = "'" Then sRecvName = Trim(Mid(sRecvName, 2, Len(sRecvName) - 2))
    End If
    
    '
    ' ersetzt durch SMTP- AUTH- Control
    ' thx to Ingo
    '
    With frmHaupt.SMTP_1
    
        .SMTPDebugMode = bDebugMode                'zumindest erstmal
        
        .MailRecipientName = sRecvName             ' "otto "
        .MailRecipientEMail = sRecv                ' "otto@web.de"
        
        .MailSenderName = sSenderName              ' "Alfred E Neumann"
        .MailSenderEMail = sSender                 ' "alfred.e.neumann@mad.tv"
        
        '.MailSubject = txt_subj.Text             ' Subject wird mit in den Body verpackt ..
        .MailBody = sMsg                           ' Mit Subject, 2x LF + Body
        
        .SMTPPort = giSmtpPort                      ' "25"
        If gbSmtpUseSSL And gsSmtpCmdSSL > "" Then
            .SMTPServer = "127.0.0.1"
        Else
            .SMTPServer = gsSmtpServer                  ' "mx.freenet.de"
        End If
        
        If gbUseSmtpAuth Then
            .SMTPAuthPass = gsPopPass                 ' "geheim"
            .SMTPAuthUser = gsPopUser                 ' "otto@freenet.de"
        Else
            .SMTPAuthPass = ""                      ' leer
            .SMTPAuthUser = ""                      ' leer
        End If
        
        SendSMTP = .SendMail(bTestOnly)           ' Und wech
        
    End With
    
Next

'
' Hard exit, rest = alte Proc ;-)
'
Exit Function

'Achtung: ggf. erst POP3 für Pop-before-SMTP


'If Not RequestSema(gtTcpSema) Then Exit Function
'
'gbFatalError = False
'
'str = ""
'txtStatus = ""
'
'
'    iPos = InStr(1, sSender, "@")
'    If iPos > 0 Then
'        domain = Mid$(sRecv, iPos + 1)
'    Else
'        domain = sRecv
'    End If
'    aktDate = "Date: " & Format(MyNow, "dd mmm yyyy hh:nn:ss ") & "+0200"
'
'    Select Case Mid$(aktDate, 10, 3)
'        Case "Mär"
'            Mid(aktDate, 10, 3) = "Mar"
'        Case "Mai"
'            Mid(aktDate, 10, 3) = "May"
'        Case "Okt"
'            Mid(aktDate, 10, 3) = "Oct"
'        Case "Dez"
'            Mid(aktDate, 10, 3) = "Dec"
'
'    End Select
'    ClearTests
'
'    On Error Resume Next
'
'    If Not frmHaupt.tcpIn.State = sckClosed Then
'        frmHaupt.tcpIn.Close
'        WaitFor mlCLOSED, gsOutText
'    End If
'
'    ClearTests
'    frmHaupt.tcpIn.Connect gsSmtpServer, 25
'
'    If Err Then
'        str = "ConnectError" & txtStatus
'        gbFatalError = True
'        Err = 0
'        GoTo errhdl
'    End If
'
'    'WaitFor mlCONNECTED, gsOutText
'    'If gsOutText = "-ER" Then
'    '    Exit Function
'    'End If
'    'On Error GoTo 0
'
'    WaitFor mlSMTPDATA, gsOutText
'
'    str = str & txtStatus
'
'    If gbFatalError Then
'          GoTo errhdl
'    End If
'
'    gbSessionClosed = False
'
'    If gsOutText = "220" Then
'        Send "HELO " + domain
'    Else
'        GoTo errhdl
'    End If
'
'    WaitFor mlSMTPDATA, gsOutText
'    str = str & txtStatus
'
'    If gsOutText = "250" Then
'        Send "MAIL FROM:<" + sSender + ">"
'    Else
'        GoTo errhdl
'    End If
'
'    WaitFor mlSMTPDATA, gsOutText
'    str = str & txtStatus
'
'    If gsOutText = "250" Then
'        Send "RCPT TO:" + sRecv
'    Else
'        GoTo errhdl
'    End If
'
'    WaitFor mlSMTPDATA, gsOutText
'    str = str & txtStatus
'
'    If gsOutText = "250" Then
'        Send "DATA"
'    Else
'        GoTo errhdl
'    End If
'
'    WaitFor mlSMTPDATA, gsOutText
'    str = str & txtStatus
'
'    If gsOutText = "354" Then
'        sMsg = "From: <" & sSender & ">" & vbCrLf & "To: <" & sRecv & ">" & vbCrLf & aktDate & vbCrLf & sMsg
'        Send sMsg & vbCrLf & "."
'    Else
'        GoTo errhdl
'    End If
'
'    WaitFor mlSMTPDATA, gsOutText
'    str = str & txtStatus
'
'    SendSMTP = True
'
'errhdl:
'
'    PopDisconnect
'    txtStatus = str
'    ReleaseSema gtTcpSema
'
End Function

'Modem- Procs

Public Function ModemConnect(lHandle As Long) As Boolean
    
    Dim rc As Long
    
    mlGlobHandle = lHandle
    
    If Not gsConnectName = "--" Then
        If Not CheckInternetConnection Then
            glConnectID = 0
            Call DebugPrint("Dialup")
            rc = InternetDial(lHandle, gsConnectName, DIAL_FORCE_UNATTENDED, glConnectID, 0)
            ModemConnect = CBool(rc = 0 And glConnectID <> 0)
        Else
            ModemConnect = True
        End If
    Else
        ModemConnect = CheckInternetConnection
    End If
    
End Function
Public Sub ModemHangUp()

If CheckInternetConnection Then
    gbLastDialupWasManually = False
    If glConnectID Then Call InternetHangUp(glConnectID, 0): Call DebugPrint("Hangup")
    glConnectID = 0
End If
End Sub


Public Function CheckInternetConnection() As Boolean
Dim lpRasConn(255) As RASCONN
Dim lpcConnections As Long
Dim lpcb As Long
Dim result As Long

On Error GoTo errhdl

CheckInternetConnection = False

If Not gbUsesModem Then
    CheckInternetConnection = True
Else
    lpRasConn(0).dwSize = 412
    lpcb = 256 * lpRasConn(0).dwSize

    result = RasEnumConnections(lpRasConn(0), lpcb, lpcConnections)
    
    If lpcConnections < 1 Then
        'es besteht keine DFÜ-Verbindung
    Else
        'es besteht mind. 1 DFÜ-Verbindung
        CheckInternetConnection = True
    End If
End If

Exit Function

errhdl:
    MsgBox "RAS / DFUE- Netzwerk nicht installiert"
    gbUsesModem = False
End Function
Public Sub GetDFUEList()
Dim s As Long, LN As Long, X As Integer
Dim r(255) As RASENTRYNAME95
Dim ConName As String

'Namen der bestehenden DFÜ-Verbindungen einlesen
r(0).dwSize = 264
s = 256 * r(0).dwSize

'macht heftig Probleme wenn kein RAS installiert .. bis zum Rechnercrash
On Error GoTo errhdl

gbUsesModem = True
Call CheckInternetConnection
If Not gbUsesModem Then Exit Sub

Call RasEnumEntries(vbNullString, vbNullString, r(0), s, LN)

If LN <> 0 Then
   frmSettings.lstDfue.Clear
   'Es besteht mindestens eine DFÜ-Verbindung
   For X = 0 To LN - 1
       ConName = StrConv(r(X).szEntryName(), vbUnicode)
       frmSettings.lstDfue.AddItem Left$(ConName, InStr(ConName, vbNullChar) - 1)
   Next X
   frmSettings.lstDfue.ListIndex = 0
Else
   'Keine DFÜ da
   MsgBox ("Keine DFÜ-Verbindung vorhanden")
End If

Exit Sub

errhdl:
   MsgBox ("Fehler beim Lesen der DFÜ-Verbindungen: " & Err.Description)

End Sub



Private Function TesteAbsender(ByVal sMailAdr As String) As Boolean
    
    On Error Resume Next
    Dim lPos As Long
    Dim lPosEnde As Long
    Dim v As Variant
    Dim va As Variant


    Call DebugPrint("Teste Absender: " & sMailAdr, 2)

    TesteAbsender = False
    gsAbsender = Replace(LCase(gsAbsender), " ", "")
    gsAbsender = Replace(LCase(gsAbsender), ",", ";")
    
    If gsAbsender = "" Then
        Call DebugPrint("Teste Absender: keine Prüfung " & sMailAdr, 2)
        TesteAbsender = True
        Exit Function
    Else
        Call DebugPrint("Teste Absender: Erlaubt=#" & gsAbsender & "#", 2)
    End If
    
    'mal sehen, ob wir die MailAddy rausfischen müssen ..
    
    sMailAdr = Trim(LCase(sMailAdr))
    lPos = InStrRev(sMailAdr, ">")
    If lPos > 0 Then
        lPosEnde = InStrRev(sMailAdr, "<", lPos)
        sMailAdr = Mid(sMailAdr, lPosEnde + 1, lPos - lPosEnde - 1)
    End If
    
    sMailAdr = Trim(LCase(sMailAdr))
    If Len(sMailAdr) < 2 Or Not sMailAdr Like "*?@?*.?*" Then
        TesteAbsender = False
        Exit Function
    End If
    
    va = Split(gsAbsender, ";")
    For Each v In va
        If sMailAdr Like v Then TesteAbsender = True
    Next
    
    If TesteAbsender Then
        Call DebugPrint("Absender erkannt. " & sMailAdr, 2)
    Else
        Call DebugPrint("Falscher Absender " & sMailAdr & " erlaubt:#" & gsAbsender & "#", 2)
    End If
    
End Function

Public Function GetINetTime() As String

  If giUseNtp = 1 Then GetINetTime = GetINetTimeTime()
  If giUseNtp = 2 Then GetINetTime = GetINetTimeSntp()

End Function

Private Function GetINetTimeSntp() As String
    
    Dim bSuccess As Boolean
    
    frmHaupt.ctlSNTP1.TimeServer = GetServerFromServer(gsNtpServer)
    If frmHaupt.ctlSNTP1.SyncTime() Then
        bSuccess = True
        gfTimeDeviation = 0
    ElseIf frmHaupt.ctlSNTP1.LastError = "No permission to set time" Then
        bSuccess = True
        gfTimeDeviation = frmHaupt.ctlSNTP1.LastLapse / 1000
        Call DebugPrint("Systemzeit konnte nicht geändert werden.", 2)
    End If
    
    If bSuccess Then GetINetTimeSntp = Date2Str(MyNow)
    
End Function

Private Function GetINetTimeTime() As String
    
    On Error Resume Next

    Dim fNtpTime As Double
    Dim LngTimeFrom1990 As Long
    Dim datUtcDate As Date
    Dim lTimeDelay As Long
    Dim ST As SYSTEMTIME
        
    
    GetINetTimeTime = ""
    giNtpErr = 0
    gsNtpData = ""
    
    If Not frmHaupt.NTP.State = sckClosed Then
        frmHaupt.NTP.Close
        Call NTPWaitFor(mlCLOSED, gsOutText)
    End If
    
    Err.Clear
    frmHaupt.NTP.RemoteHost = GetServerFromServer(gsNtpServer)
    frmHaupt.NTP.RemotePort = GetPortFromServer(gsNtpServer)
    If frmHaupt.NTP.RemotePort = 0 Then frmHaupt.NTP.RemotePort = 37
    frmHaupt.NTP.Connect
    
    If Err Then
        giNtpErr = 1
        GoTo errhdl
    End If
        
    Call NTPWaitFor(mlNTPDATA, gsOutText)
    
       
    If giNtpErr Then GoTo errhdl
    
    lTimeDelay = CLng((Timer - gfNtpDelay) / 2)
    
    If Len(gsOutText) = 4 Then
        
        fNtpTime = Asc(Left$(gsOutText, 1)) * 256 ^ 3 + _
            Asc(Mid$(gsOutText, 2, 1)) * 256 ^ 2 + _
            Asc(Mid$(gsOutText, 3, 1)) * 256 ^ 1 + _
            Asc(Right$(gsOutText, 1))
        
        On Error GoTo errhdl
        
        LngTimeFrom1990 = fNtpTime - 2840140800#
        
        datUtcDate = DateAdd("s", CDbl(LngTimeFrom1990 + lTimeDelay), #1/1/1990#)
        
        'datUtcDate = DateAdd("h", NTPOffset, datUtcDate) 'lg 12.05.2003
        'und in die Syszeit- Variablen einbauen:
        ST.wYear = Year(datUtcDate)
        ST.wMonth = Month(datUtcDate)
        ST.wDay = Day(datUtcDate)
        ST.wHour = Hour(datUtcDate)
        ST.wMinute = Minute(datUtcDate)
        ST.wSecond = Second(datUtcDate)
        
        If SetSystemTime(ST) Then
            gfTimeDeviation = 0
        Else
            gfTimeDeviation = lTimeDelay
            Call DebugPrint("Systemzeit konnte nicht geändert werden.", 2)
        End If
        GetINetTimeTime = Date2Str(MyNow)
        
    End If
    
errhdl:
    
    frmHaupt.NTP.Close
    Call NTPWaitFor(mlCLOSED, gsOutText)
End Function

Public Function URLDecode(strData As String) As String

Dim strTemp As String
Dim lPos As Long

strTemp = Trim(strData)

lPos = InStr(1, strTemp, "%")
Do While lPos > 0

  strTemp = Left(strTemp, lPos - 1) & Chr(CByte("&h" & Mid(strTemp, lPos + 1, 2))) & Mid(strTemp, lPos + 3)
  lPos = InStr(lPos + 1, strTemp, "%")

Loop

URLDecode = strTemp

End Function

Public Function URLEncode(strData As String) As String

Dim i As Integer
Dim strTemp As String
Dim strChar As String
Dim strOut As String
Dim intAsc As Integer

strTemp = Trim(strData)

For i = 1 To Len(strTemp)
   strChar = Mid(strTemp, i, 1)
   intAsc = Asc(strChar)
   If (intAsc >= 42 And intAsc <= 42) Or _
      (intAsc >= 45 And intAsc <= 46) Or _
      (intAsc >= 48 And intAsc <= 57) Or _
      (intAsc >= 97 And intAsc <= 122) Or _
      (intAsc >= 65 And intAsc <= 90) Then
      strOut = strOut & strChar
   ElseIf intAsc = 32 Then
      strOut = strOut & "+"
   Else
      strOut = strOut & "%" & Hex(intAsc)
   End If
Next i

URLEncode = strOut

End Function

Public Function Encode_UTF8(ByVal astr$) As String

  Dim n As Long
  Dim c As Long
  Dim utftext As String
  utftext = ""

  For n = 1 To Len(astr$)
    c = AscW(Mid(astr$, n, 1))
    If c < 128 Then
        utftext = utftext + Mid(astr$, n, 1)
    ElseIf ((c > 127) And (c < 2048)) Then
        utftext = utftext + Chr(((c \ 64) Or 192))              '((c>>6)|192);
        utftext = utftext + Chr(((c And 63) Or 128))            '((c&63)|128);}
    Else
        utftext = utftext + Chr(((c \ 4096) Or 224))            '((c>>12)|224);
        utftext = utftext + Chr((((c \ 64) And 63) Or 128))     '(((c>>6)&63)|128);
        utftext = utftext + Chr(((c And 63) Or 128))            '((c&63)|128);
    End If
  Next n

  Encode_UTF8 = utftext

End Function

Public Function Decode_UTF8(b() As Byte) As Byte()
  
  Dim m As Long
  Dim n As Long
  Dim c As Long
  Dim X As Long
  Dim Y As Long
  Dim Z As Long
  Dim b2() As Byte
  Dim s As Long
  Dim t As String
  
  On Error GoTo ERROR_HANDLER
  
  s = UBound(b) + 1
  ReDim b2(0 To s * 2)
  
  m = 0
  n = 0
  Do While (n < s)
  
    Z = b(n)
    n = n + 1
    
    If Z < 128 Then ' nur 1 Zeichen
      c = Z
    ElseIf Z < 224 Then ' 2 Zeichen
      Y = b(n)
      n = n + 1
      If (Z >= 192 And Y >= 128 And Y <= 191) Then ' 2 Zeichen
        c = (Z - 192) * 64 + (Y - 128)
      Else ' ungültig
        b2(m) = Z
        m = m + 2
        c = Y
      End If
    Else ' 3 Zeichen
      Y = b(n)
      n = n + 1
      X = b(n)
      n = n + 1
      If (Z >= 224 And Y >= 128 And Y <= 191 And X >= 128 And X <= 191) Then ' 3 Zeichen
        c = (Z - 224) * 4096 + (Y - 128) * 64 + (X - 128)
      Else ' ungültig
        b2(m) = Z
        m = m + 2
        b2(m) = Y
        m = m + 2
        c = X
      End If
    End If
    
    If c < 256 Then
      t = Chr(c)
    ElseIf c < 65536 Then
      t = ChrW(c)
    Else
      t = "?"
    End If
    CopyMemory ByVal VarPtr(b2(m)), ByVal StrPtr(t), 2
    m = m + 2
  Loop
  
  ReDim Preserve b2(0 To m - 1)
  Decode_UTF8 = b2

Exit Function
ERROR_HANDLER:
  MsgBox "Error: " & Err.Description, vbCritical
'  MsgBox "m: " & m & ", n: " & n & ", c: " & c & ", X: " & X & ", Y: " & Y & ", Z: " & Z & ", l: " & l

End Function

Public Function String2ByteArray(ByVal sTxt As String) As Byte()
    
    Dim s As Long
    Dim b() As Byte
    
    s = LenB(sTxt)
    
    ReDim b(0 To s - 1) As Byte
   
    Call CopyMemory(ByVal VarPtr(b(0)), ByVal StrPtr(sTxt), s)
    
    String2ByteArray = b()
    Erase b()
    
End Function

Public Function ByteArray2String(b() As Byte) As String
    
    Dim s As Long
    Dim sTmp As String
    
    s = UBound(b) - LBound(b) + 1
    
    sTmp = String(s / 2, " ")
    
    Call CopyMemory(ByVal StrPtr(sTmp), ByVal VarPtr(b(0)), s)
    
    ByteArray2String = sTmp

End Function

Private Sub NTPWaitFor(bytMode As Byte, sTestFor As String)

    'Schweinisches Quasi- Warten auf die quittung
    sTestFor = "-ER"
    
    'ohne Timeouthandling kann es fürchterlich knallen, wenn eine Msg verlorengeht!
    gbTimeOutOccurs = False
    giPopTimeOutCount = 0
    giTimeOutTimerTimeOut = 3
    frmHaupt.TimeoutTimer.Enabled = True
    
    Do
        Select Case bytMode
            Case mlNTPDATA
                If Len(gsNtpData) >= 4 Then
                    sTestFor = gsNtpData
                    Exit Do
                End If
                
            Case mlCLOSED
                If frmHaupt.NTP.State = sckClosed Then
                    sTestFor = "+OK"
                    Exit Do
                End If
        End Select
        
        If giNtpErr Then
            frmHaupt.NTP.Close
            Exit Do
        End If
        
        DoEvents
        Call Sleep(10)
        If gbTimeOutOccurs Then
            Exit Do
        End If
    Loop
    
    frmHaupt.TimeoutTimer.Enabled = False
    
End Sub

Private Function DownloadFromWeb(ByVal strUrl As String, ByVal sSaveFilePathName As String) As Long
    On Error Resume Next
    DownloadFromWeb = URLDownloadToFile(0, strUrl, sSaveFilePathName, 0, 0)
End Function

Public Function GetUTCOffset() As Double
    
    Dim udtTZI As TIME_ZONE_INFORMATION
    
    If GetTimeZoneInformation(udtTZI) = TIME_ZONE_ID_DAYLIGHT Then
        GetUTCOffset = (udtTZI.Bias + udtTZI.DaylightBias) / -60
    Else
        GetUTCOffset = (udtTZI.Bias + udtTZI.StandardBias) / -60
    End If
    
End Function

Private Function Decode_qp(ByVal sTxt As String) As String
    
    Const CODIERUNG_QUOTED_PRINTABLE As String = "=[?]ISO-8859-1[?]Q[?]*[?]="
    
    Dim sTmp As String
    Dim lPos As Long
    
    sTmp = sTxt
    If Trim(sTxt) Like CODIERUNG_QUOTED_PRINTABLE Then
        
        sTmp = Trim(sTmp)
        sTmp = Mid(sTmp, 16)
        sTmp = Left(sTmp, Len(sTmp) - 2)
        sTmp = Replace(sTmp, "_", " ")
        lPos = InStr(lPos + 1, sTmp, "=")
        Do While lPos > 0
            sTmp = Left(sTmp, lPos - 1) & GetChar(Mid(sTmp, lPos + 1, 2)) & Mid(sTmp, lPos + 3)
            lPos = InStr(lPos + 1, sTmp, "=")
        Loop
    
    End If
    Decode_qp = sTmp
    
End Function

Public Function GetChar(ByVal sHexCode As String) As String

  On Error GoTo ERROR_HANDLER
  
  If UCase(sHexCode) = "A4" Then sHexCode = "80"
  If sHexCode = vbCrLf Then
    GetChar = ""
  Else
    GetChar = Chr(CByte("&h" & sHexCode))
  End If
  
  Exit Function
ERROR_HANDLER:

End Function

Public Function GetLinkNamedLike(sTxt As String, sLinkName As String) As String

  On Error GoTo ERROR_HANDLER
  Dim lPos As Long
  Dim lPos2 As Long
  Dim lPos3 As Long
  Dim sLink As String
  
  lPos = InStr(1, sTxt, "a href=""", vbTextCompare)
  Do While lPos > 0
  
    lPos = lPos + 8
    lPos2 = InStr(lPos, sTxt, """")
    If lPos2 > 0 Then
      sLink = Mid(sTxt, lPos, lPos2 - lPos)
      
      lPos2 = InStr(lPos2, sTxt, ">")
      If lPos2 > 0 Then
        lPos3 = InStr(lPos2, sTxt, "</a>", vbTextCompare)
        If lPos3 > 0 Then
          If Mid(sTxt, lPos2 + 1, lPos3 - lPos2 - 1) Like "*" & sLinkName & "*" Then
            GetLinkNamedLike = Replace(sLink, "&amp;", "&", , , vbTextCompare)
            Exit Function
          End If
        End If
      End If
      
    End If
    lPos = InStr(lPos, sTxt, "a href=""", vbTextCompare)
    
  Loop
ERROR_HANDLER:

End Function

Public Function IsOnline() As Boolean
    IsOnline = CheckInternetConnection()
End Function

Private Function GetPortFromServer(sServer As String) As Integer
    
    If InStr(1, sServer, ":") > 0 Then
        GetPortFromServer = Val(Mid(sServer, InStr(1, sServer, ":") + 1))
    End If
    
End Function

Public Function GetServerFromServer(sServer As String) As String
    
    GetServerFromServer = sServer
    If InStr(1, sServer, ":") > 0 Then
        GetServerFromServer = Left(sServer, InStr(1, sServer, ":") - 1)
    End If
    
End Function

Public Function GetPathFromUrl(sUrl As String) As String
    
    Dim iPos As Integer
    
    iPos = InStr(1, sUrl, "://")
    If iPos > 0 Then
        iPos = InStr(iPos + 3, sUrl, "/")
    Else
        iPos = InStr(1, sUrl, "/")
    End If
    
    If iPos > 0 Then
        GetPathFromUrl = Mid(sUrl, iPos)
    Else
        GetPathFromUrl = "/"
    End If
    
End Function

Public Function GetDomainFromUrl(sUrl As String) As String
    
    Dim iPos1 As Integer
    Dim iPos2 As Integer
    Dim iPos3 As Integer
    
    iPos1 = InStr(1, sUrl, "://")
    If iPos1 > 0 Then
        iPos1 = iPos1 + 3
    Else
        iPos1 = 1
    End If
    iPos2 = InStr(iPos1, sUrl, "/")
    iPos3 = InStr(iPos1, sUrl, ":")
    
    If iPos3 > 0 And iPos2 > iPos3 Then iPos2 = iPos3
    If iPos2 = 0 Then iPos2 = iPos3
    If iPos2 = 0 Then iPos2 = 9999
    
    GetDomainFromUrl = Mid(sUrl, iPos1, iPos2 - iPos1)
    
End Function

Public Function GetPortFromUrl(sUrl As String) As String
    
    Dim iPos1 As Integer
    Dim iPos2 As Integer
    
    iPos1 = InStr(1, sUrl, "://")
    If iPos1 > 0 Then
        iPos1 = iPos1 + 3
    Else
        iPos1 = 1
    End If
    iPos2 = InStr(iPos1, sUrl, "/")
    
    If iPos2 = 0 Then iPos2 = 9999
    
    GetPortFromUrl = GetPortFromServer(Mid(sUrl, iPos1, iPos2))
    
End Function

Public Function MyNow() As Date

  MyNow = Now + gfTimeDeviation / 24 / 60 / 60

End Function

