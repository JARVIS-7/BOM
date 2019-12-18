VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ctlSMTPRelay 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   ScaleHeight     =   960
   ScaleWidth      =   3465
   Begin VB.Frame Frame1 
      Caption         =   "SMTP Controll"
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3315
      Begin VB.ListBox smtpdebug 
         Height          =   450
         ItemData        =   "ctlSMTPRelay.ctx":0000
         Left            =   120
         List            =   "ctlSMTPRelay.ctx":0002
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSWinsockLib.Winsock sckSMTP 
         Left            =   120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "ctlSMTPRelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' $author: ingo-/hjs$
' $id: V 2.0.3 date 270303 hdn$
' $version: 2.0.3$
' $file: $
'
' last modified:
' &date: 270303 scr 710624 hdn$
'
' contact: visit http://de.groups.yahoo.com/group/BOMInfo
'
'*******************************************************
'
' SMTP- Auth, ersetzt alten SMTP- Zugriff
' thx to ingo_
'
' ##############################################################################
'                               Definitionen
' ##############################################################################
Option Explicit
'
Private Const msBASE64CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const msISO_8859_1 As String = "ISO-8859-1"
Private Const msISO_8859_15 As String = "ISO-8859-15"
'
Private mbSendOk As Boolean
Private msSenderName As String
Private msSenderEMail  As String
Private msRecipientName As String
Private msRecipientEMail As String
Private msSubject As String
Private msMsg As String
Private msServer As String
Private mlPort As Long
Private msAuthUser As String
Private msAuthPass As String
Private mbSmtpDebug As Boolean
Private mbTestOnly As Boolean 'lg 29.05.03
'
Public Enum SmtpStateEnum
    [ssConnecting] = 1&
    [ssIdentifying] = 2&
    [ssAuthIdentify] = 3&
    [ssAuthUserName] = 4&
    [ssAuthPassword] = 5&
    [ssMailFrom] = 6&
    [ssRcptTo] = 7&
    [ssBeginBody] = 8&
    [ssSendBody] = 9&
    [ssClosing] = 10&
    [ssClosed] = 11&
End Enum
'
Private meSmtpState As SmtpStateEnum


' ##############################################################################
'                               Eigenschaften des benutzerdefinierten
'                               Steuerelement übernehmen und später wieder übergeben
' ##############################################################################

Public Property Let MailSenderName(sName As String)
    msSenderName = sName
End Property

Public Property Get MailSenderName() As String
    MailSenderName = msSenderName
End Property

Public Property Let MailSenderEMail(sAddr As String)
    msSenderEMail = sAddr
End Property

Public Property Get MailSenderEMail() As String
    MailSenderEMail = msSenderEMail
End Property

Public Property Let MailRecipientName(sName As String)
    msRecipientName = sName
End Property

Public Property Get MailRecipientName() As String
    MailRecipientName = msRecipientName
End Property

Public Property Let MailRecipientEMail(sAddr As String)
    msRecipientEMail = sAddr
End Property

Public Property Get MailRecipientEMail() As String
    MailRecipientEMail = msRecipientEMail
End Property

Public Property Let MailSubject(sSubject As String)
    msSubject = sSubject
End Property

Public Property Get MailSubject() As String
    MailSubject = msSubject
End Property

Public Property Let MailBody(sBody As String)
    msMsg = sBody
End Property

Public Property Get MailBody() As String
    MailBody = msMsg
End Property

Public Property Let SMTPServer(sServer As String)
    msServer = sServer
End Property

Public Property Get SMTPServer() As String
    SMTPServer = msServer
End Property

Public Property Let SMTPPort(pPort As Long)
    mlPort = pPort
End Property

Public Property Get SMTPPort() As Long
    SMTPPort = mlPort
End Property

Public Property Let SMTPAuthUser(sUser As String)
    msAuthUser = sUser
End Property

Public Property Get SMTPAuthUser() As String
    SMTPAuthUser = msAuthUser
End Property

Public Property Let SMTPAuthPass(sPass As String)
    msAuthPass = sPass
End Property

Public Property Get SMTPAuthPass() As String
    SMTPAuthPass = msAuthPass
End Property

Public Property Let SMTPDebugMode(bDebug As Boolean)
    mbSmtpDebug = bDebug
End Property

Public Property Get SMTPDebugMode() As Boolean
    SMTPDebugMode = mbSmtpDebug
End Property

Public Property Get SMTPDebugOutput() As String
    
    Dim i As Long
    Dim sTmp As String
    
    For i = 0 To smtpdebug.ListCount - 1
        sTmp = sTmp & smtpdebug.List(i) & vbCrLf
    Next 'i
    
    SMTPDebugOutput = sTmp
    
End Property
' ##############################################################################
'                               Senderoutine (Steuerelement Start)
' ##############################################################################
Public Function SendMail(Optional bTestOnly As Boolean = False) As Boolean
    
    On Error GoTo handleError
    
    sckSMTP.Close
    
    mbSendOk = False
    mbTestOnly = bTestOnly
    
    Call Encode_qp
    
    ''nicht anzeigen, dafür wird der debugoutput erzeugt und kann über get ausgelesen werden. lg 14.03.03
    'Debugfenster
    'If mbSmtpDebug Then
        'smtpdebug.Visible = True
    'Else
        'smtpdebug.Visible = False
    'End If
    
    If mbSmtpDebug Then smtpdebug.Clear
    If mbSmtpDebug Then smtpdebug.AddItem ("Start")
    
    'Verbindung öffnen
    meSmtpState = [ssConnecting]
    Call sckSMTP.Connect(msServer, mlPort)
    
    'Schleife bis die Mail versendet wurde.
    'Der restliche Versand geht dann durch Reaktionen auf die Serverantworten.
    Do Until meSmtpState = [ssClosed]
        DoEvents
    Loop
    
    'Rückmeldung i.O.
    SendMail = mbSendOk
    
Done:
On Error GoTo 0

Exit Function
    
handleError:
sckSMTP.Close
Resume Done

End Function

' ##############################################################################
'                               Verbindung schließen
' ##############################################################################
Private Sub sckSMTP_Close()

    On Error Resume Next
    sckSMTP.Close
    On Error GoTo 0
    meSmtpState = [ssClosed]
    
End Sub

' ##############################################################################
'                               SMTP Serverantworten analysieren und neue
'                               Befehle senden. Der Status der Sendung wird
'                               in der Variable smtpState zwischengespeichert
' ##############################################################################
Private Sub sckSMTP_DataArrival(ByVal bytesTotal As Long)
    
    On Error Resume Next
    Dim sData As String
    Dim iUTCOffset As Integer
    
    sckSMTP.GetData sData
    DoEvents

    If mbSmtpDebug Then smtpdebug.AddItem sData '(Val(Mid$(sData, 1, 3)))               ' Debuginfo
    Select Case Val(Mid$(sData, 1, 3))
        Case 220
            meSmtpState = [ssIdentifying]
            If Trim(msAuthUser) <> "" And Trim(msAuthPass) <> "" Then
                sckSMTP.SendData "EHLO " & GetUserName(msAuthUser) & vbCrLf
                If mbSmtpDebug Then smtpdebug.AddItem ("Ehlo")                   ' Debuginfo
            Else
                sckSMTP.SendData "HELO " & GetUserName(msSenderEMail) & vbCrLf
                If mbSmtpDebug Then smtpdebug.AddItem ("Helo")                   ' Debuginfo
            End If
        Case 235
            Select Case meSmtpState
                Case [ssAuthPassword]
                    meSmtpState = [ssMailFrom]
                    If mbSmtpDebug Then smtpdebug.AddItem ("Mailfrom")           ' Debuginfo
                    sckSMTP.SendData "MAIL FROM: <" & msSenderEMail & ">" & vbCrLf
            End Select
        Case 250
            Select Case meSmtpState
                Case [ssIdentifying]
                    If Trim(msAuthUser) <> "" And Trim(msAuthPass) <> "" Then
                        meSmtpState = [ssAuthIdentify]
                        If mbSmtpDebug Then smtpdebug.AddItem ("AUTH LOGIN")      ' Debuginfo
                        sckSMTP.SendData "AUTH LOGIN" & vbCrLf
                    Else
                        meSmtpState = [ssMailFrom]
                        If mbSmtpDebug Then smtpdebug.AddItem ("MailFrom Auth")   ' Debuginfo
                        sckSMTP.SendData "MAIL FROM: <" & msSenderEMail & ">" & vbCrLf
                    End If
                Case [ssMailFrom]
                    meSmtpState = [ssRcptTo]
                    If mbSmtpDebug Then smtpdebug.AddItem ("RCPT TO")             ' Debuginfo
                    sckSMTP.SendData "RCPT TO: <" & msRecipientEMail & ">" & vbCrLf
                Case [ssRcptTo]
                    meSmtpState = [ssBeginBody]
                    If mbTestOnly Then 'wenn nur Test hier abbrechen, lg 29.05.03
                        sckSMTP.SendData "QUIT" & vbCrLf
                        sckSMTP.Close
                        mbSendOk = True
                        meSmtpState = [ssClosed]
                    Else
                        If mbSmtpDebug Then smtpdebug.AddItem ("DATA1")           ' Debuginfo
                        sckSMTP.SendData "DATA" & vbCrLf
                    End If
                Case [ssSendBody]
                    meSmtpState = [ssClosing]
                    If mbSmtpDebug Then smtpdebug.AddItem ("QUIT")                ' Debuginfo
                    sckSMTP.SendData "QUIT"
                    sckSMTP.Close
                    mbSendOk = True
                    meSmtpState = [ssClosed]
            End Select
        Case 251
            Select Case meSmtpState
                Case [ssRcptTo]
                    meSmtpState = [ssBeginBody]
                    If mbTestOnly Then 'wenn nur Test hier abbrechen, lg 29.05.03
                        sckSMTP.SendData "QUIT" & vbCrLf
                        sckSMTP.Close
                        mbSendOk = True
                        meSmtpState = [ssClosed]
                    Else
                        If mbSmtpDebug Then smtpdebug.AddItem ("DATA2")               ' Debuginfo
                        sckSMTP.SendData "DATA" & vbCrLf
                    End If
            End Select
        Case 334
            Select Case meSmtpState
                Case [ssAuthIdentify]
                    meSmtpState = [ssAuthUserName]
                    If mbSmtpDebug Then smtpdebug.AddItem ("AUTH USER")           ' Debuginfo
                    sckSMTP.SendData Base64Encode(msAuthUser) & vbCrLf
                Case [ssAuthUserName]
                    meSmtpState = [ssAuthPassword]
                    If mbSmtpDebug Then smtpdebug.AddItem ("AUTH PASS")           ' Debuginfo
                    sckSMTP.SendData Base64Encode(msAuthPass) & vbCrLf
            End Select
        Case 354
            Select Case meSmtpState
                Case [ssBeginBody]
                    meSmtpState = [ssSendBody]
                    iUTCOffset = GetUTCOffset()
                    'changed 270303 scr 710624 hdn
                    If mbSmtpDebug Then smtpdebug.AddItem ("SEND BODY")           ' Debuginfo
                    sckSMTP.SendData _
                    "DATE: " & GetMsgDate() & " " & Format(Time, "Long Time") & " " & IIf(iUTCOffset < 0, "", "+") & Format(iUTCOffset, "00") & "00" & vbCrLf & _
                    "FROM: " & msSenderName & "<" & msSenderEMail & ">" & vbCrLf & _
                    "TO: " & msRecipientName & "<" & msRecipientEMail & ">" & vbCrLf & _
                    msMsg & vbCrLf & _
                    "." & vbCrLf
                    'subject ist schon in der message! checked 220503 IG
                    '"SUBJECT: " & msSubject & vbCrLf & vbCrLf &
            End Select
        Case Is >= 400
            meSmtpState = [ssClosed]
            If mbSmtpDebug Then smtpdebug.AddItem ("CLOSE1")                       ' Debuginfo
            sckSMTP.Close
        Case Else
            meSmtpState = [ssClosed]
            If mbSmtpDebug Then smtpdebug.AddItem ("CLOSE2")                       ' Debuginfo
            sckSMTP.Close
    End Select
End Sub

' ##############################################################################
'                               Username aus der Emailadresse extrahieren
' ##############################################################################
Private Function GetUserName(strAddress As String) As String

    On Error Resume Next
    
    If InStr(1, strAddress, "@") > 0 Then
        GetUserName = Mid$(strAddress, 1, InStr(1, strAddress, "@") - 1)
    Else
        GetUserName = strAddress
    End If
    
End Function

' ##############################################################################
'                               SMTP Fehlerroutine
' ##############################################################################
Private Sub sckSMTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    On Error Resume Next
    sckSMTP.Close
    On Error GoTo 0
    
    meSmtpState = [ssClosed]
    If mbSmtpDebug Then smtpdebug.AddItem ("ERROR" & Description)                                ' Debuginfo
    mbSendOk = False
    
End Sub

' ##############################################################################
'                               Datumsstring der Email berechnen
' ##############################################################################

Private Function GetMsgDate() As String
    
        Select Case Weekday(Now, vbMonday)
            Case 1: GetMsgDate = "Mon"
            Case 2: GetMsgDate = "Tue"
            Case 3: GetMsgDate = "Wen"
            Case 4: GetMsgDate = "Thu"
            Case 5: GetMsgDate = "Fri"
            Case 6: GetMsgDate = "Sat"
            Case 7: GetMsgDate = "Sun"
        End Select
           
        GetMsgDate = GetMsgDate & ", "
        GetMsgDate = GetMsgDate & Format$(Now, "dd")
        
        Select Case Month(Now)
            Case 1: GetMsgDate = GetMsgDate & " Jan "
            Case 2: GetMsgDate = GetMsgDate & " Feb "
            Case 3: GetMsgDate = GetMsgDate & " Mar "
            Case 4: GetMsgDate = GetMsgDate & " Apr "
            Case 5: GetMsgDate = GetMsgDate & " May "
            Case 6: GetMsgDate = GetMsgDate & " Jun "
            Case 7: GetMsgDate = GetMsgDate & " Jul "
            Case 8: GetMsgDate = GetMsgDate & " Aug "
            Case 9: GetMsgDate = GetMsgDate & " Sep "
            Case 10: GetMsgDate = GetMsgDate & " Oct "
            Case 11: GetMsgDate = GetMsgDate & " Nov "
            Case 12: GetMsgDate = GetMsgDate & " Dec "
        End Select
        GetMsgDate = GetMsgDate & Format$(Now, "YYYY")
           
End Function

' ##############################################################################
'                               Dez - Bin - Dez
' ##############################################################################
Private Function Base10ToBinary(ByVal lBase10 As Long) As String
    
    Dim iPrevResult As Integer, iCurResult As Integer
    
    If lBase10 = 0 Then
        Base10ToBinary = "0"
    Else
    
        Do
            iCurResult = Int(Log(lBase10) / Log(2))
            If iPrevResult = 0 Then iPrevResult = iCurResult + 1
            Base10ToBinary = Base10ToBinary & String$(iPrevResult - iCurResult - 1, "0") & "1"
            lBase10 = lBase10 - 2 ^ iCurResult
            iPrevResult = iCurResult
        Loop Until lBase10 = 0
        
        Base10ToBinary = Base10ToBinary & String$(iCurResult, "0")
        
    End If
    
End Function

Private Function BinaryToBase10(ByVal sBinary As String) As Long

    Dim i As Integer
    
    For i = Len(sBinary) To 1 Step -1
        BinaryToBase10 = BinaryToBase10 + Val(Mid(sBinary, i, 1)) * 2 ^ (Len(sBinary) - i)
    Next
    
End Function

' ##############################################################################
'                               Bin3x8 - Bin4x6 - Bin3x8
' ##############################################################################
Private Sub Bin3x8To4x6(ByVal Bin1Len8 As String, ByVal Bin2Len8 As String, ByVal Bin3Len8 As String, ByRef Bin1Len6 As String, ByRef Bin2Len6 As String, ByRef Bin3Len6 As String, ByRef Bin4Len6 As String)
    
    Bin1Len8 = Right("0000000" & Bin1Len8, 8)
    Bin2Len8 = Right("0000000" & Bin2Len8, 8)
    Bin3Len8 = Right("0000000" & Bin3Len8, 8)
    Bin1Len6 = Left(Bin1Len8, 6)
    Bin2Len6 = Right(Bin1Len8, 2) & Left(Bin2Len8, 4)
    Bin3Len6 = Right(Bin2Len8, 4) & Left(Bin3Len8, 2)
    Bin4Len6 = Right(Bin3Len8, 6)
    
End Sub

Private Sub Bin4x6To3x8(ByVal Bin1Len6 As String, ByVal Bin2Len6 As String, ByVal Bin3Len6 As String, ByVal Bin4Len6 As String, ByRef Bin1Len8 As String, ByRef Bin2Len8 As String, ByRef Bin3Len8 As String)
    
    Bin1Len6 = Right("00000" & Bin1Len6, 6)
    Bin2Len6 = Right("00000" & Bin2Len6, 6)
    Bin3Len6 = Right("00000" & Bin3Len6, 6)
    Bin4Len6 = Right("00000" & Bin4Len6, 6)
    Bin1Len8 = Bin1Len6 & Left(Bin2Len6, 2)
    Bin2Len8 = Right(Bin2Len6, 4) & Left(Bin3Len6, 4)
    Bin3Len8 = Right(Bin3Len6, 2) & Bin4Len6
    
End Sub

' ##############################################################################
'                               Base64 Encode - Decode
' ##############################################################################
Private Function Base64Encode(ByVal NormalString As String, Optional ByVal iBreak As Integer = 0) As String
    
    Dim i As Integer, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String
    
    If NormalString = vbNullString Then Exit Function
    
    For i = 1 To Len(NormalString) - 3 Step 3
        Bin1Len8 = Base10ToBinary(Asc(Mid(NormalString, i, 1)))
        Bin2Len8 = Base10ToBinary(Asc(Mid(NormalString, i + 1, 1)))
        Bin3Len8 = Base10ToBinary(Asc(Mid(NormalString, i + 2, 1)))
        Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
        Base64Encode = Base64Encode & Mid(msBASE64CHARS, BinaryToBase10(Bin1Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(msBASE64CHARS, BinaryToBase10(Bin2Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(msBASE64CHARS, BinaryToBase10(Bin3Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(msBASE64CHARS, BinaryToBase10(Bin4Len6) + 1, 1)
    Next
    
    NormalString = Right(NormalString, Len(NormalString) - IIf(Len(NormalString) / 3 = Int(Len(NormalString) / 3), Len(NormalString) - 3, Int(Len(NormalString) / 3) * 3))
    Bin1Len8 = Base10ToBinary(Asc(Left(NormalString, 1)))
    If Len(NormalString) >= 2 Then Bin2Len8 = Base10ToBinary(Asc(Mid(NormalString, 2, 1))) Else Bin2Len8 = "0"
    If Len(NormalString) = 3 Then Bin3Len8 = Base10ToBinary(Asc(Right(NormalString, 1))) Else Bin3Len8 = "0"
    Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
    Base64Encode = Base64Encode & Mid(msBASE64CHARS, BinaryToBase10(Bin1Len6) + 1, 1)
    Base64Encode = Base64Encode & Mid(msBASE64CHARS, BinaryToBase10(Bin2Len6) + 1, 1)
    Base64Encode = Base64Encode & IIf(Len(NormalString) >= 2, Mid(msBASE64CHARS, BinaryToBase10(Bin3Len6) + 1, 1), "=")
    Base64Encode = Base64Encode & IIf(Len(NormalString) = 3, Mid(msBASE64CHARS, BinaryToBase10(Bin4Len6) + 1, 1), "=")
    
    If iBreak > 0 Then
        i = iBreak + 1
        While i < Len(Base64Encode)
            Base64Encode = Left(Base64Encode, i - 1) & vbCrLf & Mid(Base64Encode, i)
            i = i + iBreak + 2
        Wend
    End If
    
End Function

Private Function Base64Decode(ByVal Base64String As String) As String
    
    Dim i As Integer, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String
    
    Base64String = RemoveFromString(Base64String, " ")
    Base64String = RemoveFromString(Base64String, vbCr)
    Base64String = RemoveFromString(Base64String, vbLf)
    
    If Base64String = vbNullString Then Exit Function
    
    For i = 0 To 255
        If InStr(Base64String, Chr(i)) > 0 And Not _
            ((InStr(msBASE64CHARS, Chr(i)) > 0) Or (i = Asc("="))) Then Exit Function
    Next
    
    If Not Len(Base64String) / 4 = Len(Base64String) \ 4 Then Exit Function
    
    For i = 1 To Len(Base64String) Step 4
        Bin1Len6 = Base10ToBinary(InStr(msBASE64CHARS, Mid(Base64String, i, 1)) - 1)
        Bin2Len6 = Base10ToBinary(InStr(msBASE64CHARS, Mid(Base64String, i + 1, 1)) - 1)
        If Mid(Base64String, i + 2, 1) = "=" Then Bin3Len6 = "0" Else Bin3Len6 = Base10ToBinary(InStr(msBASE64CHARS, Mid(Base64String, i + 2, 1)) - 1)
        If Mid(Base64String, i + 3, 1) = "=" Then Bin4Len6 = "0" Else Bin4Len6 = Base10ToBinary(InStr(msBASE64CHARS, Mid(Base64String, i + 3, 1)) - 1)
        Call Bin4x6To3x8(Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6, Bin1Len8, Bin2Len8, Bin3Len8)
        Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin1Len8))
        If Not Mid(Base64String, i + 2, 1) = "=" Then Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin2Len8))
        If Not Mid(Base64String, i + 3, 1) = "=" Then Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin3Len8))
    Next
    
End Function

' ##############################################################################
'                               Entfernt Zeichen aus einem String
' ##############################################################################
Private Function RemoveFromString(ByVal sTheString As String, ByVal sWhatToRemove As String) As String

    Dim lPos As Long
    
    If Len(sWhatToRemove) Then
    
        lPos = InStr(sTheString, sWhatToRemove)
        While lPos > 0
            sTheString = Mid$(sTheString, 1, lPos - 1) & Mid$(sTheString, lPos + Len(sWhatToRemove))
            lPos = InStr(sTheString, sWhatToRemove)
        Wend
        RemoveFromString = sTheString
    End If
    
End Function

' ##############################################################################
'                               Decodiert Quoted Printable
' ##############################################################################
Public Function Decode_qp(ByVal sTxt As String) As String
    
    Const CODIERUNG_QUOTED_PRINTABLE As String = "Content-Transfer-Encoding: quoted-printable"
    Dim tmp As String
    Dim pos As Long
    
    tmp = sTxt
    pos = InStr(1, tmp, CODIERUNG_QUOTED_PRINTABLE, vbTextCompare)
    If pos > 0 Then
        pos = InStr(pos + 1, tmp, "=")
        Do While pos > 0
            If Mid(tmp, pos + 1, 2) = vbCrLf Then
              tmp = Left(tmp, pos - 1) & Mid(tmp, pos + 3)
            ElseIf Mid(tmp, pos + 1, 1) = vbLf Then
              tmp = Left(tmp, pos - 1) & Mid(tmp, pos + 2)
            Else
              tmp = Left(tmp, pos - 1) & GetChar(Mid(tmp, pos + 1, 2)) & Mid(tmp, pos + 3)
            End If
            pos = InStr(pos + 1, tmp, "=")
        Loop
    End If
    Decode_qp = tmp
    
End Function

' ##############################################################################
'                               Liefert das Zeichen zu einem Hexcode
' ##############################################################################
Private Function GetChar(ByVal sHexCode As String) As String
    
    On Error GoTo ERROR_HANDLER
    
    If UCase$(sHexCode) = "A4" Then sHexCode = "80"
    GetChar = Chr$(CByte("&h" & sHexCode))

Done:
On Error GoTo 0
Exit Function

ERROR_HANDLER:
Resume Done

End Function

' ##############################################################################
'                               Codiert ein Header-Feld Quoted Printable
' ##############################################################################
Private Function EncodeHeader_qp(ByVal sTxt As String) As String
    
    Dim i As Long
    Dim sTmp As String
    Dim bytChar As Byte
    Dim bConverted As Boolean
    Dim sCharSet As String
    
    bConverted = False
    For i = 1 To Len(sTxt)
        bytChar = Asc(Mid(sTxt, i, 1))
        If bytChar > 128 Then
            bConverted = True
            sTmp = sTmp & "=" & Hex(bytChar)
        ElseIf bytChar = 128 Then '€
            bConverted = True
            sTmp = sTmp & "=A4"
        ElseIf bytChar = 61 Then '=
            sTmp = sTmp & "=3D"
        ElseIf bytChar = 63 Then '?
            sTmp = sTmp & "=3F"
        Else
            sTmp = sTmp & Chr(bytChar)
        End If
    Next i
    
    If bConverted Then
        sTmp = Replace(sTmp, " ", "_")
        If InStr(1, sTxt, "€") > 0 Then
            sCharSet = msISO_8859_15
        Else
            sCharSet = msISO_8859_1
        End If
        EncodeHeader_qp = "=?" & sCharSet & "?Q?" & sTmp & "?="
    Else
        sTmp = Replace(sTmp, "=3D", "=")
        sTmp = Replace(sTmp, "=3F", "?")
        EncodeHeader_qp = sTmp
    End If
    
End Function

' ##############################################################################
'                               Codiert den Body Quoted Printable
' ##############################################################################
Private Function EncodeBody_qp(ByVal sTxt As String, sCharSet As String) As String
    
    Dim i As Long
    Dim sTmp As String
    Dim bytChar As Byte
    
    For i = 1 To Len(sTxt)
        bytChar = Asc(Mid(sTxt, i, 1))
        If bytChar > 128 Or bytChar = 61 Then
            sTmp = sTmp & "=" & Hex(bytChar)
        ElseIf bytChar = 128 Then '€
            sTmp = sTmp & "=A4"
        Else
            sTmp = sTmp & Chr(bytChar)
        End If
    Next i
    
    If InStr(1, sTxt, "€") > 0 Then
        sCharSet = msISO_8859_15
    Else
        sCharSet = msISO_8859_1
    End If
    EncodeBody_qp = sTmp
    
End Function

' ##############################################################################
'                               Codiert eine Adresse Quoted Printable
' ##############################################################################
Private Function EncodeAddress_qp(ByVal sTxt As String) As String
    
    'MD-Marker , Function wird nicht aufgerufen
    
'    Dim sRealName As String
'    Dim sAdresse As String
'
'    If sTxt Like "*<*>*" Then
'        sRealName = Left(sTxt, InStr(1, sTxt, "<") - 1)
'        sAdresse = Mid(sTxt, InStr(1, sTxt, "<"))
'        If sRealName Like """*""" Then sRealName = Mid(sRealName, 2, Len(sRealName) - 2)
'        sRealName = EncodeHeader_qp(sRealName)
'        If Not Left(sRealName, 5) = "=?ISO" And Len(sRealName) > 0 Then sRealName = """" & sRealName & """"
'        EncodeAddress_qp = sRealName & sAdresse
'    Else
'        EncodeAddress_qp = sTxt
'    End If
    
End Function

' ##############################################################################
'                               Codiert alle relevanten Daten Quoted Printable
' ##############################################################################
Private Sub Encode_qp()

    Dim lPos1 As Long
    Dim lPos2 As Long
    Dim sLocalSubject As String
    Dim sLocalBody As String
    Dim sCharSet As String
    
    If msSenderName Like """*""" Then msSenderName = Mid(msSenderName, 2, Len(msSenderName) - 2)
    If msSenderName Like "'*'" Then msSenderName = Mid(msSenderName, 2, Len(msSenderName) - 2)
    msSenderName = EncodeHeader_qp(msSenderName)
    If Not Left(msSenderName, 5) = "=?ISO" And Len(msSenderName) > 0 Then msSenderName = """" & msSenderName & """"
    
    If msRecipientName Like """*""" Then msRecipientName = Mid(msRecipientName, 2, Len(msRecipientName) - 2)
    msRecipientName = EncodeHeader_qp(msRecipientName)
    If Not Left(msRecipientName, 5) = "=?ISO" And Len(msRecipientName) > 0 Then msRecipientName = """" & msRecipientName & """"
    
    lPos1 = InStr(1, msMsg, "subject:", vbTextCompare)
    If lPos1 > 0 Then
      lPos2 = InStr(lPos1, msMsg, vbCrLf)
      If lPos2 > 0 Then
        sLocalSubject = Trim(Mid(msMsg, lPos1 + 8, lPos2 - lPos1 - 8))
        sLocalBody = Mid(msMsg, lPos2 + 2)
      Else
        sLocalSubject = Trim(Mid(msMsg, lPos1 + 8))
        sLocalBody = ""
      End If
      sLocalSubject = EncodeHeader_qp(sLocalSubject)
      sLocalBody = EncodeBody_qp(sLocalBody, sCharSet)
      
      msMsg = "Subject: " & sLocalSubject & vbCrLf & _
               "Mime-Version: 1.0" & vbCrLf & _
               "Content-Type: text/plain; charset=""" & sCharSet & """" & vbCrLf & _
               "Content-Transfer-Encoding: quoted-printable" & vbCrLf & vbCrLf & _
               sLocalBody
    End If

End Sub

' ##############################################################################
'                               Codiert Unicode in UTF-8
' ##############################################################################
Private Function Encode_UTF8(sTxt As String) As String
    
    Dim n As Long
    Dim c As Long
    Dim sUtfTxt As String
    
    For n = 1 To Len(sTxt$)
        c = AscW(Mid(sTxt$, n, 1))
        If c < 128 Then
            sUtfTxt = sUtfTxt & Mid(sTxt$, n, 1)
        ElseIf ((c > 127) And (c < 2048)) Then
            sUtfTxt = sUtfTxt & Chr(((c \ 64) Or 192))              '((c>>6)|192);
            sUtfTxt = sUtfTxt & Chr(((c And 63) Or 128))            '((c&63)|128);}
        Else
            sUtfTxt = sUtfTxt & Chr(((c \ 4096) Or 224))            '((c>>12)|224);
            sUtfTxt = sUtfTxt & Chr((((c \ 64) And 63) Or 128))     '(((c>>6)&63)|128);
            sUtfTxt = sUtfTxt & Chr(((c And 63) Or 128))            '((c&63)|128);
        End If
    Next 'n
    
    Encode_UTF8 = sUtfTxt
    
End Function

Public Function DecodeHeader(ByVal sTxt As String) As String
    
    Dim sTmp As String
    Dim lPos As Long
    Dim lPos1 As Long
    Dim lPos1a As Long
    Dim lPos1b As Long
    Dim lPos2 As Long
    Dim lPos3 As Long
    Dim bUtf8 As Boolean
    Dim sEncodingType As String
    
    lPos1a = InStr(1, sTxt, "=?ISO-8859-", vbTextCompare)
    lPos1b = InStr(1, sTxt, "=?UTF-8", vbTextCompare)
    If lPos1a = 0 Then lPos1a = 999999999
    If lPos1b = 0 Then lPos1b = 999999999
    
    If lPos1a < lPos1b Then
        lPos1 = lPos1a
        bUtf8 = False
    ElseIf lPos1a > lPos1b Then
        lPos1 = lPos1b
        bUtf8 = True
    Else
        lPos1 = 0
    End If
    
    lPos2 = InStr(lPos1 + 3, sTxt, "?")
    lPos3 = InStr(lPos2 + 3, sTxt, "?=")
    
    Do While (lPos1 > 0 And lPos2 > 0 And lPos3 > 0)
        
        sEncodingType = UCase(Mid(sTxt, lPos2 + 1, 1))
        sTmp = Mid(sTxt, lPos2 + 3, lPos3 - lPos2 - 3)
        
        If sEncodingType = "Q" Then 'quoted printable
            sTmp = Replace(sTmp, "_", " ")
            lPos = InStr(lPos + 1, sTmp, "=")
            Do While lPos > 0
                sTmp = Left(sTmp, lPos - 1) & GetChar(Mid(sTmp, lPos + 1, 2)) & Mid(sTmp, lPos + 3)
                lPos = InStr(lPos + 1, sTmp, "=")
            Loop
        ElseIf sEncodingType = "B" Then 'base64
            sTmp = Base64Decode(sTmp)
        'Else
            'unbekannt, so lassen
        End If
        
        If bUtf8 Then sTmp = Decode_UTF8(sTmp)
    
        sTxt = Mid(sTxt, 1, lPos1 - 1) & sTmp & Mid(sTxt, lPos3 + 2 + 1)
        
        lPos1a = InStr(lPos1 + Len(sTmp), sTxt, "=?ISO-8859-", vbTextCompare)
        lPos1b = InStr(lPos1 + Len(sTmp), sTxt, "=?UTF-8", vbTextCompare)
        If lPos1a = 0 Then lPos1a = 999999999
        If lPos1b = 0 Then lPos1b = 999999999
        
        If lPos1a < lPos1b Then
            lPos1 = lPos1a
            bUtf8 = False
        ElseIf lPos1a > lPos1b Then
            lPos1 = lPos1b
            bUtf8 = True
        Else
            lPos1 = 0
        End If
        lPos2 = InStr(lPos1 + 3, sTxt, "?")
        lPos3 = InStr(lPos2 + 3, sTxt, "?=")
    Loop
    
    DecodeHeader = sTxt
    
End Function

Function Decode_UTF8(astr$)

  Dim n As Long
  Dim c As Long
  Dim chartext As String
  
  Do While n < Len(astr$)
    n = n + 1
    If Asc(Mid(astr$, n, 1)) < 128 Then
      c = Asc(Mid(astr$, n, 1))
    ElseIf Asc(Mid(astr$, n, 1)) < 224 Then
      c = (Asc(Mid(astr$, n, 1)) - 192) * 64
      n = n + 1
      c = c + Asc(Mid(astr$, n, 1)) - 128
    Else
      c = (Asc(Mid(astr$, n, 1)) - 224) * 4096
      n = n + 1
      c = c + (Asc(Mid(astr$, n, 1)) - 128) * 64
      n = n + 1
      c = c + (Asc(Mid(astr$, n, 1)) - 128)
    End If
    chartext = chartext & ChrW(c)
  Loop

  Decode_UTF8 = chartext
  
End Function

