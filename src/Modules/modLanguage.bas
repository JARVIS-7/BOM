Attribute VB_Name = "modLanguage"
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
' $author: hjs$
' $id: V 2.0.2 date 030303 hjs$
' $version: 2.0.2$
' $file: $
'
' last modified:
' &date: 030303$
'
' contact: visit http://de.groups.yahoo.com/group/BOMInfo
'
'*******************************************************
Option Explicit

Private Const msREQUIREDLANGUAGEVERSION As String = "1.0.69"

'
' Liste der Textstrings und die entsprechenden Zugriffe
' zusätzlich LOCALE- Zugriffe
'

Public gsarrLangTxt(800) As String

'lokale Operatoren
Private msLocSeparator As String
Private msLocGrouping As String

'Dezimaloperator, Datumsfunktionen

Private Declare Function GetLocaleInfo Lib "kernel32" _
    Alias "GetLocaleInfoA" (ByVal lLocale As Long, _
        ByVal lLocaleType As Long, ByVal sLCData As String, _
        ByVal lBufferLength As Long) As Long

'Public Declare Function GetSystemDefaultLangID Lib _
    "kernel32" () As Integer

'Public Declare Function VerLanguageName Lib "kernel32" _
    Alias "VerLanguageNameA" (ByVal wLang As Long, _
    ByVal szLang As String, ByVal nSize As Long) As Long

'Public Const LOCALE_SLONGDATE = &H20
'Public Const LOCALE_SSHORTDATE = &H1F
'Public Const LOCALE_SCURRENCY = &H14
'Public Const LOCALE_SINTLSYMBOL = &H15
'Public Const LOCALE_STIMEFORMAT = &H1003
'Public Const LOCALE_SDATE = &H1D
'Public Const LOCALE_SABBREVMONTHNAME1 = &H44
Private Const LOCALE_SDECIMAL As Long = &HE
Private Const LOCALE_STHOUSAND As Long = &HF
Private Const LOCALE_USER_DEFAULT As Long = &H400


Public Function GetSupportedLanguage() As String

    Dim sFile As String
    Dim sValue As String
    Dim iTmp As Integer
    
    sFile = App.Path & "\languages.ini"
    
    'String lesen
    iTmp = INIGetValue(sFile, "Langfile", "Supported", sValue)
    If iTmp = 0 Then
        MsgBox "Languagefile 'languages.ini' not found"
        End 'MD-Marker
    Else
        GetSupportedLanguage = sValue
    End If
    
End Function
Public Sub SelectLanguage(sLanguage As String)
    
    Dim i As Integer
    Dim sFile As String
    Dim oIni As clsIni
    
    Set oIni = New clsIni
    
    sFile = App.Path & "\languages.ini"
  
    Call oIni.ReadIni(sFile, sLanguage)
    
    For i = 1 To UBound(gsarrLangTxt)
        gsarrLangTxt(i) = oIni.GetValue(sLanguage, "text" & CStr(i))
    Next i
    
    Set oIni = Nothing
    
End Sub

Public Sub CheckVersionOfLanguageFile()
    
    Dim sMsg As String
    
    sMsg = gsarrLangTxt(25) & gsarrLangTxt(26)
    sMsg = Replace(sMsg, "%FILE%", "languages.ini")
    sMsg = Replace(sMsg, "%REQVER%", msREQUIREDLANGUAGEVERSION)
    sMsg = Replace(sMsg, "\n", vbCrLf)
    
    If VersionValue(msREQUIREDLANGUAGEVERSION) > VersionValue(GetLanguageFileVersion()) Then
        If vbNo = MsgBox(sMsg, vbYesNo Or vbQuestion) Then
            End 'MD-Marker
        End If
    End If
    
End Sub

Public Function GetLanguageFileVersion() As String
    
    Dim sVersion As String
    Dim sFile As String
    
    sFile = App.Path & "\languages.ini"
    
    Call INIGetValue(sFile, "Langfile", "Version", sVersion)
    
    GetLanguageFileVersion = sVersion
    
End Function

'
' Wandlungsroutinen
'

' Wandlung EBayString mit Nachkomma in Double (EbayString ist Länderspezifisch, in USA z.B. mit .)
Public Function EbayString2Float(ByVal StrVal As String) As Double

On Error Resume Next

Dim i As Long
Dim sTmp As String
Static bDecimalSpecifierWarningShown As Boolean

If msLocSeparator = "" Then
    msLocSeparator = GetDecimalSpecifier
End If
If msLocGrouping = "" Then
    msLocGrouping = GetThousandSpecifier
End If

If msLocSeparator = msLocGrouping And Not bDecimalSpecifierWarningShown Then
    bDecimalSpecifierWarningShown = True
    MsgBox gsarrLangTxt(443), vbExclamation
End If

If msLocSeparator <> gsCmdDecSeparator Then
    'Group- Separator löschen
    StrVal = Replace(StrVal, msLocSeparator, "") 'geändert, vorher wurde der Rückgabewert nicht entgegengenommen, lg 18.04.03
    StrVal = Replace(StrVal, gsCmdDecSeparator, msLocSeparator) 'dito
End If

For i = 1 To Len(StrVal)
  If IsNumeric(Mid(StrVal, i, 1)) Or Mid(StrVal, i, 1) = msLocSeparator Then
    sTmp = sTmp & Mid(StrVal, i, 1)
  End If
Next i
EbayString2Float = CDbl(sTmp)

End Function

Public Function String2Float(ByVal sTmp As String) As Double
    
    On Error Resume Next
    
    Static bDecimalSpecifierWarningShown As Boolean
    
    If msLocSeparator = "" Then
        msLocSeparator = GetDecimalSpecifier
    End If
    If msLocGrouping = "" Then
        msLocGrouping = GetThousandSpecifier
    End If
    
    If msLocSeparator = msLocGrouping And Not bDecimalSpecifierWarningShown Then
        bDecimalSpecifierWarningShown = True
        MsgBox gsarrLangTxt(443), vbExclamation
    End If
    
    sTmp = Trim(sTmp)
    If InStr(1, sTmp, " ") > 0 Then
        sTmp = Left(sTmp, InStr(1, sTmp, " ") - 1)
    End If
    
    If msLocSeparator <> msLocGrouping Then
        If sTmp Like "*" & msLocGrouping & "###" Then
            ' das ist vermutlich richtig -> 123.456
        ElseIf InStr(1, sTmp, msLocSeparator) = 0 Then
            sTmp = Replace(sTmp, msLocGrouping, msLocSeparator)
        Else
            sTmp = Replace(sTmp, msLocGrouping, "")
        End If
    End If
    
    String2Float = CDbl(sTmp)
    
End Function

'
'Wandlung Datumsabkürzung extern > Zahl
'
Public Function ConvertMonthname1(ByVal sExtDateString As String) As String
    
    Dim iPos As Integer
    Dim i As Integer
    
    On Error GoTo errExit
    
    For i = 1 To 12
        iPos = InStrRev(sExtDateString, gsarrMonthNames1(i), , vbTextCompare)
        If iPos > 0 Then
            'ok, gotcha
            sExtDateString = Left(sExtDateString, iPos - 1) & CStr(i) & Mid(sExtDateString, iPos + Len(gsarrMonthNames1(i)))
            'Exit For ' Jetzt alles umwandeln weil teilweise die Wochentage exakt wie die Monate beginnen und dann der Wochentag als Montag genommen wird.
        End If
    Next i
    
errExit:
    ConvertMonthname1 = sExtDateString
    
End Function

Public Function ConvertMonthname2(ByVal sExtDateString As String) As String
    
    Dim iPos As Integer
    Dim i As Integer
    
    On Error GoTo errExit
    
    For i = 1 To 12
        iPos = InStrRev(sExtDateString, gsarrMonthNames2(i), , vbTextCompare)
        If iPos > 0 Then
            'ok, gotcha
            sExtDateString = Left(sExtDateString, iPos - 1) & CStr(i) & Mid(sExtDateString, iPos + Len(gsarrMonthNames2(i)))
            'Exit For ' Jetzt alles umwandeln weil teilweise die Wochentage exakt wie die Monate beginnen und dann der Wochentag als Montag genommen wird.
        End If
    Next i
    
errExit:
    ConvertMonthname2 = sExtDateString
    
End Function

'
'Zugriffe auf lokales Datum, Zeit etc
'
'Public Function GetDateTimeFormat() As String
'
'   Dim lBuffLen    As Long
'   Dim sBuffer     As String
'   Dim lResult     As Long
'   Dim sDateFormat As String
'
'   On Error GoTo vbErrorHandler
'
'   If LocDateFormat <> "" Then
'        GetDateTimeFormat = LocDateFormat
'        Exit Function
'   End If
'
'   lBuffLen = 128
'   sBuffer = String$(lBuffLen, vbNullChar)
'
'   lResult = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer, lBuffLen)
'
'   If lResult > 0 Then
'      sDateFormat = Left$(sBuffer, lResult - 1)
'
'
'      GetDateTimeFormat = sDateFormat
'   Else
'      GetDateTimeFormat = "DD.MM.YY"
'   End If
'
'   lResult = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIMEFORMAT, sBuffer, lBuffLen)
'
'   If lResult > 0 Then
'      sDateFormat = Left$(sBuffer, lResult - 1)
'      GetDateTimeFormat = GetDateTimeFormat & " " & sDateFormat
'   Else
'      GetDateTimeFormat = GetDateTimeFormat & " HH:mm:ss"
'   End If
'   LocDateFormat = GetDateTimeFormat
'   Exit Function
'
'vbErrorHandler:
'   'Err.Raise Err.Number, "GetDateTimeFormat", Err.Description
'End Function

Public Function GetDecimalSpecifier() As String
   '
   ' lokalen Dezimaltrenner bestimmen
   '
   Dim lBuffLen As Long
   Dim sBuffer  As String
   Dim sDecimal As String
   Dim lResult  As Long

   On Error GoTo vbErrorHandler

   lBuffLen = 128

   sBuffer = String$(lBuffLen, vbNullChar)

   lResult = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, lBuffLen)
   sDecimal = Left$(sBuffer, lResult - 1)

   GetDecimalSpecifier = sDecimal

   Exit Function

vbErrorHandler:
   'Err.Raise Err.Number, "GetDecimalSpecifier", Err.Description
End Function

Public Function GetThousandSpecifier() As String
   '
   ' lokalen Tausendertrenner bestimmen
   '
   Dim lBuffLen As Long
   Dim sBuffer  As String
   Dim sThousand As String
   Dim lResult  As Long

   On Error GoTo vbErrorHandler

   lBuffLen = 128

   sBuffer = String$(lBuffLen, vbNullChar)

   lResult = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sBuffer, lBuffLen)
   sThousand = Left$(sBuffer, lResult - 1)

   GetThousandSpecifier = sThousand

   Exit Function

vbErrorHandler:
   'Err.Raise Err.Number, "GetThousandSpecifier", Err.Description
End Function

'
' lesen der lokalen Abkürzung des Monatsnamens
'

'Function AbbrevMonthName(ByVal MonthNum As Long) As String
'   Dim lBuffLen As Long
'   Dim sBuffer  As String
'   Dim sDecimal As String
'   Dim lResult  As Long
'
'   On Error GoTo errorhandler
'
'   lBuffLen = 128
'
'   sBuffer = String$(lBuffLen, vbNullChar)
'
'   If MonthNum >= 1 And MonthNum <= 12 Then
'
'        MonthNum = MonthNum + LOCALE_SABBREVMONTHNAME1 - 1
'
'        lResult = GetLocaleInfo(LOCALE_USER_DEFAULT, MonthNum, sBuffer, lBuffLen)
'
'
'        AbbrevMonthName = Left$(sBuffer, lResult - 1)
'
'        Exit Function
'    End If
'
'errorhandler:
'
'End Function

Public Function TimeLeft2String(fTimeLeft As Double) As String
    
    If fTimeLeft < gfRESTZEITEWIG Then
        TimeLeft2String = CStr(Int(fTimeLeft)) & gsarrLangTxt(56) & " " & Format(CDate(fTimeLeft), "hh:mm:ss")
    Else
        TimeLeft2String = "---"
    End If
    
End Function

Public Function HtmlCleanup(ByVal sTxt, Optional bInsertSpaces As Boolean = False) As String

  Dim iPos1 As Integer
  Dim iPos2 As Integer

  sTxt = Replace(sTxt, Chr(160), Chr(32))
  sTxt = Replace(sTxt, Chr(9), Chr(32))
  Do While InStr(1, sTxt, "  ") > 0
    sTxt = Trim(Replace(sTxt, "  ", " "))
  Loop
  
  iPos1 = InStr(1, sTxt, "<")
  iPos2 = InStr(iPos1 + 1, sTxt, ">")
  Do While (iPos1 > 0 And iPos2 > iPos1)
    sTxt = Left(sTxt, iPos1 - 1) & IIf(bInsertSpaces, " ", "") & Mid(sTxt, iPos2 + 1)
    iPos1 = InStr(1, sTxt, "<")
    iPos2 = InStr(iPos1 + 1, sTxt, ">")
  Loop
  Do While InStr(1, sTxt, "  ") > 0
    sTxt = Trim(Replace(sTxt, "  ", " "))
  Loop
  
  sTxt = HtmlZeichenConvert(sTxt)
  HtmlCleanup = Trim(sTxt)

End Function

Public Function HtmlZeichenConvert(ByVal sHtmlString As String) As String

  Dim i As Long
  Dim sTmp As String
  Dim sarrHtmlZeichen() As String
  
  ReDim sarrHtmlZeichen(1 To 255) As String
  
  sarrHtmlZeichen(34) = "&quot;"      ' "
  sarrHtmlZeichen(38) = "&amp;"       ' &
  sarrHtmlZeichen(39) = "&#39;"       ' '
  sarrHtmlZeichen(60) = "&lt;"        ' <
  sarrHtmlZeichen(62) = "&gt;"        ' >
  sarrHtmlZeichen(128) = "&euro;"     ' €
  sarrHtmlZeichen(160) = "&nbsp;"     '
  sarrHtmlZeichen(161) = "&iexcl;"    ' ¡
  sarrHtmlZeichen(162) = "&cent;"     ' ¢
  sarrHtmlZeichen(163) = "&pound;"    ' £
  sarrHtmlZeichen(164) = "&curren;"   ' ¤
  sarrHtmlZeichen(165) = "&yen;"      ' ¥
  sarrHtmlZeichen(166) = "&brvbar;"   ' ¦
  sarrHtmlZeichen(167) = "&sect;"     ' §
  sarrHtmlZeichen(168) = "&uml;"      ' ¨
  sarrHtmlZeichen(169) = "&copy;"     ' ©
  sarrHtmlZeichen(170) = "&ordf;"     ' ª
  sarrHtmlZeichen(171) = "&laquo;"    ' «
  sarrHtmlZeichen(172) = "&not;"      ' ¬
  sarrHtmlZeichen(173) = "&shy;"      ' ­
  sarrHtmlZeichen(174) = "&reg;"      ' ®
  sarrHtmlZeichen(175) = "&macr;"     ' ¯
  sarrHtmlZeichen(176) = "&deg;"      ' °
  sarrHtmlZeichen(177) = "&plusmn;"   ' ±
  sarrHtmlZeichen(178) = "&sup2;"     ' ²
  sarrHtmlZeichen(179) = "&sup3;"     ' ³
  sarrHtmlZeichen(180) = "&acute;"    ' ´
  sarrHtmlZeichen(181) = "&micro;"    ' µ
  sarrHtmlZeichen(182) = "&para;"     ' ¶
  sarrHtmlZeichen(183) = "&middot;"   ' ·
  sarrHtmlZeichen(184) = "&cedil;"    ' ¸
  sarrHtmlZeichen(185) = "&sup1;"     ' ¹
  sarrHtmlZeichen(186) = "&ordm;"     ' º
  sarrHtmlZeichen(187) = "&raquo;"    ' »
  sarrHtmlZeichen(188) = "&frac14;"   ' ¼
  sarrHtmlZeichen(189) = "&frac12;"   ' ½
  sarrHtmlZeichen(190) = "&frac34;"   ' ¾
  sarrHtmlZeichen(191) = "&iquest;"   ' ¿
  sarrHtmlZeichen(192) = "&Agrave;"   ' À
  sarrHtmlZeichen(193) = "&Aacute;"   ' Á
  sarrHtmlZeichen(194) = "&Acirc;"    ' Â
  sarrHtmlZeichen(195) = "&Atilde;"   ' Ã
  sarrHtmlZeichen(196) = "&Auml;"     ' Ä
  sarrHtmlZeichen(197) = "&Aring;"    ' Å
  sarrHtmlZeichen(198) = "&AElig;"    ' Æ
  sarrHtmlZeichen(199) = "&Ccedil;"   ' Ç
  sarrHtmlZeichen(200) = "&Egrave;"   ' È
  sarrHtmlZeichen(201) = "&Eacute;"   ' É
  sarrHtmlZeichen(202) = "&Ecirc;"    ' Ê
  sarrHtmlZeichen(203) = "&Euml;"     ' Ë
  sarrHtmlZeichen(204) = "&Igrave;"   ' Ì
  sarrHtmlZeichen(205) = "&Iacute;"   ' Í
  sarrHtmlZeichen(206) = "&Icirc;"    ' Î
  sarrHtmlZeichen(207) = "&Iuml;"     ' Ï
  sarrHtmlZeichen(208) = "&ETH;"      ' Ð
  sarrHtmlZeichen(209) = "&Ntilde;"   ' Ñ
  sarrHtmlZeichen(210) = "&Ograve;"   ' Ò
  sarrHtmlZeichen(211) = "&Oacute;"   ' Ó
  sarrHtmlZeichen(212) = "&Ocirc;"    ' Ô
  sarrHtmlZeichen(213) = "&Otilde;"   ' Õ
  sarrHtmlZeichen(214) = "&Ouml;"     ' Ö
  sarrHtmlZeichen(215) = "&times;"    ' ×
  sarrHtmlZeichen(216) = "&Oslash;"   ' Ø
  sarrHtmlZeichen(217) = "&Ugrave;"   ' Ù
  sarrHtmlZeichen(218) = "&Uacute;"   ' Ú
  sarrHtmlZeichen(219) = "&Ucirc;"    ' Û
  sarrHtmlZeichen(220) = "&Uuml;"     ' Ü
  sarrHtmlZeichen(221) = "&Yacute;"   ' Ý
  sarrHtmlZeichen(222) = "&THORN;"    ' Þ
  sarrHtmlZeichen(223) = "&szlig;"    ' ß
  sarrHtmlZeichen(224) = "&agrave;"   ' à
  sarrHtmlZeichen(225) = "&aacute;"   ' á
  sarrHtmlZeichen(226) = "&acirc;"    ' â
  sarrHtmlZeichen(227) = "&atilde;"   ' ã
  sarrHtmlZeichen(228) = "&auml;"     ' ä
  sarrHtmlZeichen(229) = "&aring;"    ' å
  sarrHtmlZeichen(230) = "&aelig;"    ' æ
  sarrHtmlZeichen(231) = "&ccedil;"   ' ç
  sarrHtmlZeichen(232) = "&egrave;"   ' è
  sarrHtmlZeichen(233) = "&eacute;"   ' é
  sarrHtmlZeichen(234) = "&ecirc;"    ' ê
  sarrHtmlZeichen(235) = "&euml;"     ' ë
  sarrHtmlZeichen(236) = "&igrave;"   ' ì
  sarrHtmlZeichen(237) = "&iacute;"   ' í
  sarrHtmlZeichen(238) = "&icirc;"    ' î
  sarrHtmlZeichen(239) = "&iuml;"     ' ï
  sarrHtmlZeichen(240) = "&eth;"      ' ð
  sarrHtmlZeichen(241) = "&ntilde;"   ' ñ
  sarrHtmlZeichen(242) = "&ograve;"   ' ò
  sarrHtmlZeichen(243) = "&oacute;"   ' ó
  sarrHtmlZeichen(244) = "&ocirc;"    ' ô
  sarrHtmlZeichen(245) = "&otilde;"   ' õ
  sarrHtmlZeichen(246) = "&ouml;"     ' ö
  sarrHtmlZeichen(247) = "&divide;"   ' ÷
  sarrHtmlZeichen(248) = "&oslash;"   ' ø
  sarrHtmlZeichen(249) = "&ugrave;"   ' ù
  sarrHtmlZeichen(250) = "&uacute;"   ' ú
  sarrHtmlZeichen(251) = "&ucirc;"    ' û
  sarrHtmlZeichen(252) = "&uuml;"     ' ü
  sarrHtmlZeichen(253) = "&yacute;"   ' ý
  sarrHtmlZeichen(254) = "&thorn;"    ' þ
  sarrHtmlZeichen(255) = "&yuml;"     ' ÿ
  
  sTmp = sHtmlString
  
  For i = 1 To 255
    If Len(sarrHtmlZeichen(i)) > 0 Then
       sTmp = Replace(sTmp, sarrHtmlZeichen(i), Chr(i))
       sTmp = Replace(sTmp, "&#" & CStr(i) & ";", Chr(i))
    End If
  Next i
  
  HtmlZeichenConvert = sTmp
  
  ReDim sarrHtmlZeichen(1 To 1) As String
  Erase sarrHtmlZeichen()
End Function

Public Function IsoDate2Date(ByVal sIsoDate As String) As Date
    
    On Error Resume Next
    
    If sIsoDate Like "####-##-## ##:##:##" Then
        IsoDate2Date = myDateSerial(Mid(sIsoDate, 1, 4), Mid(sIsoDate, 6, 2), Mid(sIsoDate, 9, 2)) + myTimeSerial(Mid(sIsoDate, 12, 2), Mid(sIsoDate, 15, 2), Mid(sIsoDate, 18, 2))
    Else
        IsoDate2Date = CDate(sIsoDate)
    End If
    
    On Error GoTo 0
    
End Function

Public Function Date2IsoDate(ByVal datDate As Date) As String
    
    On Error Resume Next
    Date2IsoDate = Format$(datDate, "yyyy-mm-dd hh:mm:ss")
    On Error GoTo 0
    
End Function

Public Function UnixDate2Date(ByVal sUnixDate As String) As Date
    
    On Error Resume Next
    If Val(sUnixDate) = 0 Then
        UnixDate2Date = 0
    Else
        UnixDate2Date = DateAdd("s", CLng(sUnixDate), myDateSerial(1970, 1, 1))
        UnixDate2Date = DateAdd("n", GetUTCOffset() * 60, UnixDate2Date)
    End If
    On Error GoTo 0
    
End Function

Public Function Date2UnixDate(ByVal datDate As Date) As String
    
    On Error Resume Next
    If datDate = 0 Then
        Date2UnixDate = "0"
    Else
        datDate = DateAdd("n", GetUTCOffset() * -60, datDate)
        Date2UnixDate = CStr(DateDiff("s", myDateSerial(1970, 1, 1), datDate))
    End If
    On Error GoTo 0
    
End Function

Public Function FilterNumeric(ByVal sValue As String) As String
    
    Dim i As Integer
    
    For i = 1 To Len(sValue)
        If IsNumeric(Mid(sValue, i, 1)) Then
            FilterNumeric = FilterNumeric & Mid(sValue, i, 1)
        Else
            FilterNumeric = FilterNumeric & "-"
        End If
    Next i
    
End Function

Public Function GetNumericPart(ByVal sValue As String) As String
    
    Dim i As Integer
    
    sValue = Trim(sValue)
    For i = 1 To Len(sValue)
        If IsNumeric(Mid(sValue, i, 1)) Then
            If GetNumericPart > "" Or Mid(sValue, i, 1) <> "0" Then
                GetNumericPart = GetNumericPart & Mid(sValue, i, 1)
            End If
        Else
            Exit For
        End If
    Next i
    
End Function

Public Function GetOffsetLocalFromDate(datDate As Date) As Double
    
    Static DstSwitchDates(0 To 6)
    Dim i As Integer
    Dim a As Integer, b As Integer
    
    If DstSwitchDates(0) = 0 Then
        
        For i = -1 To 1
            a = i * 2 + 3
            b = i * 2 + 4
            DstSwitchDates(a) = myDateSerial(Year(datDate) + i, 3, 31) + myTimeSerial(2, 0, 0)
            DstSwitchDates(b) = myDateSerial(Year(datDate) + i, 10, 31) + myTimeSerial(3, 0, 0)
            Do Until Weekday(DstSwitchDates(a), vbMonday) = 7
                DstSwitchDates(a) = DateAdd("d", -1, DstSwitchDates(a))
            Loop
            
            Do Until Weekday(DstSwitchDates(b), vbMonday) = 7
                DstSwitchDates(b) = DateAdd("d", -1, DstSwitchDates(b))
            Loop
        Next i
    End If
    
    For i = 0 To 5
        If datDate > DstSwitchDates(i) And datDate < DstSwitchDates(i + 1) Then
            GetOffsetLocalFromDate = i Mod 2
            Exit For
        End If
    Next
    
End Function
