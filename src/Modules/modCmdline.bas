Attribute VB_Name = "modCmdline"
Option Explicit

' Unsere "eindeutige" Markierung des BOM Fensters
Private Const msTHISAPPID As String = "Biet-O-Matic"
'
Private mlMutexHwnd As Long
Private mlPreviousHwnd As Long
Private mbInDevelopment As Boolean


'
Public Const SMTO_NORMAL As Long = &H0
Public Const WM_COPYDATA As Long = &H4A
'
Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_RESTORE As Long = &HF120&
Private Const ERROR_ALREADY_EXISTS As Long = 183&
'
Public Type COPYDATASTRUCT
   dwData As Long
   cbData As Long
   lpData As Long
End Type
'
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long




Public Function ParseCommand(ByVal sCommand As String) As Long
Dim i As Integer
Dim sDummy As String
Dim sCommandTmp As String
Dim bFound As Boolean
Dim sTmp As String

Call DebugPrint("Parsing command: " & sCommand)

If frmDummy.Timer1.Enabled Then ' wir sind die 2. Instanz
  frmDummy.Timer1.Enabled = False
  Call Send(sCommand)
  frmDummy.TimerJetzt
  Exit Function
End If

On Error Resume Next

   ' Kommandozeilen-Syntax:

   ' ADD/MOD={Wert} [BID={Wert}][USER={Wert}][GROUP={Wert}][COMMENT={Wert}][SHIP={Wert}]
   ' DEL={Wert}
   ' URL={Wert}

sCommand = Trim$(sCommand)
sCommandTmp = UCase(sCommand)

''auf URL prüfen -> URL in BOM verarbeiten
'Pos = InStr(1, sCommandTmp, "URL=") + 4
'If Pos > 4 Then
''URL auswerten:
'    sDummy = Mid(sCommand, Pos, Len(sCommand) - Pos)
'    If sDummy <> "" Then
'        ParseURL sDummy, "", ""
'        Exit Sub
'    End If
'End If

'auf DEL prüfen -> Artikel löschen
sDummy = StripParameterValue(sCommand, "DEL")
If sDummy <> "" Then 'in sDummy steckt die Artikelnummer
    'Wir hauen einen Artikel raus
    sDummy = SepariereArtNr(sDummy) 'evtl. umgebenden Schrott entfernen
    sTmp = Trim(StripParameterValue(sCommand, "ONLYIFUSER"))
    
    'Artikel im Array suchen
    For i = 1 To UBound(gtarrArtikelArray)
        If gtarrArtikelArray(i).Artikel = sDummy And (sTmp = "" Or gtarrArtikelArray(i).UserAccount = sTmp) Then
            'wech damit ;-)
            frmHaupt.RemoveArtikel i
            Exit For
        End If
    Next i
    Exit Function
End If

'auf ADD prüfen -> Artikel hinzufügen/modifizieren
sDummy = StripParameterValue(sCommand, "ADD")
If sDummy <> "" Then 'in sDummy steckt die Artikelnummer
    'Wir filetieren die Artikelnummer:
    sDummy = SepariereArtNr(sDummy) 'umgebenden Schrott entfernen
    sTmp = Trim(StripParameterValue(sCommand, "ONLYIFUSER"))

    'Artikel im Array suchen
    For i = 1 To giAktAnzArtikel
        bFound = False
        If gtarrArtikelArray(i).Artikel = sDummy Then
            If gtarrArtikelArray(i).UserAccount = sTmp Or sTmp = "" Then
                bFound = True
                Exit For
            Else
                Exit Function
            End If
        End If
    Next i

    'Artikel nicht vorhanden? -> neuer Artikel musz angelegt werden
    If Not bFound Then
        i = frmHaupt.AddArtikel(sDummy)
    End If

    'in i steckt immer noch / wieder der Index ;-)

    If i > 0 Then

        'Gebot
        sDummy = StripParameterValue(sCommand, "BID")
        If sDummy <> "" Then gtarrArtikelArray(i).Gebot = String2Float(sDummy)
        'Account
        sDummy = StripParameterValue(sCommand, "USER")
        If sDummy <> "" Then gtarrArtikelArray(i).UserAccount = frmHaupt.GetAccountFromAccount(Trim(sDummy))
        'Gruppe
        sDummy = StripParameterValue(sCommand, "GROUP")
        If sDummy <> "" Then
          Dim alteGruppe As String
          alteGruppe = gtarrArtikelArray(i).Gruppe
          gtarrArtikelArray(i).Gruppe = Trim(sDummy)
          frmHaupt.CheckBietgruppe alteGruppe
          frmHaupt.CheckBietgruppe gtarrArtikelArray(i).Gruppe
        End If
        frmHaupt.CheckSofortkaufArtikel
        'Kommentar
        sDummy = StripParameterValue(sCommand, "COMMENT")
        If sDummy <> "" Then gtarrArtikelArray(i).Kommentar = Trim(sDummy)
        'Versand
        sDummy = StripParameterValue(sCommand, "SHIPPING")
        If sDummy <> "" Then gtarrArtikelArray(i).Versand = "*" & Trim(sDummy) ' der Stern bedeutet manuelle Versandkosten, sollen nicht mehr durch eBay überschrieben werden
        If gtarrArtikelArray(i).Versand = "*" Then gtarrArtikelArray(i).Versand = ""  ' Wenn gar nichts eigetragen, dann wieder für Automatik freischalten
        'Status
        sDummy = StripParameterValue(sCommand, "STATUS")
        If sDummy <> "" Then gtarrArtikelArray(i).Status = Val(Trim(sDummy))
        
        gtarrArtikelArray(i).LastChangedId = GetChangeID()
        
    End If
    
    frmHaupt.ArtikelArrayToScreen frmHaupt.VScroll1.Value
    Exit Function
End If

'auf UPDATE prüfen -> Artikel aktualisieren
sDummy = StripParameterValue(sCommand, "UPDATE")
If sDummy <> "" Then 'in sDummy steckt die Artikelnummer
    'Wir aktualisieren einen Artikel
    sDummy = SepariereArtNr(sDummy) 'evtl. umgebenden Schrott entfernen
    sTmp = Trim(StripParameterValue(sCommand, "ONLYIFUSER"))
    
    'Artikel im Array suchen
    For i = 1 To UBound(gtarrArtikelArray)
        If gtarrArtikelArray(i).Artikel = sDummy And (sTmp = "" Or gtarrArtikelArray(i).UserAccount = sTmp) Then
            'aktualisieren
            frmHaupt.Upd_Art i, vbNullString, False
            Exit For
        End If
    Next i
    Exit Function
End If

'auf AUTH prüfen -> Username und Passwort überprüfen
If InStr(1, sCommandTmp, "/AUTH") > 0 Then
    i = 0
    sDummy = Trim(StripParameterValue(sCommand, "USER"))
    If sDummy > "" Then
      i = frmHaupt.CheckPassForAccount(sDummy, Trim(StripParameterValue(sCommand, "PASS")))
    End If
    
    sDummy = StripParameterValue(sCommand, "CALLBACKHWND")
    If sDummy <> "" Then SendCallBackData sDummy, Trim(StripParameterValue(sCommand, "CALLBACKID")) & vbCrLf & CStr(i)
    
    Exit Function
End If

'auf SAVE prüfen -> Artikel und Einstellungen speichern
If InStr(1, sCommandTmp, "/SAVE") > 0 Then frmHaupt.SaveArtikel

'auf SHOW prüfen -> BOM zeigen, aktivieren
If InStr(1, sCommandTmp, "/SHOW") > 0 Then RestoreAndActivate frmHaupt.hWnd: Exit Function

'auf HIDE prüfen -> BOM minimieren
If InStr(1, sCommandTmp, "/HIDE") > 0 Then frmHaupt.WindowState = vbMinimized: Exit Function

'auf QUIT prüfen -> BOM beenden
If InStr(1, sCommandTmp, "/QUIT") > 0 Then frmHaupt.QuitTimer.Enabled = True: Exit Function

'auf SUSPEND prüfen -> Rechner schlafen legen
If InStr(1, sCommandTmp, "/SUSPEND") > 0 Then Suspend: Exit Function

'auf AUTOMODE prüfen -> Automatik ein/ausschalten
sDummy = StripParameterValue(sCommand, "AUTOMODE")
If sDummy <> "" Then gbAutoMode = CBool(sDummy): frmHaupt.CheckAutoMode: Exit Function

'auf DUMPUSER prüfen -> Alle User auf STDOUT dumpen
sDummy = StripParameterValue(sCommand, "CALLBACKHWND")
If sDummy <> "" And InStr(1, sCommandTmp, "/DUMPUSER") > 0 Then SendCallBackData sDummy, Trim(StripParameterValue(sCommand, "CALLBACKID")) & vbCrLf & DumpAllUser(): Exit Function

'auf DUMP prüfen -> Alle Artikel auf STDOUT dumpen (CSV)
sDummy = StripParameterValue(sCommand, "CALLBACKHWND")
If sDummy <> "" And InStr(1, sCommandTmp, "/DUMP") > 0 Then SendCallBackData sDummy, Trim(StripParameterValue(sCommand, "CALLBACKID")) & vbCrLf & DumpAllItems(StripParameterValue(sCommand, "FROMID")): Exit Function

'auf URL prüfen -> URL wie bei Drag&Drop behandeln
sDummy = StripParameterValue(sCommand, "URL")
If sDummy <> "" Then frmHaupt.HandleDragDropData 0, sDummy, 0, 0, 0, 0, 0: Exit Function

'auf HELP prüfen -> Hilfe Dialog anzeigen
Dim helptext As String
If InStr(1, sCommandTmp, "/HELP") > 0 Or InStr(1, sCommand, "/?") > 0 Then
    helptext = "Biet-O-Matic (Bid-O-Matic) Command Line Parameters" & vbCrLf & _
               "************************************************" & vbCrLf & vbCrLf & _
               "Add/modify an item" & vbCrLf & _
               "/ADD=<Item Number>" & vbTab & "[BID=<Bid Price>]" & vbCrLf & _
               vbTab & vbTab & vbTab & "[GROUP=<Group String>]" & vbCrLf & _
               vbTab & vbTab & vbTab & "[USER=<User Name>]" & vbCrLf & _
               vbTab & vbTab & vbTab & "[COMMENT=<Comment String>]" & vbCrLf & _
               vbTab & vbTab & vbTab & "[SHIPPING=<Shipping Costs>]" & vbCrLf & vbCrLf & _
               "Remove an item" & vbCrLf & _
               "/DEL=<Item Number>" & vbCrLf & vbCrLf & _
               "Update an item" & vbCrLf & _
               "/UPDATE=<Item Number>" & vbCrLf & vbCrLf & _
               "Switch Automode on/off" & vbCrLf & _
               "/AUTOMODE=<1|0>" & vbCrLf & vbCrLf & _
               "Show BOM-Window if minimized" & vbCrLf & _
               "/SHOW" & vbCrLf & vbCrLf & _
               "Save items" & vbCrLf & _
               "/SAVE" & vbCrLf & vbCrLf & _
               "Quit BOM" & vbCrLf & _
               "/QUIT" & vbCrLf & vbCrLf & _
               "Dump all items on stdout as csv:" & vbCrLf & _
               "/DUMP" & vbCrLf & vbCrLf & _
               "Handle an URL" & vbCrLf & _
               "/URL=<URL String>" & vbCrLf & vbCrLf
    helptext = helptext & "Show this help dialog:" & vbCrLf & _
           "/HELP | /?" & vbCrLf & vbCrLf & _
           "-------------------------------------------------------------------" & vbCrLf & _
           "Examples:" & vbCrLf & _
           "Biet-O-Matic.exe /ADD=12345678 BID=76,50 GROUP=a;1 USER=Hans_Dampf" & vbCrLf & _
           "Biet-O-Matic.exe /ADD=12345678 COMMENT=""My first via command line added comment!""" & vbCrLf & _
           "Biet-O-Matic.exe /DEL=12345678" & vbCrLf
    MsgBox helptext, vbInformation, gsarrLangTxt(215) & " - Command Line Help"
End If

End Function

'Beseitigt unnötige Zeichen rings um die Artikelnummer
'-> nur anwendbar, wenn die Artikelnummer am Anfang durch
'nichtnummerische Zeichen eingeschlossen ist
Private Function SepariereArtNr(ByVal sArtNr As String) As String

    On Error GoTo hell
    Dim i As Integer
    Dim iPosAnf As Integer
    Dim iPosEnd As Integer
    
    'erstes numerisches Zeichen finden
    For i = 1 To Len(sArtNr) - 1
        If IsNumeric(Mid(sArtNr, i, 1)) Then
            iPosAnf = i
            Exit For
        End If
    Next i
    
    If iPosAnf > 0 Then 'hey wir haben einen Anfang
        'letztes numerisches Zeichen finden
        For i = 1 To Len(sArtNr) - iPosAnf
            If Not IsNumeric(Mid(sArtNr, iPosAnf + i, 1)) Then
                iPosEnd = iPosAnf + i
                Exit For
            End If
        Next i
        
        'jetzt die Nummer filetieren
        If iPosEnd = 0 Then iPosEnd = Len(sArtNr) + 1
        SepariereArtNr = Mid$(sArtNr, iPosAnf, iPosEnd - iPosAnf)
    Else
        SepariereArtNr = ""
    End If
    
Exit Function

hell:
Call DebugPrint("Fehler in Prozedur 'SepariereArtNr': " & Err.Description)
SepariereArtNr = ""
End Function

'liest den entsprechenden Wert für einen Parameter aus dem
'Kommandozeilen-String aus
Private Function StripParameterValue(ByVal sTxt As String, ByVal sParam As String) As String

    Dim iPos As Integer
    Dim sDummy As String
    Dim sSearchChar As String
    
    'On Error GoTo hell
    
    If Len(sTxt) > 0 Then
        iPos = InStr(1, sTxt, sParam & "=", vbTextCompare) + Len(sParam) + 1
        If iPos > Len(sParam) + 1 Then 'Parameter gibt es im String
            'alles vorher abschneiden:
            sTxt = Mid$(sTxt, iPos, Len(sTxt) - iPos + 1)
            
            If Left(sTxt, 1) = """" Then
                sSearchChar = """"
                iPos = 1
            Else
                sSearchChar = " "
                iPos = 0
            End If
            
weitersuchen:
        
            iPos = InStr(iPos + 1, sTxt, sSearchChar)
            If iPos > 1 Then
                'escaped? -> dann weitersuchen
                If Mid(sTxt, iPos - 1, 1) = "\" Then GoTo weitersuchen
            ElseIf iPos = 1 And sSearchChar = " " Then
                iPos = 0
            Else
                iPos = Len(sTxt) + 1
            End If
            
            sDummy = Trim(Mid$(sTxt, 1, iPos))
            'Anführungszeichen entfernen
            If Left(sDummy, 1) = """" And Right(sDummy, 1) = """" Then sDummy = Mid(sDummy, 2, Len(sDummy) - 2)
            sDummy = Replace(sDummy, "\ ", " ")
            sDummy = Replace(sDummy, "\""", """")
            sDummy = Replace(sDummy, "\\", "\")
            StripParameterValue = sDummy & " "
        End If
    End If
Exit Function

hell:
Call DebugPrint("Fehler in Prozedur 'StripParameterValue': " & Err.Description)
StripParameterValue = ""
End Function

Public Sub ProcessCmdline()
    
    Dim tCDS As COPYDATASTRUCT, b() As Byte, lR As Long
    
    Load frmDummy
    Call modSubclass.Subclass(frmDummy.hWnd)
    
    'wir übergeben die empfangenen Parameter an die schon laufende Instanz:
    
    'First try to find it:
    Call EnumerateWindows
    
    'If we get it:
    If (mlPreviousHwnd <> 0) Then
    
        'Send.  The app must subclass the WM_COPYDATA message
        'to get this information:
        
        b() = StrConv(Command() & " /CALLBACKHWND=" & CStr(frmDummy.hWnd), vbFromUnicode)
        tCDS.dwData = 0
        tCDS.cbData = UBound(b) + 1
        tCDS.lpData = VarPtr(b(0))
        
        frmDummy.Timer1.Enabled = True
        
        'Give in if the existing app is not responding:
        lR = SendMessageTimeout(mlPreviousHwnd, WM_COPYDATA, 0, tCDS, SMTO_NORMAL, 5000, lR)
        
    End If
    
End Sub

Private Sub RestoreAndActivate(ByVal lHwnd As Long)
    
    If Not (IsIconic(lHwnd) = 0) Then
        Call SendMessageByLong(lHwnd, WM_SYSCOMMAND, SC_RESTORE, 0)
    End If
    Call ActivateWindow(lHwnd)
    
End Sub

Public Sub TagWindow(ByVal lHwnd As Long)
    
    'Applies a window property to allow the window to
    'be clearly identified.
    Call SetProp(lHwnd, msTHISAPPID & "_APPLICATION", 1)
    
End Sub

Private Function IsThisApp(ByVal lHwnd As Long) As Boolean
    
    'Check if the windows property is set for this
    'window handle:
    If GetProp(lHwnd, msTHISAPPID & "_APPLICATION") = 1 Then
        IsThisApp = True
    End If
End Function

Private Function EnumWindowsProc(ByVal lHwnd As Long, ByVal lParam As Long) As Long

    Dim bStop As Boolean
    'Customised windows enumeration procedure.  Stops
    'when it finds another application with the Window
    'property set, or when all windows are exhausted.
    bStop = False
    If IsThisApp(lHwnd) Then
        EnumWindowsProc = 0
        mlPreviousHwnd = lHwnd
    Else
        EnumWindowsProc = 1
    End If
    
End Function

Private Function EnumerateWindows() As Boolean
    'Enumerate top-level windows:
    EnumWindows AddressOf EnumWindowsProc, 0
End Function

Private Sub ActivateWindow(ByVal lHwnd As Long)

    Call SetForegroundWindow(lHwnd)
    Call BringWindowToTop(lHwnd)
    
End Sub

Public Function InDevelopment() As Boolean
    
    'Debug.Assert code not run in an EXE.  Therefore
    'mbInDevelopment variable is never set.
    Debug.Assert InDevelopmentHack() = True
    InDevelopment = mbInDevelopment
    
End Function

Private Function InDevelopmentHack() As Boolean
    
   mbInDevelopment = True
   InDevelopmentHack = mbInDevelopment
   
End Function

Public Function WeAreAlone(Optional ByVal sMutex As String) As Boolean
    
    If sMutex = "" Then sMutex = msTHISAPPID & "_APPLICATION_MUTEX"
    
    'Don't call Mutex when in VBIDE because it will apply
    'for the entire VB IDE session, not just the app's session.
    If InDevelopment Then
        WeAreAlone = Not (App.PrevInstance)
    Else
        'Ensures we don't run a second instance even
        'if the first instance is in the start-up phase
        mlMutexHwnd = CreateMutex(ByVal 0&, 1, sMutex)
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
            Call CloseHandle(mlMutexHwnd)
        Else
            'WeAreAlone = True
            'manchmal scheitert die Mutex-Prüfung, darum hier noch einmal
            'auf App.PrevInstance prüfen (mae 280708)
            WeAreAlone = Not (App.PrevInstance)
        End If
    End If
    
End Function

Private Function EndApp()
    'MD-Marker , Function wird nicht aufgerufen
    'Call this to remove the Mutex.  It will be cleared anyway by windows, _
    but this ensures it works.
    If (mlMutexHwnd <> 0) Then
        Call CloseHandle(mlMutexHwnd)
    End If
    mlMutexHwnd = 0
End Function

Public Function IsJobCommand() As Boolean
    
    If UCase(Command()) Like "*/JOB:*.JOB*" Then IsJobCommand = True
    
End Function

