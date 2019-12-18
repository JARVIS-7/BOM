Attribute VB_Name = "modMsgBoxEx"
Option Explicit

' ===================================================================
'
'     --------------
'     MsgBoxEx (GES)
'     --------------
'
'  MessageBox mit Benutzer-definierter Buttonbeschriftung
'
' Die Vorteile:
'
' - Timer werden nicht angehalten
' - Msg-Text kann zentriert werden
' - einfaches Einbinden einer Hilfetaste
' - Buttons können frei beschriftet werden
' - Icons werden auf beiden Seiten angezeigt
' - Hintergrundfarbe kann frei gewählt werden *
' - Schriftfarbe kann frei gewählt werden     *
' - und vieles mehr ...
'
'   *  Diese Eigenschaften können eingestellt werden,
'      indem sie der Msg-Form zugewiesen werden.
'
' Autor:
' Guido Eisenbeis Software (GES), guidoeisenbeis@web.de, 2004-08-28
'
' Copyright:
' "MsgBoxEx (GES) © 2004" ist Freeware und darf frei benutzt werden,
' solange der Copyright-Hinweis in "basMsgBoxEx" erhalten bleibt.

' ===================================================================


' -----------------
' Weitere Features:
' -----------------
'
' (Fast) alle Features der "normalen" MsgBox, z.B.:
'
' - immer im Vordergrund
' - systemeigene Icons und Signaltöne
' - Schließen-Button (X-Button) ausschalten
' - Buttons können mit ShortCuts belegt werden (ALT + ...)
' - weiterschalten der Buttons mit Tab- und Pfeil-Tasten
' - Buttons können mit der Enter-Taste betätigt werden
' - automatisches Anpassen der MaxWidth an Desktop-Auflösung
' - ...

' ----------------------------------
' Was die MsgBoxEx (GES) nicht kann:
' ----------------------------------
'
' - übersetzen der Buttonbeschriftungen in andere Sprachen
' - right-to-left anzeigen auf hebr. und arab. Systemen


' -----------
' Handhabung:
' -----------

' - "frmMsgBoxEx" und "basMsgBoxEx" in ein beliebiges
'   Projekt einbinden --> MsgBoxEx aufrufen


' -------
' Aufruf:
' -------

' Der Aufruf ist dem der normalen MsgBox sehr ähnlich. Die
' meisten Argumente sind optional, dadurch ist der Aufruf
' sehr einfach.

' Beispiele:

'' 1) einfachster Aufruf
'Private Sub Command1_Click()
'   MsgBoxEx "Hier steht mein Msg-Text"
'End Sub
'
'' 2) Test für Rückgabewert
'Private Sub Command2_Click()
'   Dim ret As Long
'
'   ret = MsgBoxEx("Der Rückgabewert wird von links nach " _
'            & "rechts gezählt:" & vbNewLine & vbNewLine & _
'            "Button 1, Button 2, ...  Cancel (X-Button) = 0", _
'            "&Button 1- Bu&tton 2*- Butto&n 3- Butt&on 4-", _
'            Icon_Question, "Hier steht mein Titel-Text")
'
'   MsgBoxEx "Rückgabewert:  " & ret
'   Debug.Print "Rückgabewert:  " & ret
'End Sub
'
'' 3) Benutzer-definierte Schrift- und Hintergrund-Farbe
'Private Sub Command3_Click()
'
'   ' Hintergrundfarbe zuweisen
'   frmMsgBoxEx.BackColor = vbYellow
'
'   ' Schriftfarbe zuweisen
'   frmMsgBoxEx.ForeColor = vbRed
'
'   MsgBoxEx "Schrift- und Hintergrund-Farbe können beliebig " & _
'            "gewählt werden," & vbNewLine & vbNewLine & _
'            "um z.B. wichtige Hinweise hervorzuheben.", _
'            "&Abbrechen-&Ignorieren-&Ja stimmt*-&Wiederholen-", _
'            Icon_Exclamation
'End Sub


'    ----------------------
'    Button-Caption  Hilfe:
'    ----------------------
'
' 1) Es sind maximal 4 Buttons möglich (plus 1 Hilfe-Button).
'
' 2) Jede Button-Caption (auch die letzte!) muss mit
'    einem Minuszeichen ''-'' abgeschlossen werden.
'
'    Beispiel:   Button 1- Button 2- Button 3-
'
' 3) Der Default-Button kann rechts mit einem
'    Sternchen * markiert werden.
'
'    Beispiel:   Button 1- Button 2*- Button 3-
'
' 4) ShortCuts können ''ganz normal'' mit
'    einem ''&'' markiert werden.
'
'    Beispiel:   &Button 1- Bu&tton 2*- Butto&n 3-
'
' Leerzeichen sind erlaubt.
'
' Wird keine Button-Caption eingegeben, wird
' automatisch ein ''OK''-Button gezeigt.
'
' Der Rückgabewert wird von links nach rechts gezählt:
'    Button 1 = 1,  Button 2 = 2,  ... Cancel (X) = 0


' -------------------------------
' In "frmMsgBoxEx" wird benötigt:
' -------------------------------
'
' - 1x Command-Button (Name = "cButton", Index = 0)
'
' - und folgender Code:
'
'' Original-Code (ohne "Skin") ===================================== v
'
'' Funktionen zum Öffnen von HTML-Help (.chm-Dateien)
'Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
'Private Declare Function HtmlHelpTopic Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As String) As Long
'Private Const HH_DISPLAY_TOPIC = &H0
'
'' Funktion zum Öffnen von beliebigen Dateien
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
'' Funktion zum Schließen von Fenstern
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Const WM_CLOSE = &H10
'Private HelpHWND As Long
'
'
'Public MBExHelpFile As String ' Pfad zur Hilfe-Datei
'Public MBExParams As String   ' zusätzliche Parameter
'Public MBExRetval As Integer  ' Rückgabe-Wert für "MsgBoxEx"
'
'
'Private Sub cButton_Click(Index As Integer)
'
'   ' falls der Hilfe-Button geklickt wurde *
'   If cButton(Index).Tag = "MBExHelpButton" Then
'      ' Hier Befehle für "Hilfe" einfügen
'      Call ShowMBExHelp
'   Else
'      MBExRetval = Index
'      Unload Me
'   End If
'End Sub
''
''  * Hinweis!
''  Die Tag-Eigenschaften der Form und
''  des Hilfe-Buttons können nicht benutzt
''  werden, da sie schon in Verwendung sind!
'
'' Beispiel für den Aufruf einer Hilfe-Datei
'Private Sub ShowMBExHelp()
'
'   On Error Resume Next ' Fehlerbehandlung aus
'
'   If MBExHelpFile = "" Then Exit Sub
'   If (LCase(Right$(MBExHelpFile, 4)) = ".chm") And (MBExParams <> "") Then
'      ' Html-Hilfe mit einer bestimmten Seite öffnen
'      HelpHWND = HtmlHelpTopic(0, MBExHelpFile, _
'                  HH_DISPLAY_TOPIC, MBExParams)
'   Else
'      ' Html-Hilfe mit Startseite, oder beliebige Datei
'      ' (.txt, .doc, .htm ...) mit oder ohne Parameter öffnen
'      Call ShellExecute(Me.hWnd, "Open", _
'         MBExHelpFile, MBExParams, "", 1)
'   End If
'
'   On Error GoTo 0 ' Fehlerbehandlung ein
'
'End Sub
'
'' Steurerung per Tastatur ermöglichen
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim B As CommandButton
'   ' F1-Taste wurde gedrückt
'   If KeyCode = vbKeyF1 Then
'      KeyCode = 0
'      ' Hilfe aufrufen, falls vorhanden
'      For Each B In Me
'         If B.Tag = "MBExHelpButton" Then
'            B.SetFocus
'            B = True
'            Exit For
'         End If
'      Next B
'   ' Escape-Taste wurde gedrückt
'   ElseIf KeyCode = vbKeyEscape Then
'      KeyCode = 0
'      ' "frmMsgBoxEx" schließen, falls erlaubt
'      If Me.Tag <> "NoCloseButton" Then Unload Me
'   End If
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   ' Wenn eine Hilfe-Datei mit einer bestimmten Seite und
'   ' "HtmlHelp" geöffnet wird, muss das Fenster der Hilfe
'   ' geschlossen sein, bevor das Programm beendet wird.
'   ' Ansonsten erfolgt eine Speicherzugriffsverletzung.
'   If HelpHWND Then
'      SendMessage HelpHWND, WM_CLOSE, 0&, 0&
'      HelpHWND = 0
'   End If
'End Sub
'' ================================================================= ^


' ===================================================================
'
'  Info: MaxWidth einer "normalen" MsgBox:
'
'                                      MsgBox + Frei = Gesamt
'                             --------------------------------------
'                             |                                    |
'   Monitor     Auflösung           Twips             prozentual
'  ---------   ------------   -------------------   ----------------
'   21 Zoll     1024 x 768    9570 + 5805 = 15375   62% + 38% = 100%
'   15 Zoll      600 x 800    7395 + 4575 = 11970   62% + 38% = 100%
'
' ===================================================================


' --------------------------
' MsgBoxEx (GES) Enum-Types:
' --------------------------

Public Enum GES_ShowStyle
   ' Icons und Töne
   Icon_Critical = 8        ' Stop
   Icon_Question = 16       ' Fragezeichen
   Icon_Exclamation = 32    ' Ausrufezeichen
   Icon_Information = 64    ' Information
   ' Extended-Style
   Show_AlwaysOnTop = 128   ' immer im Vordergrund
   Show_HelpButton = 256    ' Hilfe-Button zeigen
   Show_NoCloseButton = 512 ' Schliessen-Button disablen
   ' Text-Alignment
   Text_Left = 1            ' linksbündig
   Text_Right = 2           ' rechtsbündig
   Text_Center = 4          ' zentriert
End Enum


' ---------------------------------
' Funktionen, Typen und Konstanten:
' ---------------------------------

' Ermitteln von Fenster-Schriften
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
'
Private Const LOGPIXELSY                As Long = &H5A&
'
'Private Const LNG_FONT_CAPTION          As Long = &H1&
'Private Const LNG_FONT_MENU             As Long = &H2&
'Private Const LNG_FONT_MESSAGE          As Long = &H3&
'Private Const LNG_FONT_SMCAPTION        As Long = &H4&
'Private Const LNG_FONT_STATUS           As Long = &H5&
'
Public Enum SystemFontEnum
    fntCaption = &H1& 'LNG_FONT_CAPTION
    fntMenu = &H2& 'LNG_FONT_MENU
    fntMessage = &H3& 'LNG_FONT_MESSAGE
    fntSmCaption = &H4& 'LNG_FONT_SMCAPTION
    fntStatus = &H5& 'LNG_FONT_STATUS
End Enum
'
Private Const LF_FACESIZE               As Long = &H20&
'
Private Type LOGFONT
    lfHeight                            As Long
    lfWidth                             As Long
    lfEscapement                        As Long
    lfOrientation                       As Long
    lfWeight                            As Long
    lfItalic                            As Byte
    lfUnderline                         As Byte
    lfStrikeOut                         As Byte
    lfCharSet                           As Byte
    lfOutPrecision                      As Byte
    lfClipPrecision                     As Byte
    lfQuality                           As Byte
    lfPitchAndFamily                    As Byte
    lfFaceName                          As String * LF_FACESIZE
End Type
'
Private Const SPI_GETNONCLIENTMETRICS   As Long = &H29&
'
Private Type NONCLIENTMETRICS
    cbSize                              As Long
    iBorderWidth                        As Long
    iScrollWidth                        As Long
    iScrollHeight                       As Long
    iCaptionWidth                       As Long
    iCaptionHeight                      As Long
    lfCaptionFont                       As LOGFONT
    iSMCaptionWidth                     As Long
    iSMCaptionHeight                    As Long
    lfSMCaptionFont                     As LOGFONT
    iMenuWidth                          As Long
    iMenuHeight                         As Long
    lfMenuFont                          As LOGFONT
    lfStatusFont                        As LOGFONT
    lfMessageFont                       As LOGFONT
End Type


' Ausgeben von Text
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
'
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
'
'Private Type DRAWTEXTPARAMS
'  cbSize As Long
'  iTabLength As Long
'  iLeftMargin As Long
'  iRightMargin As Long
'  uiLengthDrawn As Long
'End Type
'
' DrawText() Format Flags
'Private Const DT_TOP  As Long = &H0
Private Const DT_LEFT  As Long = &H0
Private Const DT_CENTER  As Long = &H1
Private Const DT_RIGHT  As Long = &H2
'Private Const DT_VCENTER  As Long = &H4
'Private Const DT_BOTTOM  As Long = &H8
'Private Const DT_SINGLELINE  As Long = &H20
Private Const DT_EXPANDTABS  As Long = &H40
Private Const DT_TABSTOP  As Long = &H80
'Private Const DT_EXTERNALLEADING  As Long = &H200
'Private Const DT_CALCRECT  As Long = &H400
'Private Const DT_INTERNAL  As Long = &H1000
Private Const DT_NOCLIP As Long = &H100
Private Const DT_NOPREFIX As Long = &H800
'Private Const DT_HIDEPREFIX As Long = &H100000 ' Only Windows 2000/XP
'Private Const DT_PREFIXONLY As Long = &H200000 ' Only Windows 2000/XP
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_EDITCONTROL As Long = &H2000
Private Const DT_RTLREADING As Long = &H20000
'Private Const DT_MODIFYSTRING As Long = &H10000
'Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
'Private Const DT_END_ELLIPSIS As Long = &H8000
'Private Const DT_PATH_ELLIPSIS As Long = &H4000
'Private Const DT_WORD_ELLIPSIS As Long = &H40000


' Positionieren und Anzeigemodus einer Form
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'
' SetWindowPos Flags
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOZORDER = &H4
'Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20   ' The frame changed: send WM_NCCALCSIZE
'Private Const SWP_SHOWWINDOW = &H40
'Private Const SWP_HIDEWINDOW = &H80
'Private Const SWP_NOCOPYBITS = &H100
'Private Const SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
''
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
'
' SetWindowPos() hwndInsertAfter values
Private Const HWND_TOP As Long = 0&
'Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST As Long = -1&
'Private Const HWND_NOTOPMOST = -2

' Beispiel:
'   ' Set the window position to topmost
'   SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
'         SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE


' Ermitteln der aktiven (aufrufenden) Form, die dann als
' Besitzer zum modalen Aufrufen der Msg-Form benutzt wird.
Private Declare Function GetActiveWindow Lib "user32" () As Long


' MsgBox-Icon laden und zeichnen
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
'
' MsgBox-Icon Konstanten
'Private Const IDI_APPLICATION = 32512&       ' Applikation
Private Const IDI_HAND As Long = 32513&             ' Stop
Private Const IDI_QUESTION As Long = 32514&         ' Fragezeichen
Private Const IDI_EXCLAMATION As Long = 32515&      ' Ausrufezeichen
Private Const IDI_ASTERISK As Long = 32516&         ' Information
'Private Const IDI_WINLOGO = 32517            ' Windows-Logo (XP: Applikation)
'Private Const IDI_ERROR = IDI_HAND           ' Stop
'Private Const IDI_WARNING = IDI_EXCLAMATION  ' Ausrufezeichen
'Private Const IDI_INFORMATION = IDI_ASTERISK ' Information


' MsgBox Töne abspielen
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
'
' MsgBox-Töne Konstanten
Private Const MB_ICONHAND As Long = &H10&        ' Stop
Private Const MB_ICONQUESTION As Long = &H20&    ' Fragezeichen
Private Const MB_ICONEXCLAMATION As Long = &H30& ' Ausrufezeichen
Private Const MB_ICONASTERISK As Long = &H40&    ' Information
'Private Const MB_Beep As Long = 0&               ' normaler Beep (Blub)


' Zu "RemoveSysMenu"
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
'
' System Menu Command Values
Private Const SC_RESTORE = &HF120&  ' Wiederherstellen
'Private Const SC_MOVE = &HF010&     ' Verschieben
Private Const SC_SIZE = &HF000&     ' Grösse verändern
Private Const SC_MINIMIZE = &HF020& ' Minimieren
Private Const SC_MAXIMIZE = &HF030& ' Maximieren
Private Const SC_CLOSE = &HF060&    ' Schliessen
'
Private Const MF_BYCOMMAND As Long = &H0
Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_STRING As Long = &H0&



' Hilfe-Hinweis für Button-Captions
' (bei Fehleingabe automatisch anzeigen)
Private Sub BtnCapsHelp()
    MsgBoxEx _
        vbNewLine & "        Button-Caption Hilfe:" & _
        vbNewLine & "        -------------------------" & _
        vbNewLine & vbNewLine & "1) Es sind maximal 4 Buttons" & _
        " möglich  (plus 1 Hilfe-Button)." & vbNewLine & vbNewLine _
        & vbNewLine & "2) Jede Button-Caption (auch die letzte!) " & _
        "muss mit einem Minuszeichen ''-'' abgeschlossen werden." & _
        vbNewLine & vbNewLine & _
        "Beispiel:     Button 1- Button 2- Button 3-" & vbNewLine & _
        vbNewLine & vbNewLine & "3) Der Default-Button kann " & _
        "rechts mit einem Sternchen * markiert werden." & vbNewLine & _
        vbNewLine & "Beispiel:     Button 1- Button 2*- Button 3-" & _
        vbNewLine & vbNewLine & vbNewLine & _
        "4) ShortCuts können ''ganz normal'' mit einem ''&'' " & _
        "markiert werden." & vbNewLine & vbNewLine & _
        "Beispiel:     &Button 1- Bu&tton 2*- Butto&n 3-" & _
        vbNewLine & vbNewLine & vbNewLine & _
        "Wird keine Button-Caption eingegeben, wird automatisch " & _
        "ein ''OK''-Button gezeigt." & vbNewLine & vbNewLine & _
        "Der Rückgabewert wird von links nach rechts gezählt:  " & _
        "Button 1 = 1,  Button 2 = 2,  ... Cancel (X) = 0", , _
        Icon_Information Or Text_Left _
        , " MsgBoxEx (GES)    Button-Caption Hilfe"
End Sub

Private Sub HelpFile_Help()
    MsgBoxEx _
        vbNewLine & "     HelpFile Hilfe" & vbNewLine & _
        "     ----------------" & vbNewLine & vbNewLine & _
        "Mit dem Hilfe-Button kann eine beliebige Hilfe-Datei " & _
        "(.txt, .doc, .chm ...) aufgerufen werden. Dazu können " & _
        "in ''HelpFile'' der Pfad zur Hilfe-Datei und in " & _
        "''Context'' zusätzliche Parameter angegeben werden." & _
        vbNewLine & vbNewLine & vbNewLine & "WICHTIG!" & vbNewLine & _
        vbNewLine & "Der Pfad zur Hilfe-Datei muss mit einem " & _
        "Minuszeichen ''-'' abgeschlossen werden! Fehlt das " & _
        "Minuszeichen, wird diese Hilfe aufgerufen." & vbNewLine & _
        vbNewLine & "Beispiel:     MsgBoxEx ''Hallo Welt'', , , , " & _
        "App.Path & ''\MeineHilfe.chm-''" & vbNewLine & vbNewLine & _
        vbNewLine & vbNewLine & "In ''frmMsgBoxEx'' können " & _
        "''HelpFile'' (MBExHelpFile) und ''Context'' (MBExParams) " & _
        "zum Aufrufen einer Hilfe-Datei verwendet werden." & _
        vbNewLine & vbNewLine & _
        "Beispiel:     If MBExHelpFile = '''' Then Exit Sub" & _
        vbNewLine & _
        "                  ShellExecute Me.hwnd, ''Open'', " & _
        "MBExHelpFile, MBExParams, '''', 1" & vbNewLine, _
        , Icon_Exclamation Or Text_Left, _
        "MsgBoxEx (GES)   HelpFile Hilfe                           " & _
        "Minuszeichen fehlt!"
End Sub

                         
Public Function MsgBoxEx(MsgText As Variant, _
                Optional Buttons As String, _
                Optional ShowStyle As GES_ShowStyle, _
                Optional Title As Variant, _
                Optional HelpFile As Variant, _
                Optional Context As Variant) As Integer
                         
   ' Text(-Ausgabe)-Bereich
   Dim sLine() As String           ' einzelne Zeile, getrennt an orig. Umbrüchen
   Dim MARGIN_W As Long            ' Breite des Textrands (mit/ohne Icons)
   Dim TextWidth As Long           ' endgültige Ausgabe-Weite
   Dim TextHeight As Long          ' endgültige Ausgabe-Höhe
   Dim TextRc As RECT              ' Ausgabe-Rechteck
   Dim TextAlign As Long           ' DrawText-Flag für Msg-Text-Ausrichtung
   
   ' Msg-Form
   Dim F As Form
   Dim WinWidth As Long            ' endgültige Form-Weite
   Dim WinHeight As Long           ' endgültige Form-Höhe
   Dim MinWinWidth As Long         ' mindest Form-Weite
   Dim MaxWinWidth As Long         ' maximale Form-Weite
   Dim MinWinHeight As Long        ' mindest Form-Höhe
   Dim BorderWidth As Long         ' = 2x Rahmenbreite
   Dim BorderHeight As Long        ' = 1x Titelleiste + 1x Rahmenbreite
   Dim CaptionWidth As Long        ' Weite der Titelzeile mit/ohne Border
   Dim sCaption As String          ' temp. Variable für Titel-Text
   Dim AW_hWnd As Long             ' temp. hWnd für OwnerForm-Suche
   Dim OwnerForm As Form           ' Elternfenster für "frmMsgBoxEx"
   
   ' Buttons
   Const ButtonWidth As Long = 76  ' einzelne Button-Weite
   Const ButtonHeight As Long = 24 ' einzelne Button-Höhe
   Dim AllBtsWidth As Long         ' Gesamtweite aller Buttons + Abstände
   Dim BtnCaps() As String         ' Array das die Button-Captions enthält
   Dim DefBtnNo As Integer         ' Index-Nr zum Setzen des Default-Buttons
   
   ' Icons
   Dim hIcon As Long               ' hWnd des System-Icons
   Dim MB_Icon As Long             ' ausgewähltes Icon
   Dim MB_Sound As Long            ' zugehöriger Sound
   Dim ShowIcon As Boolean         ' Icon anzeigen?
   
   ' allgemein
   Dim X As Integer
   
   
   ' Eltern-Fenter ermitteln (aufrufendes Fenster),
   ' für das modale Anzeigen der Msg-Form.
   AW_hWnd = GetActiveWindow
   For Each F In Forms
      If F.hWnd = AW_hWnd Then
         Set OwnerForm = F
         Exit For
      End If
   Next F
      
   With frmMsgBoxEx
      
      ' Variablen vorbelegen
      .MBExHelpFile = ""
      .MBExParams = ""
      .MBExRetval = 0
   
      ' Standard-Einstellung für den Titelzeilen-Text hier setzen,
      ' macht den MsgBoxEx-Aufruf (ToolTip) übersichtlicher.
      If IsMissing(Title) Then Title = App.Title
      
      ' wurde eine Hilfe-Datei angegeben?
      If IsMissing(HelpFile) Then
         HelpFile = ""
      Else
         ' alles OK, falls mit Minuszeichen abgeschlossen ...
         If Right$(HelpFile, 1) = "-" Then
            ' Minuszeichen rausfiltern
            HelpFile = Left$(HelpFile, Len(HelpFile) - 1)
         Else
            ' ... ansonsten Hilfe anzeigen
            Call HelpFile_Help
            Exit Function
         End If
      End If

      ' Msg-Form einrichten (Grund-Einstellungen)
      .AutoRedraw = True
      .BorderStyle = vbFixedDouble   ' "3 - Fester Dialog"
      .Caption = Title               ' Wichtig! (damit die Änderung
                                     ' für BorderStyle übernommen wird)
      Set .Icon = Nothing            ' SystemMenü-Icon entfernen
                                     ' (muss danach stehen)
      .ScaleMode = vbPixels          ' auf Pixels einstellen
      ' System-Schrift für Msg-Text zuweisen
      Set .Font = SystemGetFont(.hdc, fntMessage)
      .KeyPreview = True             ' Tasteneingabe für Form ermöglichen
   
      ' Button-Grundeinstellungen einrichten
      Set .cButton(0).Font = .Font
      .cButton(0).Visible = False
      .cButton(0).Width = ButtonWidth
      .cButton(0).Height = ButtonHeight
      
      ' falls Button-Captions eingegeben wurden
      If Buttons <> "" Then
         ' prüfen ob mit einem Minuszeichen abgeschloßen wurde
         If Right$(Buttons, 1) = "-" Then
            BtnCaps = Split(Buttons, "-", , vbTextCompare)
            ' prüfen ob maximal 4 Button-Captions eingegeben wurden
            If UBound(BtnCaps) <= 4 Then
               ' alle Buttons laden und beschriften
               For X = 1 To UBound(BtnCaps)
                  Load .cButton(X)
                  .cButton(X).Visible = True
                  .cButton(X).Caption = Trim$(BtnCaps(X - 1))
                  ' Default-Button setzen, falls angegeben
                  If Right$(.cButton(X).Caption, 1) = "*" Then
                     DefBtnNo = X
                     .cButton(X).Default = True
                     ' Sternchen "*" rausfiltern
                     .cButton(X).Caption = _
                        Trim$(Left$(.cButton(X).Caption, _
                           Len(.cButton(X).Caption) - 1))
                  End If
                  If .cButton(X).Caption Like "*[[]*#[]]" Then
                     .Timer.Enabled = True
                     .Timer.Tag = X
                  End If
               Next X
            Else
               ' falls mehr als 4 Button-Captions eingegeben wurden
               Call BtnCapsHelp ' Hilfe anzeigen
               Set OwnerForm = Nothing
               Exit Function
            End If
         Else
            ' Falls Fehler bei der Eingabe auftreten
            Call BtnCapsHelp ' Hilfe anzeigen
            Set OwnerForm = Nothing
            Exit Function
         End If
      Else
         ' falls keine Button-Caption eingegeben wurde,
         ' automatisch "OK"-Button anzeigen
         Load .cButton(1)
         .cButton(1).Visible = True
         .cButton(1).Caption = "OK"
      End If
      
      ' falls nur 1 Button vorhanden, diesen
      ' automatisch als Default setzen
      If .cButton.UBound = 1 Then
         .cButton(1).Default = True
         DefBtnNo = 1
      End If
      
      ' falls ein Hilfe-Button angezeigt werden soll, einen
      ' zusätzlichen, kleinen Button mit einem "?" erstellen
      If (ShowStyle And Show_HelpButton) Or (HelpFile <> "") Then
         X = .cButton.Count
         Load .cButton(X)
         .cButton(X).Width = 24
         .cButton(X).Height = 24
         .cButton(X).Font = "Arial"
         .cButton(X).FontSize = 12
         .cButton(X).FontBold = True
         .cButton(X).Caption = "?"
         .cButton(X).Visible = True
         ' Hilfe-Button markieren, Pfad zur Hilfe-Datei
         ' und Parameter zuweisen
         If IsMissing(Context) Then Context = ""
         .cButton(X).Tag = "MBExHelpButton"
         .MBExHelpFile = Trim$(CStr(HelpFile))
         .MBExParams = Trim$(CStr(Context))
         ' Hilfe-Button ist 52 Pixels kleiner
         AllBtsWidth = -52
      End If
      
      ' TabIndex-Reihenfolge der Buttons berechnen und zuweisen
      ' (Thanks to Marco Wünschmann)
      For X = 1 To .cButton.UBound
         If X >= DefBtnNo Then
            .cButton(X).TabIndex = X - DefBtnNo
         Else
            .cButton(X).TabIndex = .cButton.UBound - DefBtnNo + X
         End If
         ' Gesamtweite aller Buttons ermitteln
         ' ButtonWeite + 6 Pixels Abstand zwischen den Buttons addieren
         AllBtsWidth = AllBtsWidth + ButtonWidth + 6
      Next X
      '
      ' den Abstand hinter dem letzten Button entfernen
      AllBtsWidth = AllBtsWidth - 6
      
      
      ' BorderWidth = 2x Rahmen
      BorderWidth = (.Width / Screen.TwipsPerPixelX) - .ScaleWidth
      ' BorderHeight = 1x Titelzeile + 1x Rahmen
      BorderHeight = (.Height / Screen.TwipsPerPixelY) - .ScaleHeight
      
      ' maximale Fensterweite zuweisen (62% der Bildschirmweite)
      MaxWinWidth = (Screen.Width * 0.62) / Screen.TwipsPerPixelX
      
      ' mindest Fensterweite zuweisen (alle Button-Weiten + seitlicher Abstand)
      MinWinWidth = AllBtsWidth + 20 + BorderWidth
      
      
      ' Mindest-Weite der Msg-Form so einstellen, so dass
      ' der Titel-Text komplett angezeigt wird:
      '
      ' Font kurzzeitig auf Caption-Font umschalten
      Set .Font = SystemGetFont(.hdc, fntCaption)
      '
      ' TextWeite der Titelzeile ermitteln
      ' CaptionWidth = TextWeite der Caption + BorderHöhe (entspricht
      ' ca. der Weite des X-Button-Bereichs) + 1 Border-Seite
      CaptionWidth = .TextWidth(.Caption) + BorderHeight + (BorderWidth / 2)
      '
      ' Msg-Form auf Message-Font zurückstellen
      Set .Font = SystemGetFont(.hdc, fntMessage)
      '
      ' MinWinWidth anpassen
      If CaptionWidth > MinWinWidth Then
         If CaptionWidth < MaxWinWidth Then
            MinWinWidth = CaptionWidth
         Else
            MinWinWidth = MaxWinWidth
         End If
      End If
      
      
      ' die Weite der längsten Zeile ermitteln (an Original-Zeilenumbrüchen)
      sLine = Split(MsgText, vbNewLine, , vbTextCompare)
      For X = 0 To UBound(sLine)
         If .TextWidth(sLine(X)) > TextWidth Then
            TextWidth = .TextWidth(sLine(X))
         End If
      Next X
   
      'IconStyle herausfiltern und zuweisen
      ShowIcon = True
      If ShowStyle And Icon_Critical Then        ' Stop
         MB_Icon = IDI_HAND
         MB_Sound = MB_ICONHAND
      ElseIf ShowStyle And Icon_Exclamation Then ' Ausrufezeichen
         MB_Icon = IDI_EXCLAMATION
         MB_Sound = MB_ICONEXCLAMATION
      ElseIf ShowStyle And Icon_Information Then ' Information
         MB_Icon = IDI_ASTERISK
         MB_Sound = MB_ICONASTERISK
      ElseIf ShowStyle And Icon_Question Then    ' Fragezeichen
         MB_Icon = IDI_QUESTION
         MB_Sound = MB_ICONEXCLAMATION
      Else
         ShowIcon = False
      End If
      
      ' Randbreite zuweisen (= Rand für 1 Seite)
      MARGIN_W = 10 + (BorderWidth / 2)
      
      ' falls ein Icon angezeigt werden soll
      ' (Randbreite + IconBreite (32) + nochmal 10 Abstand)
      If ShowIcon Then MARGIN_W = MARGIN_W + 42
      
      ' Fensterweite zuweisen
      WinWidth = TextWidth + (MARGIN_W * 2)
      
      ' falls die Fensterweite breiter als MaxWinWidth (62% des Bildschirms) ist
      If WinWidth > MaxWinWidth Then WinWidth = MaxWinWidth
      
      ' falls die Fensterweite kleiner als die MinWinWidth ist
      ' (muß nach der Prüfung auf MaxWinWeidth stehen)
      If WinWidth < MinWinWidth Then WinWidth = MinWinWidth
      
      ' Text-Weite anpassen
      TextWidth = WinWidth - (MARGIN_W * 2)
      
      ' Ausgaberechteck anpassen
      TextRc.Left = MARGIN_W - 2 ' 2 Pixels "Editcontrol"-Abweichung
      TextRc.Right = TextRc.Left + TextWidth
      TextRc.Top = 12
      TextRc.Bottom = 1000 ' willkürlicher Wert > O, da die Texthöhe
                           ' erst noch ermittelt wird
     
      ' Textausrichtung herausfiltern und zuweisen
      If ShowStyle And Text_Left Then
         TextAlign = DT_LEFT
      ElseIf ShowStyle And Text_Right Then
         TextAlign = DT_RIGHT
      Else
         TextAlign = DT_CENTER
      End If
      
      ' Den Text innerhalb des Ausgabe-Rechtecks zeichen.
      ' Dabei soll er am Zeilenende umgebrochen werden
      ' (DT_WORDBREAK) und bei zusammenhängenden
      ' Zeichenfolgen, die zu lang für eine Zeile sind
      ' (DT_EDITCONTROL). Das '&'-Zeichen soll nicht
      ' als Prefix interpretiert werden (DT_NOPREFIX).
      
      ' Ermitteln der Text-Höhe
      TextHeight = DrawTextEx(.hdc, MsgText, Len(MsgText), TextRc, _
                     DT_EDITCONTROL Or DT_NOCLIP Or DT_NOPREFIX Or _
                     DT_WORDBREAK Or DT_RTLREADING Or TextAlign _
                     Or DT_TABSTOP Or DT_EXPANDTABS, ByVal 0&)
      
      ' Fenster-Höhe zuweisen
      WinHeight = BorderHeight + 12 + TextHeight + 14 + ButtonHeight + 10
      
      ' Mindest-Fenster-Höhe zuweisen
      ' (Titelleiste + VorTextabstand + NachTextabstand + Buttonhöhe + Bodenabstand)
      MinWinHeight = BorderHeight + 12 + 14 + ButtonHeight + 10
      
      ' falls ein Icon angezeigt werden soll
      If ShowIcon Then MinWinHeight = MinWinHeight + 32
         
      ' Fenster-Höhe auf Mindest-Höhe prüfen
      If WinHeight < MinWinHeight Then WinHeight = MinWinHeight
      
      ' Fenster anpassen und unsichtbar in
      ' der Bildschirmmitte positionieren
      SetWindowPos .hWnd, HWND_TOP, _
            ((Screen.Width / Screen.TwipsPerPixelX) - WinWidth) / 2, _
            ((Screen.Height / Screen.TwipsPerPixelY) - WinHeight) / 2, _
            WinWidth, WinHeight, SWP_NOACTIVATE
      
      
      ' Titel-Text rechtsbündig ausrichten?
      If TextAlign = DT_RIGHT Then
         ' Titel-Text an Variable zuweisen, damit während des
         ' Anpassens keine Veränderung an der Msg-Form geschieht.
         sCaption = .Caption
         ' CaptionWidth = gesamte FensterWeite - BorderHöhe (entspricht
         ' ca. dem Bereich des X-Buttons) + 1 Border-Seite
         CaptionWidth = WinWidth - BorderHeight - (BorderWidth / 2)
         ' Font kurzzeitig auf Caption-Font umschalten
         Set .Font = SystemGetFont(.hdc, fntCaption)
         Do
            ' Titelzeile von links mit Leerzeichen füllen ...
            sCaption = " " & sCaption
            ' ... bis der rechte Rand erreicht ist.
            If .TextWidth(sCaption) > CaptionWidth Then
               ' links 1 Leerzeichen wieder entfernen
               .Caption = Mid$(sCaption, 2)
               Exit Do
            End If
         Loop
         ' Msg-Form auf Message-Font zurückstellen
         Set .Font = SystemGetFont(.hdc, fntMessage)
      End If
      
      
      ' den ersten Button positionieren
      .cButton(1).Left = CInt((.ScaleWidth - AllBtsWidth) / 2)
      .cButton(1).Top = CInt(.ScaleHeight - ButtonHeight - 10)
      '
      ' die restlichen Buttons positionieren
      For X = 2 To .cButton.UBound
         .cButton(X).Left = .cButton(X - 1).Left + ButtonWidth + 6
         .cButton(X).Top = .cButton(1).Top
      Next X
      
      ' Grösse der Ausgabefläche wurde angepasst, also nochmals drucken
      .Cls
      Call DrawTextEx(.hdc, MsgText, Len(MsgText), TextRc, _
             DT_EDITCONTROL Or DT_NOCLIP Or DT_NOPREFIX Or _
             DT_WORDBREAK Or DT_RTLREADING Or TextAlign _
             Or DT_TABSTOP Or DT_EXPANDTABS, ByVal 0&)
         
      ' Entfernen des SystemMenüs und des Schließen-Symbols (X-Button)
      If ShowStyle And Show_NoCloseButton Then
         Call RemoveSysMenu(.hWnd, True) ' Close-Button deaktivieren
         .Tag = "NoCloseButton"
      Else
         Call RemoveSysMenu(.hWnd, False)
      End If
      
      ' soll ein Icon angezeigt werden?
      If ShowIcon Then
         '
         ' Icon laden und zeigen
         hIcon = LoadIcon(ByVal 0&, MB_Icon) ' Icon aus System-DLL laden
         DrawIcon .hdc, 10, 10, hIcon        ' in die linke Ecke malen
         DrawIcon .hdc, .ScaleWidth - 42, _
                               10, hIcon     ' in die rechte Ecke malen
         DestroyIcon hIcon                   ' Handle zerstören
         '
         ' System-Msg-Sound abspielen
         Call MessageBeep(MB_Sound)
         '
      End If
      
      ' immer im Vordergrund ?
      If ShowStyle And Show_AlwaysOnTop Then
         SetWindowPos .hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_DRAWFRAME Or SWP_NOMOVE Or SWP_NOSIZE
      End If
      
      ' Msg-Form zeigen
      .Refresh
      .Show vbModal, OwnerForm
   
      ' Rückgabewert zuweisen
      MsgBoxEx = .MBExRetval
   
   End With
   
End Function


' Funktion zum Ermitteln der System-Fonts
' für Form-Caption und Message-Text
Public Function SystemGetFont(lngHDC As Long, lngFont As SystemFontEnum) As StdFont
    Dim oLogFont As LOGFONT
    Dim oNCM     As NONCLIENTMETRICS
    Dim oReturn  As StdFont
    Dim lngReturn  As Long
    '
    Set oReturn = New StdFont
    
    oNCM.cbSize = Len(oNCM)
    '
    lngReturn = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, _
                                        oNCM.cbSize, oNCM, 0)
    '
    If lngReturn <> 0 Then
        Select Case lngFont
            Case Is = [fntCaption]:   oLogFont = oNCM.lfCaptionFont
            Case Is = [fntMenu]:      oLogFont = oNCM.lfMenuFont
            Case Is = [fntMessage]:   oLogFont = oNCM.lfMessageFont
            Case Is = fntSmCaption: oLogFont = oNCM.lfSMCaptionFont
            Case Is = [fntStatus]:    oLogFont = oNCM.lfStatusFont
            Case Else:                    GoTo ERR_HANDLER
        End Select
        '
        With oReturn
            .bold = (oLogFont.lfWeight > 400)
            .Charset = oLogFont.lfCharSet
            .italic = (oLogFont.lfItalic = 1)
            .Name = StripNull(oLogFont.lfFaceName)
            .Size = -MulMul(oLogFont.lfHeight, _
                     GetDeviceCaps(lngHDC, LOGPIXELSY), 72)
            .Strikethrough = (oLogFont.lfStrikeOut = 1)
            .underline = (oLogFont.lfUnderline = 1)
            .Weight = oLogFont.lfWeight
        End With
    End If
    '
ERR_HANDLER:
    Set SystemGetFont = oReturn
End Function
'
Public Function StripNull(strData As String) As String
    Dim strReturn As String
    Dim lngPos As Long
    '
    lngPos = InStr(strData, vbNullChar)
    If lngPos > 0 Then
        strReturn = Left(strData, (lngPos - 1))
    Else
        strReturn = strData
    End If
    '
    StripNull = strReturn
End Function
'
Private Function MulMul(arg1 As Long, arg2 As Long, arg3 As Long) As Integer
  '
  ' A weird name for a function :-)
  ' Actually, its based on the reverse of MulDiv
  ' (the multiple divide C macro) and since it
  ' returns the opposite data, I named it MulMul
  ' (though there is no multiplication in it!).
  ' I have no idea what it's real corresponding
  ' name would be in C.
   Dim tmp As Single
   '
   tmp = arg2 / arg3
   tmp = arg1 / tmp
   '
   MulMul = CInt(tmp)
   '
End Function


' Befehle aus dem Systemmenü entfernen (Minimieren, Maximieren usw.):
Private Sub RemoveSysMenu(FormHwnd As Long, ByVal DisableClose As Boolean)
                
   Dim MnuHwnd As Long
  
   MnuHwnd = GetSystemMenu(FormHwnd, False)
   
   RemoveMenu MnuHwnd, SC_RESTORE, MF_BYCOMMAND  ' Wiederherstellen
   'RemoveMenu MnuHwnd, SC_MOVE, MF_BYCOMMAND    ' Verschieben (bleibt erhalten)
   RemoveMenu MnuHwnd, SC_SIZE, MF_BYCOMMAND     ' Grösse verändern
   RemoveMenu MnuHwnd, SC_MINIMIZE, MF_BYCOMMAND ' Minimieren
   RemoveMenu MnuHwnd, SC_MAXIMIZE, MF_BYCOMMAND ' Maximieren
   
   ' "Schliessen" Befehl behalten ?
   If DisableClose Then
      RemoveMenu MnuHwnd, SC_CLOSE, MF_BYCOMMAND ' Schliessen
      RemoveMenu MnuHwnd, 1, MF_BYPOSITION       ' Trennlinie
   End If
   
   ' Copyright-Hinweis (muss erhalten bleiben!)
   ' "MsgBoxEx (GES)" ist Freeware und darf frei benutzt werden.
   AppendMenu MnuHwnd, MF_SEPARATOR, 1, vbNullString
   AppendMenu MnuHwnd, MF_STRING, 1, "MsgBoxEx (GES) © 2004"
   AppendMenu MnuHwnd, MF_STRING, 1, "guidoeisenbeis@web.de"
End Sub









