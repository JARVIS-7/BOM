VERSION 5.00
Begin VB.Form frmMsgBoxEx 
   AutoRedraw      =   -1  'True
   Caption         =   "frmMsgBoxEx"
   ClientHeight    =   1575
   ClientLeft      =   210
   ClientTop       =   1560
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMsgBoxEx.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   1200
   End
   Begin VB.CommandButton cButton 
      Caption         =   "cButton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1140
   End
End
Attribute VB_Name = "frmMsgBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   

' ================================================================= '
'
'     --------------
'     MsgBoxEx (GES)
'     --------------
'
' Autor:
' Guido Eisenbeis Software (GES), guidoeisenbeis@web.de, 2004-08-28
' ================================================================= '


' Grund-Code (ohne "Skin") ======================================== v

' Funktionen zum Öffnen von HTML-Help (.chm-Dateien)
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Private Declare Function HtmlHelpTopic Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As String) As Long
Private Const HH_DISPLAY_TOPIC As Long = &H0

' Funktion zum Öffnen von beliebigen Dateien
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Funktion zum Schließen von Fenstern
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE As Long = &H10
Private HelpHWND As Long


Public MBExHelpFile As String ' Pfad zur Hilfe-Datei
Public MBExParams As String   ' zusätzliche Parameter
Public MBExRetval As Integer  ' Rückgabe-Wert für "MsgBoxEx"


Private Sub cButton_Click(Index As Integer)

   Timer.Enabled = False
   ' falls der Hilfe-Button geklickt wurde *
   If cButton(Index).Tag = "MBExHelpButton" Then
      ' Hier Befehle für "Hilfe" einfügen
      Call ShowMBExHelp
   Else
      MBExRetval = Index
      Unload Me
   End If
End Sub
'
'  * Hinweis!
'  Die Tag-Eigenschaften der Form und
'  des Hilfe-Buttons können nicht benutzt
'  werden, da sie schon in Verwendung sind!

' Beispiel für den Aufruf einer Hilfe-Datei
Private Sub ShowMBExHelp()

   On Error Resume Next ' Fehlerbehandlung aus
      
   If MBExHelpFile = "" Then Exit Sub
   If (LCase(Right$(MBExHelpFile, 4)) = ".chm") And (MBExParams <> "") Then
      ' Html-Hilfe mit einer bestimmten Seite öffnen
      HelpHWND = HtmlHelpTopic(0, MBExHelpFile, _
                  HH_DISPLAY_TOPIC, MBExParams)
   Else
      ' Html-Hilfe mit Startseite, oder beliebige Datei
      ' (.txt, .doc, .htm ...) mit oder ohne Parameter öffnen
      Call ShellExecute(Me.hWnd, "Open", _
         MBExHelpFile, MBExParams, "", 1)
   End If
      
   On Error GoTo 0 ' Fehlerbehandlung ein

End Sub

' Steurerung per Tastatur ermöglichen
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim b As CommandButton
   ' F1-Taste wurde gedrückt
   If KeyCode = vbKeyF1 Then
      KeyCode = 0
      ' Hilfe aufrufen, falls vorhanden
      For Each b In Me
         If b.Tag = "MBExHelpButton" Then
            b.SetFocus
            b = True
            Exit For
         End If
      Next b
   ' Escape-Taste wurde gedrückt
   ElseIf KeyCode = vbKeyEscape Then
      KeyCode = 0
      ' "frmMsgBoxEx" schließen, falls erlaubt
      If Me.Tag <> "NoCloseButton" Then Unload Me
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Wenn eine Hilfe-Datei mit einer bestimmten Seite und
   ' "HtmlHelp" geöffnet wird, muss das Fenster der Hilfe
   ' geschlossen sein, bevor das Programm beendet wird.
   ' Ansonsten erfolgt eine Speicherzugriffsverletzung.
   If HelpHWND Then
      SendMessage HelpHWND, WM_CLOSE, 0&, 0&
      HelpHWND = 0
   End If
End Sub
' ================================================================= ^



'Private Sub Form_Load()
'   ' ForeColor und BackColor können an folgenden
'   ' Stellen geändert werden:
'   ' 1) an frmMsgBoxEx selbst (zur Entwicklungszeit)
'   ' 2) hier im Form-Load
'   ' 3) an beliebiger Stelle des Projekts direkt
'   '    vor dem MsgBoxEx-Aufruf
'
'   ' Hinweis!
'   ' Abhängig von der Auflösung des Monitors kann bei
'   ' farbigem Hintergrund der Msg-Text unleserlich sein!
'
'   ' Beispiel:
'   Me.BackColor = vbBlue
'   Me.ForeColor = vbWhite
'End Sub


Private Sub Timer_Timer()

  Dim TimeLeft As Integer
  TimeLeft = Val(Mid(cButton(Timer.Tag).Caption, InStrRev(cButton(Timer.Tag).Caption, "[") + 1))
  cButton(Timer.Tag).Caption = Left(cButton(Timer.Tag).Caption, InStrRev(cButton(Timer.Tag).Caption, "[") - 1) & "[" & CStr(TimeLeft - 1) & "]"
  If TimeLeft <= 1 Then
    Timer.Enabled = False
    cButton_Click Timer.Tag
  End If

End Sub
