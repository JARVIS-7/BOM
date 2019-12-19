VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHaupt 
   Caption         =   "Biet-O-Matic"
   ClientHeight    =   4470
   ClientLeft      =   1875
   ClientTop       =   2475
   ClientWidth     =   11880
   Icon            =   "frmHaupt.frx":0000
   LinkTopic       =   "frmHaupt"
   LockControls    =   -1  'True
   ScaleHeight     =   4841.154
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   12000
   Visible         =   0   'False
   Begin VB.TextBox Gebot 
      Alignment       =   1  'Rechts
      Height          =   405
      Index           =   0
      Left            =   8790
      TabIndex        =   27
      ToolTipText     =   "Was ist mir der Artikel wert? ACHTUNG: in Bietwährung ! "
      Top             =   600
      Width           =   855
   End
   Begin VB.Timer TrayFlashTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3000
      Top             =   3240
   End
   Begin VB.Timer WakeupTimer 
      Interval        =   1000
      Left            =   3480
      Top             =   3240
   End
   Begin VB.Timer ExtCmdTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   3240
   End
   Begin VB.Timer CurlTimer 
      Interval        =   100
      Left            =   4440
      Top             =   3240
   End
   Begin VB.Timer QuitTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4920
      Top             =   3240
   End
   Begin VB.TextBox VersandkostenEdit 
      Alignment       =   2  'Zentriert
      Height          =   285
      Left            =   6720
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer ShowInfo 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5400
      Top             =   3240
   End
   Begin VB.Timer TitelFlashTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   3240
   End
   Begin BietOMatic.ctlSNTP ctlSNTP1 
      Height          =   495
      Left            =   9480
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
   End
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   8880
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4215
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6773
            MinWidth        =   1746
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6773
            MinWidth        =   1746
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6773
            MinWidth        =   1746
         EndProperty
      EndProperty
   End
   Begin VB.Timer ZeilenUnloadTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   3720
   End
   Begin VB.Timer ArtikelBuffTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   3720
   End
   Begin VB.Timer ODBC_Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   3720
   End
   Begin VB.Timer AutoSave 
      Interval        =   60000
      Left            =   120
      Top             =   3720
   End
   Begin VB.Timer MailBuffTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   3720
   End
   Begin VB.CommandButton AutoStatus 
      Height          =   300
      Left            =   3120
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   150
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock NTP 
      Left            =   7320
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Zusatzfeld 
      Appearance      =   0  '2D
      BackColor       =   &H80000016&
      Height          =   375
      Left            =   120
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2115
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton NewArtikel 
      Appearance      =   0  '2D
      Caption         =   "Artikel"
      Height          =   300
      Left            =   75
      TabIndex        =   1
      ToolTipText     =   "Klicken, um eien neuen Artikel hinzuzufügen"
      Top             =   150
      Width           =   1335
   End
   Begin VB.CommandButton SortEnde 
      Appearance      =   0  '2D
      Caption         =   "Endet"
      Height          =   300
      Left            =   1530
      TabIndex        =   2
      Tag             =   "asc"
      ToolTipText     =   "Klicken, um nach der Endezeit zu sortieren"
      Top             =   150
      Width           =   1455
   End
   Begin VB.Timer Gruppe_Timer 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   4440
      Top             =   3720
   End
   Begin VB.Timer GebotTimer 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3960
      Top             =   3720
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   825
      LargeChange     =   11
      Left            =   11672
      Max             =   1
      Min             =   1
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   555
      Value           =   1
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Timer RechnerZeitTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   3720
   End
   Begin VB.Timer ArtikelTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   3720
   End
   Begin VB.PictureBox Picture1 
      Height          =   300
      Left            =   8520
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer TimeoutTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1560
      Top             =   3720
   End
   Begin VB.Timer AboutTimer 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1080
      Top             =   3720
   End
   Begin VB.TextBox Bietgruppe 
      Height          =   405
      Index           =   0
      Left            =   10065
      OLEDropMode     =   1  'Manuell
      TabIndex        =   29
      ToolTipText     =   "Bietgruppe (Nr;Anzahl) bei Artikeln gleicher Bietgruppe wird nur mitgeboten, wenn die Anzahl Artikel noch nicht ersteigert wurde"
      Top             =   600
      Width           =   555
   End
   Begin VB.Timer POPTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2040
      Top             =   3720
   End
   Begin MSWinsockLib.Winsock tcpIn 
      Left            =   6840
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   3000
      Top             =   3720
   End
   Begin VB.TextBox Artikel 
      Alignment       =   1  'Rechts
      Height          =   405
      Index           =   0
      Left            =   90
      OLEDropMode     =   1  'Manuell
      TabIndex        =   12
      ToolTipText     =   "Hier die Artikelnummer eingeben"
      Top             =   600
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artikel lesen"
            Object.Tag             =   "tbReadArtikel"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "alles speichern"
            Object.Tag             =   "tbSaveAll"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artikelinfo Update"
            Object.Tag             =   "tbUpdateArtikel"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "beobachtete Artikel lesen"
            Object.Tag             =   "tbReadEbay"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mit Serverzeit Synchronisieren"
            Object.Tag             =   "tbSyncEbayTime"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Status: Abgemeldet"
            Object.Tag             =   "tbLogin"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Automatikmodus einschalten"
            Object.Tag             =   "tbAuto"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "mnuSettings_click"
            Object.ToolTipText     =   "Einstellungsdialog öffnen"
            Object.Tag             =   "tbSettings"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Browser öffnen"
            Object.Tag             =   "tbBrowser"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artikelfenster öffnen"
            Object.Tag             =   "tbArtikel"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Über B-O-M"
            Object.Tag             =   "tbAbout"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Hilfe aufrufen"
            Object.Tag             =   "tbHelp"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Hilfe aufrufen"
            Object.Tag             =   "tbHelp"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin BietOMatic.ctlMWheel MWheel1 
      Height          =   480
      Left            =   8880
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin BietOMatic.ctlSMTPRelay SMTP_1 
      Height          =   915
      Left            =   9360
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1614
   End
   Begin VB.Shape Fokus 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   375
      Left            =   2160
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Restzeit 
      Alignment       =   2  'Zentriert
      Height          =   405
      Index           =   0
      Left            =   7620
      OLEDropMode     =   1  'Manuell
      TabIndex        =   6
      ToolTipText     =   "Zeit bis zum Gebotsablauf"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image Ecke 
      Height          =   75
      Index           =   0
      Left            =   6195
      Picture         =   "frmHaupt.frx":014A
      Top             =   930
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Blink 
      Height          =   92
      Index           =   1
      Left            =   11655
      TabIndex        =   25
      Top             =   225
      Width           =   99
   End
   Begin VB.Label Blink 
      Height          =   92
      Index           =   0
      Left            =   11475
      TabIndex        =   24
      Top             =   225
      Width           =   99
   End
   Begin VB.Label EndeZeit 
      Height          =   405
      Index           =   0
      Left            =   1530
      OLEDropMode     =   1  'Manuell
      TabIndex        =   20
      ToolTipText     =   "Endezeit; Rot: ist beendet"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Waehrung 
      Height          =   405
      Index           =   0
      Left            =   9735
      OLEDropMode     =   1  'Manuell
      TabIndex        =   19
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Titel 
      Height          =   405
      Index           =   0
      Left            =   3120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   16
      ToolTipText     =   "Mit einem Klick auf den Titel in die Artikelbeschreibung!"
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Gruppe"
      Height          =   222
      Left            =   10065
      TabIndex        =   15
      ToolTipText     =   "Bietgruppe; bei Artikeln gleicher Bietgruppe wird nur mitgeboten, wenn noch kein Artikel der Gruppe ersteigert wurde"
      Top             =   194
      Width           =   585
   End
   Begin VB.Line Line12 
      X1              =   121.212
      X2              =   11700
      Y1              =   584.837
      Y2              =   584.837
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   121.212
      X2              =   11700
      Y1              =   1129.603
      Y2              =   1129.603
   End
   Begin VB.Label Status 
      Height          =   405
      Index           =   0
      Left            =   10740
      OLEDropMode     =   1  'Manuell
      TabIndex        =   10
      ToolTipText     =   "Nach dem Bieten: OK wenn Gebot erfolgreich, sonst ERR Klick für Statusbericht"
      Top             =   600
      Width           =   891
   End
   Begin VB.Label Label11 
      Caption         =   "Status"
      Height          =   222
      Left            =   10830
      TabIndex        =   9
      ToolTipText     =   "Nach dem Bieten: OK wenn Gebot erfolgreich, sonst ERR Klick für Statusbericht"
      Top             =   194
      Width           =   540
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Zentriert
      Caption         =   "Mein Gebot"
      Height          =   222
      Left            =   8790
      TabIndex        =   8
      ToolTipText     =   "Was ist mir der Artikel wert? ACHTUNG: in Bietwährung ! "
      Top             =   194
      Width           =   885
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Zentriert
      Caption         =   "Restzeit"
      Height          =   225
      Left            =   7560
      TabIndex        =   7
      Top             =   195
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Zentriert
      Caption         =   "Akt.Preis "
      Height          =   225
      Left            =   6240
      TabIndex        =   5
      ToolTipText     =   "letztes Gebot, als die Artikelinfo gelesen wurde; Klick = Artikelupdate"
      Top             =   195
      Width           =   1215
   End
   Begin VB.Label Versandkosten 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   0
      Left            =   6240
      OLEDropMode     =   1  'Manuell
      TabIndex        =   0
      Top             =   795
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Preis 
      Alignment       =   2  'Zentriert
      Height          =   390
      Index           =   0
      Left            =   6240
      OLEDropMode     =   1  'Manuell
      TabIndex        =   13
      ToolTipText     =   "letztes Gebot, als die Artikelinfo gelesen wurde; Klick = Artikelupdate"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File    "
      Begin VB.Menu mnuSave 
         Caption         =   "&Speichern"
         Begin VB.Menu mnuSaveSettings 
            Caption         =   "Nur Einstellungen"
         End
         Begin VB.Menu mnuSaveArtikel 
            Caption         =   "Nur Artikel "
         End
         Begin VB.Menu mnuSaveAll 
            Caption         =   "&Alles"
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu mnuLesen 
         Caption         =   "&Lesen"
         Begin VB.Menu mnuReadArtikel 
            Caption         =   "&Artikel aus Datei lesen"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Ende"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Aktionen    "
      Begin VB.Menu mnuUpdateArtikel 
         Caption         =   "Artikelinfo Update"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuMyEbay 
         Caption         =   "Beobachtete Artikel lesen"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuDeleteArtikel 
         Caption         =   "Alle Artikel entfernen"
      End
      Begin VB.Menu mnuCleanupArtikel 
         Caption         =   "Abgelaufene Artikel entfernen"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuCleanupArtikel2 
         Caption         =   "Abgebrochene Artikel entfernen"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSync 
         Caption         =   "Zeit synchronisieren"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuPasswd 
         Caption         =   "Passwort prüfen"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuAuto 
         Caption         =   "Automatik ein/ aus"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Suchen"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearchContinue 
         Caption         =   "Weitersuchen"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Fenster    "
      Begin VB.Menu mnuSettings 
         Caption         =   "Einstellungen"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuBrowser 
         Caption         =   "Browser"
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArtikel 
         Caption         =   "Artikeleingabe"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "Sprache"
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnul 
         Caption         =   "mnul"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Info"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Hilfe"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "&Versionsprüfung"
      End
      Begin VB.Menu mnuReleasenotes 
         Caption         =   "Versionshinweise"
      End
      Begin VB.Menu mnuCurrUpdate 
         Caption         =   "&Währungsupdate"
      End
      Begin VB.Menu mnuHomepage 
         Caption         =   "H&omepage"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuicon 
      Caption         =   "Taskbar"
      Visible         =   0   'False
      Begin VB.Menu mnuMax 
         Caption         =   "Öffnen"
      End
      Begin VB.Menu mnuAbout2 
         Caption         =   "Info"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "Beenden"
      End
   End
   Begin VB.Menu mnuMouse 
      Caption         =   "TitleMouse"
      Visible         =   0   'False
      Begin VB.Menu mnuComment 
         Caption         =   "Kommentar eingeben"
      End
      Begin VB.Menu mnuBid 
         Caption         =   "jetzt Bieten"
      End
      Begin VB.Menu mnuAkt 
         Caption         =   "Aktualisieren"
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "Browser"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Löschen"
      End
      Begin VB.Menu mnuAccount 
         Caption         =   "Account wählen"
      End
      Begin VB.Menu mnuSendItemTo 
         Caption         =   "Senden an"
      End
      Begin VB.Menu mnuEditShipping 
         Caption         =   "Versandkosten bearbeiten"
      End
      Begin VB.Menu mnuProductSearch 
         Caption         =   "Produktsuche"
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   1
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnuls 
            Caption         =   "mnuls"
            Index           =   10
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Tools"
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   1
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnult 
            Caption         =   "mnult"
            Index           =   10
            Visible         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "frmHaupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' $author: hjs$
' $id: V 2.0.4 date 170403 hjs$
' $version: 2.0.4$
' $file: $
'
' last modified:
' &date: 170403 scr 710622 hdn$
'
' contact: visit http://de.groups.yahoo.com/group/BOMInfo
'
'*******************************************************
Option Explicit
'
'
' Die Hauptmaske von BOM
'
' wenn der Code etwas merchwürdich ist .. daran denken: alles historisch gewuchert :-)
'
' Modifiziere, was immer Du magst ..
'
' ACHTUNG: die automatische Funktion "Bieten" in Timer1 ist auskommentiert ;-)
'
'ModulLocals
'
Private Const mlMAXBIETVERSUCHE As Long = 5& 'Nur für Prozedur_Bieten

Private mbEBayTimeIsSync As Boolean
Private msScratch As String
Private miRowCount As Integer
Private mbQuietExit As Boolean
Private miPopTimerCount As Integer
Private miOdbcTimerCount As Integer
Private mbInitDone As Boolean
Private miGlobArtikel As Integer
Private mbPopupShown As Boolean
Private miStartShowArtikel As Integer
Private mbCheckDone As Boolean
Private miStartWidth As Integer
Private miStartHeight As Integer

Private miGebotsIndex As Integer
Private miGruppeIndex As Integer
Private mbFormLoaded As Boolean

Private mbStopUpdate As Boolean

Private msCaptionCache As String

Private mbAlreadyRunning As Boolean  'Nur für Proz_CheckSofortkaufArtikel

'Artikelupdate
Private miArtikelCycleCount As Integer
Private miSaveCount As Integer
Private miMaxSaveCount As Integer

'Special
Private mbIsOn As Boolean
Private mbIsSync As Boolean
Private mbIsLoggingIn As Boolean
Private mbIsBidding As Boolean

'Errormeldung beim Bieten
Private miErrStatus As Integer

'WindowState sichern
Private mlPrevWindowState As Long

'Resizer
Private moResize As clsDoResize
Private mbIsUnloadingZeilen As Boolean
Private mbIsMinimizing As Boolean
Private mbIsRestoring As Boolean
Private mbRestoreQueued As Boolean
Private mbMinimizeQueued As Boolean

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd&) As Long

Private Declare Function SetWindowText _
        Lib "user32.dll" Alias "SetWindowTextA" _
        (ByVal hWnd As Long, ByVal lpString As String) _
        As Long
        
'** Prog in Taskbar
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias _
                               "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As _
                               NOTIFYICONDATA) As Boolean

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private mtTrayIcon As NOTIFYICONDATA

'Const für Taskbar
Private Const NIM_ADD As Long = &H0
Private Const NIM_MODIFY As Long = &H1
Private Const NIM_DELETE As Long = &H2
Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_ICON As Long = &H2
Private Const NIF_TIP As Long = &H4

Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
'Const WM_LBUTTONDBLCLK = &H203
'Const WM_RBUTTONDOWN = &H204
'Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONUP As Long = &H205

Private Type udtEzeitPos
  ts_Ezeit As Date
  ts_AnzPos As Integer
End Type

Private mbAktPaused As Boolean 'Aktualisierung anhalten

'MouseOver
Private miMouseIndex As Integer
Private miShowIndex As Integer
Private miTmpIndex As Integer

Private mbPopChecked As Boolean
Private mbWaitForFirstPop As Boolean

Private Sub Artikel_GotFocus(Index As Integer)
    Call SetFocusRect(Index)
End Sub

'Private Sub AboutTimer_Timer()
'
'On Error Resume Next
'
'AboutTimer.Enabled = False
'
'If AboutTimerCount = 0 And gbShowSplash Then
'    If Not gbSettingsIsUp Then
'        frmAbout.Show vbModal, Me
'    End If
'End If
'
'If Not CheckDone And Not gbUsesModem And gbCheckForUpdate And Not gbAutoStart Then
'    CheckDone = True
'    mnuVersion_Click
'End If
'
'End Sub


Private Sub Artikel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call SetFocusRect(Index)
    If KeyCode = vbKeyDelete And Shift = vbCtrlMask Then KeyCode = 0
End Sub

Private Sub Artikel_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo errhdl
Dim lRet As VbMsgBoxResult

If KeyCode = vbKeyDelete And Shift = vbCtrlMask Then 'Strg-del   '46 , 2
    
    If gbConfirmDelete Then
        lRet = MsgBox(gsarrLangTxt(51) & "?", vbYesNo Or vbQuestion)
    Else
        lRet = vbYes
    End If
    
    If lRet = vbYes Then
        Artikel(Index).Text = ""
        Artikel_LostFocus Index
    Else
        Call ArtikelArrayToScreen(VScroll1.Value)
        Artikel(Index).SetFocus
    End If

End If

If KeyCode = vbKeyReturn Then 'Enter  '13
    Artikel_LostFocus Index
End If

If KeyCode = vbKeyDown Then '40
    If Index < giMaxRow Then
        Artikel(Index + 1).SetFocus
    Else
        If VScroll1.Value < giAktAnzArtikel - giMaxRow Then
            VScroll1.Value = VScroll1.Value + 1
        End If
    End If
End If

If KeyCode = vbKeyUp Then  '38
    If Index > 0 Then
        Artikel(Index - 1).SetFocus
    Else
        If VScroll1.Value > 1 Then
            VScroll1.Value = VScroll1.Value - 1
        End If
    End If
End If

If KeyCode = vbKeyU And Shift = vbCtrlMask Then 'Strg-U  '85  , 2

    If Titel(Index).Caption = "" And (Index + VScroll1.Value <= giAktAnzArtikel) Then
        'nur aus der Anzeige gelöscht, undo möglich
        Call ArtikelArrayToScreen(VScroll1.Value)
        Artikel(Index).SetFocus
    End If
End If

If KeyCode = vbKeyPageUp Then 'Page up  33

    If VScroll1.Value > giMaxRow Then
        VScroll1.Value = VScroll1.Value - giMaxRow
    Else
        VScroll1.Value = VScroll1.Min
    End If
End If

If KeyCode = vbKeyPageDown Then 'Page down  34

    If VScroll1.Value + giMaxRow < VScroll1.Max Then
        VScroll1.Value = VScroll1.Value + giMaxRow
    Else
        VScroll1.Value = VScroll1.Max
    End If
End If

errhdl:

End Sub

Private Sub Artikel_LostFocus(Index As Integer)
Dim sTmp As String
Dim iArtIndex As Integer

On Error Resume Next

'prüfen ob ok
sTmp = CStr(Val(Trim(Artikel(Index))))
If sTmp = " " Or Val(sTmp) = 0 Then
    sTmp = ""
    Artikel(Index) = sTmp
End If

iArtIndex = Index + VScroll1.Value

If sTmp = "" Then
    If iArtIndex <= giAktAnzArtikel Then
        RemoveArtikel (iArtIndex)
        Artikel(Index).SetFocus
    End If
Else
    If iArtIndex > giAktAnzArtikel Then
        'neuer Artikel
        Call AddArtikel(sTmp)
    Else
        'Update auf die Zeile ?
        If gtarrArtikelArray(iArtIndex).Artikel <> sTmp Then
            Call InitArtikel(iArtIndex)
            gtarrArtikelArray(iArtIndex).Artikel = sTmp
            Call Update_Artikel(iArtIndex)
            '1.7.5 Resort
            Call Sortiere
        End If
    End If
End If

End Sub

Private Sub Artikel_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

' Index für ArtikelArray übergeben KOM 3.9.03
Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub ArtikelBuffTimer_Timer()
    
    Dim sTxt As String
    Dim fGebot As Double
    Dim sGruppe As String
    Dim sUser As String
    Dim sKommentar As String
    Dim sArr() As String
    
    If giSuspendState = 0 And mbIsBidding = False Then
        If ReadArtikelBuff(sTxt) Then
            'Einketten
            sArr() = Split(sTxt, vbTab, -1, vbBinaryCompare)
            
            If UBound(sArr()) > (LBound(sArr()) - 1) Then sTxt = sArr(0)
            If UBound(sArr()) > LBound(sArr()) Then fGebot = CDbl(sArr(1))
            If UBound(sArr()) > (LBound(sArr()) + 1) Then sGruppe = sArr(2)
            If UBound(sArr()) > (LBound(sArr()) + 2) Then sUser = sArr(3)
            If UBound(sArr()) > (LBound(sArr()) + 3) Then sKommentar = sArr(4)
            
            Call AddArtikel(sTxt, fGebot, sGruppe, sUser, sKommentar)
            Erase sArr()
        End If
        ArtikelBuffTimer.Enabled = CBool(ArtikelBuffTimer.Tag)
    End If
    
End Sub

Private Sub AutoSave_Timer()
    
    If giSuspendState = 0 And mbIsBidding = False Then
        miSaveCount = miSaveCount + 1
        'miMaxSaveCount wird in Form_Load erst nach der Passwortabfrage gesetzt, _
        sonst ist die Artikeldatei leer wenn der User mit der Eingabe zu lange wartet!
        If miMaxSaveCount > 0 Then
            If miSaveCount >= miMaxSaveCount Then
                miSaveCount = 0
                Call mnuSaveArtikel_Click
            End If
        End If
    End If
    
End Sub

Public Sub AutoStatus_Click()
    
    If gbEmpUserEnd Then
        gbAutoMode = True
        gbEmpUserEnd = False
    End If
    
    If gbAutoMode = False Then
        gbAutoMode = True
        miArtikelCycleCount = 0 'sh 6.11.2003
    Else
        gbAutoMode = False
    End If
    
    Call CheckAutoMode
    
End Sub

Private Sub Bietgruppe_GotFocus(Index As Integer)
    Call SetFocusRect(Index)
End Sub

Private Sub Bietgruppe_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call SetFocusRect(Index)
End Sub

Private Sub Bietgruppe_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo errhdl

If KeyCode = vbKeyDown Then '40
    If Index < giMaxRow Then
        Bietgruppe(Index + 1).SetFocus
    Else
        If VScroll1.Value < giAktAnzArtikel - giMaxRow Then
            Call Bietgruppe_LostFocus(Index)
            VScroll1.Value = VScroll1.Value + 1
        End If
    End If
End If

If KeyCode = vbKeyUp Then '38
    If Index > 0 Then
        Bietgruppe(Index - 1).SetFocus
    Else
        If VScroll1.Value > 1 Then
            Call Bietgruppe_LostFocus(Index)
            VScroll1.Value = VScroll1.Value - 1
        End If
    End If
End If

If KeyCode = vbKeyReturn Then '13
    Call Bietgruppe_LostFocus(Index)
End If

If KeyCode >= vbKey0 Then '48
    miGruppeIndex = Index
    Gruppe_Timer.Enabled = True
End If

If KeyCode = vbKeyPageUp Then '33
    Call Bietgruppe_LostFocus(Index)
    If VScroll1.Value > giMaxRow Then
        VScroll1.Value = VScroll1.Value - giMaxRow
    Else
        VScroll1.Value = VScroll1.Min
    End If
End If

If KeyCode = vbKeyPageDown Then '34
    Call Bietgruppe_LostFocus(Index)
    If VScroll1.Value + giMaxRow < VScroll1.Max Then
        VScroll1.Value = VScroll1.Value + giMaxRow
    Else
        VScroll1.Value = VScroll1.Max
    End If
End If

errhdl:
End Sub

Private Sub Bietgruppe_LostFocus(Index As Integer)
Static bAlreadyRunning As Boolean
Dim sTmp As String
Dim sGruppeAlt As String

If Not bAlreadyRunning Then

    Gruppe_Timer.Enabled = False
    
    If Artikel(0).Tag = "in" Then
    
        bAlreadyRunning = True
        
        'prüfen ob ok
        sTmp = Trim(Bietgruppe(Index))
        
        If giAktAnzArtikel >= Index + VScroll1.Value Then
            Bietgruppe(Index).Text = sTmp
            
            With gtarrArtikelArray(Index + VScroll1.Value)
                If .Status = [asCancelGroup] And .Gruppe <> Bietgruppe(Index).Text Then 'lg 31.07.03
                    If ResetStatusCancel() Then .Status = [asNixLos]
                End If
                
                If .Status = [asBuyOnlyCanceled] And .Gruppe <> Bietgruppe(Index).Text Then 'lg 31.07.03
                    If ResetStatusCancel() Then .Status = [asBuyOnly]
                End If
                
                sGruppeAlt = .Gruppe
                .Gruppe = Bietgruppe(Index).Text
                
                If sGruppeAlt <> .Gruppe Then
                    .LastChangedId = GetChangeID()
                End If
                Call CheckBietgruppe(sGruppeAlt)
                Call CheckBietgruppe(.Gruppe)
            End With 'gtarrArtikelArray(Index + VScroll1.Value)
            
            Call CheckSofortkaufArtikel
            
            Call ArtikelArrayToScreen(VScroll1.Value)
        
        Else
            Bietgruppe(Index).Text = ""
        End If
        
    End If 'Artikel(0).Tag = "in"
    bAlreadyRunning = False
End If 'Not bAlreadyRunning
End Sub

Private Function ResetStatusCancel() As Boolean

  If vbYes = MsgBox(gsarrLangTxt(752), vbYesNo) Then ResetStatusCancel = True
  
End Function

Public Sub CheckAutoMode()
Dim i As Integer
Dim bOk As Boolean

DebugPrint "Automode " & IIf(gbAutoMode, "on", "off")

gbBeendenNachAuktionAktiv = True

If (gbAutoMode) And (gsUser = "" Or gsPass = "") Then
    If FormLoaded("frmAbout") Then frmAbout.Hide
    MsgBox gsarrLangTxt(2)
    gbAutoMode = False
End If

If Not mbEBayTimeIsSync And gbAutoMode Then
    If Not gbAutoStart Then
        Call TimeSync
    Else
        Call Zeitsync
    End If
End If

'Mal sehen ob überhaupt schon die Zeiten geladen sind ..
If gbAutoMode Then
    bOk = False
    For i = 1 To giAktAnzArtikel
        If gtarrArtikelArray(i).Titel <> "" Then
            bOk = True
            Exit For
        End If
    Next i
    If Not bOk Then
        Call LoadArtikel
    End If
    Call SetIcon(Me.hWnd, MyLoadResPicture(202, 16))
    Call updTaskbar(gsarrLangTxt(47) & ": " & gsUser & " " & gsarrLangTxt(255), True)
    Call CheckSofortkaufArtikel
Else
    Blink(0).BackColor = &H8000000F
    Blink(1).BackColor = &H8000000F
    Call SetIcon(Me.hWnd, MyLoadResPicture(201, 16))
    Call updTaskbar(gsarrLangTxt(47) & ": " & gsUser & " " & gsarrLangTxt(256), True)
    Call PanelText(StatusBar1, 3, "")
    Call ResetWakupTime
End If

Call ArtikelArrayToScreen(VScroll1.Value)
mnuAuto.Checked = gbAutoMode
Timer1.Enabled = gbAutoMode
POPTimer.Enabled = gbAutoMode
ArtikelTimer.Enabled = gbAutoMode
ExtCmdTimer.Enabled = gbAutoMode
TrayFlashTimer.Enabled = True

gbUseWinShutdown = gbFileWinShutdown

If gbAutoMode Then
    Call SetToolbarImage(Toolbar1.Buttons(8), 10)
    Toolbar1.Buttons(8).ToolTipText = gsarrLangTxt(63)
    AutoStatus.Caption = gsarrLangTxt(61)
    AutoStatus.BackColor = vbGreen
Else
    Call SetToolbarImage(Toolbar1.Buttons(8), 13)
    Toolbar1.Buttons(8).ToolTipText = gsarrLangTxt(64)
    AutoStatus.Caption = gsarrLangTxt(62)
    AutoStatus.BackColor = vbRed
    If gbUsePop Then
        Call PanelText(StatusBar1, 2, "")
    End If
End If

'For i = 0 To IIf(Artikel(0).Tag = "in", giMaxRow, 0)
'    If gbAutoMode = False Then
'        Artikel(i).Enabled = True
'    Else
'       Artikel(i).Enabled = False
'    End If
'Next i

If gbUsePop And gbAutoMode Then
    If Not mbPopChecked Then
        Call PanelText(StatusBar1, 2, gsarrLangTxt(83))
        bOk = PopTest
        DebugPrint "PopTest " & IIf(bOk, "ok", "failed")
        Call PanelText(StatusBar1, 2, "")
        If bOk Then
            Call PanelText(StatusBar1, 2, Replace(gsarrLangTxt(84), "%MIN%", CStr(giPopZyklus)), True, vbGreen)
            mbPopChecked = True
            If Not mbWaitForFirstPop Then miPopTimerCount = giPopZyklus
        Else
            Call PanelText(StatusBar1, 2, gsarrLangTxt(85), False, vbRed)
            mbPopChecked = False
        End If
    End If
Else
    Call PanelText(StatusBar1, 2, "")
    mbPopChecked = False
End If

End Sub

Private Sub Bietgruppe_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)
    
End Sub

Private Sub CurlTimer_Timer()

    Dim i As Integer
    Dim sUrl As String
    Dim sData As String
    Dim sArtikel As String
    
    If giSuspendState = 0 Then
        CurlTimer.Enabled = False
        If PollPendingCurls(sUrl, sData) Then
            If InStr(1, sUrl, Replace(gsCmdViewItem, "[Item]", "")) > 0 Then
                sArtikel = GetItemFromUrl(sUrl)
                If sArtikel > "" Then
                    i = ItemToIndex(sArtikel)
                      If i > 0 Then
                          Call Update_Artikel(i, sData)
                          Call ArtikelArrayToScreen(miStartShowArtikel)
                      End If
                End If
            End If
        End If
    End If 'giSuspendState = 0
    CurlTimer.Enabled = True
    
End Sub

Private Sub Ecke_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetFocusRect(Index)
    If Artikel(Index).Text <> "" Then
        miGlobArtikel = Index + VScroll1.Value
        If miGlobArtikel <= UBound(gtarrArtikelArray()) Then
            Call mnuComment_Click
        End If
    End If
    
End Sub

Private Sub EndeZeit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetFocusRect(Index)
    Call VersandkostenUebernehmen
    
    If Not Artikel(Index).Text = "" Then
        If Button = vbRightButton Then '2
            Call ShowContextMenu(Index)
            'miGlobArtikel = Index + VScroll1.Value
            'Me.PopupMenu mnuMouse
        ElseIf Button = vbLeftButton Then '1
            miGlobArtikel = Index + VScroll1.Value
            'DebugPrint StatusBar1.SimpleText
            If giUserAnzahl > 0 And miGlobArtikel <= giAktAnzArtikel Then
                If gtarrArtikelArray(miGlobArtikel).UserAccount = "" Then
                    Call PanelText(StatusBar1, 3, "Account: " & gsarrLangTxt(47), True)
                Else
                    Call PanelText(StatusBar1, 3, "Account: " & gtarrArtikelArray(miGlobArtikel).UserAccount, True)
                End If
            End If
        End If
    End If
    
End Sub

Private Sub EndeZeit_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Index für ArtikelArray übergeben KOM 3.9.03
    Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)
    
End Sub

Private Sub ExtCmdTimer_Timer()

    Static lCnt As Long
    Dim i As Integer
    
    If giSuspendState = 0 And mbIsBidding = False Then
        
        ExtCmdTimer.Enabled = False
        lCnt = lCnt + 1
        For i = 1 To giAktAnzArtikel
            With gtarrArtikelArray(i)
                If gsExtCmdPreCmd > "" And .Gebot > 0 And .Status <= [asNixLos] And _
                    .ExtCmdPreDone = False And .EndeZeit - MyNow <= myTimeSerial(0, 0, glExtCmdPreTime) And _
                    .EndeZeit - MyNow > myTimeSerial(0, 0, glExtCmdPreTime) - myTimeSerial(0, 0, glExtCmdTimeWindow) Then
                    
                    Call CallExtCmd(i, gsExtCmdPreCmd)
                    .ExtCmdPreDone = True
                End If
                
                If gsExtCmdPostCmd > "" And .Titel > "" And .ExtCmdPostDone = False And _
                    MyNow - .EndeZeit >= myTimeSerial(0, 0, glExtCmdPostTime) And _
                    MyNow - .EndeZeit < myTimeSerial(0, 0, glExtCmdPostTime) + myTimeSerial(0, 0, glExtCmdTimeWindow) Then
                    
                    Call CallExtCmd(i, gsExtCmdPostCmd)
                    .ExtCmdPostDone = True
                End If
            End With 'gtarrArtikelArray(i)
        Next i
        
        If gsExtCmdPeriodicCmd > "" And lCnt >= glExtCmdPeriodicTime And glExtCmdPeriodicTime > 0 Then
            
            On Error Resume Next
            Call Shell(gsExtCmdPeriodicCmd)
            lCnt = 0
            On Error GoTo 0
        End If
        ExtCmdTimer.Enabled = gbAutoMode
    End If 'giSuspendState = 0
    
End Sub

Private Sub Form_Activate()
    
    If mbFormLoaded Then
        Me.WindowState = giStartupSize
        mbFormLoaded = False
    End If
    
End Sub

Private Sub Form_Initialize()

    Set moResize = New clsDoResize
 
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call VersandkostenUebernehmen
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    gbWarSchonWach = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Not gbExplicitEnd And gbTrayAction Then 'aufgelöst, lg 29.05.03
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.WindowState = vbMinimized
    End If
End If

End Sub

Private Sub Gebot_GotFocus(Index As Integer)
  Call SetFocusRect(Index)
End Sub

Private Sub Gebot_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call SetFocusRect(Index)
End Sub

Private Sub Gebot_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo errhdl

If KeyCode = vbKeyDown Then '40
    If Index < giMaxRow Then
        If Gebot(Index + 1).Enabled Then
            Gebot(Index + 1).SetFocus
        Else
            Index = Index + 1
            If Index < giMaxRow Then
                If Gebot(Index + 1).Enabled Then
                    Gebot(Index + 1).SetFocus
                End If
            End If
        End If
    Else
        If VScroll1.Value < giAktAnzArtikel - giMaxRow Then
            Call Gebot_LostFocus(Index)
            VScroll1.Value = VScroll1.Value + 1
        End If
    End If

End If
If KeyCode = vbKeyUp Then '38
    If Index > 0 Then
        If Gebot(Index - 1).Enabled Then
            Gebot(Index - 1).SetFocus
        Else
            If Index > 1 Then
                Index = Index - 1
                If Gebot(Index - 1).Enabled Then
                    Gebot(Index - 1).SetFocus
                End If
            End If
        End If
    Else
        If VScroll1.Value > 1 Then
            Call Gebot_LostFocus(Index)
            VScroll1.Value = VScroll1.Value - 1
        End If
    End If
End If

If KeyCode = vbKeyReturn Then '13
    Call Gebot_LostFocus(Index)
End If

If KeyCode = vbKeyEscape Then '27
    Call ArtikelArrayToScreen(Index + VScroll1.Value, False, True)
End If

If (KeyCode = vbKeyDelete Or KeyCode = vbKeyBack) And Trim(Gebot(Index).Text = "") Then
    Call Gebot_LostFocus(Index)
End If

If KeyCode >= vbKey0 Then '48
   GebotTimer.Enabled = True
   miGebotsIndex = Index
End If

If KeyCode = vbKeyPageUp Then '33
    Call Gebot_LostFocus(Index)
    If VScroll1.Value > giMaxRow Then
        VScroll1.Value = VScroll1.Value - giMaxRow
    Else
        VScroll1.Value = VScroll1.Min
    End If
End If
If KeyCode = vbKeyPageDown Then '34
    Call Gebot_LostFocus(Index)
    If VScroll1.Value + giMaxRow < VScroll1.Max Then
        VScroll1.Value = VScroll1.Value + giMaxRow
    Else
        VScroll1.Value = VScroll1.Max
    End If
End If

errhdl:
End Sub

Private Sub Gebot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetFocusRect(Index)
End Sub

Private Sub Bietgruppe_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetFocusRect(Index)
End Sub

Private Sub Artikel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetFocusRect(Index)
End Sub

Private Sub Gebot_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub GebotTimer_Timer()
    If giSuspendState = 0 Then Call Gebot_LostFocus(miGebotsIndex)
End Sub

Private Sub Gruppe_Timer_Timer()
    If giSuspendState = 0 Then Call Bietgruppe_LostFocus(miGruppeIndex)
End Sub

Private Sub Mailbufftimer_Timer()
        
    Dim sBuffer As String
    Dim sSendTo As String
    
    If giSuspendState = 0 Then
        
        If ReadMailBuff(sBuffer) Then
            'Mail verschicken
            If sBuffer Like "To: *" & vbCrLf & "*" Then
                sSendTo = Mid(sBuffer, 5, InStr(1, sBuffer, vbCrLf) - 5)
                sBuffer = Mid(sBuffer, InStr(1, sBuffer, vbCrLf) + 2)
            Else
                sSendTo = gsSendEndTo
            End If
            Call SendSMTP(gsSendEndFromRealname & "<" & gsSendEndFrom & ">", sSendTo, sBuffer)
        End If
        MailBuffTimer.Enabled = CBool(MailBuffTimer.Tag)
        
    End If 'giSuspendState = 0
    
End Sub

Private Sub mnuAccount_Click()

    If giUserAnzahl > 0 Then
        If gtarrArtikelArray(miGlobArtikel).Status = [asEnde] Or _
            gtarrArtikelArray(miGlobArtikel).Status = [asOK] Then
            
            MsgBox gsarrLangTxt(735), vbExclamation Or vbOKOnly
        Else
            giArtChoose = miGlobArtikel
            frmChooseUser.Show vbModal, Me
        End If
    Else
        MsgBox gsarrLangTxt(736), vbExclamation Or vbOKOnly
    End If
    
End Sub

Private Sub mnuCleanupArtikel_Click()
    
    Dim i As Integer
    Dim lRet As VbMsgBoxResult
    
    lRet = MsgBox(gsarrLangTxt(266), vbYesNo Or vbQuestion, gsarrLangTxt(59))
    If lRet = vbYes Then
        Call frmProgress.InitProgress(0, giAktAnzArtikel)
        DoEvents
        
        For i = giAktAnzArtikel To 1 Step -1
            If GetRestzeitFromItem(i) = 0 Then Call RemoveArtikel(i, False, False)
            frmProgress.Step
        Next
        frmProgress.TerminateProgress
    End If
    Call RemoveArtikel(0, True, True)  ' Jetzt noch die Anzeige refreshen
          
End Sub

Private Sub mnuCleanupArtikel2_Click()
        
    Dim i As Integer
    Dim lRet As VbMsgBoxResult
    
    lRet = MsgBox(gsarrLangTxt(750), vbYesNo Or vbQuestion, gsarrLangTxt(751))
    If lRet = vbYes Then
        frmProgress.InitProgress 0, giAktAnzArtikel
        DoEvents
        For i = giAktAnzArtikel To 1 Step -1
            If gtarrArtikelArray(i).Status = [asCancelGroup] Or gtarrArtikelArray(i).Status = [asBuyOnlyCanceled] Then
                Call RemoveArtikel(i, False, False)
            End If
            frmProgress.Step
        Next
        frmProgress.TerminateProgress
    End If
    Call RemoveArtikel(0, True, True)  ' Jetzt noch die Anzeige refreshen
    
End Sub

Private Sub mnuDeleteArtikel_Click()

    Dim i As Integer
    Dim lRet As VbMsgBoxResult
    
    lRet = MsgBox(gsarrLangTxt(753), vbYesNo Or vbQuestion Or vbDefaultButton2, gsarrLangTxt(754))
    If lRet = vbYes Then
        frmProgress.InitProgress 0, giAktAnzArtikel
        DoEvents
        For i = giAktAnzArtikel To 1 Step -1
            Call RemoveArtikel(i, False, False)
            frmProgress.Step
        Next
        frmProgress.TerminateProgress
    End If
    Call RemoveArtikel(0, True, True)  ' Jetzt noch die Anzeige refreshen

End Sub

Private Sub mnuComment_Click()
    
    frmKommentar.SetArtikelID = miGlobArtikel
    Load frmKommentar
    frmKommentar.Show vbModal, Me
    
    Call ArtikelArrayToScreen(VScroll1.Value)
    
End Sub
Private Sub mnuAbout2_Click()
Call mnuAbout_Click
End Sub

Private Sub mnuAkt_Click()
    Call Preis_MouseDown(miGlobArtikel - VScroll1.Value, 1, 0, 0, 0)
End Sub

Private Sub mnuArtikel_Click()
    frmNeuerArtikel.Show  'MD-Marker , nicht Modal anzeigen ? L: Natürlich nicht, BOM soll doch nebenbei bedienbar bleiben!!!
End Sub

Private Sub mnuAuto_Click()
     
   gbAutoMode = Not gbAutoMode
   Call CheckAutoMode
    
End Sub

Private Sub mnuBid_Click()
    
    Dim lRet As VbMsgBoxResult
    Dim bOk As Boolean
    Dim sUser As String
    Dim sMsg As String
    Dim bByPass As Boolean
    Dim sItem As String
    
    Call Gebot_LostFocus(miGlobArtikel - VScroll1.Value)  'Übernahme der Werte vor dem Bieten, lg 19.04.03
    
    If LenB(gsUser) Or LenB(gsPass) Or giDefaultUser > 0 Or giUserAnzahl > 0 Then
        
        With gtarrArtikelArray(miGlobArtikel)
            If .Gebot = 0 And Not StatusIstBuyItNowStatus(.Status) Then
                MsgBox gsarrLangTxt(3) & .Artikel & gsarrLangTxt(4)
            Else
                
                'sh luser auf Bietaccount setzen
                If .eBayUser = "" Then
                    
                    If giUserAnzahl > 0 Then
                        If .UserAccount <> "" And .UserAccount <> gsUser Then
                            sUser = .UserAccount
                        Else
                            sUser = gsUser
                        End If
                    End If
                Else
                    sUser = .eBayUser
                End If
                
                If StatusIstBuyItNowStatus(.Status) Then
                    lRet = MsgBox(gsarrLangTxt(5) & sUser & vbCrLf & gsarrLangTxt(732) & ": " & .Artikel & vbCrLf & .Titel & " ?", vbYesNo)
                Else
                    lRet = MsgBox(gsarrLangTxt(5) & sUser & vbCrLf & gsarrLangTxt(6) & .Artikel & vbCrLf & .Titel & vbCrLf & Format(.Gebot, "###,##0.00") & " " & .WE & " ?", vbYesNo)
                End If
                If lRet = vbYes Then
                    
                    bByPass = True
                    If Not CheckInternetConnection Then
                        Call Ask_Online
                        If Not IsOnline Then
                            bByPass = False
                        End If
                    End If
                    
                    If bByPass Then
                        sItem = .Artikel ' Artikel-Nr merken
                        mbIsBidding = True
                        bOk = Bieten(.Artikel, .Gebot, .eBayUser, .eBayPass, miGlobArtikel, True, .UseToken, StatusIstBuyItNowStatus(.Status))
                        mbIsBidding = False
                        
                        miGlobArtikel = ItemToIndex(sItem)
                        If gbUpdateAfterManualBid And miGlobArtikel > 0 Then Call Upd_Art(miGlobArtikel, False)
                                            
                        sMsg = IIf(bOk, gsarrLangTxt(7), gsarrLangTxt(8))
                        
                        If gbQuietAfterManualBid Then
                            If Not bOk Then sMsg = gsarrLangTxt(95)
                            Call PanelText(StatusBar1, 2, sMsg, True, IIf(bOk, vbGreen, vbRed))
                        Else
                            MsgBox sMsg
                            If Not bOk Then Call ShowStatus(sItem)
                        End If
                        
                        If gbUsesModem And gbLastDialupWasManually Then Call Ask_Offline
                   End If 'bByPass
                End If 'lRet = vbYes
            End If 'gtarrArtikelArray(miGlobArtikel).Gebot = 0 And Not
        End With 'gtarrArtikelArray(miGlobArtikel)
    Else
        MsgBox gsarrLangTxt(2)
    End If 'LenB(gsUser) Or LenB(gsPass) Or giDefaultUser > 0 Or giUserAnzahl > 0
    
End Sub

Private Sub mnuBrowser_Click()

gsGlobalUrl = "http://" & gsMainUrl
Call ShowBrowser(Me.hWnd)

End Sub

Private Sub mnuBrowse_Click()

Dim bBrowserTmp As Boolean

bBrowserTmp = gbOpenBrowserOnClick
gbOpenBrowserOnClick = True
Call Titel_MouseDown(miGlobArtikel - VScroll1.Value, 1, 0, 0, 0)
gbOpenBrowserOnClick = bBrowserTmp

End Sub

Private Sub mnuDel_Click()
Dim iIdx As Integer
Dim lRet As VbMsgBoxResult

iIdx = miGlobArtikel - VScroll1.Value

If gbConfirmDelete Then
    lRet = MsgBox(gsarrLangTxt(51) & "?", vbYesNo Or vbQuestion, gsarrLangTxt(51))
Else
    lRet = vbYes
End If

If lRet = vbYes Then
    Artikel(iIdx).Text = ""
    Call Artikel_LostFocus(iIdx)
End If

End Sub

Private Sub mnuEditShipping_Click()
    Call VersandkostenBearbeiten(miGlobArtikel - VScroll1.Value)
End Sub

Private Sub mnuHomepage_Click()
    Call ExecuteDoc(Me.hWnd, gsBOMUrlHP)
End Sub

Private Sub mnul_Click(Index As Integer)
Dim i As Integer

For i = 1 To 9
    mnul(i).Checked = False
Next i

mnul(Index).Checked = True
gsAktLanguage = mnul(Index).Caption
Call SelectLanguage(gsAktLanguage)
Call SetLanguage

End Sub

Private Sub mnuCurrUpdate_Click()
    
    Call UpdateCurrencies
    Call ArtikelArrayToScreen(VScroll1.Value)
    
End Sub

Private Sub mnuls_Click(Index As Integer)
    
    Call ShowTool(gsarrAnsPsLinkA(Index), gsarrAnsPsEncodingA(Index), miGlobArtikel, gbarrAnsPsEditA(Index), gsarrAnsPsNameA(Index))
    
End Sub

Private Sub mnult_Click(Index As Integer)
    
    Call ShowTool(gsarrAnsToolLinkA(Index), gsarrAnsToolEncodingA(Index), miGlobArtikel, gbarrAnsToolEditA(Index), gsarrAnsToolNameA(Index))
    
End Sub

Public Sub ShowTool(ByVal sUrl As String, ByVal sEncoding As String, iIdx As Integer, bEdit As Boolean, sToolName As String)

  Dim vntKeyname As Variant
  Dim vntKeywert As Variant
  Dim sTmp As String
  Dim i As Integer
  
  With gtarrArtikelArray(iIdx)
    vntKeyname = Array("url", "seller", "item", "highbidder", "title", "location", "price", "currency", "group", "comment", "bid", "endtime", "bidcount", "minbid", "timeleft", "timenext", "status")
    vntKeywert = Array("http://" & gsScript4 & gsScriptCommand4 & gsCmdViewItem, .Verkaeufer, .Artikel, .Bieter, .Titel, .Standort, Format(.AktPreis, "###,##0.00"), .WE, .Gruppe, .Kommentar, Format(.Gebot, "###,##0.00"), Date2Str(.EndeZeit), .AnzGebote, Format(.MinGebot, "###,##0.00"), Abs(Date2UnixDate(.EndeZeit) - Date2UnixDate(MyNow)), Date2UnixDate(MyNow + gfRestzeitZaehler) - Date2UnixDate(MyNow), IIf(.Status = [asOK], 1, 0))
  End With
  
  
  For i = LBound(vntKeywert) To UBound(vntKeywert)
    If InStr(1, sUrl, "[" & vntKeyname(i) & "]", vbTextCompare) > 0 Then
      If bEdit Then
        sTmp = InputBox(sToolName, , vntKeywert(i))
        If sTmp = "" And vntKeywert(i) > "" Then Exit Sub
        vntKeywert(i) = sTmp
      End If
      If sEncoding = "utf-8" Then vntKeywert(i) = Encode_UTF8(CStr(vntKeywert(i)))
      If sUrl Like "http*" Then vntKeywert(i) = URLEncode(CStr(vntKeywert(i)))
      sUrl = Replace(sUrl, "[" & vntKeyname(i) & "]", vntKeywert(i), , , vbTextCompare)
    End If
  Next i
  
  If sUrl > "" Then
    If sUrl Like "http*" Then
      gsGlobalUrl = sUrl
      ShowBrowser (Me.hWnd)
    ElseIf sUrl Like "setclipboard:*" Then
      Clipboard.Clear
      Clipboard.SetText Mid(sUrl, 14)
    Else
      ExecuteDoc Me.hWnd, sUrl, "", True
    End If
  End If

End Sub

Private Sub mnuReleasenotes_Click()
    
    Call ExecuteDoc(Me.hWnd, App.Path & "\WhatsNew.txt")
    
End Sub

Private Sub mnuSearch_Click()
    
    Dim sTmp As String
    
    sTmp = InputBox(gsarrLangTxt(746), gsarrLangTxt(744), gsSearchTerm)
    
    If sTmp > "" Then
        gsSearchTerm = sTmp
        Call Search
    End If
    
End Sub

Private Sub mnuSearchContinue_Click()
    
    If gsSearchTerm = "" Then
        Call mnuSearch_Click
    Else
        Call Search
    End If
    
End Sub

Private Sub Search()

  Dim sTmp As String
  Dim lIdx As Long
  Dim vntDynArray As Variant
  Dim lLoopCount As Long
  Dim lMinVisible As Long
  Dim lMaxVisible As Long
  Dim lLastSearchPosition As Long
  
  If giAktAnzArtikel = 0 Then Exit Sub
  
  lLastSearchPosition = glSearchPosition
  
  Do While (True)
    glSearchPosition = glSearchPosition + 1
    If glSearchPosition > giAktAnzArtikel Then glSearchPosition = 1: lLoopCount = lLoopCount + 1
    If lLoopCount > 2 Then Exit Sub
    
    vntDynArray = GetDynArrayFromArtikelZeile(gtarrArtikelArray(glSearchPosition))
    sTmp = Join(vntDynArray, vbTab)
    
    If UCase(sTmp) Like "*" & UCase(gsSearchTerm) & "*" Then
      
      lMinVisible = VScroll1.Value
      lMaxVisible = VScroll1.Value + giMaxRow
      
      If glSearchPosition < lMinVisible Then VScroll1.Value = VScroll1.Value - (lMinVisible - glSearchPosition)
      If glSearchPosition > lMaxVisible Then VScroll1.Value = VScroll1.Value + (glSearchPosition - lMaxVisible)
      
      lIdx = glSearchPosition - VScroll1.Value
      
      Do While TitelFlashTimer.Enabled ' evtl. Timer von vorheriger Suche schnell ablaufen lassen und solange warten
        TitelFlashTimer.Interval = 1
        DoEvents
      Loop
      
      TitelFlashTimer.Tag = "6," & CStr(lIdx) & ",0"
      TitelFlashTimer.Interval = 100
      TitelFlashTimer.Enabled = True
      
      Exit Sub
    End If

  Loop
  
End Sub

Private Sub mnuSendItemTo_Click()
    
    Call SendenAn(miGlobArtikel)
    
End Sub

Private Sub mnuVersion_Click()
      
    Call CheckUpdate(Me)
    
End Sub

Private Sub mnuExit2_Click()
    Call mnuExit_Click
End Sub

Private Sub mnuUpdateArtikel_Click()
Call LoadArtikel
End Sub

Private Sub mnuMax_Click()
    If Not gbAboutIsUp Then Call FromTaskbar
End Sub

Private Sub mnuMyEbay_Click()
    Call LoadMyEbay
End Sub

Private Sub mnuPasswd_Click()
    Call LogIn
End Sub

Private Sub mnuSettings_Click()
Call Einstellungen
End Sub

Private Sub mnuSync_Click()
Call TimeSync
End Sub

Private Sub NewArtikel_Click()
Call frmNeuerArtikel.Show  'MD-Marker , nicht Modal anzeigen ? L: Natürlich nicht, BOM soll doch nebenbei bedienbar bleiben!!!
End Sub

Private Sub NTP_Connect()
  gfNtpDelay = Timer
End Sub

Private Sub NTP_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    On Error Resume Next
    
    NTP.GetData strData, vbString
    gsNtpData = gsNtpData & strData

End Sub

Private Sub NTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    giNtpErr = 1 'True
End Sub

Private Sub ODBC_Timer_Timer()
    
    If giSuspendState = 0 And mbIsBidding = False Then
    
        ODBC_Timer.Interval = 60000
        miOdbcTimerCount = miOdbcTimerCount + 1
        Zusatzfeld.Text = "ODBC in " & CStr(giOdbcZyklus - miOdbcTimerCount) & " Min"
        
        If miOdbcTimerCount < giOdbcZyklus Then Exit Sub  'MD-Marker
        
        miOdbcTimerCount = 0
        
        ODBC_Timer.Enabled = False
                
        With Zusatzfeld
            If ODBC_Check Then .BackColor = vbGreen
            
            If Not gsOdbcStopRead Then
                DoEvents
                .Text = "ODBC Read"
                Call ODBC_ReadNew
                If gsOdbcStopRead Then
                    .BackColor = vbRed
                    .Text = "ODBC STOP"
                Else
                    .BackColor = vbGreen
                End If
            End If
            
            If Not gsOdbcStopRead Then
                .Text = "ODBC Remove"
                Call ODBC_RemoveArtikel
            End If
            
            If Not gsOdbcStopRead Then
                .Text = "ODBC Update"
                Call ODBC_UpdateArtikel
                If Not gsOdbcStopRead Then
                    .Text = " ODBC ok"
                    'Zusatzfeld.BackColor = &H8000000F
                End If
            End If
        End With 'Zusatzfeld
        
        Call ArtikelArrayToScreen(VScroll1.Value)
        
        gsOdbcStopRead = False
        ODBC_Timer.Enabled = True
    End If 'giSuspendState = 0
    
End Sub

Private Sub Preis_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      
    Call SetFocusRect(Index)
    Call VersandkostenUebernehmen
    Call Gebot_LostFocus(giLastGebotEditedIndex)
    
    If Button = vbLeftButton Then '1
        If Artikel(Index).Text <> "" Then
            If Not CheckInternetConnection Then
                Call Ask_Online
                If Not IsOnline Then
                    Exit Sub  'MD-Marker
                End If
            End If
            
            Call Upd_Art(Index + VScroll1.Value)
            Call ArtikelArrayToScreen(VScroll1.Value)
            
            If gbUsesModem And gbLastDialupWasManually Then Ask_Offline
        End If
    ElseIf Button = vbRightButton Then '2
        Call ShowContextMenu(Index)
    End If
    
End Sub

Private Sub Preis_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)
    
End Sub

Private Sub QuitTimer_Timer()
    
    If giSuspendState = 0 Then
        QuitTimer.Enabled = False
        gbWarnenBeimBeenden = False
        mbQuietExit = True
        Call mnuExit_Click
    End If
    
End Sub

Private Sub RechnerZeitTimer_Timer()
    
    Static UpdateCheckCounter As Long
    Dim sTmp As String
    
    If giSuspendState = 0 Then
    
        sTmp = gsarrLangTxt(215) & _
            IIf(gbShowTitleVersion, " " & GetBOMVersion() & IIf(gbNewBOMVersionAvailable, " *", ""), "") & _
            IIf(gbShowTitleAuctionHome, "      " & gsarrLangTxt(319) & gsAuctionHome, "") & _
            IIf(gbShowTitleDefaultUser, "      " & gsarrLangTxt(47) & ": " & gsUser, "") & _
            IIf(gbShowTitleTimeLeft And gbAutoMode, "      " & gsarrLangTxt(87) & " " & TimeLeft2String(gfRestzeitZaehler), "") & _
            IIf(gbShowTitleDateTime, "      " & gsarrLangTxt(86) & " " & Date2Str(MyNow, gbShowWeekday, gsSpecialDateFormat), "")
            
        If msCaptionCache <> sTmp Then
            msCaptionCache = sTmp
            'Me.Caption = sTmp
            Call SetWindowText(Me.hWnd, sTmp)
        End If
        
        If Not gbAutoMode And Not gbCountDownInAutomodeOnly Then Call ArtikelArrayToScreen(VScroll1.Value, True)
        
        If gbCleanStatus Then Call PanelReset
        
        Call WatchTrayIcon
        Call WatchArtikelZeilen
        
        UpdateCheckCounter = UpdateCheckCounter + 1
        If UpdateCheckCounter > glCheckForUpdateInterval * 60 * 60 Then
            If gfRestzeitZaehler > myTimeSerial(0, 1, 0) Or gbAutoMode = False Then
                If CheckInternetConnection Then Call CheckUpdate(Me, True): UpdateCheckCounter = 0
            Else
                UpdateCheckCounter = UpdateCheckCounter - 90
            End If
        End If
        
    End If 'giSuspendState = 0
    
End Sub

Private Sub Restzeit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      
    Call SetFocusRect(Index)
    Call VersandkostenUebernehmen
    If Button = vbRightButton Then Call ShowContextMenu(Index)
    
End Sub

Private Sub Restzeit_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)
    
End Sub

Private Sub SortEnde_Click()
    
    'Sort nach Ende ;-)
    gsSortOrder = IIf(gsSortOrder = "asc", "desc", "asc")
    Call Sortiere
    
End Sub

Private Sub Sortiere()
On Error GoTo errhdl

If giAktAnzArtikel < 2 Then
    ArtikelArrayToScreen VScroll1.Value
    Exit Sub
End If

'ab in den Sorter ..
Call QuickSortDate(gtarrArtikelArray, 1, UBound(gtarrArtikelArray), (gsSortOrder = "asc")) 'lg 10.07.2003

'und die Ausgabe wieder aufbereiten ..

ArtikelArrayToScreen VScroll1.Value

errhdl:

End Sub

Private Sub LoadArtikel()
'Infos für Artikel Updaten
Dim i As Integer
Dim iAktRow As Integer
Dim iAktAnzahlArtikelNext As Integer 'sh 23.10.03

If Not CheckInternetConnection Then
    Ask_Online
    If Not IsOnline Then
        Exit Sub
    End If
End If

If Not Check_ebayUp Then Exit Sub

ArtikelTimer.Enabled = False

Screen.MousePointer = vbHourglass
Toolbar1.Buttons(8).Enabled = False
'Toolbar1.Buttons(4).Enabled = False
SetToolbarImage Toolbar1.Buttons(4), 3
Toolbar1.Buttons(4).ToolTipText = gsarrLangTxt(606)
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(6).Enabled = False
AutoStatus.Enabled = False

'lg 24.05.03
Toolbar1.Buttons(12).Enabled = False
NewArtikel.Enabled = False
SortEnde.Enabled = False
mnuUpdateArtikel.Enabled = False
mnuAuto.Enabled = False
mnuMyEbay.Enabled = False
mnuDeleteArtikel.Enabled = False
mnuCleanupArtikel.Enabled = False
mnuCleanupArtikel2.Enabled = False
mnuReadArtikel.Enabled = False
mnuArtikel.Enabled = False
For i = 0 To IIf(Artikel(0).Tag = "in", giMaxRow, 0)
       Artikel(i).Enabled = False
Next i
'lg 24.05.03

On Error Resume Next
mbStopUpdate = False

'sh 25.10.03 Nur nächsten anstehenden aktualisieren
If Not gbAAnext Then
  iAktAnzahlArtikelNext = giAktAnzArtikel: i = 1
Else   'abgelaufene aussortieren
  iAktAnzahlArtikelNext = AnzValidArtikel(False): i = iAktAnzahlArtikelNext
End If

SortEnde.Enabled = False
For iAktRow = i To iAktAnzahlArtikelNext
  Call PanelText(StatusBar1, 2, gsarrLangTxt(34) & " " & gsarrLangTxt(31) & " " & CStr(iAktRow) & " / " & CStr(giAktAnzArtikel))
  If Not gtarrArtikelArray(iAktRow).Artikel = "" _
  And Not gtarrArtikelArray(iAktRow).PostUpdateDone _
  And Not mbStopUpdate Then
  
    Call Upd_Art(iAktRow, vbNullString, False)

    If iAktRow - VScroll1.Value <= giMaxRow Then
      ArtikelArrayToScreen VScroll1.Value
    End If
  End If
Next iAktRow

SortEnde.Enabled = True
If gbAAnext Then gbAAnext = False

Call SetToolbarImage(Toolbar1.Buttons(4), 8)
Toolbar1.Buttons(4).ToolTipText = gsarrLangTxt(37)
Call PanelText(StatusBar1, 2, "")

'Label9_Click

Screen.MousePointer = vbNormal
Toolbar1.Buttons(8).Enabled = True
Toolbar1.Buttons(4).Enabled = True
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(6).Enabled = True
AutoStatus.Enabled = True
'lg 24.05.03
Toolbar1.Buttons(12).Enabled = True
NewArtikel.Enabled = True
SortEnde.Enabled = True
mnuUpdateArtikel.Enabled = True
mnuAuto.Enabled = True
mnuMyEbay.Enabled = True
mnuDeleteArtikel.Enabled = True
mnuCleanupArtikel.Enabled = True
mnuCleanupArtikel2.Enabled = True
mnuReadArtikel.Enabled = True
mnuArtikel.Enabled = True
For i = 0 To IIf(Artikel(0).Tag = "in", giMaxRow, 0)
       Artikel(i).Enabled = True
Next i

ArtikelTimer.Enabled = gbAutoMode
miArtikelCycleCount = 0

If gbUsesModem And gbLastDialupWasManually Then Call Ask_Offline

End Sub

Private Function sucheEnde(str As String, AktRow As Integer, EndeZeit As Date) As String

Dim lPos As Long
Dim lEpochPos As Long
Dim lEpochSecs As Long
Dim sLocStr As String
Dim lPosTmp As Long
Dim lBeschreibungPos As Long
Dim fOffsetLocal As Double
Dim fOffsetLocalKorrektur As Double

On Error Resume Next
'SaveToFile str, ".\Bieten.html"    'nur zum Debuggen

lBeschreibungPos = InStr(1, str, gsAnsDescriptionBegin)
If lBeschreibungPos = 0 Then lBeschreibungPos = 2000000

lPos = InStr(1, str, gsAnsInvalid)

If lPos > 0 And lPos < lBeschreibungPos Then
    sucheEnde = "Invalid"
    Exit Function
End If

lPos = InStr(1, str, gsAnsEndTime)
If lPos > lBeschreibungPos Then lPos = 0
If lPos = 0 Then lPos = InStr(1, str, gsAnsEndTime2) 'nach 2tem Endezeit String suchen
If lPos = 0 Then lPos = InStr(1, str, gsAnsEndTime) 'nochmal nach 1tem Endezeit String suchen und diesmal lBeschreibungPos ignorieren
If lPos = 0 Then lEpochPos = InStr(1, str, gsAnsEndTimeEpoch)  'nach Epochen Endezeit String suchen

If lPos = 0 And lEpochPos = 0 Then
'    sucheEnde = "Fehler"
    sucheEnde = "Ok"
    EndeZeit = gdatENDEZEITNOTFOUND
    Exit Function
End If

If lEpochPos > 0 Then ' Epochen-Timestamp
    sLocStr = Mid(str, lEpochPos + Len(gsAnsEndTimeEpoch), 10)
    lEpochSecs = Val(sLocStr)
    If lEpochSecs > 1000000000 And lEpochSecs < 2000000000 Then
        EndeZeit = UnixDate2Date(lEpochSecs)
    Else
        sucheEnde = "Fehler"
        Exit Function
    End If
Else ' normale Endzeit-Angabe
    lPosTmp = lPos
    
    lPos = InStr(lPosTmp, str, gsAnsTime1_1)
    If lPos > 0 Then fOffsetLocal = gsAnsOffsetLocal1_1
    
    If lPos = 0 Or (lPos - lPosTmp) > gsAnsEndTimeMaxLen Then
        'keine Sommerzeit?
        lPos = InStr(lPosTmp, str, gsAnsTime1_2)
        If lPos > 0 Then fOffsetLocal = gsAnsOffsetLocal1_2
    End If
    
    If lPos = 0 Then
        sucheEnde = "Fehler"
        Exit Function
    End If
    
    sLocStr = Mid(str, lPos - 100, 100)
    sLocStr = HtmlCleanup(sLocStr, True)
    sLocStr = ConvertMonthname1(sLocStr)
    
    EndeZeit = DateAdd("n", 60 * (GetUTCOffset() - fOffsetLocal), Str2Date(sLocStr, gsAnsDateFormat1))
    
    If EndeZeit > 0 And gsAnsTime1_1 = gsAnsTime1_2 Then
      fOffsetLocalKorrektur = -GetOffsetLocalFromDate(EndeZeit)
      EndeZeit = DateAdd("h", fOffsetLocalKorrektur, EndeZeit)
    End If
End If

If EndeZeit < MyNow() Then
  sucheEnde = "Beendet"
Else
  sucheEnde = "Ok"
End If

End Function
Private Function sucheTitel(sTxt As String) As String
Dim lPos As Long
Dim lPosStart As Long
Dim sTmp As String
Dim sChar As String
Dim i As Integer

On Error Resume Next
'SaveToFile sTxt, ".\Artikel.html"
sucheTitel = ""
lPos = InStr(1, sTxt, gsAnsTitle)
lPosStart = InStr(lPos, sTxt, gsAnsTitleStart) + Len(gsAnsTitleStart)
lPos = InStr(lPosStart, sTxt, gsAnsTitleEnd)

sTmp = Mid(sTxt, lPosStart, lPos - lPosStart)
lPosStart = InStr(1, sTmp, "&#")
Do While lPosStart > 0 'sonderzeichen!
    sucheTitel = sucheTitel & Left(sTmp, lPosStart - 1)
    lPos = InStr(lPosStart, sTmp, ";") - 1
    If lPos = -1 Then Exit Do ' damit es keine Endlosschleife wird, lg 07.06.03
    sChar = Mid(sTmp, lPosStart + 2, lPos - lPosStart - 1)
    sChar = Chr(sChar)
    sucheTitel = sucheTitel & sChar
    sTmp = Mid(sTmp, lPos + 2, Len(sTmp) - lPos)
    lPosStart = InStr(1, sTmp, "&#")
Loop
sucheTitel = sucheTitel & sTmp

sucheTitel = HtmlCleanup(sucheTitel)
sucheTitel = HtmlZeichenConvert(sucheTitel)


'Sonderzeichen rauswerfen
For i = 1 To Len(sucheTitel)
    If Mid$(sucheTitel, i, 1) < " " Then Mid$(sucheTitel, i, 1) = " "
Next i

sucheTitel = Left(Trim(sucheTitel), 300)

End Function

Private Function sucheBieter(sTxt As String, iNumBids As Integer, sBieterAlt As String) As String
Dim lPos As Long
Dim lPos2 As Long
Dim lPosStart As Long
Dim sTmp As String
Dim lBeschreibungPos As Long

On Error Resume Next

lBeschreibungPos = InStr(1, sTxt, gsAnsDescriptionBegin)
If lBeschreibungPos = 0 Then lBeschreibungPos = 2000000

sucheBieter = ""
sTmp = ""

lPosStart = InStr(1, sTxt, gsAnsDutch)
If lPosStart > 0 Then
    'Powerauktion
    'Anzahl raussuchen
    sTmp = sucheMenge(sTxt)
    sucheBieter = gsarrLangTxt(267) & sTmp & " " & gsarrLangTxt(31)
    Exit Function
End If

lPosStart = InStr(1, sTxt, gsAnsPrivat)
If lPosStart > 0 And lPosStart < lBeschreibungPos Then
    'Privatauktion
    sucheBieter = gsarrLangTxt(268)
    Exit Function
End If

lPosStart = InStr(1, sTxt, gsAnsAnonBidder)
If lPosStart > 0 And lPosStart < lBeschreibungPos Then
    'anonymer Bieter
    lPosStart = InStr(lPosStart + Len(gsAnsAnonBidder), sTxt, gsAnsAnonBidderStart)
    lPos = InStr(lPosStart + Len(gsAnsAnonBidderStart), sTxt, gsAnsAnonBidderEnd)
    sucheBieter = gsarrLangTxt(747)
    If lPosStart > 0 And lPos > lPosStart Then
      lPosStart = lPosStart + Len(gsAnsAnonBidderStart)
      sucheBieter = Mid(sTxt, lPosStart, lPos - lPosStart)
    End If
    Exit Function
End If

lPos = 0
If lPos = 0 Or lPos > lBeschreibungPos Then lPos = InStr(1, sTxt, gsAnsWinner)
If lPos = 0 Or lPos > lBeschreibungPos Then lPos = InStr(1, sTxt, gsAnsBidder)
If lPos = 0 Or lPos > lBeschreibungPos Then lPos = InStr(1, sTxt, gsAnsBuyer)

If lPos > 0 And iNumBids > 0 Then
    lPosStart = InStr(lPos, sTxt, gsAnsUserIDStart) + Len(gsAnsUserIDStart)   '+ 2
    If lPosStart > Len(gsAnsUserIDStart) + 2 Then
        lPos = InStr(lPosStart, sTxt, gsAnsUserIdEnd)
        lPos2 = InStr(lPosStart, sTxt, gsAnsUserIdEnd2)
        If lPos2 > 0 And lPos2 < lPos Then lPos = lPos2
        sTmp = Mid(sTxt, lPosStart, lPos - lPosStart)
    End If
    sucheBieter = URLDecode(sTmp)
End If

' wenn es vorher eine Privat- oder Powerauktion war und jetzt gar nichts erkannt wurde
' ist es evtl. eine abgelaufene Privat- oder Powerauktion. Wir nehmen den vorherigen Wert
lPos = InStr(1, sBieterAlt, gsarrLangTxt(267), vbTextCompare) + _
      InStr(1, sBieterAlt, gsarrLangTxt(268), vbTextCompare) + _
      InStr(1, sBieterAlt, gsarrLangTxt(747), vbTextCompare)
      
If sucheBieter = "" And lPos > 0 Then sucheBieter = sBieterAlt

End Function

Private Function sucheAnzGebote(sTxt As String) As Long

Dim lPos As Long
Dim lPosEnd  As Long
Dim sTmp As String

On Error Resume Next

If sucheAnzGebote = 0 Then
  lPos = InStr(1, sTxt, gsAnsHistory)
  If lPos > 0 Then
    lPos = InStr(lPos, sTxt, gsAnsNumBids)
  End If
  If lPos > 0 Then
    lPos = InStr(lPos, sTxt, gsAnsNumBidsStart) + Len(gsAnsNumBidsStart)
    lPosEnd = InStr(lPos, sTxt, gsAnsNumBidsEnd)
    sTmp = Mid(sTxt, lPos, lPosEnd - lPos)
    sTmp = HtmlCleanup(sTmp)
    sucheAnzGebote = Val(sTmp)
  End If
End If

If sucheAnzGebote = 0 Then
  lPos = InStr(1, sTxt, gsAnsHistory)
  If lPos > 0 Then
    lPos = InStr(lPos, sTxt, gsAnsNumBids2)
  End If
  If lPos > 0 Then
    lPos = InStr(lPos, sTxt, gsAnsNumBidsStart2) + Len(gsAnsNumBidsStart2)
    lPosEnd = InStr(lPos, sTxt, gsAnsNumBidsEnd2)
    sTmp = Mid(sTxt, lPos, lPosEnd - lPos)
    sTmp = HtmlCleanup(sTmp)
    sucheAnzGebote = Val(sTmp)
  End If
End If

If sucheAnzGebote = 0 Then
  lPos = InStr(1, sTxt, gsAnsHistory)
  If lPos > 0 Then
    lPos = InStr(lPos, sTxt, gsAnsNumBids3)
  End If
  If lPos > 0 Then
    lPos = InStr(lPos, sTxt, gsAnsNumBidsStart3) + Len(gsAnsNumBidsStart3)
    lPosEnd = InStr(lPos, sTxt, gsAnsNumBidsEnd3)
    sTmp = Mid(sTxt, lPos, lPosEnd - lPos)
    sTmp = HtmlCleanup(sTmp)
    sucheAnzGebote = Val(sTmp)
  End If
End If

If sucheAnzGebote = 0 Then
  lPos = InStr(1, sTxt, gsAnsDutch)
  If lPos > 0 Then 'Powerauktion
    sucheAnzGebote = -1
  End If
End If

End Function

Private Function sucheVK(sTxt As String) As String

Dim lPos As Long
Dim lPosStart As Long
Dim sTmp As String

On Error Resume Next

sTmp = ""

lPos = InStr(1, sTxt, gsAnsAskSeller)

If lPos > 0 Then
    lPosStart = InStr(lPos, sTxt, gsAnsAskSellerStart) + Len(gsAnsAskSellerStart)
    lPos = InStr(lPosStart, sTxt, gsAnsAskSellerEnd)

    If lPosStart > Len(gsAnsAskSellerStart) Then
        sTmp = Mid(sTxt, lPosStart, lPos - lPosStart)
    End If
End If

sucheVK = URLDecode(sTmp)

End Function

Private Function sucheAktGebot(sTxt As String, sWe As String, bFlag As Boolean) As String

Dim lPosStart As Long
Dim lPosEnd As Long
Dim lBeschreibungPos As Long
Dim i As Integer
Dim sTmp As String

On Error GoTo errhdl

lBeschreibungPos = InStr(1, sTxt, gsAnsDescriptionBegin)
If lBeschreibungPos = 0 Then lBeschreibungPos = 2000000

For i = LBound(gsarrAnsPriceA) To UBound(gsarrAnsPriceA)
  lPosStart = InStr(1, sTxt, gsarrAnsPriceA(i))
  If (lPosStart > 0) Then
    lPosStart = InStr(lPosStart + Len(gsarrAnsPriceA(i)), sTxt, gsarrAnsPriceStartA(i))
    If (lPosStart > 0) Then
      lPosStart = lPosStart + Len(gsarrAnsPriceStartA(i))
      lPosEnd = InStr(lPosStart, sTxt, gsarrAnsPriceEndA(i))
      If lPosEnd > lPosStart And lPosEnd < lBeschreibungPos Then
        sTmp = Mid(sTxt, lPosStart, lPosEnd - lPosStart)
        If sTmp Like "*[0-9]*" And Len(sTmp) < 100 Then
          
          sucheAktGebot = getBetragUndWaehrung(sTmp, sWe)
          
          If gsarrAnsPriceTypeA(i) > 0 Then ' der Preis ist für Sofort-Kauf markiert
            lPosStart = InStr(1, sTxt, gsAnsBuyOnly) ' Sofortkauf markieren
            If lPosStart > 0 And lPosStart < lBeschreibungPos Then bFlag = True
          End If

          Exit Function
        End If
      End If
    End If
  End If
Next i

sucheAktGebot = "0,00"
sWe = "?"
If gbTest Then
  DebugPrint "--- kein AktGebot gefunden -----------------------------------------------------" & vbCrLf & _
        sTxt & vbCrLf & Date2Str(MyNow) & "   --------------------------------------------------------------------------------"
End If

Exit Function

errhdl:
  sucheAktGebot = "?,??"
  sWe = "?"
End Function

Private Function sucheVersand(sTxt As String, sWe As String) As String

  Const iMax As String = 3

  Dim i As Integer
  Dim iShippingMode As Integer
  
  For i = 0 To 2
    iShippingMode = giShippingMode + i
    If iShippingMode > iMax Then iShippingMode = iShippingMode - iMax
    sucheVersand = sucheVersandWrapped(sTxt, sWe, iShippingMode)
   
    If (sucheVersand > "" And sucheVersand <> "?,??") Or (sWe > "" And sWe <> "?") Then Exit Function
  Next i

End Function

Private Function sucheVersandWrapped(sTxt As String, sWe As String, iShippingMode As Integer) As String
On Error GoTo errhdl
Dim lPosStart As Long
Dim lPosStart2 As Long
Dim lPosEnd As Long

Dim sAnsShipping As String
Dim sAnsShippingStart As String
Dim sAnsShippingEnd As String

If iShippingMode = 1 Then
  sAnsShipping = gsAnsShipping1
  sAnsShippingStart = gsAnsShippingStart1
  sAnsShippingEnd = gsAnsShippingEnd1
ElseIf iShippingMode = 2 Then
  sAnsShipping = gsAnsShipping2
  sAnsShippingStart = gsAnsShippingStart2
  sAnsShippingEnd = gsAnsShippingEnd2
Else
  sAnsShipping = gsAnsShipping3
  sAnsShippingStart = gsAnsShippingStart3
  sAnsShippingEnd = gsAnsShippingEnd3
End If

lPosStart = InStr(1, sTxt, sAnsShipping)
If lPosStart > 0 Then
  lPosStart = lPosStart + Len(sAnsShipping)
  lPosStart2 = InStr(lPosStart, sTxt, sAnsShippingStart)
  If lPosStart2 > lPosStart + 500 Then lPosStart = 0 Else lPosStart = lPosStart2
  If lPosStart > 0 Then
    lPosStart = lPosStart + Len(sAnsShippingStart)
    lPosEnd = InStr(lPosStart, sTxt, sAnsShippingEnd)
    If lPosEnd > 0 And lPosEnd - lPosStart < 100 Then
      sucheVersandWrapped = getBetragUndWaehrung(Mid(sTxt, lPosStart, lPosEnd - lPosStart), sWe)
      Exit Function
    End If
  End If
End If

errhdl:
  sucheVersandWrapped = "?,??"
  sWe = "?"
End Function

Private Function getBetragUndWaehrung(ByVal sTxt As String, ByRef sWe As String) As String
  On Error GoTo errhdl
  Dim i As Integer
  Dim sTmp As String

  sWe = ""
  sTxt = HtmlCleanup(sTxt)
  
  For i = 1 To Len(sTxt)
    If Mid(sTxt, i, 1) Like "[0-9,.]" Then
      sTmp = sTmp & Mid(sTxt, i, 1)
    ElseIf Asc(Mid(sTxt, i, 1)) >= 32 Then
      sWe = sWe & Mid(sTxt, i, 1)
    End If
  Next i
  sWe = Trim(sWe)
  
  Select Case sWe ' Währungzeichen aufbereiten
    Case "EUR": sWe = Chr(128) '"€"
    Case "&euro;": sWe = Chr(128) '"€"
    Case "GBP": sWe = Chr(163) '"£"
    Case "US $", "USD": sWe = "$"
    Case "AU $", "AUD": sWe = "AU$"
    Case "C $", "CAD": sWe = "C$"
  End Select
  
  getBetragUndWaehrung = sTmp

Done:
On Error GoTo 0
Exit Function

errhdl:
    getBetragUndWaehrung = "0,00" 'TODO
    sWe = "?"
    Resume Done
    
End Function

Private Function sucheMenge(sTxt As String) As Long
Dim lPos As Long
Dim lPosStart As Long
Dim sTmp As String

On Error GoTo errExit

sucheMenge = -1

lPos = InStr(1, sTxt, gsAnsQuant)
If lPos > 0 Then
    lPosStart = InStr(lPos, sTxt, gsAnsQuantStart) + Len(gsAnsQuantStart)
    lPos = InStr(lPosStart, sTxt, gsAnsQuantEnd)
    sTmp = Mid(sTxt, lPosStart, lPos - lPosStart)
    Do While Not IsNumeric(Left(sTmp, 1)) And Len(sTmp) > 0
      sTmp = Mid(sTmp, 2)
    Loop
    sucheMenge = GetNumericPart(sTmp)
End If

errExit:

End Function

Private Function sucheStandort(sTxt As String) As String
Dim lPos As Long
Dim lPosStart As Long

On Error GoTo errExit

lPos = InStr(1, sTxt, gsAnsLocation)
If lPos > 0 Then
    lPosStart = InStr(lPos, sTxt, gsAnsLocationStart) + Len(gsAnsLocationStart)
    lPos = InStr(lPosStart, sTxt, gsAnsLocationEnd)
    If lPos = 0 Then lPos = lPosStart + 100
    sucheStandort = HtmlCleanup(Left(Mid(sTxt, lPosStart, lPos - lPosStart), 100))
    lPos = InStr(1, sucheStandort, "<")
    If lPos > 0 Then sucheStandort = Left(sucheStandort, lPos - 1)
    sucheStandort = Replace(sucheStandort, vbTab, " ")
    sucheStandort = Replace(sucheStandort, vbCr, " ")
    sucheStandort = Replace(sucheStandort, vbLf, " ")
    sucheStandort = Trim(HtmlZeichenConvert(sucheStandort))
    sucheStandort = Replace(sucheStandort, ", ,", ",")
    sucheStandort = Replace(sucheStandort, ",,", ",")
    Do While Right(sucheStandort, 1) = vbCr Or Right(sucheStandort, 1) = vbLf Or Right(sucheStandort, 1) = vbTab Or Right(sucheStandort, 1) = " "
        sucheStandort = Left(sucheStandort, Len(sucheStandort) - 1)
    Loop
End If

errExit:

End Function

Private Function sucheMinGebot(sTxt As String, sWe As String) As String
On Error GoTo errhdl
Dim lPosStart As Long
Dim lPosEnd As Long

lPosStart = InStr(1, sTxt, gsAnsMinBid)
If lPosStart > 0 Then
  lPosStart = lPosStart + Len(gsAnsMinBid)
  lPosStart = InStr(lPosStart, sTxt, gsAnsMinBidStart)
  If lPosStart > 0 Then
    lPosStart = lPosStart + Len(gsAnsMinBidStart)
    lPosEnd = InStr(lPosStart, sTxt, gsAnsMinBidEnd)
    If lPosEnd > 0 Then
      sucheMinGebot = getBetragUndWaehrung(Mid(sTxt, lPosStart, lPosEnd - lPosStart), sWe)
      Exit Function
    End If
  End If
End If

errhdl:
  sucheMinGebot = "0,00"
  sWe = "?"
End Function

Public Function sucheBewertung(sTxt As String) As String
Dim lPosStart As Long
Dim lPosEnd As Long
Dim sTmp As String

On Error GoTo errhdl

lPosStart = InStr(1, sTxt, gsAnsAssessment, vbTextCompare)
lPosStart = InStr(lPosStart + 1, sTxt, gsAnsAssessmentAnzStart, vbTextCompare)
lPosEnd = InStr(lPosStart, sTxt, gsAnsAssessmentAnzEnd, vbTextCompare)
If lPosStart > 0 And lPosEnd > 0 And lPosStart < lPosEnd Then
  sTmp = HtmlCleanup(Mid$(sTxt, lPosStart + Len(gsAnsAssessmentAnzStart), lPosEnd - lPosStart - Len(gsAnsAssessmentAnzStart)))
Else
  sucheBewertung = "?/?"
  Exit Function
End If

If gsAnsAssessmentPercentStart > "" Then
  lPosStart = InStr(lPosEnd, sTxt, gsAnsAssessmentPercentStart, vbTextCompare)
  lPosEnd = InStr(lPosStart + 1, sTxt, gsAnsAssessmentPercentEnd, vbTextCompare)
  If lPosStart > 0 And lPosEnd > 0 And lPosStart < lPosEnd Then
    sTmp = DelSZ(sTmp) & "/" & HtmlCleanup(Mid$(sTxt, lPosStart + Len(gsAnsAssessmentPercentStart), lPosEnd - lPosStart - Len(gsAnsAssessmentPercentStart))) & "%"
  Else
    sucheBewertung = DelSZ(sTmp) & "/?"
    Exit Function
  End If
End If
sucheBewertung = DelSZ(sTmp)
Exit Function

errhdl:
Err.Clear
sucheBewertung = "?/?"

End Function

Private Sub LoadMyEbay(Optional bDontAsk As Boolean = False)
'MeinEbay laden
Dim sServer As String
Dim sKommando As String
Dim sBuffer As String
Dim lPos As Long
Dim bTimerEnabled As Boolean
Dim lRet As VbMsgBoxResult
Dim sArtikelTmp As String
Dim i As Integer
Dim lPosEnd As Long
Dim sTmp As String
Dim bLoginTried As Boolean
Dim sReadItems As String
Dim sNotiz As String


On Error GoTo errExit
Call PanelText(StatusBar1, 2, StatusBar1.Panels(2).Text)

If Not bDontAsk Then

    lRet = MsgBox(gsarrLangTxt(9) & vbCrLf & vbCrLf & _
                gsarrLangTxt(10) & vbCrLf & _
                gsarrLangTxt(11) & vbCrLf & _
                gsarrLangTxt(12) & vbCrLf _
                , vbYesNoCancel, gsarrLangTxt(38))
    
    
    If lRet = vbCancel Then
        Exit Sub
    End If
    
    If Not CheckInternetConnection Then
        Ask_Online
        If Not IsOnline Then
            Exit Sub
        End If
    End If

Else
    lRet = vbYes
End If 'bDontAsk

If gsEbayLocalPass = "" Then
    LogIn
    bLoginTried = True
End If

If gsEbayLocalPass = "" Then
    Exit Sub
End If

If Not bDontAsk Then

    Screen.MousePointer = vbHourglass
    
    bTimerEnabled = Timer1.Enabled
    Timer1.Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    'LogIn.Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    AutoStatus.Enabled = False
    
    For i = 0 To IIf(Artikel(0).Tag = "in", giMaxRow, 0)
           Artikel(i).Enabled = False
    Next i

End If 'dont ask

Call PanelText(StatusBar1, 2, gsarrLangTxt(88))
sServer = "https://" & gsScript1 & gsScriptCommand1
'Test
'gsEbayLocalPass = URLEncode(gsEbayLocalPass)
'gsUser = URLEncode(gsUser)
sBuffer = ""

sKommando = gsCmdWatchList
sKommando = Replace(sKommando, "[User]", gsUser)
sKommando = Replace(sKommando, "[lPass]", gsEbayLocalPass)
sBuffer = ShortPost(sServer & sKommando)

Call PanelText(StatusBar1, 2, "")
If Check_Wartung(sBuffer) Then GoTo errExit

If InStr(1, sBuffer, gsAnsLoginOk, vbTextCompare) _
 + InStr(1, sBuffer, gsAnsLoginOk2, vbTextCompare) = 0 Then
  gsEbayLocalPass = ""
  If Not bLoginTried Then
    Call LogIn
    sBuffer = ShortPost(sServer & sKommando)
  End If
End If

If gsCmdWatchList2 > "" And Not gbReadEndedItems Then
  sKommando = gsCmdWatchList2
  sKommando = Replace(sKommando, "[User]", gsUser)
  sKommando = Replace(sKommando, "[lPass]", gsEbayLocalPass)
  sBuffer = ShortPost(sServer & sKommando)
End If

If gsCmdBidList > "" Then
  sKommando = gsCmdBidList
  sKommando = Replace(sKommando, "[User]", gsUser)
  sKommando = Replace(sKommando, "[lPass]", gsEbayLocalPass)
  sBuffer = sBuffer & ShortPost(sServer & sKommando)
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InStr(1, sBuffer, gsAnsWatchList) = 0 Then
  gsWatchListType = IIf(gsWatchListType = "", "2", "")
  Call ReadAllKeywords
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sReadItems = "|"

If FindeBereich(sBuffer, gsAnsBidStart, "", gsAnsBidEnd, sTmp) Then
  lPos = 1
  lPos = FindeBereich(sTmp, gsAnsBidItemStart1, gsAnsBidItemPreEnd1, gsAnsBidItemEnd1, sArtikelTmp, lPos)
  Do While lPos
    lPosEnd = InStr(1, sArtikelTmp, gsAnsBidItemEnded)
    If lPosEnd = 0 Then
      If FindeBereich(sArtikelTmp, gsAnsBidItemStart2, ansBidItemPreEnd2, gsAnsBidItemEnd2, sArtikelTmp) Then
        Call DebugPrint("BidItem: " & sArtikelTmp & IIf(lPosEnd, " (beendet)", ""))
        Call PanelText(StatusBar1, 2, gsarrLangTxt(34) & " " & gsarrLangTxt(31) & " " & sArtikelTmp)
        Call AddArtikel(sArtikelTmp)
        sReadItems = sReadItems & sArtikelTmp & "|"
      End If
    End If
    lPos = FindeBereich(sTmp, gsAnsBidItemStart1, gsAnsBidItemPreEnd1, gsAnsBidItemEnd1, sArtikelTmp, lPos)
  Loop
End If

If FindeBereich(sBuffer, gsAnsWatchStart, "", gsAnsWatchEnd, sTmp) Then
  lPos = 1
  lPos = FindeBereich(sTmp, gsAnsWatchItemStart1, gsAnsWatchItemPreEnd1, gsAnsWatchItemEnd1, sArtikelTmp, lPos)
  Do While lPos
    lPosEnd = InStr(1, sArtikelTmp, gsAnsWatchItemEnded)
    If lPosEnd = 0 Then
      If FindeBereich(sArtikelTmp, gsAnsWatchItemStart2, gsAnsWatchItemPreEnd2, gsAnsWatchItemEnd2, sArtikelTmp) Then
        Call DebugPrint("WatchItem: " & sArtikelTmp & IIf(lPosEnd, " (beendet)", ""))
        Call PanelText(StatusBar1, 2, gsarrLangTxt(34) & " " & gsarrLangTxt(31) & " " & sArtikelTmp)
        Call AddArtikel(sArtikelTmp)
        sReadItems = sReadItems & sArtikelTmp & "|"
      End If
    End If
    lPos = FindeBereich(sTmp, gsAnsWatchItemStart1, gsAnsWatchItemPreEnd1, gsAnsWatchItemEnd1, sArtikelTmp, lPos)
  Loop
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


lPos = 1
lPos = FindeBereich(sBuffer, gsAnsNoteStart, "", gsAnsNoteEnd, sTmp, lPos)
Do While lPos
  If FindeBereich(sTmp, gsAnsNoteTextStart, "", gsAnsNoteTextEnd, sNotiz) Then
    If FindeBereich(sTmp, gsAnsNoteLineIDStart, "", gsAnsNoteLineIDEnd, sArtikelTmp) Then
      Call DebugPrint("Notiz zu Artikel " & sArtikelTmp & " : " & sNotiz)
      For i = 1 To giAktAnzArtikel
        If gtarrArtikelArray(i).Artikel = sArtikelTmp Then
          sNotiz = Replace(sNotiz, "<br>", " ", , , vbTextCompare)
          If InStr(1, gtarrArtikelArray(i).Kommentar, sNotiz, vbTextCompare) = 0 Then
            If gtarrArtikelArray(i).Kommentar > "" Then gtarrArtikelArray(i).Kommentar = gtarrArtikelArray(i).Kommentar & " | "
            gtarrArtikelArray(i).Kommentar = gtarrArtikelArray(i).Kommentar & sNotiz
            Exit For
          End If
        End If
      Next
    End If
  End If
  lPos = FindeBereich(sBuffer, gsAnsNoteStart, "", gsAnsNoteEnd, sTmp, lPos)
Loop

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If lRet = vbNo Then ' Jetzt die nicht eingelesenen entfernen

  frmProgress.InitProgress 0, giAktAnzArtikel
  DoEvents
  For i = giAktAnzArtikel To 1 Step -1
    If InStr(1, sReadItems, "|" & gtarrArtikelArray(i).Artikel & "|") = 0 Then RemoveArtikel i, False, False
    frmProgress.Step
  Next
  frmProgress.TerminateProgress
  Call RemoveArtikel(0, True, True)  ' Jetzt noch die Anzeige refreshen

End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

errExit:
On Error Resume Next
Call PanelText(StatusBar1, 2, "")

Call ArtikelArrayToScreen(VScroll1.Value)

If Not bDontAsk Then

Screen.MousePointer = vbNormal

Toolbar1.Buttons(4).Enabled = True
'LoadArtikel

For i = 0 To IIf(Artikel(0).Tag = "in", giMaxRow, 0)
       Artikel(i).Enabled = True
Next i

Screen.MousePointer = vbNormal
Toolbar1.Buttons(8).Enabled = True
Toolbar1.Buttons(4).Enabled = True
Toolbar1.Buttons(5).Enabled = True
'LogIn.Enabled = True
Toolbar1.Buttons(6).Enabled = True
AutoStatus.Enabled = True
Sortiere

Timer1.Enabled = (bTimerEnabled Or Timer1.Enabled) And gbAutoMode

If gbUsesModem And gbLastDialupWasManually Then Ask_Offline

End If 'dont ask

End Sub

Public Sub LogIn(Optional ByVal sLoginUser As String, Optional ByVal sLoginPass As String, Optional bLoginUseSecurityToken As Boolean = False)
'User- PasswdCheck und Login
Dim sServer As String
Dim sKommando As String
Dim sBuffer As String
Dim sBuffer2 As String
Dim lPos As Long
Dim bTimerEnabled As Boolean
Dim oHtmlForm As clsHtmlForm
Dim tmpgsAnsSummary As String

'wenn nicht explizit mit User + Pass aufgerufen, dann StandardUser nehmen (mae 050717)
If sLoginUser = "" Or sLoginPass = "" Then
    sLoginUser = gsUser
    sLoginPass = gsPass
    bLoginUseSecurityToken = gbUseSecurityToken
End If

If sLoginUser = "" Or gsPass = "" Then
    MsgBox gsarrLangTxt(2)
    Exit Sub
End If


If Not CheckInternetConnection Then
    Call Ask_Online
    If Not IsOnline Then Exit Sub
End If

Screen.MousePointer = vbHourglass

bTimerEnabled = Timer1.Enabled
Timer1.Enabled = False
mbIsLoggingIn = True

'If Check_ebayUp Then
    
    With Toolbar1
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
        .Buttons(6).Enabled = False
        .Buttons(7).Enabled = False
        .Buttons(8).Enabled = False
    End With
    On Error GoTo errExit
    AutoStatus.Enabled = False
    
    
    'DebugPrint "Login mit gsUser: " & sLoginUser, "Pass: " & sLoginPass
    
    If gbUseIECookies Then ' wenn wir die IE-Kekse nehmen, ist evtl. jemand anderes eingeloggt, also zunächst ein Logout
    
        Call PanelText(StatusBar1, 2, "Logout")
        
        sServer = "https://" & gsScript5 & gsScriptCommand5
        
        'Logout
        sKommando = gsCmdLogOff
        sBuffer = ShortPost(sServer & sKommando, , , sLoginUser)
        
    Else ' wir nehmen unsere eigenen Kekse, die werden pro User verwaltet, also ist niemand anderes eingeloggt, wir prüfen, ob wir schon eingeloggt sind
        
        Call PanelText(StatusBar1, 2, "Check Login")
        
        sServer = "https://" & gsMainUrl
        
        'wir testen mal den Login-Status
        sKommando = gsCmdSummary
        sBuffer = ShortPost(sServer & sKommando, , , sLoginUser)
              
    End If
    
    tmpgsAnsSummary = Replace(gsAnsSummary, "[User]", sLoginUser)
    
    lPos = InStr(1, sBuffer, tmpgsAnsSummary, vbTextCompare)
'    If lPos = 0 Then 'ausgelogt, jetzt einloggen
'
'        Call PanelText(StatusBar1, 2, "Login " & sLoginUser)
'
'        If bLoginUseSecurityToken Then
'            sLoginPass = sLoginPass & GetSecurityToken(sLoginUser)
'        End If
'
'        sServer = "http://" & gsScript5 & gsScriptCommand5
'
'        sKommando = gsCmdLogIn
'        sBuffer = ShortPost(sServer & sKommando, , , sLoginUser)
'
'        gsEbayLocalPass = ""
'        Call SetToolbarImage(Toolbar1.Buttons(7), 12)  'ausgeloggt
'        Toolbar1.Buttons(7).ToolTipText = gsarrLangTxt(70) & ": " & gsarrLangTxt(76)
'
'        Set oHtmlForm = New clsHtmlForm
'
'        With oHtmlForm
'            .Clear
'            Call .ReadForm(sBuffer, gsAnsLoginFrm)     'das Html-Formular einlesen
'
'            If .FormFound Then
'                If .GetFieldType(gsAnsPassField) = "" Then 'es ist die Logout-Bestätigungs-Seite => nochmal aufrufen
'                    sBuffer = ShortPost(sServer & sKommando, , , sLoginUser)
'                    .Clear
'                    Call .ReadForm(sBuffer, gsAnsLoginFrm)     'das Html-Formular einlesen
'                End If
'            End If
'
'            If .FormFound Then
'
'                Call .PutField(gsAnsPassField, sLoginPass)
'                If .GetFieldType(gsAnsUserField) = "text" Then Call .PutField(gsAnsUserField, sLoginUser)
'                If .GetFieldType(gsAnsUserField2) = "text" Then Call .PutField(gsAnsUserField2, sLoginUser)
'                If .GetFieldType(gsAnsUserField3) = "text" Then Call .PutField(gsAnsUserField3, sLoginUser)
'                If .GetFieldType("runid") <> "" Then Call .PutField("runid", "")
'                sServer = .GetAction()            'wo müssen wir das Zeug hinschicken?
'                If Not sServer Like "http*" Then sServer = "http" & IIf(gbUseSSL, "s", "") & "://" & gsScript5 & gsScriptCommand5 & sServer
'
'                If gbUseSSL Then
'                    If Not sServer Like "https*" Then .PutField gsAnsPassField, "" 'Passwort nie unverschlüsselt senden!
'                End If
'                sKommando = .GetFields(gsSiteEncoding)  'und welche Daten?
'                sKommando = sKommando & .ClickImage(gsAnsLoginSubmitImage1)
'                sBuffer = ShortPost(sServer, sKommando, , sLoginUser)
'
'            Else ' keine SignInForm gefunden => old style login
'
'                sServer = "https://" & gsScript5 & gsScriptCommand5
'                sKommando = gsCmdLogIn2
'                sKommando = Replace(sKommando, "[User]", URLEncode(sLoginUser))
'                sKommando = Replace(sKommando, "[Pass]", URLEncode(sLoginPass))
'                sBuffer = ShortPost(sServer, sKommando, , sLoginUser)
'
'            End If
'
'            'jetzt eingeloggt?
'            lPos = InStr(1, sBuffer, gsAnsLoginOk, vbTextCompare) _
'                 + InStr(1, sBuffer, gsAnsLoginOk2, vbTextCompare)
'            If lPos = 0 Then
'
'                .Clear
'                Call .ReadForm(sBuffer, gsAnsLoginFrm)     'das Html-Formular einlesen
'
'                If .FormFound Then
'                    If .GetFieldType(gsAnsTokenField) = "text" Then
'
'                        Call .PutField(gsAnsTokenField, GetSecurityToken(sLoginUser))
'
'                        sServer = .GetAction()            'wo müssen wir das Zeug hinschicken?
'                        sKommando = .GetFields(gsSiteEncoding) 'und welche Daten?
'                        If sServer Like "https*" Or Not gbUseSSL Then sBuffer = ShortPost(sServer, sKommando, , sLoginUser)
'
'                        Call SetUseTokenByAccount(sLoginUser)
'
'                    End If
'                End If
'
'                'jetzt eingeloggt?
'                lPos = InStr(1, sBuffer, gsAnsLoginOk, vbTextCompare) _
'                     + InStr(1, sBuffer, gsAnsLoginOk2, vbTextCompare)
'
'            End If 'lPos = 0
'        End With 'oHtmlForm
'    End If
    
'    If lPos = 0 Then ' entweder sind wir noch nicht drin oder eBay zeigt wieder irgendeinen
'                     ' Werbe- oder Passwortfragen-Krams an, wir testen einfach nochmal:
'        sServer = "http://" & gsScript1 & gsScriptCommand1
'
'        'wir testen mal den Login-Status
'        sKommando = gsCmdWatchList
'        sBuffer2 = ShortPost(sServer & sKommando, , , sLoginUser)
'
'        lPos = InStr(1, sBuffer2, gsAnsLoginOk, vbTextCompare) _
'             + InStr(1, sBuffer2, gsAnsLoginOk2, vbTextCompare)
'    End If
    
    If lPos = 0 Then
        If Check_Wartung(sBuffer) Then GoTo errExit
        
        'tja, war nix
        'Pass_Check.Caption = "Status: Fehler bei User/ Passwortprüfung, kein Login"
        'Pass_Check.BackColor = vbRed
        'und Fehlermaske aufpoppen ;-)
        'erstmal wegsichern ;-)
        Call SetToolbarImage(Toolbar1.Buttons(7), 12)
        Toolbar1.Buttons(7).ToolTipText = gsarrLangTxt(70) & ": " & gsarrLangTxt(77)
        Call PanelText(StatusBar1, 2, gsarrLangTxt(70) & ": " & gsarrLangTxt(77), False, vbRed)
        
        Call SaveToFile(sBuffer, msScratch)
        On Error Resume Next
        'Unload frmBrowser
        gsGlobalUrl = msScratch
        'zwangsweise nach vorne holen ;-)
        'Call ShowBrowser(Me.hWnd)
        'frmBrowser.Show vbModal
        Call ExecuteDoc(Me.hWnd, gsGlobalUrl)
        
    Else
            
        gsEbayLocalPass = "default"
        Call PanelText(StatusBar1, 2, gsarrLangTxt(70) & ": " & gsarrLangTxt(89), True, vbGreen)
        
        Call SetToolbarImage(Toolbar1.Buttons(7), 16)
        Toolbar1.Buttons(7).ToolTipText = gsarrLangTxt(70) & ": " & sLoginUser & " " & gsarrLangTxt(78) 'Status: [LoginUser] angemeldet
        
    End If
    
'End If 'Check_ebayUp

errExit:

Screen.MousePointer = vbNormal
Set oHtmlForm = Nothing
Timer1.Enabled = (bTimerEnabled Or Timer1.Enabled) And gbAutoMode
mbIsLoggingIn = False

Toolbar1.Buttons(4).Enabled = True
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(6).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Toolbar1.Buttons(8).Enabled = True
AutoStatus.Enabled = True

If gbUsesModem And gbLastDialupWasManually Then Call Ask_Offline

'mae: auskommentiert, da es bei fehlerhaftem ReLogin den Automodus ausschaltet
'If gsEbayLocalPass = "" Then
'     gbAutoMode = False
'     CheckAutoMode
'End If

End Sub

Function LogIn2(ByVal sTxt As String) As String

  Dim sLoginUser As String
  Dim sLoginPass As String
  Dim bLoginUseSecurityToken As Boolean
  Dim sServer As String
  Dim sKommando As String
  Dim sSaveStr As String
  Dim oHtmlFrm As clsHtmlForm
  
  sSaveStr = sTxt
  
  sLoginUser = gsUser
  sLoginPass = gsPass
  bLoginUseSecurityToken = gbUseSecurityToken

  sServer = "https://" & gsScript5 & gsScriptCommand5
  sKommando = gsCmdLogIn

' DebugPrint "Login mit User: " & sLoginUser, "Pass: " & sLoginPass

  Set oHtmlFrm = New clsHtmlForm

  oHtmlFrm.Clear
  oHtmlFrm.ReadForm sTxt, gsAnsLoginFrm    'das Html-Formular einlesen
  
  If oHtmlFrm.FormFound Then
    If oHtmlFrm.GetFieldType(gsAnsPassField) = "" Then 'es ist die Logout-Bestätigungs-Seite => nochmal aufrufen
      sTxt = ShortPost(sServer & sKommando, , , sLoginUser)
      oHtmlFrm.Clear
      oHtmlFrm.ReadForm sTxt, gsAnsLoginFrm    'das Html-Formular einlesen
    End If
  End If

  If oHtmlFrm.FormFound Then
  
    If bLoginUseSecurityToken Then
      sLoginPass = sLoginPass & GetSecurityToken(sLoginUser)
    End If

    oHtmlFrm.PutField gsAnsPassField, sLoginPass
    If oHtmlFrm.GetFieldType(gsAnsUserField) = "text" Then oHtmlFrm.PutField gsAnsUserField, sLoginUser
    If oHtmlFrm.GetFieldType(gsAnsUserField2) = "text" Then oHtmlFrm.PutField gsAnsUserField2, sLoginUser
    If oHtmlFrm.GetFieldType(gsAnsUserField3) = "text" Then oHtmlFrm.PutField gsAnsUserField3, sLoginUser
    sServer = oHtmlFrm.GetAction()            'wo müssen wir das Zeug hinschicken?
    If Not sServer Like "https*" Then oHtmlFrm.PutField gsAnsPassField, "" 'Passwort nie unverschlüsselt senden!
    sKommando = oHtmlFrm.GetFields(gsSiteEncoding) 'und welche Daten?
    sTxt = ShortPost(sServer, sKommando, , sLoginUser)
  
    oHtmlFrm.Clear
    oHtmlFrm.ReadForm sTxt, gsAnsLoginFrm    'das Html-Formular einlesen
      
    If oHtmlFrm.FormFound Then
      If oHtmlFrm.GetFieldType(gsAnsTokenField) = "text" Then
      
        oHtmlFrm.PutField gsAnsTokenField, GetSecurityToken(sLoginUser)
        
        sServer = oHtmlFrm.GetAction()            'wo müssen wir das Zeug hinschicken?
        sKommando = oHtmlFrm.GetFields(gsSiteEncoding) 'und welche Daten?
        If sServer Like "https*" Then sTxt = ShortPost(sServer, sKommando, , sLoginUser)
        
        SetUseTokenByAccount sLoginUser
      
      End If
    End If
  
  Else
    sTxt = sSaveStr
  End If
  
  LogIn2 = sTxt

End Function

Private Sub TimeSync() 'umbenannt, vorher EbayTimeSync_Click(), lg 29.05.03
'Ebay TimeSync
Dim bTimerEnabled As Boolean
Dim bCheckEnabled As Boolean
Dim sRechnerzeit As String
Dim bOk As Boolean
Dim sNetTime As String

On Error GoTo errhdl

If Not CheckInternetConnection Then
    Ask_Online
    If Not IsOnline Then
        Exit Sub
    End If
End If

Screen.MousePointer = vbHourglass

bTimerEnabled = Timer1.Enabled
Timer1.Enabled = False
bCheckEnabled = Toolbar1.Buttons(8).Enabled
Toolbar1.Buttons(8).Enabled = False
Toolbar1.Buttons(4).Enabled = False
Toolbar1.Buttons(5).Enabled = False
'LogIn.Enabled = False
Toolbar1.Buttons(6).Enabled = False
AutoStatus.Enabled = False

sRechnerzeit = Date2Str(MyNow)
sNetTime = Zeitsync()
bOk = sNetTime <> ""

If Me.WindowState <> vbMinimized And Not gbAutoMode Then
  'nicht bei Minimized und auch nicht im Automode, gibt Probleme
  If bOk Then
    If Not gbKeinHinweisNachZeitsync Then 'lg 29.05.03
      If FormLoaded("frmAbout") Then frmAbout.Hide
      MsgBox gsarrLangTxt(14) & vbCrLf & vbCrLf & gsarrLangTxt(15) & vbTab & sRechnerzeit & vbCrLf & gsarrLangTxt(16) & vbTab & Date2Str(MyNow) & IIf(gfTimeDeviation > 0, " ( + " & Format(gfTimeDeviation, "##0.0") & " s )", ""), vbInformation
    End If
  Else
    If FormLoaded("frmAbout") Then frmAbout.Hide
    MsgBox gsarrLangTxt(17), vbCritical
  End If
End If

errhdl:

Timer1.Enabled = (bTimerEnabled Or Timer1.Enabled) And gbAutoMode
Toolbar1.Buttons(8).Enabled = bCheckEnabled
Toolbar1.Buttons(4).Enabled = True
Toolbar1.Buttons(5).Enabled = True
'LogIn.Enabled = True
Toolbar1.Buttons(6).Enabled = True
AutoStatus.Enabled = bCheckEnabled

Screen.MousePointer = vbNormal

If gbUsesModem And gbLastDialupWasManually Then Ask_Offline

End Sub

Private Sub Einstellungen()

On Error GoTo errhdl
Load frmSettings
frmSettings.Show vbModal, Me
Unload frmSettings
Call CheckSofortkaufArtikel
Call ArtikelArrayToScreen(VScroll1.Value)
Exit Sub
errhdl:
MsgBox "Err at EinstellungenClick: " & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim sTestFile As String
Dim sTestFileOld As String

ReDim Preserve gtarrArtikelArray(Artikel.UBound) As udtArtikelZeile
ReDim gtarrRemovedArtikelArray(0 To 0) As udtArtikelZeile

miMouseIndex = -1
miTmpIndex = -1
miShowIndex = -1
Call SetFocusRect(-3)
giLastGebotEditedIndex = -1

gfLzMittel = 2
giDebugLevel = 1 'Vorerst auf 1, anderer Wert kommt aus der INI

mbIsOn = False
gbShowSplash = True

Call SetIcon(Me.hWnd, MyLoadResPicture(201, 16))

gsTempPfad = GetLongTempPath()
If gsTempPfad = "" Then gsTempPfad = "C:"
msScratch = gsTempPfad & "\scratch.html"

gbTest = True 'Testdrucke

On Error Resume Next

Call GetWinVersion

'und Testdrucke ..
If gbTest Then
    sTestFileOld = gsAppDataPath & "\testfile.tst"
    sTestFile = gsAppDataPath & "\History.log"
    
    If Dir(sTestFile) = "" And Dir(sTestFileOld) > "" Then
      Name sTestFileOld As sTestFile
    End If
    
    Err.Clear
    DebugPrint "Programmstart BOM " & GetBOMVersion()
End If

miStartWidth = Me.Width


Call InitSemas
Call InitRingBuffs
Call InitCurrencies

Call ReadAllSettings
Call ReadLocalServStrings
Call ReadAllKeywords

Call SelectLanguage(gsAktLanguage)

If IsJobCommand() Then
    Call DoTheJob(Mid(Command(), InStr(1, Command(), "/JOB:", vbTextCompare) + 5, InStr(InStr(1, Command(), "/JOB:", vbTextCompare) + 5, Command(), ".JOB", vbTextCompare) - InStr(1, Command(), "/JOB:", vbTextCompare) - 1))
    mbQuietExit = True
    Unload Me
    End 'MD-Marker
End If

If App.PrevInstance Then
    If vbNo = MsgBox(gsarrLangTxt(1), vbYesNo) Then
        gbExplicitEnd = True
        End 'MD-Marker
    End If
End If

'Versionsprüfung Ini-Files
Call CheckIniVersion

'1.8.0 Startpasswort
If gbPassAtStart Then
    frmStartPass.Show vbModal
End If

Call ShowSplashOnce 'frmAbout wenigstens einmalig anzeigen

VScroll1.LargeChange = giMaxRowSetting + 1 'seitenweise scrollen

gbAAnext = IIf(gbAutoAktualisierenNext And gbAutoAktualisieren, True, False) 'sh 25.10.03
mbAktPaused = False

Call SetInitLineSize

Artikel(0).Tag = "out"
Call SwapZeilen("in")

Call InitToolbar

Call SetLanguage
Call SetSupportedLanguages
Toolbar1.Enabled = gbShowToolbar
Toolbar1.Visible = gbShowToolbar

Zusatzfeld.Visible = gbOperaField

Call SetFormSize
Call SetFont(Me)
'Wir merken uns alle Objekte in der Originalgrösse ..
On Error Resume Next

For i = 0 To Me.Controls.Count - 1
    With Me.Controls(i)
        If Not TypeOf Me.Controls(i).Container Is ToolBar Then
            Call moResize.AddControl(Me.Controls(i), rszStickBottom, Me)
        End If
    End With
 Next i

miStartWidth = Me.Width
miStartHeight = Me.Height

' 1.8.0 Wheelmouse
' vorsicht bei Debug!
If Not InDevelopment Then
    If gbUseWheel Then
        MWheel1.hWndCapture = Me.hWnd
        MWheel1.EnableWheel
        gbWheelUsed = True
    End If
    Call modSubclass.Subclass(frmDummy.hWnd)
End If

Set goCookieHandler = New clsCookieHandler

mbEBayTimeIsSync = CBool(giUseTimeSync <> 1) 'nur wenn nicht auf "einmal" steht

With Me
    Call .Move(glPosLeft, glPosTop, glPosWidth, glPosHeight)
    .WindowState = giStartupSize
End With

Call ReadArtikelIni(False)
miMaxSaveCount = 10 'Minuten AutoSave

If gbUsesOdbc Then
    miOdbcTimerCount = giOdbcZyklus + 1 'sofort loslaufen
    ODBC_Timer.Enabled = True
    
    With Zusatzfeld
        .Visible = True
        .Enabled = False
        .BackColor = &H8000000F
        .Text = "ODBC " & Replace(gsarrLangTxt(103), "%MIN%", giOdbcZyklus) 'lg 27.05.03
    End With
End If

i = ReadSetting("Fenster", "ScrollValue", 1)
If i > VScroll1.Max Then i = VScroll1.Max
VScroll1.Value = i
Call ArtikelArrayToScreen(VScroll1.Value)

Call AddTrayIcon
Call updTaskbar(gsarrLangTxt(47) & ": " & gsUser & gsarrLangTxt(256))

AutoStatus.Caption = gsarrLangTxt(62)
AutoStatus.BackColor = vbRed

mbInitDone = True

Me.OLEDropMode = vbOLEDropManual

TimeoutTimer.Interval = 1000 ' 1 sec

Call Form_Resize
Me.Show
If gbNewItemWindowOpenOnStartup Then
  frmNeuerArtikel.Show
  DoEvents
  Me.SetFocus
End If
mbFormLoaded = True
If gbAboutIsUp Then frmAbout.SetFocus

If Not mbCheckDone And Not gbUsesModem Then
    mbCheckDone = True
    If gbAutoUpdateCurrencies Then
        Call UpdateCurrencies
        Call ArtikelArrayToScreen(VScroll1.Value)
        DoEvents
    End If
    If gbCheckForUpdate Then Call CheckUpdate(Me, True)
End If

If gbAutoLogin Then
    If gsUser > "" And gsPass > "" Then
        Call LogIn
    Else
        If Not gbAutoStart Then ' bei Autostart gibts die Fehlermeldung in CheckAutoMode
            If FormLoaded("frmAbout") Then frmAbout.Hide
            MsgBox gsarrLangTxt(2)
        End If
    End If
End If

If gbAutoAktualisieren And giAktAnzArtikel > 0 Then 'lg 29.05.03
    Call LoadArtikel
End If

If gbAutoStart Then gbAutoMode = True ''lg 06.08.03

If (giUseTimeSync And 4) > 0 Then 'Programmstart, lg 29.05.03
    Call TimeSync
End If

If gbUsePop Then mbWaitForFirstPop = True

Call CheckAutoMode ' wir wollen die Timer erst nach dem TimeSync loslaufen lassen aber vorher schon den Automode setzen!

Call CheckIEStatus

RechnerZeitTimer.Enabled = True

If Command() > "" Then ParseCommand Command()

If Not gbAutoMode Then mbWaitForFirstPop = False

End Sub

Private Function Bieten(ByVal sItem As String, _
                        ByVal sMaxGebot As String, _
                        ByVal sEBayUser As String, _
                        ByVal sEBayPass As String, _
                        ByVal iAktRow As Integer, _
                        Optional bSofortBieten As Boolean = False, _
                        Optional bUseToken As Boolean = False, _
                        Optional bIsBuyItNow As Boolean = False) As Boolean
                        
Dim sServer  As String
Dim sKommando As String
Dim sKommandoPublic As String
Dim sBuffer As String
Dim lPos As Long
Dim sSaveFile As String
Dim sMaxBid As String
Dim sReferer As String
Dim sTmp As String
Dim oHtmlForm As clsHtmlForm
Dim lAufruf As Long
Dim lVersuch As Long
Dim lDurchgang As Long
Dim sUserAccount As String
Dim datEndeZeit As Date

sUserAccount = gtarrArtikelArray(iAktRow).UserAccount
datEndeZeit = gtarrArtikelArray(iAktRow).EndeZeit

Bieten = False
miErrStatus = [asErr]

On Error Resume Next

'die alten Savefiles löschen: lg 08.06.03
sSaveFile = Dir(gsTempPfad & "\Art-" & sItem & "-*.html")
Do While sSaveFile <> ""
  Kill gsTempPfad & "\" & sSaveFile
  sSaveFile = Dir()
Loop

lAufruf = 0
lVersuch = 0
lDurchgang = 0

Call DebugPrint("Biete " & sMaxGebot & " auf Artikel " & sItem)

' User und Passwort bestimmen (Multi-Account)
If giUserAnzahl > 0 Then
  If sUserAccount <> "" Then
    'Bietaccount gesetzt
    If sUserAccount <> gsUser Then
      'Bietaccount nicht Standarduser
      sEBayUser = sUserAccount
      sEBayPass = gtarrUserArray(UsrAccToIndex(sUserAccount)).UaPass
      bUseToken = gtarrUserArray(UsrAccToIndex(sUserAccount)).UaToken
    Else
      'Bietaccount = Explicit Standarduser
      sEBayUser = gsUser
      sEBayPass = gsPass
      bUseToken = gbUseSecurityToken
    End If
  Else
    'Bietaccount nicht gesetzt -> Standarduser
    If giDefaultUser <> 0 Then
      sEBayUser = gtarrUserArray(giDefaultUser).UaUser
      sEBayPass = gtarrUserArray(giDefaultUser).UaPass
      bUseToken = gtarrUserArray(giDefaultUser).UaToken
    ElseIf sEBayUser = "" Or sEBayPass = "" Then ' Evtl. explizit übergeben von SpeedTest?
      Bieten = False
      Exit Function
    End If
  End If
Else
  If sEBayUser = "" Or sEBayPass = "" Then
    'Kein Useraccount definiert
    Bieten = False
    Exit Function
  Else
    'User und Passwort übergeben, z.B. per Kommandozeile
    'weiterlaufen lassen
  End If
End If

' 1.0.5 ODBC
If sEBayUser = "" Or sEBayPass = "" Then
   sEBayUser = gsUser
   sEBayPass = gsPass
   bUseToken = gbUseSecurityToken
End If

Call DebugPrint("User: " & sEBayUser)

Call PanelText(StatusBar1, 2, gsarrLangTxt(90) & sItem)

NochmalVonVorne:
lDurchgang = lDurchgang + 1

'Dezimalstelle des Gebots anpassen
Select Case gsCmdDecSeparator
    Case "."
        lPos = InStr(1, sMaxGebot, ",")
        If lPos > 0 Then Mid(sMaxGebot, lPos, 1) = "."
    Case ","
        lPos = InStr(1, sMaxGebot, ".")
        If lPos > 0 Then Mid(sMaxGebot, lPos, 1) = ","
End Select

sMaxBid = URLEncode(sMaxGebot)

'Kommando zerlegen
sBuffer = ""
If bIsBuyItNow Then
  sTmp = gsCmdBuyItNow
Else
  sTmp = gsCmdMakeBid
End If

sTmp = Replace(sTmp, "[Item]", sItem)
sTmp = Replace(sTmp, "[MaxBid]", sMaxBid)

sKommando = sBuffer & sTmp

sServer = "https://" & gsScript4 & gsScriptCommand4
sReferer = sServer & "ViewItem&item=" & sItem

sServer = "https://" & gsScript3 & gsScriptCommand3

Call DebugPrint("Sende Anfrage zum Server: " & sServer & sKommando)

sBuffer = ShortPost(sServer & sKommando, , sReferer, sEBayUser)

'File für Errormeldungen wegretten
lAufruf = lAufruf + 1
sSaveFile = gsTempPfad & "\Art-" & sItem & "-1.html"
Call SaveToFile(StripJavaScript(sBuffer), sSaveFile)
sSaveFile = gsTempPfad & "\Art-" & sItem & "-" & CStr(lAufruf) & "-1.html"
Call SaveToFile(StripJavaScript(sBuffer), sSaveFile)
If (lDurchgang = 1) Then
  sSaveFile = gsTempPfad & "\Art-" & sItem & "-status.html"
  Call SaveToFile(StripJavaScript(sBuffer), sSaveFile)
End If

Set oHtmlForm = New clsHtmlForm

'Jetzt die Bietschleife
Do While lVersuch < mlMAXBIETVERSUCHE
  lVersuch = lVersuch + 1
  DoEvents

  If InStr(1, sBuffer, gsAnsBidAccepted) + _
     InStr(1, sBuffer, gsAnsBidAccepted2) + _
     InStr(1, sBuffer, gsAnsBidAccepted3) + _
     InStr(1, sBuffer, gsAnsBidAccepted4) > 0 Then  ' Wir haben das Gebot erfolgreich platziert
     
    If FindeBereich(sBuffer, gsAnsTimeLeft, gsAnsTimeLeftStart, gsAnsTimeLeftEnd, sTmp) > 0 Then
      sTmp = sTmp & gsAnsTimeLeftEnd
      FindeBereich sTmp, gsAnsTimeLeftStart, "", gsAnsTimeLeftEnd, sTmp
      sTmp = Replace(sTmp, vbCr, "")
      sTmp = Replace(sTmp, vbLf, "")
      sTmp = Replace(sTmp, vbTab, "")
      sTmp = Trim(sTmp)
    Else
      sTmp = "unbekannt"
    End If
    Call DebugPrint("Gebot angenommen, Restzeit: " & sTmp)
    miErrStatus = [asOK]
    Bieten = True
    Call PanelText(StatusBar1, 2, "")
    Exit Function
  End If

  If InStr(1, sBuffer, gsAnsBidOutBid) > 0 Then ' Wir wurden überboten
    Call DebugPrint("Überboten")
    miErrStatus = [asUeberboten]
    Bieten = False
    Call PanelText(StatusBar1, 2, "")
    Exit Function
  End If

  If InStr(1, sBuffer, gsAnsBidReserveNotMet) > 0 Then ' Mindestpreis nicht erreicht
    Call DebugPrint("Mindestpreis nicht erreicht")
    miErrStatus = [asUeberboten]
    Bieten = False
    Call PanelText(StatusBar1, 2, "")
    Exit Function
  End If

  If InStr(1, sBuffer, gsAnsBidErrMinBid) > 0 Then ' Unser Gebot ist zu niedrig
    Call DebugPrint("Gebot zu niedrig")
    miErrStatus = [asUeberboten]
    Bieten = False
    Call PanelText(StatusBar1, 2, "")
    Exit Function
  End If
  
  If InStr(1, sBuffer, gsAnsBidErrEnded) + InStr(1, sBuffer, gsAnsBidErrEnded2) > 0 Then ' Der Artikel ist abgelaufen
    Call DebugPrint("Artikel beendet")
    miErrStatus = [asErr]
    Bieten = False
    Call PanelText(StatusBar1, 2, "")
    Exit Function
  End If

  If InStr(1, sBuffer, gsAnsBidErrNotAvail) > 0 Then  ' Der Artikel ist nicht mehr verfügbar
    Call DebugPrint("Artikel nicht mehr verfügbar")
    miErrStatus = [asErr]
    Bieten = False
    Call PanelText(StatusBar1, 2, "")
    Exit Function
  End If

  If InStr(1, sBuffer, gsAnsSignInError) > 0 Then ' User/Pass falsch
    Call DebugPrint("Anmeldung fehlerhaft")
    miErrStatus = [asErr]
    Bieten = False
    Call PanelText(StatusBar1, 2, "")
    Exit Function
  End If

  If InStr(1, sBuffer, gsAnsBidErrGeneral) > 0 Then ' Sonstiger Fehler
    Call DebugPrint("Sonstiger Fehler")
    miErrStatus = [asErr]
    Bieten = False
    Call PanelText(StatusBar1, 2, "")
    Exit Function
  End If

  oHtmlForm.Clear
  oHtmlForm.ReadForm sBuffer, IIf(bIsBuyItNow, gsAnsBuyForm, gsAnsBidForm)       'das Html-Formular einlesen
  If Not oHtmlForm.FormFound Then oHtmlForm.ReadForm sBuffer, gsAnsLoginFrm   'evtl. Login-Formular einlesen
  If Not oHtmlForm.FormFound Then oHtmlForm.ReadForm sBuffer, IIf(bIsBuyItNow, gsAnsBuyForm2, gsAnsBidForm2)   'nächster Versuch (egun) / Belehrungsseite ('bay)
  If Not oHtmlForm.FormFound Then
    'zuerst ein LogOff
    sServer = "https://" & gsScript5 & gsScriptCommand5
    sKommando = gsCmdLogOff
    sBuffer = ShortPost(sServer & sKommando, , , sEBayUser)
    Call DebugPrint("Neustart")
    GoTo NochmalVonVorne                 'hier geht was völlig schief, wir probieren es noch mal von vorne
  End If
  
  If oHtmlForm.GetFieldType(gsAnsPassField) = "password" Then ' Wenn es eine Passwort-Eingabe gibt
    
    If bUseToken Then
      sEBayPass = sEBayPass & GetSecurityToken(sEBayUser)
    End If
    
    oHtmlForm.PutField gsAnsPassField, sEBayPass
    If oHtmlForm.GetFieldType(gsAnsUserField) = "text" Then oHtmlForm.PutField gsAnsUserField, sEBayUser
    If oHtmlForm.GetFieldType(gsAnsUserField2) = "text" Then oHtmlForm.PutField gsAnsUserField2, sEBayUser
    If oHtmlForm.GetFieldType(gsAnsUserField3) = "text" Then oHtmlForm.PutField gsAnsUserField3, sEBayUser
    If oHtmlForm.GetFieldType("runid") <> "" Then oHtmlForm.PutField "runid", ""
    
  ElseIf oHtmlForm.GetFieldType(gsAnsTokenField) = "text" Then
        
    oHtmlForm.PutField gsAnsTokenField, GetSecurityToken(sEBayUser)
    Call SetUseTokenByAccount(sEBayUser)
          
  End If
    
  'Wenn User ausgefüllt ist und nicht unser User ist, dann Logout
  If (oHtmlForm.GetField(gsAnsUserField) <> "" And LCase(oHtmlForm.GetField(gsAnsUserField)) <> LCase(sEBayUser)) Or _
   (oHtmlForm.GetField(gsAnsUserField2) <> "" And LCase(oHtmlForm.GetField(gsAnsUserField2)) <> LCase(sEBayUser)) Or _
   (oHtmlForm.GetField(gsAnsUserField3) <> "" And LCase(oHtmlForm.GetField(gsAnsUserField3)) <> LCase(sEBayUser)) Then
     
    Call DebugPrint("Achtung, falscher User(" & oHtmlForm.GetField(gsAnsUserField) & oHtmlForm.GetField(gsAnsUserField2) & oHtmlForm.GetField(gsAnsUserField3) & ") - Logout")
    
    'Probieren den "Nicht Ihr Mitgliedsname"-Link zu ermitteln
    sServer = GetLinkNamedLike(sBuffer, gsAnsLinkChangeUser)
    If sServer <> "" Then
      If InStr(1, sServer, "?") > 0 Then
        sKommando = Mid(sServer, InStr(1, sServer, "?") + 1)
        sServer = Left(sServer, InStr(1, sServer, "?") - 1)
      Else
        sKommando = ""
      End If
      
      Call DebugPrint("Change User")
      GoTo PostIt ' Wir melden uns direkt ab
    End If
    
    'Plain old LogOff
    Call DebugPrint("Normal Logout")
    sServer = "https://" & gsScript5 & gsScriptCommand5
    sKommando = gsCmdLogOff
    sBuffer = ShortPost(sServer & sKommando, , , sEBayUser)
    GoTo NochmalVonVorne
    
  End If
  
  If oHtmlForm.GetFieldType(gsAnsBidField) = "text" Then oHtmlForm.PutField gsAnsBidField, sMaxGebot
  
  sServer = oHtmlForm.GetAction()             'wo müssen wir das Zeug hinschicken?
  If Not sServer Like "http*" Then sServer = "http" & IIf(gbUseSSL, "s", "") & "://" & gsScript3 & IIf(sServer Like "/*", "", gsScriptCommand3) & sServer
  If gbUseSSL Then
    If Not sServer Like "https*" Then oHtmlForm.PutField gsAnsPassField, ""  'Passwort nie unverschlüsselt senden!
    If Not sServer Like "https*" Then oHtmlForm.PutField gsAnsTokenField, ""  'Token nie unverschlüsselt senden!
  End If
  sKommando = oHtmlForm.GetFields(gsSiteEncoding) 'und welche Daten?
    
  sKommando = sKommando & oHtmlForm.ClickImage(gsAnsLoginSubmitImage1)
  sKommando = sKommando & oHtmlForm.ClickImage(gsAnsLoginSubmitImage2)
  sKommando = sKommando & oHtmlForm.ClickImage(gsAnsLoginSubmitImage3)
  
  If bIsBuyItNow Then
    sKommando = sKommando & oHtmlForm.ClickImage(gsAnsBuyStep1SubmitImage)
    sKommando = sKommando & oHtmlForm.ClickImage(gsAnsBuyStep2SubmitImage)
  Else
    sKommando = sKommando & oHtmlForm.ClickImage(gsAnsBidStep1SubmitImage)
    sKommando = sKommando & oHtmlForm.ClickImage(gsAnsBidStep2SubmitImage)
  End If
  
  If oHtmlForm.GetField(gsAnsPassField) <> "" Then oHtmlForm.PutField gsAnsPassField, "xxx"
  sKommandoPublic = oHtmlForm.GetFields(gsSiteEncoding) 'und welche Daten?
  
  If InStr(1, sBuffer, gsAnsBidConfirm) > 0 Then  ' Wir sollen bestätigen und der User stimmt, wir warten bis kurz vor Ende und schlagen dann zu
    If Not bSofortBieten And gfVorlaufSnipe > 0 Then
      Call DebugPrint("Warten auf den Snipe")
      Do Until MyNow >= DateAdd("s", -Int(gfLzMittel + 0.5 + gfVorlaufSnipe), datEndeZeit)
        DoEvents
        Call Sleep(10)
        
        ' Wir wollen eine aktuelle Laufzeit haben und da wir grad noch Zeit haben, machen wir noch ein wenig Traffic
        If MyNow < DateAdd("s", -glVorlaufGebot, datEndeZeit) And Int(Timer) Mod 10 = 0 Then
          Call ShortPost(Replace("http://" & gsScript4 & gsScriptCommand4 & gsCmdViewItem, "[Item]", sItem))
        End If
        
      Loop
    End If
    Call PanelText(StatusBar1, 2, gsarrLangTxt(91) & sItem)
  End If
  
PostIt:
  Call DebugPrint("Sende Anfrage zum Server: " & sServer & "?" & sKommandoPublic)
 
  sBuffer = ShortPost(sServer, sKommando, sReferer, sEBayUser)    'GET
  
  lAufruf = lAufruf + 1
  sSaveFile = gsTempPfad & "\Art-" & sItem & "-2.html"
  Call SaveToFile(StripJavaScript(sBuffer), sSaveFile)
  sSaveFile = gsTempPfad & "\Art-" & sItem & "-" & CStr(lAufruf) & "-2-" & CStr(lVersuch) & ".html"
  Call SaveToFile(sBuffer, sSaveFile)  ' ohne StripJavaScript, debuggt sich einfach besser!
  If (lDurchgang = 1) Then
    sSaveFile = gsTempPfad & "\Art-" & sItem & "-status.html"
    Call SaveToFile(StripJavaScript(sBuffer), sSaveFile)
  End If
  
Loop

Call PanelText(StatusBar1, 2, "")

End Function

Private Function ZeitPrüfung(ByVal iIdx As Integer, ByVal bRefreshOnly As Boolean) As Boolean
Dim fDz As Double
Dim fTimeVal As Double
Dim fDiff As Double
Dim fNowVal As Double
Dim iAktRow As Integer
Dim bShowIt As Boolean
Dim fUtcOffset As Double

ZeitPrüfung = False

'iAktRow ist Index auf Zeile, 0..Max
'iIdx = index auf ArtikelArray!

iAktRow = iIdx - VScroll1.Value
bShowIt = iAktRow >= 0 And iAktRow <= giMaxRow And Me.WindowState <> vbMinimized

On Error GoTo errhdl

fUtcOffset = GetUTCOffset()
If fUtcOffset <> gtarrArtikelArray(iIdx).TimeZone And gtarrArtikelArray(iIdx).EndeZeit <> myDateSerial(1999, 9, 9) Then
  gtarrArtikelArray(iIdx).EndeZeit = DateAdd("n", 60 * (fUtcOffset - gtarrArtikelArray(iIdx).TimeZone), gtarrArtikelArray(iIdx).EndeZeit)
  gtarrArtikelArray(iIdx).TimeZone = fUtcOffset
  Call ArtikelArrayToScreen(iIdx, False, True)
End If

'Restzeit eintragen
fDz = gtarrArtikelArray(iIdx).EndeZeit
fNowVal = MyNow

If fDz >= fNowVal Then

    fDiff = fDz - fNowVal
       
    'Restzeit wieder grün machen wenn noch genug Zeit ist
    If fDiff > gfVorlaufGebotTimeVal And gtarrArtikelArray(iIdx).Status <> [asEnde] Then
        If bShowIt Then
            Restzeit(iAktRow).BackColor = -2147483633
        End If
    End If
       
    'nur bieten, falls noch nicht geboten wurde, lg 16.05.03
    If fDiff <= gfVorlaufGebotTimeVal And gtarrArtikelArray(iIdx).Status <= [asNixLos] Then

        'nu gehts los
        If bShowIt Then
            Restzeit(iAktRow).BackColor = vbYellow
            Restzeit(iAktRow).Caption = gsarrLangTxt(92)
            EndeZeit(iAktRow).BackColor = vbRed
        End If
        
        If Not bRefreshOnly Then
            gtarrArtikelArray(iIdx).Status = [asBieten]
        End If
        
        ZeitPrüfung = True
        'raus um Tempo zu machen
        Exit Function
    End If
    
    'jetzt Sofort-Kauf???
    If gtarrArtikelArray(iIdx).Status = [asBuyOnlyBuyItNow] Then
    
        'nu gehts los
        If bShowIt Then
            Restzeit(iAktRow).BackColor = vbYellow
            Restzeit(iAktRow).Caption = gsarrLangTxt(92)
            EndeZeit(iAktRow).BackColor = vbRed
        End If
        
        ZeitPrüfung = True
        'raus um Tempo zu machen
        Exit Function
    End If
   
    
    
    If bShowIt Then
        'Restzeit nur aktualisieren, wenn Artikel gefunden wurde
        If gtarrArtikelArray(iIdx).Status <> [asNotFound] Then
           Restzeit(iAktRow).Caption = TimeLeft2String(fDiff) 'lg 31.07.03
        End If
        'Farbe nur auf grau falls noch nicht geboten wurde, lg 16.05.03
        If gtarrArtikelArray(iIdx).Status <= [asNixLos] Then
          Restzeit(iAktRow).BackColor = -2147483633
        End If
    End If
    
    On Error Resume Next
    fTimeVal = gfRestzeitBerechner ' TimeValue(Aktion.Caption) ' lg 31.07.03
    'nächste Auktion nur setzen, falls noch nicht geboten wurde, lg 16.05.03
    If fDiff < fTimeVal And gtarrArtikelArray(iIdx).Gebot > 0 And gtarrArtikelArray(iIdx).Status <= [asNixLos] Then
        gfRestzeitBerechner = fDiff
        If gbAutoMode Then
            'mae 050617: über welchen Account wird als nächstes geboten?
            gsNextUser = gtarrArtikelArray(iIdx).UserAccount
            If gsNextUser = "" Then gsNextUser = gsUser 'StandardUser
        End If
    End If
Else    'Auktion ist beendet
    If bShowIt Then
        Restzeit(iAktRow).Caption = gsarrLangTxt(93)
        Restzeit(iAktRow).BackColor = vbRed
    End If
    
    If gtarrArtikelArray(iIdx).Status <= [asNixLos] Then
        gtarrArtikelArray(iIdx).Status = [asEnde]
        Call DebugPrint("Art: " & gtarrArtikelArray(iIdx).Titel & "(" & gtarrArtikelArray(iIdx).Artikel & ") Ende")
    End If
    
End If

errhdl:

End Function

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub Do_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  On Error Resume Next
  Dim sBuffer As String
  Dim i As Long
  
  For i = 1 To 255
    If Data.GetFormat(i) Then
      sBuffer = Data.GetData(i)
      If sBuffer > "" Then Exit For
    End If
  Next i
  
  Call HandleDragDropData(Index, sBuffer, Effect, Button, Shift, X, Y)

End Sub

Public Sub HandleDragDropData(iIdx As Integer, sData As String, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    Dim i As Integer
    Dim sTmp As String
    Dim sArtikel As String
    Dim fGebot As Double
    Dim sGruppe As String
    Dim sUser As String
    Dim arr As Variant
    
    sTmp = sData
    Call DebugPrint("HandleDragDropData: " & sData, 3)
    
    sArtikel = GetItemFromUrl(sTmp)
    If Len(sArtikel) > 0 Then ' schon was da
        arr = Array(sArtikel)
    Else
        arr = ResolveItemUrl(sTmp)
    End If
    
    For i = LBound(arr) To UBound(arr)
    
        sArtikel = GetItemFromUrl(arr(i))
        
        If Len(sArtikel) > 0 Then
            'Prüfe, ob Shift Taste gedrückt war, auf einen Artikel fallengelassen wurde und Index gültig
            If Shift = vbShiftMask And iIdx > 0 And iIdx <= UBound(gtarrArtikelArray()) Then
                'Ja, Index ist für ArtikelArray gültig also Gebot und Gruppe übernehmen
                fGebot = 0
                sGruppe = ""
                fGebot = gtarrArtikelArray(iIdx).Gebot
                sGruppe = gtarrArtikelArray(iIdx).Gruppe
                sUser = gtarrArtikelArray(iIdx).UserAccount
                Call InsertArtikelBuff(sArtikel & vbTab & CStr(fGebot) & vbTab & sGruppe & vbTab & sUser)
            Else
                'Einfach so übernehmen
                Call InsertArtikelBuff(sArtikel)
            End If
        End If
        
    Next i
    
End Sub

Private Sub Form_Resize()

On Error GoTo errhdl

If mbInitDone Then
    Call SetFocusRect(-2)
    If Me.WindowState = vbMinimized Then 'lg 22.07.03
        If gbMinToTray Then Call ToTaskbar
        If gbShowToolbar Then Call SwapToolbar("out")
        Call SwapZeilen("out")
        If gbShowDebugWindow Then frmDebug.Hide
    Else
        Call SwapZeilen("in")
        If gbShowToolbar Then Call SwapToolbar("in")
        Call moResize.Resize(Me.Width, Me.Height, miStartWidth, miStartHeight)
        mlPrevWindowState = Me.WindowState
        Call PanelRepaint
        Call VScroll1_Change
        If gbShowDebugWindow Then frmDebug.Show
    End If
    
    With Me
        If .WindowState = vbNormal Then
            glPosLeft = .Left
            glPosTop = .Top
            glPosWidth = .Width
            glPosHeight = .Height
        End If
        If mbMinimizeQueued Then
            mbMinimizeQueued = False
            mbIsMinimizing = True
            .WindowState = vbMinimized
            mbIsMinimizing = False
        End If
    End With
    
    Call SetFocusRect(-1)
End If 'mbInitDone
errhdl:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next 'MD-Marker , 20090406
    Dim lRet  As VbMsgBoxResult
    Dim sTmp As String
    Dim obj As Object
    Dim sData As String
    Dim bChangeFlag As Boolean
    
    If gbWarnenBeimBeenden Then
        If IstWasZuTun(sTmp) Then
            lRet = MsgBox(gsarrLangTxt(71) & " " & sTmp & "." & vbCrLf & gsarrLangTxt(28), vbQuestion Or vbYesNo)
        End If
    End If
    
    If lRet = vbNo Then
        Cancel = 1
        gbExplicitEnd = False
    Else
        'vorbereiten zum Programm beeden
        Call SaveWindowSettings
        
        sData = BuildArtikelCSV2()
        bChangeFlag = Not CBool(gsLastSavedCrc = Crc32(sData & vbCrLf))
        
        If mbQuietExit Then bChangeFlag = False
        
        If AutoSave.Enabled Then
            lRet = vbYes
        Else
            If bChangeFlag Then
                lRet = MsgBox(gsarrLangTxt(18), vbYesNoCancel Or vbQuestion)
            End If
        End If
        
        If lRet = vbCancel Then
            Cancel = 1
            gbExplicitEnd = False
        Else
            
            'Alle Timer anhalten
            For Each obj In frmHaupt.Controls
                If TypeOf obj Is Timer Then obj.Enabled = False
            Next
            
            If lRet = vbYes Then Call WriteArtikelCsv2(sData)
            
            If Not InDevelopment Then
                Call modSubclass.UnSubclass(frmDummy.hWnd)
                If gbWheelUsed Then
                    MWheel1.DisableWheel
                End If
            End If
                       
            Set moResize = Nothing
            Call RemoveCookies
            Set goCookieHandler = Nothing
            Call ShrinkLogfile
            
            If gbTest Then Call DebugPrint("Programmende ")
            
            'Alle Forms entladen , bis auf frmHaupt
            For Each obj In Forms
                If obj.Name <> Me.Name Then Unload obj
            Next
            
            If gbUsesModem Then Call ModemHangUp
            
            If mtTrayIcon.hWnd Then
                Call Shell_NotifyIcon(NIM_DELETE, mtTrayIcon)
            End If
            
            'Abschließend , noch die frmHaupt entladen.
            Unload Me: Set frmHaupt = Nothing
            End 'und jetzt ist aber endgültig Schluss!
            
        End If 'lRet = vbCancel
    End If 'lRet = vbNo
    
End Sub
Private Sub Gebot_LostFocus(Index As Integer)
On Error Resume Next
Dim sTmp As String
Dim fVal As Double

GebotTimer.Enabled = False

If Index >= 0 Then
    If Artikel(0).Tag = "in" Then
        With Gebot(Index)
            'prüfen ob ok
            sTmp = Trim(.Text)
            If CDbl(sTmp) = 0 Then
                sTmp = ""
            Else
                fVal = String2Float(sTmp)
                sTmp = Format$(fVal, "###,##0.00")
            End If
            
            If (Index + VScroll1.Value) > giAktAnzArtikel Then
                .Text = ""
            Else
                .Text = sTmp
                .FontBold = True
                .FontItalic = False
                
                If Not gbAutoMode And .Text <> "" Then
                    Select Case Waehrung(Index).Caption
                        Case "€"
                            .ToolTipText = Format(1.9558 * CDbl(sTmp), "#,##0.00") & " DM"
                        Case "£"
                            .ToolTipText = "ca. " & Format(CDbl(.Text) * gcolWeValues("GBP"), "#,##0.00") & " €"
                        Case "$"
                            .ToolTipText = "ca. " & Format(CDbl(.Text) * gcolWeValues("USD"), "#,##0.00") & " €"
                        Case "CHF"
                            .ToolTipText = "ca. " & Format(CDbl(.Text) * gcolWeValues("CHF"), "#,##0.00") & " €"
                        Case "AU$"
                            .ToolTipText = "ca. " & Format(CDbl(.Text) * gcolWeValues("AUD"), "#,##0.00") & " €"
                        Case "C$"
                            .ToolTipText = "ca. " & Format(CDbl(.Text) * gcolWeValues("CAD"), "#,##0.00") & " €"
                        Case Else
                            .ToolTipText = gsarrLangTxt(79)
                    End Select
                End If
                
                With gtarrArtikelArray(Index + VScroll1.Value)
                    If .Gebot <> String2Float(Gebot(Index).Text) Then
                        .Gebot = String2Float(Gebot(Index).Text)
                        
                        If .Status = [asCancelGroup] Or .Status = [asLowBid] Then
                            If ResetStatusCancel() Then .Status = [asNixLos]
                        End If
                        
                        If .Status = [asBuyOnlyCanceled] Then
                            If ResetStatusCancel() Then .Status = [asBuyOnly]
                        End If
                        
                        Call ArtikelArrayToScreen(VScroll1.Value)
                        .LastChangedId = GetChangeID()
                    End If
                    
                    Call CheckBietgruppe(.Gruppe)
                End With 'gtarrArtikelArray(Index + VScroll1.Value)
                
                Call CheckSofortkaufArtikel
                giLastGebotEditedIndex = -1
            End If '(Index + VScroll1.Value) > giAktAnzArtikel
        End With 'Gebot(Index)
    End If 'Artikel(0).Tag = "in"
End If 'Index >= 0
End Sub

Private Sub StarteBrowser()
On Error Resume Next
Dim sTmp As String, sBuffer As String
Dim sKommando As String
Dim lPos As Long

If gsEbayLocalPass <> "" Then
    sBuffer = ""

    sTmp = gsCmdWatchList
    lPos = InStr(1, sTmp, "[User]")
    If lPos > 0 Then
        sBuffer = Left$(sTmp, lPos - 1) & URLEncode(gsUser)
        sTmp = Mid$(sTmp, lPos + 6)
    End If
    lPos = InStr(1, sTmp, "[lPass]")
    If lPos > 0 Then
        sBuffer = sBuffer & Left$(sTmp, lPos - 1) & URLEncode(gsEbayLocalPass)
        sTmp = Mid$(sTmp, lPos + 7)
    End If

    sKommando = sBuffer & sTmp

    gsGlobalUrl = "http://" & gsScript1 & gsScriptCommand1 & sKommando
    '& "MyEbayItemsBiddingOn&userid=" & gsUser & "&pass=" & gsEbayLocalPass & "&first=N&sellerSort=3&bidderSort=3&watchSort=3&dayssince=2&p1=0&p2=0&p3=0&p4=0&p5=0"
Else
    gsGlobalUrl = "http://" & gsMainUrl
End If

Call ExecuteDoc(Me.hWnd, gsGlobalUrl)
'ShowBrowser (Me.hWnd)
'frmBrowser.Show

End Sub

Private Sub Artikel_Change(Index As Integer)
    Static bInArtikelChange As Boolean
    If bInArtikelChange Then Exit Sub
    
    If Artikel(0).Tag = "in" Then
        bInArtikelChange = True
        EndeZeit(Index).Caption = ""
        Titel(Index).Caption = ""
        Preis(Index).Caption = ""
        Versandkosten(Index).Caption = ""
        Restzeit(Index).Caption = ""
        Restzeit(Index).BackColor = -2147483633
        Status(Index).Caption = ""
        Status(Index).BackColor = -2147483633
        Gebot(Index).Text = ""
        Gebot(Index).ToolTipText = ""
        Bietgruppe(Index).Text = ""
        EndeZeit(Index).BackColor = -2147483633
        Waehrung(Index).Caption = ""
        Preis(Index).ToolTipText = "Click =" & gsarrLangTxt(50)
        Ecke(Index).Visible = False
        On Error Resume Next
        If Trim(Artikel(Index).Text) = "" Then
            Artikel_LostFocus Index
        End If
    End If
    bInArtikelChange = False
End Sub

Private Sub Gebot_Change(Index As Integer)
  Gebot(Index).FontBold = False
  Gebot(Index).BackColor = vbWindowBackground
  giLastGebotEditedIndex = Index
End Sub

Private Sub mnuAbout_Click()

    'Zuerst das entsprechende Property setzen und dann
    'Call frmAbout.SetSpendeActiv
    Load frmAbout 'die Form laden , anschließend
    frmAbout.Show vbModal, Me 'anzeigen
    
End Sub

Private Sub mnuExit_Click()
    gbExplicitEnd = True
    Unload Me
End Sub

Private Sub mnuHelp_Click()

    If gsarrLangTxt(418) <> "" And UCase(Dir(gsarrLangTxt(418))) = UCase(gsarrLangTxt(418)) Then
        Call ExecuteDoc(Me.hWnd, gsarrLangTxt(418))
    Else
        MsgBox Replace(gsarrLangTxt(419), "%FILE%", gsarrLangTxt(418)), vbExclamation, gsarrLangTxt(215) & " - " & gsarrLangTxt(44)
    End If
End Sub

Private Sub mnuReadArtikel_Click()
    
    Dim lRet As VbMsgBoxResult
    
    lRet = MsgBox(gsarrLangTxt(19) & vbCrLf & vbCrLf & gsarrLangTxt(10) & _
        vbCrLf & gsarrLangTxt(11) & vbCrLf & gsarrLangTxt(12) & vbCrLf, vbInformation Or vbYesNoCancel)
    
    If lRet <> vbCancel Then
        'ReadEbayIni
        Call ReadArtikelIni(CBool(lRet = vbYes))
        Call Sortiere
    End If
    
End Sub

Public Sub SaveArtikel()

  mnuSaveArtikel_Click

End Sub

Private Sub mnuSaveAll_Click()
    
    Call mnuSaveArtikel_Click
    Call mnuSaveSettings_Click
    
End Sub

Private Sub mnuSaveArtikel_Click()

    Call WriteArtikelCsv2
    Call PanelText(StatusBar1, 1, gsarrLangTxt(31) & gsarrLangTxt(94), True)
    
End Sub

Private Sub mnuSaveSettings_Click()

    Call SaveAllSettings
    Call PanelText(StatusBar1, 1, gsarrLangTxt(32) & gsarrLangTxt(94), True)
    
End Sub

Private Sub POPTimer_Timer()

  If giSuspendState > 0 Then Exit Sub

  'schlägt alle 60 sec zu
  Dim i As Integer
  Dim fDz As Double
  Dim sTmp As String
  Dim sSender As String
  Dim sRecv As String
  Dim vSubjects As Variant
  Dim col As Collection
  Dim col2 As New Collection
  Dim sShutdownFlag As String
  Dim bSuspendFlag As Boolean
  Dim bAktualisieren As Boolean
  Dim bRet As Boolean
  
  On Error Resume Next
  
  POPTimer.Enabled = False
  
  If gbUsePop Then
    miPopTimerCount = miPopTimerCount + 1
    If miPopTimerCount >= giPopZyklus Then
      miPopTimerCount = 0
          'mal sehen ob keine Auktion ansteht
      fDz = gfRestzeitZaehler
      If fDz > myTimeSerial(0, 5, 0) Then
        If IsOnline Then
          bAktualisieren = ArtikelTimer.Enabled
          Set col = GetPop()
          tcpIn.Close
          While (col.Count > 0)
              sTmp = col.Item(1)(LBound(col.Item(1)) + 0)
              sSender = col.Item(1)(LBound(col.Item(1)) + 1)
              sRecv = col.Item(1)(LBound(col.Item(1)) + 2)
              col.Remove (1)
              If col2.Count = 0 Then
                vSubjects = Split(sTmp, gsPopSubjectDelimiter)
                For i = LBound(vSubjects) To UBound(vSubjects)
                  sTmp = vSubjects(i)
                  col2.Add Array(sTmp, sSender, sRecv)
                Next i
              Else
                If sSender <> col2.Item(1)(LBound(col.Item(1)) + 1) Then
                  BearbeiteMailAuftrag col2, [saShutdown], bSuspendFlag
                  Set col2 = New Collection
                End If
                vSubjects = Split(sTmp, gsPopSubjectDelimiter)
                For i = LBound(vSubjects) To UBound(vSubjects)
                  sTmp = vSubjects(i)
                  col2.Add Array(sTmp, sSender, sRecv)
                Next i
              End If
          Wend
          If col2.Count > 0 Then Call BearbeiteMailAuftrag(col2, sShutdownFlag, bSuspendFlag)
          
          CheckAlleBietgruppen
         
          If bAktualisieren Then
              ArtikelTimer.Enabled = True
          End If
          
        End If 'isonline
      End If 'fDz
   End If 'timer abgelaufen
  
    If bSuspendFlag Then
      bRet = ShowUpdateBox(Me, [ftCountDown1], 103, IIf(Me.WindowState = vbMinimized Or Me.Visible = False, 2, 1), "Suspend", 60, gsarrLangTxt(208), gsarrLangTxt(209), gsarrLangTxt(210), "-", gsarrLangTxt(359), True)
      If bRet Then Call Suspend
    End If
  
    If sShutdownFlag <> "" Then
       Select Case sShutdownFlag
           Case "0"
              'Shutdown- Flag setzen
              gbUseWinShutdown = True
           Case "1"
              'Shutdown- Flag rücksetzen
              gbUseWinShutdown = False
           Case "2"
              'shutdown now
              bRet = ShowUpdateBox(Me, [ftCountDown1], 103, IIf(Me.WindowState = vbMinimized Or Me.Visible = False, 2, 1), "ShutDown", 60, gsarrLangTxt(208), gsarrLangTxt(209), gsarrLangTxt(210), "-", gsarrLangTxt(359), True)
              If bRet Then
                gbWarnenBeimBeenden = False
                mbQuietExit = True
                Call ShutDownWin
                Call mnuExit_Click
              Else
                gbUseWinShutdown = False
              End If
           Case "3"
              'end now
              bRet = ShowUpdateBox(Me, [ftCountDown2], 103, IIf(Me.WindowState = vbMinimized Or Me.Visible = False, 2, 1), gsarrLangTxt(354), 10, gsarrLangTxt(355), gsarrLangTxt(356), gsarrLangTxt(357), gsarrLangTxt(358), gsarrLangTxt(359), True)
              If bRet Then
                gbWarnenBeimBeenden = False
                mbQuietExit = True
                Call mnuExit_Click
              End If
            End Select
    End If 'sShutdownFlag
    
  End If 'usepop
  
  POPTimer.Enabled = gbAutoMode

End Sub

Private Function BearbeiteMailAuftrag(col As Collection, ByRef sShutdownFlag As String, ByRef bSuspendFlag As Boolean)

  On Error Resume Next
  
  'schlägt alle 60 sec zu
  Dim sBuffer As String
  Dim sBufferOrg As String
  Dim sKommando As String
  Dim sArtikel As String
  Dim sGebot As String
  Dim sGruppe As String
  Dim sAccount As String
  Dim lPosStart As Long
  Dim lPos As Long
  Dim i As Integer
  Dim bVorhanden As Boolean
  Dim sQuittText As String
  Dim bOk As Boolean
  Dim fStrVal As Double
  Dim iMsgCnt As Integer
  Dim bSaveMe As Boolean
  Dim sRecvReply As String
  Dim sArtikelString As String
  Dim sArtikelFileVersion As String
  Dim sGruppeAlt As String
  Dim sTmp As String
  Dim oRC4 As clsRC4
  Dim sSubject As String
  Dim sMailText As String
  
  
  sShutdownFlag = ""
  bVorhanden = False

  While (col.Count > 0)
      sBuffer = col.Item(1)(LBound(col.Item(1)) + 0)
      sRecvReply = col.Item(1)(LBound(col.Item(1)) + 1)
      col.Remove (1)
      
      iMsgCnt = iMsgCnt + 1
      sQuittText = sQuittText & vbCrLf & "# ************ Message " & iMsgCnt & " *******************" & vbCrLf
  
      DebugPrint "POP- command: " & sBuffer, 2
      
      sAccount = ""
      sGruppe = ""
      sGebot = ""
      sArtikel = ""

      bVorhanden = False
      
      sBuffer = RTrim(sBuffer) & " "
      
      'Alles Upcase ..
      sBufferOrg = sBuffer
      sBuffer = UCase(sBuffer)
      'String zerlegen
      
      sKommando = "nix"
      
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "READCSV") + 4
          If lPos > 4 Then
              sKommando = "Read"
              sArtikel = "0"
          End If
      End If
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "ADDMOD") + 4
          If lPos > 4 Then
              sKommando = "addmod"
          End If
      End If
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "ADD") + 3
          If lPos > 3 Then
              sKommando = "add"
          End If
      End If
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "MOD") + 3
          If lPos > 3 Then
              sKommando = "mod"
          End If
      End If
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "DEL") + 3
          If lPos > 3 Then
              sKommando = "del"
          End If
      End If
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "STATUS") + 4
          If lPos > 4 Then
              sKommando = "status"
              lPos = InStr(1, sBuffer, "REFR")
              If lPos > 0 Then
                  sArtikel = "refresh"
              Else
                  sArtikel = "x"
              End If
          End If
      End If
      
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "SEND") + 4
          If lPos > 4 Then
              sKommando = "send"
              sArtikel = "0"
          End If
      End If
      
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "LOAD") + 4
          If lPos > 4 Then
              sKommando = "load"
              sArtikel = "0"
          End If
      End If
      

      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "SHUTDOWN") + 4
          If lPos > 4 Then
              sKommando = "shutdown"
              sArtikel = "0"
              'Zusatzkommandos?
              lPos = InStr(1, sBuffer, "REMO") + 4
              If lPos > 4 Then
                  sArtikel = "1"
              End If
              lPos = InStr(1, sBuffer, "NOW") + 3
              If lPos > 3 Then
                  sArtikel = "2"
              End If
          End If
      End If
      
      
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "SUSPEND") + 4
          If lPos > 4 Then
              sKommando = "suspend"
              sArtikel = "0"
          End If
      End If
      
      
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "END") + 3
          If lPos > 3 Then
              sKommando = "end"
              sArtikel = "3"
          End If
      End If
      
      
      If sKommando = "nix" Then
          lPos = InStr(1, sBuffer, "ENCRYPTION_NEEDED")
          If lPos > 0 Then
              sKommando = "Encryption_needed"
              sArtikel = "0"
          End If
      End If
      

      lPos = InStr(1, sBuffer, "ART") + 3
      
      If lPos > 3 Then
          lPosStart = InStr(lPos, sBuffer, " ")
          Do While lPos = lPosStart
              lPos = lPos + 1
              lPosStart = InStr(lPos, sBuffer, " ")
          Loop
          sArtikel = Trim(Mid(sBuffer, lPos, lPosStart - lPos))
      End If
      
      ' Mal sehen, ob und wo ein Gebotsbetrag steht ..
      lPos = InStr(1, sBuffer, "EURO") + 4
      
      If lPos <= 4 Then
           lPos = InStr(1, sBuffer, "EUR") + 3
      End If
      If lPos <= 3 Then
          lPos = InStr(1, sBuffer, "EU") + 2
      End If
      If lPos <= 2 Then
          lPos = InStr(1, sBuffer, "€") + 1
      End If

      If lPos > 1 Then
          lPosStart = InStr(lPos, sBuffer, " ")
          Do While lPos = lPosStart
              lPos = lPos + 1
              lPosStart = InStr(lPos, sBuffer, " ")
          Loop
          sGebot = Trim(Mid(sBuffer, lPos, lPosStart - lPos))
          
          'und richtig zusammenstauchen:
          fStrVal = String2Float(sGebot)
          sGebot = CStr(fStrVal)

      End If
      
      lPos = InStr(1, sBuffer, "GRUPPE") + 6
      If lPos > 6 Then
          lPosStart = InStr(lPos, sBuffer, " ")
          Do While lPos = lPosStart
              lPos = lPos + 1
              lPosStart = InStr(lPos, sBuffer, " ")
          Loop
          sGruppe = Trim(Mid(sBufferOrg, lPos, lPosStart - lPos))
      End If

      lPos = InStr(1, sBuffer, "ACCOUNT") + 7
      If lPos > 7 Then
          lPosStart = InStr(lPos, sBuffer, " ")
          Do While lPos = lPosStart
              lPos = lPos + 1
              lPosStart = InStr(lPos, sBuffer, " ")
          Loop
          sAccount = Trim(Mid(sBufferOrg, lPos, lPosStart - lPos))
      End If

      'wenigstens der Artikel muss vorhanden sein ..
      If Not sArtikel = "" Then
        bOk = False
        
        If sKommando = "addmod" Then
          For i = 1 To giAktAnzArtikel
              If gtarrArtikelArray(i).Artikel = sArtikel Then
                   bVorhanden = True
                   Exit For
              End If
          Next i
          sKommando = IIf(bVorhanden, "mod", "add")
          bVorhanden = False
        End If
        
        Select Case LCase(sKommando)
          Case "add"
              'wir wollen das Gebot nicht doppelt haben:
              For i = 1 To giAktAnzArtikel
                  If gtarrArtikelArray(i).Artikel = sArtikel Then
                       bVorhanden = True
                       sQuittText = sQuittText & "Quittung- Art " & sArtikel & " nicht eingefügt, schon vorhanden. Bitte MOD benutzen!"
                       Exit For
                  End If
              Next i
              
              If Not bVorhanden And gbPopNeedsUsername And GetAccountFromAccount(sAccount) = "" Then
                  bVorhanden = True
                  sQuittText = sQuittText & "Quittung- Art " & sArtikel & " nicht eingefügt. Bitte ACCOUNT angeben!"
              End If
              
              If Not bVorhanden Then
                  i = AddArtikel(sArtikel)
                  If i > 0 Then
                      gtarrArtikelArray(i).Gebot = String2Float(sGebot)
                      gtarrArtikelArray(i).Gruppe = sGruppe
                      gtarrArtikelArray(i).UserAccount = GetAccountFromAccount(sAccount)
                  
                      bSaveMe = True
                      
                      If sGruppe = "" Then
                          sQuittText = sQuittText & "Quittung- Art " & sArtikel & " Nr " & CStr(giAktAnzArtikel) & " Gebot " & sGebot & gtarrArtikelArray(i).WE & " ohne Gruppe eingefügt"
                      Else
                          sQuittText = sQuittText & "Quittung- Art " & sArtikel & " Nr " & CStr(giAktAnzArtikel) & " Gebot " & sGebot & gtarrArtikelArray(i).WE & " Gruppe " & sGruppe & " eingefügt"
                      End If
                      If sAccount > "" Then
                          sQuittText = sQuittText & ", Account: " & sAccount
                      End If
                      If gtarrArtikelArray(i).EndeZeit = myDateSerial(1999, 9, 9) Then
                          sQuittText = sQuittText & " - Art " & sArtikel & " ist nicht vorhanden"
                          gtarrArtikelArray(i).Gebot = 0
                      Else
                          sQuittText = sQuittText & vbCrLf & gtarrArtikelArray(i).Titel
                      End If
                  End If
                  gtarrArtikelArray(i).LastChangedId = GetChangeID()
                  ArtikelArrayToScreen VScroll1.Value
              End If
          Case "mod"
              For i = 1 To giAktAnzArtikel
                  If gtarrArtikelArray(i).Artikel = sArtikel Then
                       bVorhanden = True
                       gtarrArtikelArray(i).Gebot = String2Float(sGebot)
                       If Not sGruppe = "" Then
                          sGruppeAlt = gtarrArtikelArray(i).Gruppe
                          gtarrArtikelArray(i).Gruppe = sGruppe
                          CheckBietgruppe sGruppeAlt
                          CheckBietgruppe sGruppe
                       End If
                       CheckSofortkaufArtikel
                       If Not sAccount = "" Then
                          gtarrArtikelArray(i).UserAccount = GetAccountFromAccount(sAccount)
                       End If
                       
                       If gtarrArtikelArray(i).Status = [asCancelGroup] Or _
                          gtarrArtikelArray(i).Status = [asLowBid] Then
                              gtarrArtikelArray(i).Status = [asNixLos]
                       End If
                       If gtarrArtikelArray(i).Status = [asBuyOnlyCanceled] Then
                              gtarrArtikelArray(i).Status = [asBuyOnly]
                       End If
                       bSaveMe = True
                       'Update_Artikel (i)
                       gtarrArtikelArray(i).LastChangedId = GetChangeID()
                       ArtikelArrayToScreen VScroll1.Value

                       If sGruppe = "" Then
                           sQuittText = sQuittText & "Quittung- Art " & sArtikel & " Gebot " & sGebot & gtarrArtikelArray(i).WE & " ohne Gruppe geändert"
                       Else
                           sQuittText = sQuittText & "Quittung- Art " & sArtikel & " Gebot " & sGebot & gtarrArtikelArray(i).WE & " Gruppe " & sGruppe & " geändert"
                       End If
                       If sAccount > "" Then
                           sQuittText = sQuittText & ", Account: " & sAccount
                       End If
                       sQuittText = sQuittText & vbCrLf & gtarrArtikelArray(i).Titel
                       
                       Exit For
            
                  End If
              Next i
              
              If Not bVorhanden Then
                  sQuittText = sQuittText & "Quittung- Art " & sArtikel & " wurde bei Biet-O-Matic nicht gefunden, nicht geändert"
              End If
          Case "del"
              For i = giAktAnzArtikel To 1 Step -1
                  'alle Artikel löschen
                  If sArtikel = "*" Then
                      bVorhanden = True
                      sQuittText = sQuittText & "Quittung- Art " & gtarrArtikelArray(1).Artikel & " wurde gelöscht" & vbCrLf
                      sQuittText = sQuittText & gtarrArtikelArray(i).Titel & vbCrLf
                      RemoveArtikel 1, False, False
                  'abgelaufene Artikel löschen
                  ElseIf LCase(sArtikel) = "ended" Then
                      If GetRestzeitFromItem(i) = 0 Then
                          bVorhanden = True
                          sQuittText = sQuittText & "Quittung- Art " & gtarrArtikelArray(i).Artikel & " wurde gelöscht" & vbCrLf
                          sQuittText = sQuittText & gtarrArtikelArray(i).Titel & vbCrLf
                          RemoveArtikel i, False, False
                      End If
                  'bestimmten Artikel löschen
                  Else
                      If gtarrArtikelArray(i).Artikel = sArtikel Then
                          bVorhanden = True
                          sQuittText = sQuittText & "Quittung- Art " & sArtikel & " wurde gelöscht" & vbCrLf
                          sQuittText = sQuittText & gtarrArtikelArray(i).Titel & vbCrLf
                          RemoveArtikel i
                          Exit For
                       End If
                  End If
                  RemoveArtikel 0, True, True
              Next i
              If Not bVorhanden Then
                  sQuittText = sQuittText & "Quittung- Art " & sArtikel & " wurde bei Biet-O-Matic nicht gefunden, nicht gelöscht"
              Else
                  ArtikelArrayToScreen VScroll1.Value
              End If
              
          Case "status"
              
              For i = 1 To giAktAnzArtikel
                  If Not gtarrArtikelArray(i).Artikel = "" Then
                  
                       If sArtikel = "refresh" Then
                          Update_Artikel i
                       End If
                       
                       If i > 1 Then
                          sQuittText = sQuittText & vbCrLf & vbCrLf
                       End If
                       
                       sQuittText = sQuittText & "Art " & gtarrArtikelArray(i).Artikel & ": " & gtarrArtikelArray(i).Titel & vbTab & " Max. Gebot " & Format(gtarrArtikelArray(i).Gebot, "###,##0.00") & " " & gtarrArtikelArray(i).WE & vbTab
                       
                       If gtarrArtikelArray(i).Gruppe <> "" Then
                          sQuittText = sQuittText & " Gruppe " & gtarrArtikelArray(i).Gruppe & vbTab
                       End If
                       If gtarrArtikelArray(i).Status > [asNixLos] And Not StatusIstBuyItNowStatus(gtarrArtikelArray(i).Status) Then
                          sQuittText = sQuittText & " - beendet "
                       Else
                          sQuittText = sQuittText & vbCrLf & vbTab & " Ende: " & CStr(gtarrArtikelArray(i).EndeZeit)
                       End If

                       Select Case gtarrArtikelArray(i).Status
                          Case [asNixLos]:
                              sQuittText = sQuittText & " - läuft noch " & TimeLeft2String(GetRestzeitFromItem(i))
                          Case [asErr]:
                              sQuittText = sQuittText & " - Auktion verloren"
                          Case [asOK]:
                              sQuittText = sQuittText & " - Auktion gewonnen"
                          Case [asLowBid]:
                              sQuittText = sQuittText & " - Gebot zu niedrig"
                          Case [asCancelGroup]:
                              sQuittText = sQuittText & " - Auktion gecancelt"
                          Case [asBuyOnlyCanceled]:
                              sQuittText = sQuittText & " - Sofortkauf gecancelt"
                          Case [asHoldGroup]:
                              sQuittText = sQuittText & " - Auktion auf Hold"
                          Case [asBuyOnlyOnHold]:
                              sQuittText = sQuittText & " - Sofortkauf auf Hold"
                          Case [asBuyOnlyDelegated]:
                              sQuittText = sQuittText & " - Sofortkauf delegiert"
                          Case [asDelegatedBom]:
                              sQuittText = sQuittText & " - Auktion delegiert"
                          Case [asCancelBid]:
                              sQuittText = sQuittText & " - kein Gebot eingetragen"
                          Case [asBuyOnly]:
                              sQuittText = sQuittText & " - kein Gebot möglich, nur Sofortkauf"
                          Case [asAdvertisement]:
                              sQuittText = sQuittText & " - kein Gebot möglich, nur Preisanzeige"
                          Case [asUeberboten]
                              sQuittText = sQuittText & " - Auktion verloren, überboten"
                       End Select
                       If StatusIstBuyItNowStatus(gtarrArtikelArray(i).Status) Or gtarrArtikelArray(i).Status = [asAdvertisement] Then
                          sQuittText = sQuittText & ", Preis " & Format(gtarrArtikelArray(i).AktPreis, "###,##0.00") & " " & gtarrArtikelArray(i).WE
                       Else
                          sQuittText = sQuittText & ", " & CStr(gtarrArtikelArray(i).AnzGebote) & " Gebot(e), Akt. Preis " & Format(gtarrArtikelArray(i).AktPreis, "###,##0.00") & " " & gtarrArtikelArray(i).WE
                       End If
                       sQuittText = sQuittText & ", Account: " & gtarrArtikelArray(i).UserAccount
                       If Trim(gtarrArtikelArray(i).Kommentar) <> "" Then sQuittText = sQuittText & vbCrLf & "Kommentar: " & gtarrArtikelArray(i).Kommentar
                  End If
              Next i
          Case "load"
              LoadMyEbay True
              sQuittText = sQuittText & "Quittung- für " & sKommando & " Artikel gelesen, aktuell: " & CStr(giAktAnzArtikel) & " Artikel"
          Case "shutdown"
              sShutdownFlag = sArtikel
              sQuittText = sQuittText & "Quittung- für " & sKommando & " Flag = " & sShutdownFlag
          Case "end"
              sShutdownFlag = sArtikel
              sQuittText = sQuittText & "Quittung- für " & sKommando
          Case "suspend"
              bSuspendFlag = True
              sQuittText = sQuittText & "Quittung- für " & sKommando
          Case "send"
              sSubject = "b-o-m: readcsv"
              sMailText = BuildArtikelCSV2()
              
              If gbPopSendEncryptedAcknowledgment Then
                  Set oRC4 = New clsRC4
                  sSubject = Replace(oRC4.EncryptString(sSubject, gsPass, True), vbCrLf, "")
                  sMailText = oRC4.EncryptString(sMailText, gsPass, True)
                  Set oRC4 = Nothing
              End If
              For i = 1 To 3
                  bOk = SendSMTP(gsSendEndFromRealname & "<" & gsSendEndFrom & ">", sRecvReply, "subject: " & sSubject & vbCrLf & vbCrLf & sMailText)
                  tcpIn.Close
                  If bOk Then
                      Exit For
                  End If
              Next i
              sQuittText = sQuittText & "Quittung- für " & sKommando & IIf(bOk, " Ok", " SendSMTP fehlgeschlagen ")
          Case "read"
              lPos = InStr(1, sBuffer, "READ") + 4
              lPos = InStr(lPos, sBuffer, vbCrLf) + 1
              sArtikelString = Mid(sBufferOrg, lPos, Len(sBuffer) - lPos)

              'wir prüfen jetzt die Version der Artikel.csv und lesen das entsprechende Format ein, lg 04.05.03
              If InStr(sArtikelString, "Artikeldatei") = 0 Then
                  sTmp = sArtikelString
                  sTmp = Replace(sTmp, " ", "")
                  sTmp = Replace(sTmp, ".", "")
                  sTmp = Replace(sTmp, vbCr, "")
                  sTmp = Replace(sTmp, vbLf, "")
                  Set oRC4 = New clsRC4
                  sTmp = oRC4.DecryptString(sTmp, gsPass, True)
                  Set oRC4 = Nothing
                  If InStr(sTmp, "Artikeldatei") > 0 Then
                      sArtikelString = sTmp
                  End If
              End If
              
              sArtikelFileVersion = GetArtikelFileVersion(Left(sArtikelString, 100))
              If VersionValue(sArtikelFileVersion) >= VersionValue("2.4.0") Then
                AddCsvArtikel2 sArtikelString 'ab Version 2.4.0
              Else
                AddCsvArtikel2 sArtikelString, True 'bis Version 2.3.0
              End If
              
              ArtikelArrayToScreen VScroll1.Value
              bSaveMe = True
              sQuittText = sQuittText & "Quittung- für " & sKommando & " Artikel gelesen, aktuell: " & giAktAnzArtikel & " Artikel"
          Case "encryption_needed"
              sQuittText = sQuittText & "Quittung- keine Aktion, Befehl war nicht verschlüsselt, bitte auf " & gsBOMUrlHP & "hp/security.php verschlüsseln."
          Case Else
              sQuittText = sQuittText & "Quittung- für " & sKommando & " keine Aktion, Befehl unbekannt "
        End Select
      Else
        sQuittText = sQuittText & "Quittung- für " & sKommando & " keine Aktion, Artikel war leer "
      End If 'Artikelnr ok
      
      DebugPrint "POP- Auftrag " & sKommando & " ausgewertet ", 2
      
  'End If 'getpop
  
  Wend 'while getpop
  
  If sKommando = "send" And iMsgCnt = 1 Then sQuittText = ""
  
  'Quittung abschicken
  If Not sQuittText = "" And Not sRecvReply = "" Then
      'na ja, wir versuchen es halt 3 mal
      
      sQuittText = "# *************** " & sQuittText & vbCrLf & "# ******************** " & vbCrLf

      sSubject = "[BOM-Meldung] Quittung für " & CStr(iMsgCnt) & " Aufträge, vom " & CStr(MyNow)
      sMailText = sQuittText
      If gbPopSendEncryptedAcknowledgment Then
          Set oRC4 = New clsRC4
          sSubject = Replace(oRC4.EncryptString(sSubject, gsPass, True), vbCrLf, "")
          sMailText = oRC4.EncryptString(sMailText, gsPass, True)
          Set oRC4 = Nothing
      End If
      For i = 1 To 3
          bOk = SendSMTP(gsSendEndFromRealname & "<" & gsSendEndFrom & ">", sRecvReply, "subject: " & sSubject & vbCrLf & vbCrLf & sMailText)
          tcpIn.Close
          If bOk Then
              Exit For
          End If
      Next i
      
      DebugPrint "POP- Auftrag quittiert, OK? " & bOk, 2

      If bSaveMe Then
          WriteArtikelCsv2
      End If
  End If

End Function

Public Function GetAccountFromAccount(sAccount As String) As String
    
    Dim i As Integer
    
    For i = LBound(gtarrUserArray()) To UBound(gtarrUserArray())
        If UCase(gtarrUserArray(i).UaUser) = UCase(sAccount) Then
            sAccount = gtarrUserArray(i).UaUser
            GetAccountFromAccount = sAccount
            Exit Function  'MD-Marker
        End If
    Next i
    sAccount = "User '" & sAccount & "' ist nicht vorhanden"
    
End Function

Public Function CheckPassForAccount(sAccount As String, sPassword As String) As Integer
    
    Dim i As Integer
    
    For i = LBound(gtarrUserArray()) To UBound(gtarrUserArray())
        If UCase(gtarrUserArray(i).UaUser) = UCase(sAccount) Then
            If gtarrUserArray(i).UaPass = sPassword Or gtarrUserArray(i).UaPass = DecodePass(sPassword) Then
                CheckPassForAccount = 1
                If gtarrUserArray(i).UaUser = gsUser Then CheckPassForAccount = 2 ' Standard-User
                Exit For
            End If
        End If
    Next i
    
End Function

Sub SetUseTokenByAccount(Account As String)

  Dim i As Integer
  For i = LBound(gtarrUserArray) To UBound(gtarrUserArray)
    If UCase(gtarrUserArray(i).UaUser) = UCase(Account) Then
      gtarrUserArray(i).UaToken = True
      Exit Sub
    End If
  Next i

End Sub

Private Sub Status_Click(Index As Integer)
    
    Dim sItem As String
    
    Call SetFocusRect(Index)
    Call VersandkostenUebernehmen
    
    sItem = Artikel(Index).Text
    If sItem <> "" Then
        DoEvents
        Call ShowStatus(sItem)
    End If
    
End Sub

Sub ShowStatus(sItem As String)

    Dim sSaveFile As String
    Dim lFileLength As Long

    sSaveFile = gsTempPfad & "\Art-" & sItem & "-status.html"
    On Error Resume Next
    lFileLength = FileLen(sSaveFile)
    If lFileLength > 0 Then
        gsGlobalUrl = sSaveFile
        'Call ShowBrowser(Me.hWnd)
        Call ExecuteDoc(Me.hWnd, gsGlobalUrl)
        Exit Sub
    End If
    sSaveFile = gsTempPfad & "\Art-" & sItem & "-2.html"
    lFileLength = FileLen(sSaveFile)
    If lFileLength > 0 Then
        gsGlobalUrl = sSaveFile
        'Call ShowBrowser(Me.hWnd)
        Call ExecuteDoc(Me.hWnd, gsGlobalUrl)
        Exit Sub
    End If
    sSaveFile = gsTempPfad & "\Art-" & sItem & "-1.html"
    lFileLength = FileLen(sSaveFile)
    If lFileLength > 0 Then
        gsGlobalUrl = sSaveFile
        'Call ShowBrowser(Me.hWnd)
        Call ExecuteDoc(Me.hWnd, gsGlobalUrl)
    End If

End Sub

Private Sub Status_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)
    
End Sub

Private Sub TimeoutTimer_Timer()
    
    If giSuspendState = 0 Then
    
        giPopTimeOutCount = giPopTimeOutCount + 1
        If giPopTimeOutCount > giTimeOutTimerTimeOut Then
            TimeoutTimer.Enabled = False
            gbTimeOutOccurs = True
        End If
    End If 'giSuspendState = 0
    
End Sub

'
' Haupttimer, Zeitueberwachung + bieten
'
Private Sub Timer1_Timer()

On Error Resume Next

Dim bOk As Boolean
Dim iAktRow As Integer
Dim fDz As Double
Dim i As Integer
Dim fModemTime As Double
Dim fLaufzeit As Double
Dim iShowIt As Integer
Dim iCnt As Integer
Dim sGroup As String
Dim bOnlineTmp As Boolean
Dim bDummy As Boolean
Dim fArtikelzeit As Double
Dim fNowVal As Double
Dim bDoNotHangUp As Boolean
Dim sMailText As String 'lg 14.05.03
Dim bWasZuTun As Boolean
Dim fStartZeitpunkt As Double
Dim bWasBuyItNowStatus As Boolean
Dim bDoNotFallAsleep As Boolean
Dim sSubject As String
Dim oRC4 As clsRC4
Dim sItem As String
Static fLastSendCsv As Double

If giSuspendState = 0 Then
     
    Timer1.Enabled = False
    fStartZeitpunkt = GetSystemUptime()
    
    gfRestzeitBerechner = gfRESTZEITEWIG
    bWasZuTun = False
    If Blink(0).BackColor = &H8000000F Then
        Blink(1).BackColor = &H8000000F
        Blink(0).BackColor = vbGreen
    Else
        Blink(0).BackColor = &H8000000F
        Blink(1).BackColor = vbGreen
    End If
    
    For iAktRow = 1 To giAktAnzArtikel
    
      ' neu: Countdown weiterführen, auch wenn schon geboten wurde, lg 16.05.03
      If (GetRestzeitFromItem(iAktRow) > 0 Or gtarrArtikelArray(iAktRow).PostUpdateDone = 0) And gtarrArtikelArray(iAktRow).Titel > "" Then
        
        'sh Titel(15)/Taskbar(bis klick) blinken 2min vor Auctionsende wenn minimiert hochholen
        If gbAutoWarnNoBid And Not gbWarningflag Then
          If GetRestzeitFromItem(iAktRow) <= myTimeSerial(0, 2, 0) And _
             GetRestzeitFromItem(iAktRow) > myTimeSerial(0, 0, glVorlaufGebot + 1) And _
             gtarrArtikelArray(iAktRow).Gebot = 0 Then
            If Me.WindowState = vbMinimized Then
              Call FromTaskbar
            Else
              Call FlashIt(Me)
            End If
          End If
        End If
        
        bWasZuTun = (gtarrArtikelArray(iAktRow).Gebot > 0 And GetRestzeitFromItem(iAktRow) > 0) Or (gtarrArtikelArray(iAktRow).Status = [asBuyOnlyBuyItNow]) Or bWasZuTun 'lg 25.05.03
        
        If ZeitPrüfung(iAktRow, False) _
        Then
          gbWarningflag = False
          If gtarrArtikelArray(iAktRow).Gebot > 0 Or gtarrArtikelArray(iAktRow).Status = [asBuyOnlyBuyItNow] _
          Then
            DebugPrint "Bieten Art: " & gtarrArtikelArray(iAktRow).Titel & "(" & gtarrArtikelArray(iAktRow).Artikel & ") Start mit Account: " & IIf(gtarrArtikelArray(iAktRow).UserAccount = "" And giDefaultUser <> 0, gtarrUserArray(giDefaultUser).UaUser, gtarrArtikelArray(iAktRow).UserAccount)
            
            gbBeendenNachAuktionAktiv = True
            gbSuspendNachAuktionAktiv = True
            
            If gbPlaySoundOnBid Then PlaySound gsSoundOnBid
            
            sItem = gtarrArtikelArray(iAktRow).Artikel ' Artikel-Nr merken
            
            mbIsBidding = True
            bOk = Bieten(gtarrArtikelArray(iAktRow).Artikel, _
                         gtarrArtikelArray(iAktRow).Gebot, _
                         gtarrArtikelArray(iAktRow).eBayUser, _
                         gtarrArtikelArray(iAktRow).eBayPass, _
                         iAktRow, , gtarrArtikelArray(iAktRow).UseToken, StatusIstBuyItNowStatus(gtarrArtikelArray(iAktRow).Status))
            
            mbIsBidding = False
            iAktRow = ItemToIndex(sItem) ' und den Index neu berechnen falls der User zwischenzeitlich an der Liste rumgefummelt hat!
            If iAktRow = 0 Then Exit For ' uups, der Artikel wurde gelöscht, Schleife abbrechen
            ' TODO: Das Verlassen der Schleife wenn der Artikel gelöscht wurde ist nicht optimal, die Mailbenachrichtigung geht
            ' nicht mehr raus, die gruppierten Artikel werden nicht aktualisiert usw. Besser eine Kopie des Eintrags anlegen
            ' und damit arbeiten. Dann aber schauen, dass die Anzeige-Operationen nicht mehr ausgeführt werden wenn Index 0 ist.
            
            If gbPlaySoundOnBid And bOk Then PlaySound gsSoundOnBidSuccess
            If gbPlaySoundOnBid And Not bOk Then PlaySound gsSoundOnBidFail
            
            DebugPrint "Bieten Art: " & gtarrArtikelArray(iAktRow).Titel & "(" & gtarrArtikelArray(iAktRow).Artikel & ") Ende, Bieten mit Account: " & IIf(gtarrArtikelArray(iAktRow).UserAccount = "" And giDefaultUser <> 0, gtarrUserArray(giDefaultUser).UaUser, gtarrArtikelArray(iAktRow).UserAccount) & " okay? " & bOk
            
            sMailText = "Subject: " & gsarrLangTxt(420) & " "
                
            iShowIt = iAktRow - VScroll1.Value
            If iShowIt > giMaxRow Or Me.WindowState = vbMinimized Then iShowIt = -1
           
            bWasBuyItNowStatus = StatusIstBuyItNowStatus(gtarrArtikelArray(iAktRow).Status)
            
            If bOk Then
              gtarrArtikelArray(iAktRow).Status = [asOK]
              
              If iShowIt >= 0 Then
                Status(iShowIt).Caption = "OK"
                Status(iShowIt).BackColor = vbGreen
              End If
              
              'Alle Gebote für Artikel mit gleicher Bietgruppe zurücksetzen
              If gtarrArtikelArray(iAktRow).Gruppe <> "" Then
                  'Counter und Gruppe auseinanderbasteln
                  sGroup = GetGruppe(gtarrArtikelArray(iAktRow).Gruppe)
                  iCnt = GetAnzahlVonGruppe(gtarrArtikelArray(iAktRow).Gruppe)
                  
                  If iCnt > 0 Then iCnt = iCnt - 1
                  
                  For i = 1 To giAktAnzArtikel
                    If sGroup = GetGruppe(gtarrArtikelArray(i).Gruppe) _
                    And (gtarrArtikelArray(i).Status <= [asNixLos] Or gtarrArtikelArray(i).Status = [asBuyOnly]) _
                    And Not i = iAktRow Then
                        iShowIt = i - VScroll1.Value
                        If iShowIt > giMaxRow Then iShowIt = -1
    
                        If iCnt <= 0 Then
                            
                            If gtarrArtikelArray(i).Status = [asBuyOnly] Then
                              gtarrArtikelArray(i).Status = [asBuyOnlyCanceled]
                            Else
                              gtarrArtikelArray(i).Status = [asCancelGroup]
                            End If
                            
                            DebugPrint "Cancel Art: " & gtarrArtikelArray(i).Titel & "(" & gtarrArtikelArray(iAktRow).Artikel & ")"
                            
                            If iShowIt >= 0 Then
                              Status(iShowIt).Caption = "Cancel"
                              Status(iShowIt).BackColor = vbYellow
                            End If
                            
                        Else
                            gtarrArtikelArray(i).Gruppe = sGroup & ";" & CStr(iCnt)
                            DebugPrint "Gruppe " & sGroup & " count " & iCnt & " Art: " & gtarrArtikelArray(iAktRow).Titel & "(" & gtarrArtikelArray(iAktRow).Artikel & ")"
                            
                           If iShowIt >= 0 Then
                                Bietgruppe(iShowIt).Text = gtarrArtikelArray(i).Gruppe
                           End If
                        End If 'counter auf 0
                        gtarrArtikelArray(i).LastChangedId = GetChangeID()
                    End If
                  Next i
                  CheckBietgruppe sGroup
                  CheckSofortkaufArtikel
              End If
           
              If gbSendAuctionEnd And ((Not gbArtikelRefreshPost) Or bWasBuyItNowStatus) Then
                If bWasBuyItNowStatus Then
                  sMailText = Replace(sMailText, "[status]", gsarrLangTxt(421)) & gsarrLangTxt(428) 'gekauft
                Else
                  sMailText = Replace(sMailText, "[status]", gsarrLangTxt(421)) & gsarrLangTxt(423) 'erfolgreich beboten
                End If
              Else
                sMailText = "" 'ggf. beim RefreshPost senden
              End If
           
            Else
              gtarrArtikelArray(iAktRow).Status = miErrStatus
              
              If iShowIt >= 0 Then
                If miErrStatus = [asErr] Then
                    Status(iShowIt).Caption = gsarrLangTxt(96)
                Else
                    Status(iShowIt).Caption = gsarrLangTxt(95)
                End If
                Status(iShowIt).BackColor = vbRed
                Bietgruppe(iShowIt).Text = ""
              End If
              
              'lokale Bietgruppe löschen, da es nicht geklappt hat
              sGroup = gtarrArtikelArray(iAktRow).Gruppe
              gtarrArtikelArray(iAktRow).Gruppe = ""
           
              If gbSendAuctionEndNoSuccess And ((Not (gbArtikelRefreshPost And gbArtikelRefreshPost2)) Or bWasBuyItNowStatus) Then
                If bWasBuyItNowStatus Then
                  sMailText = Replace(sMailText, "[status]", gsarrLangTxt(422)) & gsarrLangTxt(429) 'nicht gekauft
                Else
                  sMailText = Replace(sMailText, "[status]", gsarrLangTxt(422)) & gsarrLangTxt(424) 'ueberboten
                End If
              Else
                sMailText = ""
              End If
              
              If sGroup > "" Then
                CheckBietgruppe sGroup
                CheckSofortkaufArtikel
              End If
           
            End If
          Else 'kein Gebot zur Bietzeit eingetragen
          
            gtarrArtikelArray(iAktRow).Status = [asCancelBid]
            gtarrArtikelArray(iAktRow).Gruppe = ""
    
            iShowIt = iAktRow - VScroll1.Value
            If iShowIt > giMaxRow Or Me.WindowState = vbMinimized Then iShowIt = -1
           
            If iShowIt >= 0 Then
                Restzeit(iShowIt).Caption = gsarrLangTxt(93)
                Restzeit(iShowIt).BackColor = vbRed
                Status(iShowIt).Caption = gsarrLangTxt(97)
                Bietgruppe(iShowIt).Text = ""
            End If
          
          End If 'Gebot <> ""
          gtarrArtikelArray(iAktRow).LastChangedId = GetChangeID()
        End If 'Zeitprüfung abgelaufen
      End If 'Restzeit abgelaufen
      
      If sMailText <> "" Then
        sMailText = Replace(sMailText, "\n", vbCrLf)
        sMailText = Replace(sMailText, vbCrLf, "")
        sMailText = sMailText & vbCrLf & vbCrLf & gsarrLangTxt(430) & vbCrLf
        sMailText = Replace(sMailText, "[url]", "http://" & gsScript4 & gsScriptCommand4 & gsCmdViewItem)
        sMailText = Replace(sMailText, "[Item]", gtarrArtikelArray(iAktRow).Artikel)
        sMailText = Replace(sMailText, "[item]", gtarrArtikelArray(iAktRow).Artikel)
        sMailText = Replace(sMailText, "[price]", Format(gtarrArtikelArray(iAktRow).AktPreis, "###,##0.00"))
        sMailText = Replace(sMailText, "[currency]", gtarrArtikelArray(iAktRow).WE)
        sMailText = Replace(sMailText, "[title]", gtarrArtikelArray(iAktRow).Titel)
        sMailText = Replace(sMailText, "[group]", gtarrArtikelArray(iAktRow).Gruppe)
        sMailText = Replace(sMailText, "[comment]", gtarrArtikelArray(iAktRow).Kommentar)
        sMailText = Replace(sMailText, "[bid]", Format(gtarrArtikelArray(iAktRow).Gebot, "###,##0.00"))
        sMailText = Replace(sMailText, "[endtime]", Date2Str(gtarrArtikelArray(iAktRow).EndeZeit))
        sMailText = Replace(sMailText, "[highbidder]", gtarrArtikelArray(iAktRow).Bieter)
        sMailText = Replace(sMailText, "[bidcount]", CStr(gtarrArtikelArray(iAktRow).AnzGebote))
        sMailText = Replace(sMailText, "[minbid]", Format(gtarrArtikelArray(iAktRow).MinGebot, "###,##0.00"))
        sMailText = Replace(sMailText, "\n", vbCrLf)
        InsertMailBuff sMailText
        sMailText = ""
      End If
      
      If gbAutoMode = False Then
        Exit Sub
      End If
      
    Next iAktRow
    
    If gfRestzeitBerechner = gfRESTZEITEWIG Then
        PanelText StatusBar1, 3, gsarrLangTxt(71) & " ---"
        updTaskbar (gsarrLangTxt(47) & ": " & gsUser & " / " & gsarrLangTxt(71) & " ---")
        ResetWakupTime
    Else
      'verschoben aus Zeitprüfung
      PanelText StatusBar1, 3, gsarrLangTxt(71) & " " & TimeLeft2String(gfRestzeitBerechner), IIf(gfRestzeitBerechner < myTimeSerial(0, 5, 0), vbYellow, &H8000000F)
      updTaskbar (gsarrLangTxt(47) & ": " & gsUser & " / " & gsarrLangTxt(71) & " " & TimeLeft2String(gfRestzeitBerechner))
      SetWakeupTime (MyNow + gfRestzeitBerechner - myTimeSerial(0, giWakeOnAuction, 0))
      If (gfRestzeitBerechner < myTimeSerial(0, giPreventSuspend, 0)) Then ResetSystemIdleTimer
    End If
    
    gfRestzeitZaehler = gfRestzeitBerechner
    fDz = gfRestzeitZaehler 'lg 31.07.03
    
    
    'ggf Artikle POS- Update, nur wenn genug Zeit bis zur nächsten Auktion
    bDoNotHangUp = False
    bDoNotFallAsleep = False
    
    If gbArtikelRefreshPost _
    And fDz > myTimeSerial(0, 0, 30) Then
    
        ShrinkLogfile
    
        For iAktRow = 1 To giAktAnzArtikel
    
          ' Auch nicht potentiell gewonnen updaten wenn ArtikelRefreshPost2, lg 29.05.03
          If Not gtarrArtikelArray(iAktRow).PostUpdateDone And _
             Not gtarrArtikelArray(iAktRow).Titel = "" And _
             ( _
                gtarrArtikelArray(iAktRow).Status = [asOK] Or _
               (gtarrArtikelArray(iAktRow).Status = [asUeberboten] And gbArtikelRefreshPost2) Or _
               (gtarrArtikelArray(iAktRow).Status = [asErr] And gbArtikelRefreshPost2) Or _
               (gtarrArtikelArray(iAktRow).Status = [asCancelBid] And gbArtikelRefreshPost2) Or _
               (gtarrArtikelArray(iAktRow).Status = [asEnde] And gbArtikelRefreshPost2) _
             ) _
             Then
          
            bDoNotHangUp = True
            bDoNotFallAsleep = True
            'Prüfen nach Ablauf
            'frühestens 10 sec nach Ende der Auktion
            fArtikelzeit = gtarrArtikelArray(iAktRow).EndeZeit
            fNowVal = MyNow - myTimeSerial(0, 0, 10)
            If fArtikelzeit < fNowVal Then
                
                sMailText = "Subject: " & gsarrLangTxt(420) & " "
                
                'los nu!
                DebugPrint "PostUpdate Art: " & gtarrArtikelArray(iAktRow).Titel & " (" & gtarrArtikelArray(iAktRow).Artikel & ")"
                
                gtarrArtikelArray(iAktRow).PostUpdateDone = True
                
                sItem = gtarrArtikelArray(iAktRow).Artikel ' Artikel-Nr merken
                SaveToFile StripJavaScript(Update_Artikel(iAktRow)), gsTempPfad & "\Art-" & gtarrArtikelArray(iAktRow).Artikel & "-3.html"
                iAktRow = ItemToIndex(sItem) ' und den Index neu berechnen falls der User zwischenzeitlich an der Liste rumgefummelt hat!
                If iAktRow = 0 Then Exit For ' uups, der Artikel wurde gelöscht, Schleife abbrechen
                
                'bin ich Höchstbietender (was sagt der Preis)?
                
                If gtarrArtikelArray(iAktRow).Gebot > 0 Then 'nur wenn überhaupt ein Gebot eingetragen war
                
                    If gtarrArtikelArray(iAktRow).Status = [asOK] And _
                       gtarrArtikelArray(iAktRow).AktPreis <= gtarrArtikelArray(iAktRow).Gebot Then
                       
                        'ok, ich habs :-)
                        DebugPrint "Gewonnen Art: " & gtarrArtikelArray(iAktRow).Titel & " (" & gtarrArtikelArray(iAktRow).Artikel & ")"
                        
                        If gbSendAuctionEnd Then
                          sMailText = Replace(sMailText, "[status]", gsarrLangTxt(421)) & gsarrLangTxt(425) 'ersteigert fuer
                        Else
                          sMailText = ""
                        End If
                    
                    Else
                        DebugPrint "verloren Art: " & gtarrArtikelArray(iAktRow).Titel & " (" & gtarrArtikelArray(iAktRow).Artikel & ") Hoechstbieter " & gtarrArtikelArray(iAktRow).Bieter
                    
                        'shit, ggf die cancels stornieren
                                        
                        'Alle Gebote für Artikel mit gleicher Bietgruppe zurücksetzen
                        If gtarrArtikelArray(iAktRow).Gruppe <> "" And gtarrArtikelArray(iAktRow).Status = [asOK] Then
                        
                            'hier erst umsetzen, lg 29.05.03
                            gtarrArtikelArray(iAktRow).Status = [asUeberboten]
                            
                            'Counter und Gruppe auseinanderbasteln
                            sGroup = GetGruppe(gtarrArtikelArray(iAktRow).Gruppe)
                            iCnt = GetAnzahlVonGruppe(gtarrArtikelArray(iAktRow).Gruppe)
                  
                            For i = 1 To giAktAnzArtikel
                                'Gruppe mit Counter = 1
                                If iCnt = 1 Then
                                    If sGroup = GetGruppe(gtarrArtikelArray(i).Gruppe) _
                                    And (gtarrArtikelArray(i).Status = [asCancelGroup] Or gtarrArtikelArray(i).Status = [asBuyOnlyCanceled]) _
                                    And Not i = iAktRow Then
                            
                                        DebugPrint "Reset Cancel:  Gruppe " & sGroup & " count " & iCnt & " Art: " & gtarrArtikelArray(i).Titel & "(" & gtarrArtikelArray(iAktRow).Artikel & ")"
                                
                                        If gtarrArtikelArray(i).Status = [asBuyOnlyCanceled] Then
                                          gtarrArtikelArray(i).Status = [asBuyOnly] 'noch nix passiert ;-)
                                        Else
                                          gtarrArtikelArray(i).Status = [asNixLos] 'noch nix passiert ;-)
                                        End If
                                        gtarrArtikelArray(i).Gruppe = gtarrArtikelArray(iAktRow).Gruppe
                                        gtarrArtikelArray(i).LastChangedId = GetChangeID()
                                    End If 'Gruppe passt
                                Else
                                'Gruppe mit Counter > 1
                                    If sGroup = GetGruppe(gtarrArtikelArray(i).Gruppe) _
                                    And (gtarrArtikelArray(i).Status <= [asNixLos] Or gtarrArtikelArray(i).Status = [asBuyOnly]) Then
                                        gtarrArtikelArray(i).Gruppe = sGroup & ";" & CStr(iCnt)
                                        gtarrArtikelArray(i).LastChangedId = GetChangeID()
                                    End If
                                End If
                            Next i
                            CheckBietgruppe gtarrArtikelArray(iAktRow).Gruppe
                            'lokale Bietgruppe löschen, da es nicht geklappt hat
                            gtarrArtikelArray(iAktRow).Gruppe = ""
                        Else
                          'und hier auch umsetzen, lg 29.05.03
                          If gtarrArtikelArray(iAktRow).Status = [asOK] Then gtarrArtikelArray(iAktRow).Status = [asUeberboten]
                        End If 'Gruppe nicht leer
                        
                        If gbSendAuctionEndNoSuccess Then
                          sMailText = Replace(sMailText, "[status]", gsarrLangTxt(422)) & gsarrLangTxt(427)  'ueberboten mit
                        Else
                          sMailText = ""
                        End If
                    
                    End If 'höchstbieter
                   
                Else 'kein Gebot => kein Mail
                    sMailText = ""
                End If 'Gebot>""
                gtarrArtikelArray(iAktRow).LastChangedId = GetChangeID()
                ArtikelArrayToScreen VScroll1.Value
                CheckSofortkaufArtikel
            End If 'Artikel abgelaufen + 10 sec
          End If 'noch kein Update gemacht
    
          If sMailText <> "" Then
            sMailText = Replace(sMailText, "\n", vbCrLf)
            sMailText = Replace(sMailText, vbCrLf, "")
            sMailText = sMailText & vbCrLf & vbCrLf & gsarrLangTxt(430) & vbCrLf
            sMailText = Replace(sMailText, "[url]", "http://" & gsScript4 & gsScriptCommand4 & gsCmdViewItem)
            sMailText = Replace(sMailText, "[Item]", gtarrArtikelArray(iAktRow).Artikel)
            sMailText = Replace(sMailText, "[item]", gtarrArtikelArray(iAktRow).Artikel)
            sMailText = Replace(sMailText, "[price]", Format(gtarrArtikelArray(iAktRow).AktPreis, "###,##0.00"))
            sMailText = Replace(sMailText, "[currency]", gtarrArtikelArray(iAktRow).WE)
            sMailText = Replace(sMailText, "[title]", gtarrArtikelArray(iAktRow).Titel)
            sMailText = Replace(sMailText, "[group]", gtarrArtikelArray(iAktRow).Gruppe)
            sMailText = Replace(sMailText, "[comment]", gtarrArtikelArray(iAktRow).Kommentar)
            sMailText = Replace(sMailText, "[bid]", Format(gtarrArtikelArray(iAktRow).Gebot, "###,##0.00"))
            sMailText = Replace(sMailText, "[endtime]", Date2Str(gtarrArtikelArray(iAktRow).EndeZeit))
            sMailText = Replace(sMailText, "[highbidder]", gtarrArtikelArray(iAktRow).Bieter)
            sMailText = Replace(sMailText, "[bidcount]", CStr(gtarrArtikelArray(iAktRow).AnzGebote))
            sMailText = Replace(sMailText, "[minbid]", Format(gtarrArtikelArray(iAktRow).MinGebot, "###,##0.00"))
            sMailText = Replace(sMailText, "\n", vbCrLf)
            InsertMailBuff sMailText
            sMailText = ""
          End If
          
        Next iAktRow
    End If 'Refresh nach dem Bieten
    
    If fDz > myTimeSerial(0, giWakeOnAuction, 30) And fDz < myTimeSerial(0, giWakeOnAuction + 1, 0) Then
        gbWarSchonWach = False
    End If
    If fDz > myTimeSerial(0, giWakeOnAuction, 0) And fDz < myTimeSerial(0, giWakeOnAuction, 30) Then
        gbWarSchonWach = True ' Wenn wir in dieser Zeit wach sind, waren wir schon vor unserem Timer wach
    End If
    If fDz < myTimeSerial(0, giWakeOnAuction + 1, 0) Then
        bDoNotFallAsleep = True 'lohnt nicht mehr schlafen zu gehen
    End If
    If fDz < myTimeSerial(0, giPreventSuspend, 0) Then
        bDoNotFallAsleep = True 'lohnt nicht mehr schlafen zu gehen
    End If
      
    If fDz < myTimeSerial(0, 5, 0) And gbBeepBeforeAuction Then
        If fDz > myTimeSerial(0, 4, 55) Then
            Beep
        End If
    End If
    
    If mbWaitForFirstPop Then
      mbWaitForFirstPop = False
      bDoNotHangUp = True ' wir wollen nochmal den waszutun-Check durchlaufen
      miPopTimerCount = giPopZyklus
      Call POPTimer_Timer
    End If
    
    If gbUsesModem Then
        'mal sehen ob wir on- oder Offline  müssen ..
    
        fModemTime = myTimeSerial(0, 1, 0) * glVorlaufModem
        
        If fModemTime >= fDz Then
            'CheckInternetConnection wird von ModemConnec aufgrufen
            bOnlineTmp = IsOnline
            ModemConnect Me.hWnd
            
            If IsOnline And Not bOnlineTmp And gbTestConnect Then
                'Provider- Umlenkung aushebeln:
                bDummy = Check_ebayUp
            End If
        End If
        If IsOnline And fModemTime < fDz And (Not bDoNotHangUp Or Not bDoNotFallAsleep) Then
            Mailbufftimer_Timer 'erst alle Mails rausschicken! lg 10.07.2003
            If Not gbKeepDialupAlive Then
              If gbLastDialupWasManually Then
                Ask_Offline
              Else
                ModemHangUp
              End If
            End If
        End If
    
    End If
    
    If giVorlaufLan > 0 Then
        'mal sehen ob wir on- oder Offline  müssen ..
    
        fModemTime = myTimeSerial(0, 1, 0) * giVorlaufLan
        
        If fModemTime >= fDz Then
            If Not mbIsOn Then
                bOk = Check_ebayUp
                mbIsOn = True
            End If
        End If
        
        If mbIsOn And fModemTime < fDz Then
            mbIsOn = False
        End If
    End If
    
    'x Minuten vor Auktion automatisch einloggen:
    If giReLogin > 0 Then
        fModemTime = myTimeSerial(0, 1, 0) * giReLogin
        If Not gbAutoLogged And fModemTime >= fDz Then
            LogIn gsNextUser, gtarrUserArray(UsrAccToIndex(gsNextUser)).UaPass, gtarrUserArray(UsrAccToIndex(gsNextUser)).UaToken
            gbAutoLogged = True ' Endlos-Einloggen verhindern
        End If
        If gbAutoLogged And fModemTime < fDz Then
            gbAutoLogged = False
        End If
    End If
    
    If (giUseTimeSync And 2) > 0 Then
        'ggf. 2 Min vorher ein Timesync
        fModemTime = myTimeSerial(0, 2, 0)
        
        If fModemTime >= fDz And Not mbIsSync Then
            Zeitsync
            mbIsSync = True
        End If
        
        If mbIsSync And fModemTime < fDz Then
            mbIsSync = False
        End If
    End If
    
    If gbUsesOdbc Then
        'ggf. 2 Min vorher ODBC abschalten
        fModemTime = myTimeSerial(0, 2, 0)
        If fModemTime >= fDz Then
            If ODBC_Timer.Enabled Then
                ODBC_Timer.Enabled = False
                gsOdbcStopRead = True
            End If
        Else
           If Not ODBC_Timer.Enabled Then
                ODBC_Timer.Enabled = True
                gsOdbcStopRead = False
           End If
        End If
    End If
    
    If glSendCsvInterval > 0 Then
        If fDz > myTimeSerial(0, 2, 0) Then
            If fLastSendCsv = 0 Then fLastSendCsv = MyNow()
            If fLastSendCsv + myTimeSerial(0, glSendCsvInterval, 0) < MyNow() Then
                fLastSendCsv = MyNow()
              
                sSubject = "b-o-m: readcsv"
                sMailText = BuildArtikelCSV2()
                If gbPopSendEncryptedAcknowledgment Then
                    Set oRC4 = New clsRC4
                    sSubject = Replace(oRC4.EncryptString(sSubject, gsPass, True), vbCrLf, "")
                    sMailText = oRC4.EncryptString(sMailText, gsPass, True)
                    Set oRC4 = Nothing
                End If
                Call SendSMTP(gsSendEndFromRealname & "<" & gsSendEndFrom & ">", gsSendCsvTo, "subject: " & sSubject & vbCrLf & vbCrLf & sMailText)
                tcpIn.Close
            End If
        End If
    End If
    
    If Not bWasZuTun And gbUseWinShutdown And Not bDoNotHangUp Then
        Call Mailbufftimer_Timer 'erst alle Mails rausschicken!
          bOk = ShowUpdateBox(Me, [ftCountDown1], 103, IIf(Me.WindowState = vbMinimized Or Me.Visible = False, 2, 1), "ShutDown", 60, gsarrLangTxt(208), gsarrLangTxt(209), gsarrLangTxt(210), "-", gsarrLangTxt(359), True)
          If bOk Then
            Call DebugPrint("Nichts mehr zu tun -> Herunterfahren")
            mbQuietExit = True
            Call ShutDownWin
            Call mnuExit_Click
        Else
            gbUseWinShutdown = False
        End If
        gbBeendenNachAuktionAktiv = False
    End If
    
    If Not bDoNotFallAsleep And Not gbWarSchonWach And gbSuspendNachAuktionAktiv Then
        gbSuspendNachAuktionAktiv = False
        Call Mailbufftimer_Timer 'erst alle Mails rausschicken!
        Call Resuspend(Me)
    End If
    
    If Not bWasZuTun And gbBeendenNachAuktion And Not bDoNotHangUp And gbBeendenNachAuktionAktiv Then
        Call Mailbufftimer_Timer 'erst alle Mails rausschicken!
        '...
        bOk = ShowUpdateBox(Me, [ftCountDown2], 103, IIf(Me.WindowState = vbMinimized Or Me.Visible = False, 2, 1), gsarrLangTxt(354), 10, gsarrLangTxt(355), gsarrLangTxt(356), gsarrLangTxt(357), gsarrLangTxt(358), gsarrLangTxt(359), True)
        If bOk Then
            Call DebugPrint("Nichts mehr zu tun -> Programmende")
            gbExplicitEnd = True
            mbQuietExit = True
            Call mnuExit_Click
        End If
        gbBeendenNachAuktionAktiv = False
    End If
    
    fLaufzeit = GetSystemUptime() - fStartZeitpunkt
    
    If fLaufzeit > 0.99 Then fLaufzeit = 0.99 'Fix Timer auf 0 bei > 0.999 && < 1
    If fLaufzeit <= 0 Then fLaufzeit = 0
    Timer1.Interval = 1000 - (fLaufzeit * 1000)
    Timer1.Enabled = gbAutoMode
End If 'giSuspendState = 0
End Sub

Private Function AVAx() As Integer
Dim i As Integer
Dim iValide As Integer
Dim iAvaTmp As Integer
    
    Select Case giArtAktOptions
        Case 0 'Alle Artikel
            AVAx = AnzValidArtikel(True)
        Case 1 'die nächsten x Artikel
            
            iValide = AnzValidArtikel(True)   'Anzahl gültige
            If giArtAktOptionsValue > iValide Then
                AVAx = iValide
            Else
                AVAx = giArtAktOptionsValue
            End If
        
        Case 2 'Alle innerhalb x min minimum 5
            i = 1
            Do
                If GetRestzeitFromItem(i) > 0 And _
                    GetRestzeitFromItem(i) <= myTimeSerial(0, giArtAktOptionsValue, 0) And _
                    gtarrArtikelArray(i).Status >= [asNixLos] Then
                    iAvaTmp = iAvaTmp + 1
                End If
                i = i + 1
            Loop Until i = giAktAnzArtikel
            AVAx = iAvaTmp
    End Select
  
End Function

Private Sub ArtikelTimer_Timer()
    
    On Error GoTo errhdl
    
    Dim i As Integer
    Dim j As Integer
    Dim fDz As Double
    Dim iAktRow As Integer
    Dim iShowIt As Integer
    Dim iAva As Integer
    Dim iAnzX As Integer
    Dim iAktStart As Integer
    
    If giSuspendState = 0 And mbIsLoggingIn = False Then
        'Update der Artikelinfo
        ArtikelTimer.Enabled = False
        
        If IsOnline Then
            
            If gbAktualisierenXvor Then
                iAva = AVAx
            End If
            
            fDz = gfRestzeitZaehler 'und bitte 3 Min vor Gebot aufhören .. 'sh warum 3 min?
            
            If fDz > myTimeSerial(0, 0, glVorlaufGebot + gfLzMittel + 5) And Not mbAktPaused Then 'wenn restzeit > Vorlauf + dauer akt. 1 Artikel + Sicherheit
                If (giUseTimeSync And 8) > 0 Then
                    glTimeSyncCounter = glTimeSyncCounter + 1
                End If
                
                'Nachts um 2 ggf. ein Time- Update
                If (((giUseTimeSync And 1) > 0) And (Time > myTimeSerial(2, 5, 0) And Time < myTimeSerial(2, 7, 0))) Or _
                    (((giUseTimeSync And 8) > 0) And (glTimeSyncCounter / 60 >= glTimeSyncIntervall)) Then
                    
                    glTimeSyncCounter = 0
                    Call Zeitsync
                End If
                
                '** wir haben noch zeit ..
                If gbGeboteAktualisieren Then
                    miArtikelCycleCount = miArtikelCycleCount + 1
                    If gbAktualisierenXvor And fDz >= myTimeSerial(0, giAktXminvor, gfLzMittel * iAva) Then
                        If miArtikelCycleCount >= (giAktXminvorCycle * 60) Then '### neu artikelauswahl
                            miArtikelCycleCount = 0
                            If iAva > 0 Then
                                iAktStart = AnzValidArtikel(False)
                                If gsSortOrder = "asc" Then
                                    iAnzX = iAktStart + (iAva)
                                Else
                                    iAnzX = iAktStart - (iAva)
                                End If
                                i = iAktStart
                                j = 1
                                Do
                                    SortEnde.Enabled = False
                                    If (Not gtarrArtikelArray(i).Artikel = "" _
                                        And (gtarrArtikelArray(i).Status <= [asNixLos] Or gtarrArtikelArray(i).Status = [asBuyOnly])) Then
                                        
                                        Call PanelText(StatusBar1, 2, gsarrLangTxt(34) & " " & gsarrLangTxt(31) & " " & CStr(i) & " / " & CStr(giAktAnzArtikel))
                                        Call Upd_Art(i, vbNullString, False)
                                        
                                        If i - VScroll1.Value <= giMaxRow Then
                                            Call ArtikelArrayToScreen(i, , True)
                                        End If
                                        
                                    End If
                                    If gsSortOrder = "asc" Then
                                        i = i + 1
                                    Else
                                        i = i - 1
                                    End If
                                    j = j + 1
                                Loop Until i = iAnzX Or fDz < myTimeSerial(0, giAktXminvor, 0)
                                SortEnde.Enabled = True
                                Call PanelText(StatusBar1, 2, "")
                            Else
                                GoTo errhdl
                            End If
                            GoTo errhdl
                        Else
                            Call PanelText(StatusBar1, 2, gsarrLangTxt(740) & " " & _
                                Format(CDate(myTimeSerial(0, giAktXminvorCycle, 0) - myTimeSerial(0, 0, miArtikelCycleCount)), "hh:mm:ss"), True)
                                GoTo errhdl
                        End If
                        GoTo errhdl
                    Else 'gbAktualisierenXvor And fDz >= myTimeSerial
                        If miArtikelCycleCount >= giArtikelRefreshCycle Then
                            miArtikelCycleCount = 0
                            If giAktualisierenOpt = 1 Then
                                iAktStart = AnzValidArtikel(False)
                                SortEnde.Enabled = False
                                Call Upd_Art(iAktStart)
                                
                                If iAktStart - VScroll1.Value <= giMaxRow Then
                                    Call ArtikelArrayToScreen(iAktStart)
                                End If
                                
                                SortEnde.Enabled = True
                                GoTo errhdl
                            ElseIf giAktualisierenOpt = 0 Then 'gbAutoAktualisierenNext
                                'update auf den akt. Wert
                                '1.7.5 hjs
                                If miRowCount <= 0 Then miRowCount = 1
                                If miRowCount > giAktAnzArtikel Then
                                    miRowCount = 1
                                End If
                                
                                iAktRow = miRowCount
                                If Not gtarrArtikelArray(miRowCount).Artikel = "" _
                                    And (gtarrArtikelArray(miRowCount).Status <= [asNixLos] Or gtarrArtikelArray(miRowCount).Status = [asBuyOnly]) Then
                                    
                                    SortEnde.Enabled = False
                                    Call Upd_Art(miRowCount)
                                    SortEnde.Enabled = True
                                    iShowIt = miRowCount - VScroll1.Value
                                    If iShowIt > giMaxRow Or Me.WindowState = vbMinimized Then iShowIt = -1
                                    
                                    If iShowIt >= 0 Then
                                        Call ArtikelArrayToScreen(VScroll1.Value)
                                    End If
                                End If
                                
                                Do
                                    miRowCount = miRowCount + 1
                                    If miRowCount > giAktAnzArtikel Then
                                        miRowCount = 1
                                    End If
                                    
                                    If miRowCount = iAktRow Then
                                        Exit Do
                                    End If
                                Loop While (gtarrArtikelArray(miRowCount).Artikel = "" Or GetRestzeitFromItem(miRowCount) = 0 Or (gtarrArtikelArray(miRowCount).Status > [asNixLos] And gtarrArtikelArray(miRowCount).Status <> [asBuyOnly]))
                                GoTo errhdl
                            End If 'gbAutoAktualisierenNext
                            If Not Check_ebayUp Then
                                GoTo errhdl
                            End If
                        Else 'miArtikelCycleCount >= giArtikelRefreshCycle
                            Call PanelText(StatusBar1, 2, gsarrLangTxt(740) & " " & _
                                Format(CDate(myTimeSerial(0, 0, giArtikelRefreshCycle) - myTimeSerial(0, 0, miArtikelCycleCount)), "hh:mm:ss"), True)
                            GoTo errhdl
                        End If 'miArtikelCycleCount >= giArtikelRefreshCycle
                        If Not Check_ebayUp Then
                            GoTo errhdl
                        End If
                    End If 'gbAktualisierenXvor And fDz >= myTimeSerial
                    If Not Check_ebayUp Then
                        GoTo errhdl
                    End If
                End If 'gebote aktualisieren
            End If 'fDz > myTimeSerial(0, 0, glVorlaufGebot + gfLzMittel + 5)
        End If 'IsOnline
    End If 'giSuspendState = 0
errhdl:
ArtikelTimer.Enabled = gbAutoMode 'And gbGeboteAktualisieren

End Sub

Public Sub SetStatus(sState As String, Optional bErrFlag As Boolean)
  
  Dim sTmp As String
  Dim lBGColor As Long
  Dim bResetable As Boolean
  
  If bErrFlag Then
      sTmp = "Offline"
      lBGColor = &H8000000F
      bResetable = True
  Else
      sTmp = "Online"
      lBGColor = vbGreen
  End If
  
  If gbEBayWartung Then
      sTmp = gsarrLangTxt(22)
      lBGColor = vbRed
  Else
      If sState = "Idle" Then
        sTmp = sState
      Else
        sTmp = sState & " -- " & sTmp
      End If
  End If
  
  PanelText StatusBar1, 1, sTmp, bResetable, lBGColor

End Sub

Private Sub tcpIn_Close()
    
    gbSessionClosed = True
    
End Sub

Private Sub tcpIn_DataArrival(ByVal bytesTotal As Long)
    
    Dim strData As String
    
    On Error Resume Next
    
    tcpIn.GetData strData, vbString
    gsWholeThing = gsWholeThing & strData
    gsThisChunk = strData                    'For testing content
    gsResponseState = Left$(gsWholeThing, 3)   ' +OK and +ER tests
    gsDotLine = Right$(gsWholeThing, 5)        ' EOM tests
    giSmtpResponse = Val(gsResponseState)     ' SmtP
    
End Sub

Private Sub tcpIn_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    gsStatusTxt = gsStatusTxt & "***POPManager: TCP/IP Error " & CStr(Number) & ": " & Description
    gbFatalError = True
   
End Sub


Private Function Update_Artikel(ByVal iAktRow As Integer, Optional sResultPage As String = "", Optional bWait As Boolean = True) As String
Dim sServer As String
Dim sUser As String
Dim sKommando As String
Dim sBuffer As String
Dim fMinGebot As Double
Dim iAnzGebote As Integer
Dim datEndeZeit As Date
Dim sTxt As String
Dim sWe As String
Dim bFlag As Boolean
Dim sTmp As String
Dim sTmp2 As String
Dim sMailText As String
Dim iAnzUpdateVersuche As Integer
Dim sItem As String
Const iMaxUpdateVersuche As Integer = 3

    sItem = gtarrArtikelArray(iAktRow).Artikel ' Artikel-Nr merken
    
    If gbUpdateAnonymous Then
      sUser = "anonymous"
    Else
      sUser = gtarrArtikelArray(iAktRow).UserAccount
    End If

    If sResultPage = "" Then
   
        sServer = "https://" & gsScript4 & gsScriptCommand4
    
        sBuffer = ""
    
        sTmp = gsCmdViewItem
        sTmp = Replace(sTmp, "[Item]", sItem)
        
        sKommando = sBuffer & sTmp
        
        If Not gbUseCurl Then bWait = True
        If Not gbConcurrentUpdates Then bWait = True
        sBuffer = ShortPost(sServer & sKommando, , , sUser, bWait)
        If Not bWait Then DoEvents: Exit Function
        
    Else
    
      sBuffer = sResultPage
    
    End If
    
    Do While InStr(1, sBuffer, gsAnsSwitchToAnonymous) > 0 And iAnzUpdateVersuche < iMaxUpdateVersuche
      iAnzUpdateVersuche = iAnzUpdateVersuche + 1
      gbUpdateAnonymous = True
      sServer = "https://" & gsScript4 & gsScriptCommand4
      sBuffer = ""
      sTmp = gsCmdViewItem
      sTmp = Replace(sTmp, "[Item]", sItem)
      sKommando = sBuffer & sTmp
      sBuffer = ShortPost(sServer & sKommando, , , "anonymous", True)
    Loop
    
    Call KeyAutoSwitch(sBuffer)
    
    iAktRow = ItemToIndex(sItem) ' und den Index neu berechnen falls der User zwischenzeitlich an der Liste rumgefummelt hat!
    If iAktRow = 0 Then Exit Function ' uups, der Artikel wurde gelöscht, raus hier!
    
    If Check_Wartung(sBuffer) Or Len(sBuffer) < 100 Then
        Call KeyAutoSwitch
        gtarrArtikelArray(iAktRow).UpdateInProgressSince = 0
        Exit Function
    End If
    
    Update_Artikel = sBuffer
    
    On Error Resume Next
    bFlag = False
    
    sTxt = sucheEnde(sBuffer, iAktRow, datEndeZeit)
    
    If sTxt = "Fehler" Then

        Call DebugPrint("Fehler bei Artikelupdate, versuche Login: ")

        sTmp = LogIn2(sBuffer)
        sTmp2 = sucheEnde(sTmp, iAktRow, datEndeZeit)

        If sTmp2 <> "Fehler" Then
          sBuffer = sTmp
          sTxt = sTmp2
          Call DebugPrint("OK")
        Else
          Call DebugPrint("Hat leider nicht geholfen.")
        End If

    End If
    
    If datEndeZeit = gdatENDEZEITNOTFOUND And sucheTitel(sBuffer) = "" Then sTxt = "Invalid"
    
    If sTxt = "Invalid" Then 'Ungültiger Artikel
        With gtarrArtikelArray(iAktRow)
             If GetRestzeitFromItem(iAktRow) > 0 Then 'vorzeitig beendet?
                If .NotFound >= 0 And .NotFound < giReloadTimes Then
                  .NotFound = .NotFound + 1
                Else 'Artikel wurde einmal zu oft nicht gefunden
                  .NotFound = -1
                  .Status = [asNotFound]
                End If
             Else 'Artikel ist wirklich nicht mehr existent
                If .Titel = "" Then .Titel = gsarrLangTxt(253)
                '.Titel = IIf(InStr(1, .Titel, gsarrLangTxt(253), vbTextCompare) > 0, .Titel, gsarrLangTxt(253) & vbNewLine & .Titel)
                .NotFound = -1
                .Status = [asNotFound]
             End If
        End With
    ElseIf sTxt = "Fehler" Then 'Fehler beim Zugriff auf den Artikel
        'gtarrArtikelArray(iAktRow).Titel = IIf(InStr(1, gtarrArtikelArray(iAktRow).Titel, gsarrLangTxt(254), vbTextCompare) > 0, gtarrArtikelArray(iAktRow).Titel, gsarrLangTxt(254) & ": --> " & gtarrArtikelArray(iAktRow).Titel)
        gtarrArtikelArray(iAktRow).Status = [asAccessErr]
    Else
        If sTxt = "Beendet" And gtarrArtikelArray(iAktRow).Status <= [asNixLos] Then
            
            gtarrArtikelArray(iAktRow).Status = [asEnde]
                
        ElseIf sTxt = "Ok" Then
            If (gtarrArtikelArray(iAktRow).Status = [asEnde] _
              Or gtarrArtikelArray(iAktRow).Status = [asAccessErr] _
              Or gtarrArtikelArray(iAktRow).Status = [asNotFound]) Then
        
                gtarrArtikelArray(iAktRow).Status = [asNixLos]
            End If
            gtarrArtikelArray(iAktRow).NotFound = 0
        
        End If
        
        gtarrArtikelArray(iAktRow).EndeZeit = datEndeZeit
        gtarrArtikelArray(iAktRow).TimeZone = GetUTCOffset
    
        gtarrArtikelArray(iAktRow).Titel = sucheTitel(sBuffer)
        gtarrArtikelArray(iAktRow).AktPreis = EbayString2Float(sucheAktGebot(sBuffer, sWe, bFlag))
        gtarrArtikelArray(iAktRow).WE = sWe
        gtarrArtikelArray(iAktRow).Bewertung = sucheBewertung(sBuffer)
        gtarrArtikelArray(iAktRow).Standort = sucheStandort(sBuffer)
        gtarrArtikelArray(iAktRow).MindestpreisNichtErreicht = InStr(1, sBuffer, gsAnsBuyerReserve)
        gtarrArtikelArray(iAktRow).Ueberarbeitet = InStr(1, sBuffer, gsAnsRevised)
        gtarrArtikelArray(iAktRow).Verkaeufer = Trim(sucheVK(sBuffer))
        'Versandkosten nur von eBay überschreiben falls noch nicht manuell gesetzt!
        If Left(gtarrArtikelArray(iAktRow).Versand, 1) <> "*" Then
          sTmp = Trim(sucheVersand(sBuffer, sWe))
          sTmp2 = Trim(Format(EbayString2Float(sTmp), "###,##0.00"))
          If FilterNumeric(sTmp) = FilterNumeric(sTmp2) Then ' numerische Versandkosten
            If sWe <> gtarrArtikelArray(iAktRow).WE Then sTmp2 = sTmp2 & " " & sWe
            gtarrArtikelArray(iAktRow).Versand = sTmp2
          Else
            gtarrArtikelArray(iAktRow).Versand = Trim(sucheVersand(sBuffer, sWe) & " " & sWe)
          End If
        End If
        If gtarrArtikelArray(iAktRow).Versand Like "[?]*" Then gtarrArtikelArray(iAktRow).Versand = "?"
        If bFlag Then 'Nur Sofortkaufen
            
            If gtarrArtikelArray(iAktRow).Status <= [asNixLos] Or _
               gtarrArtikelArray(iAktRow).Status = [asEnde] Then gtarrArtikelArray(iAktRow).Status = [asBuyOnly]
               
            gtarrArtikelArray(iAktRow).Bieter = sucheMenge(sBuffer)
            If Val(gtarrArtikelArray(iAktRow).Bieter) < 0 Then gtarrArtikelArray(iAktRow).Bieter = "1"
            gtarrArtikelArray(iAktRow).Bieter = gtarrArtikelArray(iAktRow).Bieter & " " & gsarrLangTxt(31)
            
        Else 'nicht nur Sofortkaufen
            
            'Anzahl Gebote und Bieter
            iAnzGebote = sucheAnzGebote(sBuffer)
            gtarrArtikelArray(iAktRow).AnzGebote = iAnzGebote
            gtarrArtikelArray(iAktRow).Bieter = Trim(sucheBieter(sBuffer, iAnzGebote, gtarrArtikelArray(iAktRow).Bieter))
            
            'evtl. Powerauktion ???
            If gtarrArtikelArray(iAktRow).Status <= 0 Or gtarrArtikelArray(iAktRow).Status = [asBuyOnly] Then
                If InStr(1, sBuffer, gsAnsDutch) > 0 Then 'Powerauktion
                    gtarrArtikelArray(iAktRow).Status = [asPower]
                ElseIf InStr(1, sBuffer, gsAnsAdvertisement) > 0 Then 'Preisanzeige
                    gtarrArtikelArray(iAktRow).Status = [asAdvertisement]
                Else '=> Normal
                    gtarrArtikelArray(iAktRow).Status = [asNixLos]
                End If
            End If
           
            'ca.-Preis weil unbekannte Währung ?
            If gtarrArtikelArray(iAktRow).Status = [asSellerAway] Or gtarrArtikelArray(iAktRow).Status = [asNotFound] Then 'Left(sPreis, Len(ansApproxBid)) = ansApproxBid Then
                
'                    sPreis = Mid(sPreis, Len(ansApproxBid) + 2, Len(sPreis) - Len(ansApproxBid) - 1)
          
            Else 'Währung bekannt
            
                fMinGebot = EbayString2Float(sucheMinGebot(sBuffer, sWe))
                If fMinGebot = 0 Then fMinGebot = gtarrArtikelArray(iAktRow).AktPreis
                If gbSendIfLow _
                  And gtarrArtikelArray(iAktRow).Gebot > 0 _
                  And gtarrArtikelArray(iAktRow).Gebot < fMinGebot _
                  And gtarrArtikelArray(iAktRow).Gebot >= gtarrArtikelArray(iAktRow).MinGebot Then
                    sMailText = "Subject: " & gsarrLangTxt(790)
                End If
                gtarrArtikelArray(iAktRow).MinGebot = fMinGebot
            
            End If
        End If
        'Verkäufer zurzeit abwesend ???
        If (gtarrArtikelArray(iAktRow).Status <= 0 Or gtarrArtikelArray(iAktRow).Status = [asBuyOnly]) And _
           gtarrArtikelArray(iAktRow).AktPreis = 0 And InStr(1, sBuffer, gsAnsSellerAway) > 0 Then
          gtarrArtikelArray(iAktRow).Status = [asSellerAway]
        End If
    End If
    
    If sMailText <> "" Then
        sMailText = Replace(sMailText, "\n", vbCrLf)
        sMailText = Replace(sMailText, vbCrLf, "")
        sMailText = sMailText & vbCrLf & vbCrLf & gsarrLangTxt(791) & vbCrLf
        sMailText = Replace(sMailText, "[url]", "http://" & gsScript4 & gsScriptCommand4 & gsCmdViewItem)
        sMailText = Replace(sMailText, "[Item]", gtarrArtikelArray(iAktRow).Artikel)
        sMailText = Replace(sMailText, "[item]", gtarrArtikelArray(iAktRow).Artikel)
        sMailText = Replace(sMailText, "[price]", Format(gtarrArtikelArray(iAktRow).AktPreis, "###,##0.00"))
        sMailText = Replace(sMailText, "[currency]", gtarrArtikelArray(iAktRow).WE)
        sMailText = Replace(sMailText, "[title]", gtarrArtikelArray(iAktRow).Titel)
        sMailText = Replace(sMailText, "[group]", gtarrArtikelArray(iAktRow).Gruppe)
        sMailText = Replace(sMailText, "[comment]", gtarrArtikelArray(iAktRow).Kommentar)
        sMailText = Replace(sMailText, "[bid]", Format(gtarrArtikelArray(iAktRow).Gebot, "###,##0.00"))
        sMailText = Replace(sMailText, "[endtime]", Date2Str(gtarrArtikelArray(iAktRow).EndeZeit))
        sMailText = Replace(sMailText, "[highbidder]", gtarrArtikelArray(iAktRow).Bieter)
        sMailText = Replace(sMailText, "[bidcount]", CStr(gtarrArtikelArray(iAktRow).AnzGebote))
        sMailText = Replace(sMailText, "[minbid]", Format(gtarrArtikelArray(iAktRow).MinGebot, "###,##0.00"))
        sMailText = Replace(sMailText, "\n", vbCrLf)
        InsertMailBuff sMailText
    End If

    Call KeyAutoSwitch
    gtarrArtikelArray(iAktRow).UpdateInProgressSince = 0
    
    If (gtarrArtikelArray(iAktRow).Status > [asOK] _
        Or (gtarrArtikelArray(iAktRow).Status > 0 And Not (gbArtikelRefreshPost Or gbArtikelRefreshPost2))) _
        And GetRestzeitFromItem(iAktRow) = 0 Then
       gtarrArtikelArray(iAktRow).PostUpdateDone = True
    End If
    gtarrArtikelArray(iAktRow).LastChangedId = GetChangeID()

End Function

Private Function Check_ebayUp() As Boolean

On Error GoTo errhdl

Dim i As Integer
Dim sServer As String
Dim sKommando As String
Dim sTmp As String

sServer = "https://" & gsScript2 & gsScriptCommand2
sKommando = gsCmdTimeShow

Do While sTmp = "" And i < 10
    sTmp = ShortPost(sServer & sKommando)
    i = i + 1
Loop

If sTmp = "" Or Check_Wartung(sTmp) Then
    Check_ebayUp = False
Else
    Check_ebayUp = True
End If

Exit Function

errhdl:

Check_ebayUp = False

End Function

Public Function sync_ebaytime() As String

On Error GoTo errhdl

Dim sServer As String
Dim sKommando As String
Dim sTmp As String
Dim sDatum As String
Dim datDatum As Date
Dim lPos As Long
Dim fLap As Double
Dim fLapTime As Double
Dim lOffsetLocal  As Long
Dim fOffsetLocalKorrektur As Double

    'wir messen die Laufzeit
    fLap = Timer
    
    sServer = "https://" & gsScript2 & gsScriptCommand2
    sKommando = gsCmdTimeShow
    sTmp = ShortPost(sServer & sKommando)
    
    If sTmp = "" Or Check_Wartung(sTmp) Then GoTo errhdl
    
    'die fLapTime ist allerdings nur Sekundengenau, also Aufrunden
    
    fLap = Timer - fLap + 1
    fLapTime = myTimeSerial(0, 0, 1) * (fLap / 2)
    
    'Zeit lesen
    lPos = InStr(1, sTmp, gsAnsTime2_1)
    If lPos > 0 Then lOffsetLocal = gsAnsOffsetLocal2_1
    
    If lPos = 0 Then
        lPos = InStr(1, sTmp, gsAnsTime2_2)
        If lPos > 0 Then lOffsetLocal = gsAnsOffsetLocal2_2
    End If
    If lPos > 0 Then
        sDatum = Mid(sTmp, lPos - 50, 50)
        sDatum = ConvertMonthname2(sDatum)
        
        datDatum = DateAdd("n", 60 * (GetUTCOffset() - lOffsetLocal), Str2Date(sDatum, gsCmdTimeShowFormat))
        
        If datDatum > 0 And gsAnsTime2_1 = gsAnsTime2_2 Then
          fOffsetLocalKorrektur = -GetOffsetLocalFromDate(datDatum)
          datDatum = DateAdd("h", fOffsetLocalKorrektur, datDatum)
        End If
        
        If datDatum > myDateSerial(2003, 1, 1) Then
            On Error Resume Next
            Date = DateValue(datDatum)
            Time = TimeValue(datDatum) + fLapTime
            If Err.Number = 0 Then
              gfTimeDeviation = 0
            Else
              gfTimeDeviation = fLapTime
              DebugPrint "Systemzeit konnte nicht geändert werden.", 2
            End If
            On Error GoTo errhdl
            sync_ebaytime = Date2Str(MyNow)
            mbEBayTimeIsSync = True
        End If
    Else
        GoTo errhdl
    End If
    Exit Function
errhdl:
End Function

Public Sub Ask_Online()

If Not CheckInternetConnection Then
  If MsgBoxEx(gsarrLangTxt(20) & gsConnectName, gsarrLangTxt(408) & "*-" & gsarrLangTxt(409) & " [" & CStr(giDialupRequestTimeout) & "]-") = 1 Then
    ModemConnect Me.hWnd
    If IsOnline Then gbLastDialupWasManually = True: gbKeepDialupAlive = False
  End If
End If
End Sub
Public Sub Ask_Offline()
gbKeepDialupAlive = False
gbLastDialupWasManually = False
If MsgBoxEx(gsarrLangTxt(21), gsarrLangTxt(408) & " [" & CStr(giDialupRequestTimeout) & "]-" & gsarrLangTxt(409) & "-") = 1 Then
    ModemHangUp
Else
    gbKeepDialupAlive = True
End If
End Sub

Private Sub Titel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    Dim sTmp As String

    Call SetFocusRect(Index)
    Call VersandkostenUebernehmen
    Call Gebot_LostFocus(giLastGebotEditedIndex)
    
    If Button = vbLeftButton Then '1
        
        If Not Artikel(Index) = "" Then
            If Not CheckInternetConnection Then Call Ask_Online
            
            If IsOnline Then
                If gbOpenBrowserOnClick Then
                    
                    sTmp = gsCmdViewItem
                    sTmp = Replace(sTmp, "[Item]", Artikel(Index).Text)
                    
                    DoEvents
                    gsGlobalUrl = "http://" & gsScript4 & gsScriptCommand4 & sTmp
                    
                    'Call ShowBrowser(Me.hWnd)
                    Call ExecuteDoc(Me.hWnd, gsGlobalUrl)
                End If
                
                Call Preis_MouseDown(Index, 1, 0, 0, 0)
                
                If gbUsesModem And gbLastDialupWasManually Then Call Ask_Offline
            End If
        End If
    ElseIf Button = vbRightButton Then
        Call ShowContextMenu(Index)
    End If
    
End Sub

Private Sub Titel_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Index für ArtikelArray übergeben KOM 3.9.03
    Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)
    
End Sub

Private Sub FromTaskbar()

    On Error Resume Next

    App.TaskVisible = True
    Me.WindowState = mlPrevWindowState
    Me.Show
    
End Sub

Private Sub ToTaskbar()

    On Error Resume Next
    
    Me.Hide
    App.TaskVisible = False

End Sub

Private Sub updTaskbar(sTxt As String, Optional bForceUpdate As Boolean = False)

On Error Resume Next

If mtTrayIcon.hWnd Then
    mtTrayIcon.szTip = sTxt & Chr$(0)
    mtTrayIcon.hIcon = frmDummy.Picture1(0)
    If bForceUpdate Then Call Shell_NotifyIcon(NIM_MODIFY, mtTrayIcon)
End If

End Sub

Private Sub updTaskbarIcon()

On Error Resume Next

If mtTrayIcon.hWnd Then
    mtTrayIcon.hIcon = frmDummy.Picture1(0)
    Call Shell_NotifyIcon(NIM_MODIFY, mtTrayIcon)
End If

End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    
    Dim Msg As Long
      
    Msg = X / Screen.TwipsPerPixelX
    Select Case Msg
        Case WM_LBUTTONUP
            If Not gbAboutIsUp Then
              If frmHaupt.WindowState = vbMinimized Then
                If mbIsUnloadingZeilen Or mbIsMinimizing Then
                  mbRestoreQueued = True
                Else
                  mbIsRestoring = True
                  Call FromTaskbar
                  mbIsRestoring = False
                End If
              Else
                If mbIsRestoring Then
                  mbMinimizeQueued = True
                Else
                  mbIsMinimizing = True
                  frmHaupt.WindowState = vbMinimized
                  mbIsMinimizing = False
                End If
              End If
            End If
        Case WM_RBUTTONUP
            Call SetForegroundWindow(Me.Picture1.hWnd)  'lg 19.05.03
            Call ShowPopupMenu_TrayIcon
        Case Else
    End Select
    
    If mtTrayIcon.hWnd Then
        Call Shell_NotifyIcon(NIM_MODIFY, mtTrayIcon)
    End If
    
End Sub

Public Sub ArtikelArrayToScreen(iPosStart As Integer, Optional bNurRestzeit As Boolean = False, _
                                 Optional bNurDieseZeile As Boolean = False)

On Error GoTo errhdl

Dim i As Integer
Dim iCnt As Integer
Dim iMaxCnt As Integer
Dim sTmp As String
Dim sUser As String
Dim sVersand As String
Dim m As Integer
Dim n As Integer
Dim mm As Long
Dim nn As Long

If Me.WindowState <> vbMinimized Then
    
    If Not bNurDieseZeile Then
        If iPosStart < 1 Then iPosStart = 1
        iCnt = 0
        iMaxCnt = UBound(gtarrArtikelArray)
        If iPosStart + giMaxRow < iMaxCnt Then iMaxCnt = iPosStart + giMaxRow
        If VScroll1.Value <> iPosStart Then
            VScroll1.Value = iPosStart
            Exit Sub
        End If
    Else
        i = iPosStart
        iMaxCnt = iPosStart
        iCnt = iPosStart - VScroll1.Value
    End If
    
    For i = iPosStart To iMaxCnt
        If Not bNurRestzeit Then
            Artikel(iCnt).Text = gtarrArtikelArray(i).Artikel
            EndeZeit(iCnt).Caption = Date2Str(gtarrArtikelArray(i).EndeZeit, gbShowWeekday, gsSpecialDateFormat)
            If gtarrArtikelArray(i).Gebot > 0 Then
                Gebot(iCnt) = Format(gtarrArtikelArray(i).Gebot, "###,##0.00")
            Else
                Gebot(iCnt) = ""
            End If
            Gebot(iCnt).BackColor = GetGebotColorFromItem(i)
            Gebot(iCnt).FontBold = True
            Gebot(iCnt).Enabled = Not gbUsesOdbc
            Titel(iCnt) = gtarrArtikelArray(i).Titel
            
            If gbRevisedInTitle And gtarrArtikelArray(i).Ueberarbeitet Then Titel(iCnt).Caption = Titel(iCnt).Caption & " (" & gsarrLangTxt(749) & ")"
            'Kommentar in der Liste ohne Zeilenumbrüche
            If gbCommentInTitle And gtarrArtikelArray(i).Kommentar > "" Then Titel(iCnt).Caption = Titel(iCnt).Caption & " [" & Replace(gtarrArtikelArray(i).Kommentar, vbCrLf, " ", 1, -1, vbBinaryCompare) & "]"
            Call TitelKuerzen(Titel(iCnt))
            Preis(iCnt).Caption = Format(gtarrArtikelArray(i).AktPreis, "###,##0.00") & " " & gtarrArtikelArray(i).WE
            If gtarrArtikelArray(i).MindestpreisNichtErreicht Then Preis(iCnt).Caption = Replace(gsReservedPriceMarker, "%PRICE%", Preis(iCnt).Caption)
            sVersand = Trim(IIf(Left(gtarrArtikelArray(i).Versand, 1) = "*" Or gtarrArtikelArray(i).Versand = "", Mid(gtarrArtikelArray(i).Versand, 2), Format(gtarrArtikelArray(i).Versand, "###,##0.00")))
            If sVersand Like "*#[.,]##" And InStr(1, sVersand, gtarrArtikelArray(i).WE) = 0 Then sVersand = sVersand & " " & gtarrArtikelArray(i).WE
            If gbShowShippingCosts Then
              Versandkosten(iCnt).Caption = sVersand
              sVersand = "" ' nullen weil wird schon direkt angezeigt
            Else
              Versandkosten(iCnt).Caption = ""
              sVersand = gsarrLangTxt(600) & ": " & sVersand & gsToolTipSeparator ' aufbereiten für ToolTip
            End If
            If gtarrArtikelArray(i).AktPreis = 0 Then Preis(iCnt).Caption = "?"
            Bietgruppe(iCnt).Text = gtarrArtikelArray(i).Gruppe
            Bietgruppe(iCnt).Enabled = Not gbUsesOdbc
            
            Waehrung(iCnt).Caption = gtarrArtikelArray(i).WE
        End If 'Ende Not bNurRestzeit
        
        If gtarrArtikelArray(i).UpdateInProgressSince > 0 And Abs(DateDiff("s", MyNow, gtarrArtikelArray(i).UpdateInProgressSince)) < 60 Then
          Titel(iCnt).BackColor = vbYellow
        Else
          Titel(iCnt).BackColor = &H8000000F
        End If
        
        If gtarrArtikelArray(i).EndeZeit = myDateSerial(1999, 9, 9) Then
            EndeZeit(iCnt).Caption = ""
            EndeZeit(iCnt).BackColor = &H8000000F
            Gebot(iCnt).Text = ""
            Gebot(iCnt).BackColor = vbWindowBackground
            Preis(iCnt).Caption = ""
            Preis(iCnt).BackColor = &H8000000F
            Versandkosten(iCnt).Caption = ""
            Versandkosten(iCnt).BackColor = &H8000000F
            Bietgruppe(iCnt).Text = ""
            Status(iCnt).Caption = ""
            Status(iCnt).BackColor = &H8000000F
        Else
            If gtarrArtikelArray(i).EndeZeit < MyNow Then
                EndeZeit(iCnt).BackColor = vbRed
            Else
                EndeZeit(iCnt).BackColor = vbGreen
            End If
            
            'Restzeit bestimmen
            Call ZeitPrüfung(i, True)
        End If
        
        ' Akt. Preis grün färben, wenn user Höchstbieter KOM 29.08.03
        If (gtarrArtikelArray(i).eBayUser > "") Then
            sUser = gtarrArtikelArray(i).eBayUser
        ElseIf (gtarrArtikelArray(i).UserAccount > "") Then
            sUser = gtarrArtikelArray(i).UserAccount
        Else
            sUser = gsUser
        End If
        If LCase(gtarrArtikelArray(i).Bieter) = LCase(sUser) And sUser <> "" Then
            Preis(iCnt).BackColor = vbGreen
            If gbShowShippingCosts Then
              Versandkosten(iCnt).BackColor = &H8000000F
            Else
              Versandkosten(iCnt).BackColor = vbGreen
            End If
        Else
            Preis(iCnt).BackColor = &H8000000F
            Versandkosten(iCnt).BackColor = &H8000000F
        End If
             
        Gebot(iCnt).Enabled = True
        Select Case gtarrArtikelArray(i).Status
            Case [asNixLos]
                EndeZeit(iCnt).BackColor = vbGreen
                Restzeit(iCnt).BackColor = &H8000000F
                If gtarrArtikelArray(i).NotFound > 0 Then
                    Status(iCnt).Caption = gtarrArtikelArray(i).NotFound & gsarrLangTxt(412)
                    Status(iCnt).BackColor = RGB(255, 128, 18) 'orange
                Else
                    Status(iCnt).Caption = ""
                    Status(iCnt).BackColor = &H8000000F
                End If
            Case [asErr]
                Status(iCnt).Caption = gsarrLangTxt(96)
                Status(iCnt).BackColor = vbRed
            Case [asOK]
                Status(iCnt).Caption = gsarrLangTxt(432)
                Status(iCnt).BackColor = vbGreen
            Case [asLowBid]
                Status(iCnt).Caption = ""
                Status(iCnt).BackColor = &H8000000F
            Case [asBieten]
                Status(iCnt).Caption = gsarrLangTxt(92)
                Status(iCnt).BackColor = vbYellow
            Case [asCancelGroup]
                Status(iCnt).Caption = gsarrLangTxt(433)
                Status(iCnt).BackColor = vbYellow
            Case [asHoldGroup]
                Status(iCnt).Caption = gsarrLangTxt(434)
                Status(iCnt).BackColor = vbYellow
            Case [asDelegatedBom]
                Status(iCnt).Caption = gsarrLangTxt(435)
                Status(iCnt).BackColor = vbYellow
            Case [asCancelBid]
                Status(iCnt).Caption = gsarrLangTxt(97)
                Status(iCnt).BackColor = vbYellow
            Case [asBuyOnly]
                Status(iCnt).Caption = gsarrLangTxt(98)
                Status(iCnt).BackColor = vbYellow
                Gebot(iCnt).Text = ""
                Gebot(iCnt).Enabled = False
            Case [asBuyOnlyOnHold]
                Status(iCnt).Caption = gsarrLangTxt(434) & " / " & gsarrLangTxt(98)
                Status(iCnt).BackColor = vbYellow
                Gebot(iCnt).Text = ""
                Gebot(iCnt).Enabled = False
            Case [asBuyOnlyCanceled]
                Status(iCnt).Caption = gsarrLangTxt(433) & " / " & gsarrLangTxt(98)
                Status(iCnt).BackColor = vbYellow
                Gebot(iCnt).Text = ""
                Gebot(iCnt).Enabled = False
            Case [asBuyOnlyDelegated]:
                Status(iCnt).Caption = gsarrLangTxt(435) & " / " & gsarrLangTxt(98)
                Status(iCnt).BackColor = vbYellow
                Gebot(iCnt).Text = ""
                Gebot(iCnt).Enabled = False
            Case [asAdvertisement]
                Status(iCnt).Caption = gsarrLangTxt(436)
                Status(iCnt).BackColor = vbYellow
                Gebot(iCnt).Text = ""
                Gebot(iCnt).Enabled = False
            Case [asEnde]
                Status(iCnt).Caption = gsarrLangTxt(99)
                Status(iCnt).BackColor = &H8000000F
            Case [asPower]
                Status(iCnt).Caption = gsarrLangTxt(73)
                Status(iCnt).BackColor = vbYellow
            Case [asUeberboten]
                Status(iCnt).Caption = gsarrLangTxt(95)
                Status(iCnt).BackColor = vbRed
            Case [asNotFound]
                EndeZeit(iCnt).BackColor = vbRed
                Status(iCnt).Caption = gsarrLangTxt(413)
                Status(iCnt).BackColor = vbRed
                Restzeit(iCnt).Caption = "        ? ? ?"
                Restzeit(iCnt).BackColor = vbRed
            Case [asSellerAway]
                Status(iCnt).Caption = gsarrLangTxt(431)
                Status(iCnt).BackColor = RGB(255, 128, 18) 'orange
            Case [asAccessErr]
                Status(iCnt).Caption = gsarrLangTxt(254)
                Status(iCnt).BackColor = RGB(255, 128, 18) 'orange
        End Select
        
        Gebot(iCnt).BackColor = GetGebotColorFromItem(i)
        
        If Not bNurRestzeit Then
            If Preis(iCnt).Caption <> "" Then 'lg 08.06.03
                If StatusIstBuyItNowStatus(gtarrArtikelArray(i).Status) Then
                    Preis(iCnt).ToolTipText = gtarrArtikelArray(i).Bieter & gsToolTipSeparator & _
                                               gsarrLangTxt(251) & ": " & gtarrArtikelArray(i).Verkaeufer & " (" & gtarrArtikelArray(i).Bewertung & ")" & gsToolTipSeparator & _
                                               sVersand & gtarrArtikelArray(i).Standort
                
                ElseIf gtarrArtikelArray(i).Status = [asPower] Or gtarrArtikelArray(i).AnzGebote < 0 Then
                    Preis(iCnt).ToolTipText = gtarrArtikelArray(i).Bieter & gsToolTipSeparator & _
                                               gsarrLangTxt(251) & ": " & gtarrArtikelArray(i).Verkaeufer & " (" & gtarrArtikelArray(i).Bewertung & ")" & gsToolTipSeparator & _
                                               sVersand & gtarrArtikelArray(i).Standort
                
                ElseIf gtarrArtikelArray(i).Status = [asAdvertisement] Then
                    Preis(iCnt).ToolTipText = gsarrLangTxt(251) & ": " & gtarrArtikelArray(i).Verkaeufer
    
                ElseIf gtarrArtikelArray(i).AnzGebote = 0 Then
                    Preis(iCnt).ToolTipText = gtarrArtikelArray(i).AnzGebote & " " & gsarrLangTxt(80) & gsToolTipSeparator & _
                                               gsarrLangTxt(251) & ": " & gtarrArtikelArray(i).Verkaeufer & " (" & gtarrArtikelArray(i).Bewertung & ")" & gsToolTipSeparator & _
                                               sVersand & gtarrArtikelArray(i).Standort
                
                ElseIf gtarrArtikelArray(i).Bieter = gsarrLangTxt(268) Then
                    Preis(iCnt).ToolTipText = gtarrArtikelArray(i).AnzGebote & " " & gsarrLangTxt(80) & gsToolTipSeparator & _
                                               gtarrArtikelArray(i).Bieter & gsToolTipSeparator & _
                                               gsarrLangTxt(251) & ": " & gtarrArtikelArray(i).Verkaeufer & " (" & gtarrArtikelArray(i).Bewertung & ")" & gsToolTipSeparator & _
                                               sVersand & gtarrArtikelArray(i).Standort
                ElseIf gtarrArtikelArray(i).Bieter > "" Then
                    Preis(iCnt).ToolTipText = gtarrArtikelArray(i).AnzGebote & " " & gsarrLangTxt(80) & gsToolTipSeparator & _
                                               gsarrLangTxt(252) & ": " & gtarrArtikelArray(i).Bieter & gsToolTipSeparator & _
                                               gsarrLangTxt(251) & ": " & gtarrArtikelArray(i).Verkaeufer & " (" & gtarrArtikelArray(i).Bewertung & ")" & gsToolTipSeparator & _
                                               sVersand & gtarrArtikelArray(i).Standort
                Else
                    Preis(iCnt).ToolTipText = gtarrArtikelArray(i).AnzGebote & " " & gsarrLangTxt(80) & gsToolTipSeparator & _
                                               gsarrLangTxt(251) & ": " & gtarrArtikelArray(i).Verkaeufer & " (" & gtarrArtikelArray(i).Bewertung & ")" & gsToolTipSeparator & _
                                               sVersand & gtarrArtikelArray(i).Standort
                    
                End If
                'Allgemeine Tooltips
                If gtarrArtikelArray(i).MindestpreisNichtErreicht Then
                  Preis(iCnt).ToolTipText = Preis(iCnt).ToolTipText & gsToolTipSeparator & gsarrLangTxt(602)
                End If
                If Not gbRevisedInTitle And gtarrArtikelArray(i).Ueberarbeitet Then
                  Preis(iCnt).ToolTipText = Preis(iCnt).ToolTipText & gsToolTipSeparator & gsarrLangTxt(749)
                End If
                
                If gbEditShippingOnClick And gbShowShippingCosts Then
                  Versandkosten(iCnt).ToolTipText = gsarrLangTxt(601)
                Else
                  Versandkosten(iCnt).ToolTipText = Preis(iCnt).ToolTipText
                End If
    
            Else
              Preis(iCnt).ToolTipText = "click = " & gsarrLangTxt(50)
            End If
            
            If Gebot(iCnt).Text <> "" Then
                Select Case Waehrung(iCnt).Caption
                Case "€"
                   Gebot(iCnt).ToolTipText = ""
                Case "£"
                   Gebot(iCnt).ToolTipText = "ca. " & Format(CDbl(Gebot(iCnt)) * gcolWeValues("GBP"), "#,##0.00") & " €"
                Case "$"
                   Gebot(iCnt).ToolTipText = "ca. " & Format(CDbl(Gebot(iCnt)) * gcolWeValues("USD"), "#,##0.00") & " €"
                Case "CHF"
                   Gebot(iCnt).ToolTipText = "ca. " & Format(CDbl(Gebot(iCnt)) * gcolWeValues("CHF"), "#,##0.00") & " €"
                Case "AU$"
                   Gebot(iCnt).ToolTipText = "ca. " & Format(CDbl(Gebot(iCnt)) * gcolWeValues("AUD"), "#,##0.00") & " €"
                Case "C$"
                   Gebot(iCnt).ToolTipText = "ca. " & Format(CDbl(Gebot(iCnt)) * gcolWeValues("CAD"), "#,##0.00") & " €"
                Case Else
                    Gebot(iCnt).ToolTipText = gsarrLangTxt(79)
                End Select
            Else
              Gebot(iCnt).ToolTipText = gsarrLangTxt(81)
            End If
            
            If gtarrArtikelArray(i).Kommentar <> "" Then
              'in den Tooltips die Zeilenumbrüche filtern
              Titel(iCnt).ToolTipText = Replace(gtarrArtikelArray(i).Kommentar, vbCrLf, " ", 1, -1, vbBinaryCompare)
              Ecke(iCnt).ToolTipText = Titel(iCnt).ToolTipText
              Ecke(iCnt).Visible = True
            Else
              Titel(iCnt).ToolTipText = gsarrLangTxt(82)
              Ecke(iCnt).ToolTipText = ""
              Ecke(iCnt).Visible = False
            End If
            
            sTmp = Format(gtarrArtikelArray(i).AktPreis, "###,##0.00") & " " & gtarrArtikelArray(i).WE
            sTmp = Right(sTmp, Len(gtarrArtikelArray(i).WE) + 5)
            Mid(sTmp, 1, 1) = "#"
            Mid(sTmp, 3, 2) = "##"
            
            For m = Len(Preis(iCnt).Caption) - Len(sTmp) + 1 To 1 Step -1
              If Mid(Preis(iCnt).Caption, m, Len(sTmp)) Like sTmp Then Exit For
            Next m
            For n = Len(Versandkosten(iCnt).Caption) - Len(sTmp) + 1 To 1 Step -1
              If Mid(Versandkosten(iCnt).Caption, n, Len(sTmp)) Like sTmp Then Exit For
            Next n
            If m > 0 And n > 0 Then
              ' xx = Gesamtbreite - Textbreite / 2 + Teiltextbreite
              mm = (Preis(iCnt).Width - Me.TextWidth(Preis(iCnt).Caption)) / 2 + Me.TextWidth(Left(Preis(iCnt).Caption, m))
              nn = (Preis(iCnt).Width - Me.TextWidth(Versandkosten(iCnt).Caption)) / 2 + Me.TextWidth(Left(Versandkosten(iCnt).Caption, n))
              Versandkosten(iCnt).Left = Preis(iCnt).Left + mm - nn
            Else
              Versandkosten(iCnt).Left = Preis(iCnt).Left
            End If
            If Versandkosten(iCnt).Caption > "" Then
              Call VersandKuerzen(Versandkosten(iCnt), Preis(iCnt))
            End If
            
        End If 'Ende Not bNurRestzeit
        iCnt = iCnt + 1
    Next i
    
    If Not bNurRestzeit And Not bNurDieseZeile Then
        If giAktAnzArtikel = 0 Then iCnt = 0
        If iCnt - 1 < giMaxRow Then
          ' Rest leeren
          For i = iCnt To giMaxRow
            Artikel(i).Text = ""
            EndeZeit(i).Caption = ""
            EndeZeit(i).BackColor = &H8000000F
            Gebot(i).Text = ""
            Gebot(i).BackColor = vbWindowBackground
            Gebot(i).Enabled = True
            Titel(i).Caption = ""
            Preis(i).Caption = ""
            Preis(i).BackColor = &H8000000F
            Versandkosten(i).Caption = ""
            Versandkosten(i).BackColor = &H8000000F
            Bietgruppe(i).Text = ""
            Status(i).Caption = ""
            Status(i).BackColor = &H8000000F
            Restzeit(i).Caption = ""
            Restzeit(i).BackColor = &H8000000F
          Next i
        End If
        
        NewArtikel.Caption = CStr(giAktAnzArtikel) & " " & gsarrLangTxt(31)
        VScroll1.Visible = CBool(VScroll1.Max > VScroll1.Min)
    End If
    
End If 'Me.WindowState <> vbMinimized

errhdl:

End Sub

Sub TitelKuerzen(ByVal lbTitel As Label)

  If Len(lbTitel) > 1000 Then lbTitel = Left(lbTitel, 1000)

  Dim v As Variant
  v = Split(lbTitel.Caption, " ")
  
  Dim z1 As String, z2 As String
  Dim i As Integer
  For i = LBound(v) To UBound(v)
    If z1 > "" Then z1 = z1 & " "
    If Me.TextWidth(z1 & v(i)) > lbTitel.Width Then
      Exit For
    End If
    z1 = z1 & v(i)
  Next i
  
  z2 = Mid(lbTitel.Caption, Len(z1) + 1)
  
  If Me.TextWidth(z2) < lbTitel.Width Then Exit Sub
  
  For i = Len(z2) - 1 To 1 Step -1
    If Me.TextWidth(Mid(z2, 1, i) & "... ") < lbTitel.Width Then
      lbTitel.Caption = Left(lbTitel.Caption, Len(lbTitel.Caption) - (Len(z2) - i)) & "..."
      Exit Sub
    End If
  Next i

End Sub

Sub VersandKuerzen(ByVal lbVersand As Label, ByVal lbPreis As Label)

  If Len(lbVersand.Caption) > 50 Then lbVersand.Caption = Left(lbVersand.Caption, 50)

  Dim Z As String
  Dim i As Integer
  
  Z = lbVersand.Caption
  If Me.TextWidth(Z) < lbPreis.Width Then Exit Sub

  For i = Len(Z) - 1 To 1 Step -1
    If Me.TextWidth(Mid(Z, 1, i) & "... ") < lbPreis.Width Then
      lbVersand.Caption = Left(lbVersand.Caption, Len(lbVersand.Caption) - (Len(Z) - i)) & "..."
      Exit Sub
    End If
  Next i

End Sub

Private Sub ReadArtikelIni(bDoAppend As Boolean)

  Dim iFileNr As Integer
  Dim sReadFile As String
  Dim lAktRow As Long
  Dim sArtikelFileVersion As String
  On Error GoTo errhdl
  
  If bDoAppend Then
    lAktRow = giAktAnzArtikel + 1
    ReDim Preserve gtarrArtikelArray(lAktRow)
  Else
    giAktAnzArtikel = 0
    lAktRow = 1
    ReDim gtarrArtikelArray(1)
  End If
  
  'und den ganzen Filekrams reinlutschen:
  
  RestoreArtikelCsv
  
  iFileNr = FreeFile
  
  Open gsAppDataPath & "\Artikel.csv" For Binary As #iFileNr
  sReadFile = Space$(LOF(iFileNr))
  Get #iFileNr, , sReadFile
  Close #iFileNr
  
  gsLastSavedCrc = Crc32(sReadFile)
  
  If Len(sReadFile) > 10 Then
    'wir prüfen jetzt die Version der Artikel.csv und lesen das entsprechende Format ein, lg 04.05.03
    sArtikelFileVersion = GetArtikelFileVersion(Left(sReadFile, 100))
    
    If VersionValue(sArtikelFileVersion) >= VersionValue("2.4.0") Then
      AddCsvArtikel2 sReadFile 'ab Version 2.4.0
    Else
      AddCsvArtikel2 sReadFile, True 'bis Version 2.3.0
    End If
  End If

errhdl:

 Err.Clear
 On Error Resume Next
 Close #iFileNr
 
 VScroll1.Max = giAktAnzArtikel - giMaxRow
 If VScroll1.Max < VScroll1.Min Then VScroll1.Max = VScroll1.Min
 
 ReDim Preserve gtarrArtikelArray(giAktAnzArtikel)

 ArtikelArrayToScreen 1
 
End Sub

Public Function AddArtikel(sArtikel As String, Optional fGebot As Double = 0, Optional sGruppe As String = "", Optional sUser As String = "", Optional sKommentar As String = "") As Long

Dim i As Integer
On Error GoTo errhdl

sArtikel = Trim(sArtikel)
If sArtikel = "" Or sArtikel = "0" Then Exit Function
If gbBlacklistDeletedItems And CheckItemBlacklist(sArtikel) Then Exit Function

'wir wollen das Gebot nicht doppelt haben:
For i = 1 To giAktAnzArtikel
    If gtarrArtikelArray(i).Artikel = sArtikel Then
        
        If fGebot > 0 Then gtarrArtikelArray(i).Gebot = fGebot
        If sGruppe > "" Then gtarrArtikelArray(i).Gruppe = sGruppe
        If sUser > "" Then gtarrArtikelArray(i).UserAccount = sUser
        If sKommentar > "" Then gtarrArtikelArray(i).Kommentar = sKommentar
        gtarrArtikelArray(i).LastChangedId = GetChangeID()
        Call ArtikelArrayToScreen(miStartShowArtikel)
        Exit Function
    End If
Next i

'wir hängen einen neuen Artikel ein ..
       
If giAktAnzArtikel + 1 > UBound(gtarrArtikelArray()) Then
    ReDim Preserve gtarrArtikelArray(giAktAnzArtikel + 1) As udtArtikelZeile
End If

'wir erhöhen den Zähler erst nach dem ReDim falls es schiefgeht
giAktAnzArtikel = giAktAnzArtikel + 1


'Mal sehen wohin damit ..
Call InitArtikel(giAktAnzArtikel)
With gtarrArtikelArray(giAktAnzArtikel)
    .Artikel = sArtikel
    .Gebot = fGebot     ' Übernahme von Gruppe und Gebot bei Shift Drop
    .Gruppe = sGruppe   ' Übernahme von Gruppe und Gebot bei Shift Drop
    .UserAccount = sUser
    .Kommentar = sKommentar
    .LastChangedId = GetChangeID()
End With

With VScroll1
    .Max = giAktAnzArtikel - giMaxRow
    If .Max < .Min Then .Max = .Min
    .Visible = CBool(.Max > .Min)
End With

If Not CheckInternetConnection Then
  Call Ask_Online
End If

If IsOnline Then
    Call Update_Artikel(giAktAnzArtikel)
    AddArtikel = ItemToIndex(sArtikel)
    If AddArtikel > 0 Then
        If gtarrArtikelArray(AddArtikel).EndeZeit < MyNow Then
            gtarrArtikelArray(AddArtikel).PostUpdateDone = True
            If gbBlockEndedItems Then RemoveArtikel CInt(AddArtikel), False, False
        ElseIf gtarrArtikelArray(AddArtikel).Status = [asBuyOnly] Then
            If gbBlockBuyItNowItems Then RemoveArtikel CInt(AddArtikel), False, False
        End If
    End If
End If

'ab in den Sorter ..
Call QuickSortDate(gtarrArtikelArray(), 1, UBound(gtarrArtikelArray()), CBool(gsSortOrder = "asc")) 'lg 10.07.2003

AddArtikel = ItemToIndex(sArtikel)

Call ArtikelArrayToScreen(miStartShowArtikel)

If gbUsesModem Then
    If gbLastDialupWasManually Then Call Ask_Offline
End If

Exit Function

errhdl:
Err.Clear

End Function

Public Sub RemoveArtikel(iIdx As Integer, Optional bDoArtikelArrayToScreen As Boolean = True, Optional bDoCheckBietGruppen As Boolean = True)
Dim i As Integer
Dim sTmp As String

'wir schmeissen einen Artikel raus ..
'iIdx = Zeiger auf zu löschenden Artikel!

'Wenn nichts drin, dann raus ..
If iIdx > giAktAnzArtikel Then Exit Sub

If iIdx > 0 Then
  giAktAnzArtikel = giAktAnzArtikel - 1
  
  On Error GoTo WEITER
  
  sTmp = Dir(gsTempPfad & "\Art-" & gtarrArtikelArray(iIdx).Artikel & "-*.html")
  Do While sTmp > ""
    Kill gsTempPfad & "\" & sTmp
    sTmp = Dir
  Loop
WEITER:
  
  If gbLogDeletedItems Then WriteDeletedItemLog iIdx
  If gbBlacklistDeletedItems Then WriteBlacklistedItemLog iIdx
  
  ReDim Preserve gtarrRemovedArtikelArray(0 To UBound(gtarrRemovedArtikelArray) + 1)
  gtarrRemovedArtikelArray(UBound(gtarrRemovedArtikelArray)).Artikel = gtarrArtikelArray(iIdx).Artikel
  gtarrRemovedArtikelArray(UBound(gtarrRemovedArtikelArray)).LastChangedId = GetChangeID()
  
  If giAktAnzArtikel < 1 Then
      InitArtikel (1)
      ArtikelArrayToScreen (1)
      Exit Sub
  End If
  For i = iIdx To giAktAnzArtikel
      gtarrArtikelArray(i) = gtarrArtikelArray(i + 1)
  Next i
  
  ReDim Preserve gtarrArtikelArray(giAktAnzArtikel)
End If

If bDoArtikelArrayToScreen Then
  VScroll1.Max = giAktAnzArtikel - giMaxRow
  If VScroll1.Max < VScroll1.Min Then VScroll1.Max = VScroll1.Min
  VScroll1.Visible = VScroll1.Max > VScroll1.Min
  If (miStartShowArtikel + giMaxRow) > giAktAnzArtikel Then
      ChangeView (giAktAnzArtikel - giMaxRow)
  Else
      ArtikelArrayToScreen (miStartShowArtikel)
  End If
End If

If bDoCheckBietGruppen Then CheckAlleBietgruppen

End Sub

Public Function CheckItemBlacklist(ByVal sArtikel As String) As Boolean

    Static sTmp As String
    Static dTimestamp As Date
    
    Dim sFile As String
    sFile = gsAppDataPath & "\BlacklistedItems.log"
    
    If Dir(sFile) = "" Then Exit Function
    If dTimestamp <> FileDateTime(sFile) Then
        sTmp = vbCrLf & ReadFromFile(sFile)
        dTimestamp = FileDateTime(sFile)
    End If
        
    If InStr(1, sTmp, vbCrLf & sArtikel & vbCrLf) > 0 Then CheckItemBlacklist = True
  
End Function

Private Sub ChangeView(iStartRow As Integer)
If iStartRow < 1 Then iStartRow = 1
VScroll1.Value = iStartRow
Call ArtikelArrayToScreen(miStartShowArtikel)
End Sub

Private Sub TitelFlashTimer_Timer()
    
    On Error Resume Next
    Dim tmpArray As Variant
    
    If giSuspendState = 0 Then
    
        tmpArray = Split(TitelFlashTimer.Tag, ",")
        tmpArray(0) = tmpArray(0) - 1 ' FlashesLeft
        'tmpArray(1)                   ' Index
        tmpArray(2) = 1 - tmpArray(2) ' FlashStatus
        
        TitelFlashTimer.Tag = Join(tmpArray, ",")
        
        If tmpArray(0) = 0 Then TitelFlashTimer.Enabled = False
        Titel(tmpArray(1)).BackColor = IIf(tmpArray(2) > 0, vbYellow, &H8000000F)
        
    End If 'giSuspendState = 0
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Tag
        Case "tbSaveAll"
            Call mnuSaveAll_Click
        Case "tbReadArtikel"
            Call mnuReadArtikel_Click
        Case "tbUpdateArtikel"
            If Toolbar1.Buttons(4).Image = 3 Then
                mbStopUpdate = True
            Else
                Call mnuUpdateArtikel_Click
            End If
        Case "tbReadEbay"
            Call mnuMyEbay_Click
        Case "tbSyncEbayTime"
            Call mnuSync_Click
        Case "tbLogin"
            Call mnuPasswd_Click
        Case "tbAuto"
            Call mnuAuto_Click
        Case "tbSettings"
            Call mnuSettings_Click
        Case "tbBrowser"
            Call StarteBrowser
        Case "tbArtikel"
            Call mnuArtikel_Click
        Case "tbAbout"
            Call mnuAbout_Click
        Case "tbHelp"
            Call mnuHelp_Click
        Case Else
    End Select
    
End Sub

Private Sub TrayFlashTimer_Timer()
    
    Static iCurrentIcon As Integer
    Static iTimeElapsedSinceLastIconChange As Integer
    
    Dim iTestTime As Integer
    Dim iNextIcon As Integer
    
    iNextIcon = iCurrentIcon
    If gbAutoMode Then
        If iCurrentIcon = 1 Then
            If frmDummy.Picture1(2).Tag > "" Then
                iNextIcon = 2
                iTestTime = giTrayIconDisplayTimeOnlineMode1
            Else
                TrayFlashTimer.Enabled = False
            End If
        Else
            iNextIcon = 1
            iTestTime = giTrayIconDisplayTimeOnlineMode2
        End If
    Else
        If iCurrentIcon = 3 Then
            If frmDummy.Picture1(4).Tag > "" Then
                iNextIcon = 4
                iTestTime = giTrayIconDisplayTimeOfflineMode1
            Else
                TrayFlashTimer.Enabled = False
            End If
        Else
            iNextIcon = 3
            iTestTime = giTrayIconDisplayTimeOfflineMode2
        End If
    End If
    
    iTimeElapsedSinceLastIconChange = iTimeElapsedSinceLastIconChange + 1
    If iTestTime < iTimeElapsedSinceLastIconChange Then
        iTimeElapsedSinceLastIconChange = 0
        Set frmDummy.Picture1(0) = frmDummy.Picture1(iNextIcon)
        If iNextIcon <> iCurrentIcon Then Call updTaskbarIcon
        iCurrentIcon = iNextIcon
    End If
    
End Sub

Private Sub Versandkosten_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    Call SetFocusRect(Index)
    If gbShowShippingCosts And gbEditShippingOnClick Then
        If Button = vbLeftButton Then '1
            Call VersandkostenBearbeiten(Index)
        Else
            Call Preis_MouseDown(Index, Button, Shift, X, Y)
        End If
    Else
        Call Preis_MouseDown(Index, Button, Shift, X, Y)
    End If
    
End Sub

Private Sub VersandkostenEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iIdx As Integer
    
    If KeyCode = vbKeyEscape Then '27
        VersandkostenEdit.Visible = False
        VersandkostenEdit.Tag = ""
    ElseIf KeyCode = vbKeyReturn Then '13
        iIdx = CInt(VersandkostenEdit.Tag) - VScroll1.Value
        Call VersandkostenUebernehmen
        On Error Resume Next
        Gebot(iIdx).SetFocus
        On Error GoTo 0
    End If
    
End Sub

Private Sub VersandkostenEdit_LostFocus()
Call VersandkostenUebernehmen
End Sub

Private Sub VersandkostenBearbeiten(iIdx As Integer)
    
    Dim i As Integer
    
    If Preis(iIdx).Caption <> "" Then
        i = iIdx + VScroll1.Value
        With VersandkostenEdit
            If .Visible = True Then Call VersandkostenUebernehmen
            
            .Text = IIf(Left(gtarrArtikelArray(i).Versand, 1) = "*" Or gtarrArtikelArray(i).Versand = "", Mid(gtarrArtikelArray(i).Versand, 2), Format(gtarrArtikelArray(i).Versand, "###,##0.00") & IIf(gtarrArtikelArray(i).Versand Like "*#*" And InStr(1, Format(gtarrArtikelArray(i).Versand, "###,##0.00"), gtarrArtikelArray(i).WE) = 0, " " & gtarrArtikelArray(i).WE, ""))
            Call .Move(Versandkosten(iIdx).Left, Versandkosten(iIdx).Top, Preis(iIdx).Width, Versandkosten(iIdx).Height)
            .FontName = Versandkosten(iIdx).FontName
            .FontSize = Versandkosten(iIdx).FontSize
            .Tag = iIdx + VScroll1.Value
            .TabIndex = Versandkosten(iIdx).TabIndex
            .Visible = True
            .SetFocus
        End With
    End If
    
End Sub

Private Sub VersandkostenUebernehmen()
    
    Dim i As Integer
    Dim sTagSave As String
    Dim sVersandkostenSave As String
    
    With VersandkostenEdit
        sTagSave = .Tag
        If .Visible Then
            If sTagSave > "" Then .Tag = "-" ' TagSave muss sein wegen SetFocusRect - ' brauchen wir, weil sonst bei Klick auf ein nicht-Edit-Feld einer anderen Zeile der Fokus innerhalb der vorherigen Zeile bleibt
            .Visible = False
            .Refresh
            .Tag = sTagSave
            If .Tag > "" Then
                i = CInt(.Tag)
                
                With gtarrArtikelArray(i)
                    sVersandkostenSave = .Versand
                    .Versand = "*" & Trim(VersandkostenEdit.Text)  ' der Stern bedeutet manuelle Versandkosten, sollen nicht mehr durch eBay überschrieben werden
                    If .Versand Like "*#[,.]##*" And Right(.Versand, Len(.WE)) = .WE Then .Versand = Trim(Left(.Versand, Len(.Versand) - Len(.WE))) 'Währung rausziehen
                    If .Versand = "*" Then .Versand = ""  'Wenn gar nichts eigetragen, dann wieder für Automatik freischalten
                    If sVersandkostenSave = Mid(.Versand, 2) Then ' nichts geändert
                        .Versand = sVersandkostenSave
                    Else
                        .LastChangedId = GetChangeID()
                    End If
                End With 'gtarrArtikelArray(i)
                
                Call VScroll1_Change
                .Tag = ""
            End If
        End If
    End With 'VersankostenEdit
    
End Sub

Private Sub VScroll1_Change()
  
    If VersandkostenEdit.Visible Then Call VersandkostenUebernehmen
    
    miStartShowArtikel = VScroll1.Value
    
    If miStartShowArtikel > giAktAnzArtikel - giMaxRow Then
        miStartShowArtikel = giAktAnzArtikel - giMaxRow
    End If
    If miStartShowArtikel <= 0 Then
        miStartShowArtikel = 1
    End If
    Call ArtikelArrayToScreen(miStartShowArtikel)
End Sub

Private Sub InitArtikel(iIdx As Integer)

On Error Resume Next

    With gtarrArtikelArray(iIdx)
        .Artikel = ""
        .EndeZeit = myDateSerial(1999, 9, 9)
        .Titel = ""
        .Gebot = 0
        .MinGebot = 0
        .AktPreis = 0
        .Gruppe = ""
        .Status = [asNixLos]
        .WE = ""
        .AnzGebote = 0
        .Bieter = ""
        .PostUpdateDone = False '0
        .Kommentar = ""
        .Versand = ""
        .Verkaeufer = ""
        .eBayUser = ""
        .eBayPass = ""
        .UseToken = False '0
        .UserAccount = ""
        .NotFound = 0
        .Bewertung = ""
        .Standort = ""
        .MindestpreisNichtErreicht = False '0
        .Ueberarbeitet = False '0
        .UpdateInProgressSince = 0
        .ExtCmdPreDone = False '0
        .ExtCmdPostDone = False '0
        .LastChangedId = 0
        .TimeZone = 0
    End With
    
End Sub


Public Function CheckSpeed(Optional ByVal iTestPasses As Integer = 10, Optional sUsername = "", Optional sPassword = "", Optional bToken As Boolean = False) As Double
Dim fLap As Double
Dim i  As Integer

On Error Resume Next

'ruft 10 Mal die "Bieten"- Maske mit ungültigem Artikel

fLap = Timer

'später für Vorlaufberechnung/test bietserver

For i = 1 To iTestPasses '10
    Call Bieten(gsTestArtikel, "1", sUsername, sPassword, 0, 1, bToken)
    If gbStopTests Then Exit For
Next i

fLap = (Timer - fLap) * 3 'weil später 2 Durchgänge laufen

CheckSpeed = fLap / iTestPasses ' sec je durchgang

End Function

Private Function Check_Wartung(sTxt As String) As Boolean
Dim lPos As Long

On Error Resume Next

Check_Wartung = False

lPos = InStr(1, sTxt, gsAnsAskSeller) + InStr(1, sTxt, gsAnsWatchStart)
If lPos > 0 Then
  lPos = 0 ' scheint eine Artikelseite zu sein und keine Wartung
Else
  lPos = InStr(1, sTxt, gsAnsMaintenance)
End If

If lPos > 0 Then

    Check_Wartung = True
    
    If Not gbEBayWartung Then
        Call DebugPrint("Wartung hat begonnen ")
    End If
    
    gbEBayWartung = True
    If Not mbPopupShown And Not gbGeboteAktualisieren And Not gbAutoStart Then
        mbPopupShown = True
        MsgBox gsarrLangTxt(22) & gsarrLangTxt(23)
    Else
        Call PanelText(StatusBar1, 1, gsarrLangTxt(22), False, vbRed, vbBlack)
    End If
Else
  If gbEBayWartung Then
     Call DebugPrint("Wartung beendet ")
  End If
  gbEBayWartung = False
  mbPopupShown = False
End If

End Function

Private Sub MWheel1_WheelScroll(Shift As Integer, zDelta As Integer, X As Single, Y As Single)
    
    Dim iNewScrollValue As Integer
    With VScroll1
        iNewScrollValue = .Value - zDelta
        If iNewScrollValue > (giAktAnzArtikel - giMaxRow) Then iNewScrollValue = (giAktAnzArtikel - giMaxRow)
        If iNewScrollValue < 1 Then iNewScrollValue = 1
        .Value = iNewScrollValue
    End With
    
End Sub

'
' Neue Routinen zum Zeilengenerieren
'
Private Sub SetInitLineSize()

On Error GoTo ERROR_HANDLER
Dim ERROR_TEXT As String
Dim ERROR_LOGGED As Boolean
Dim ERROR_DESC As String

Dim lHeight As Integer
Dim LFontSize As Integer

lHeight = giDefaultHeight
LFontSize = giDefaultFontSize

ERROR_TEXT = "Set FontName = """ & gsGlobFontName & """"
Me.FontName = gsGlobFontName

ERROR_TEXT = "Set FontSize = " & LFontSize
Me.FontSize = LFontSize

ERROR_TEXT = "Set Height = " & lHeight
Artikel(0).Height = lHeight

ERROR_TEXT = "???"

Artikel(0).Font.Size = LFontSize
Artikel(0).FontName = gsGlobFontName

'EndeZeit(0).Top = Artikel(0).Top
EndeZeit(0).Height = lHeight
EndeZeit(0).Font.Size = LFontSize
EndeZeit(0).FontName = gsGlobFontName

Titel(0).Height = lHeight
Titel(0).Font.Size = LFontSize
Titel(0).FontName = gsGlobFontName

Preis(0).Height = lHeight
Preis(0).Font.Size = LFontSize
Preis(0).FontName = gsGlobFontName

Versandkosten(0).Height = lHeight / 2
Versandkosten(0).Font.Size = LFontSize
Versandkosten(0).FontName = gsGlobFontName
Versandkosten(0).Top = Preis(0).Top + Preis(0).Height - Versandkosten(0).Height

Restzeit(0).Height = lHeight
Restzeit(0).Font.Size = LFontSize
Restzeit(0).FontName = gsGlobFontName

Status(0).Height = lHeight
Status(0).FontSize = LFontSize
Status(0).FontName = gsGlobFontName

Gebot(0).Height = lHeight
Gebot(0).FontSize = LFontSize
Gebot(0).FontName = gsGlobFontName

Bietgruppe(0).Height = lHeight
Bietgruppe(0).FontSize = LFontSize
Bietgruppe(0).FontName = gsGlobFontName

Waehrung(0).Height = lHeight
Waehrung(0).FontSize = LFontSize
Waehrung(0).FontName = gsGlobFontName

Line1(0).Y1 = Artikel(0).Top + Artikel(0).Height + 50
Line1(0).Y2 = Line1(0).Y1

Ecke(0).Top = Titel(0).Top + Titel(0).Height - Ecke(0).Height
Ecke(0).Left = Titel(0).Left + Titel(0).Width - Ecke(0).Width + 10

Exit Sub

ERROR_HANDLER:

ERROR_DESC = Err.Description
If Not ERROR_LOGGED Then
  DebugPrint "Error on " & ERROR_TEXT & " : " & ERROR_DESC
End If
Err.Clear
ERROR_LOGGED = True
Resume Next

End Sub

Private Sub EntferneZeile(iIdx As Integer)

'Index = Index auf zu entfernende Zeile!
' Index 1. Zeile = 0
If iIdx > 0 Then
    'Artikel
    Unload Artikel(iIdx)
    Unload Line1(iIdx)
    Unload EndeZeit(iIdx)
    Unload Titel(iIdx)
    Unload Preis(iIdx)
    Unload Versandkosten(iIdx)
    Unload Restzeit(iIdx)
    Unload Status(iIdx)
    Unload Gebot(iIdx)
    Unload Bietgruppe(iIdx)
    Unload Waehrung(iIdx)
    Unload Ecke(iIdx)
End If
End Sub

Private Sub GeneriereZeile(iIdx As Integer)
'Index = Index auf neue Zeile!
' Index 1. Zeile = 0
If iIdx > 0 Then
    'Artikel
    Load Artikel(iIdx)
    Artikel(iIdx).Top = Artikel(iIdx - 1).Top + Artikel(iIdx - 1).Height + 100
    Artikel(iIdx).Left = Artikel(iIdx - 1).Left
    Artikel(iIdx).Visible = True
    
    Load Line1(iIdx)
    Line1(iIdx).Y1 = Artikel(iIdx).Top + Artikel(iIdx).Height + 50
    Line1(iIdx).Y2 = Line1(iIdx).Y1
    Line1(iIdx).Visible = True
    
    Load EndeZeit(iIdx)
    EndeZeit(iIdx).Top = Artikel(iIdx).Top
    EndeZeit(iIdx).BackColor = &H8000000F
    EndeZeit(iIdx).Visible = True
    
    Load Titel(iIdx)
    Titel(iIdx).Top = Artikel(iIdx).Top
    Titel(iIdx).BackColor = &H8000000F
    Titel(iIdx).Visible = True
    Titel(iIdx).ZOrder vbSendToBack '1
    
    Load Preis(iIdx)
    Preis(iIdx).Top = Artikel(iIdx).Top
    Preis(iIdx).BackColor = &H8000000F
    Preis(iIdx).Visible = True
    
    Load Versandkosten(iIdx)
    Versandkosten(iIdx).Top = Preis(iIdx).Top + Preis(iIdx).Height - Versandkosten(iIdx).Height
    Versandkosten(iIdx).BackColor = &H8000000F
    Versandkosten(iIdx).Visible = True
    
    Preis(iIdx).ZOrder vbSendToBack '1
    
    Load Restzeit(iIdx)
    Restzeit(iIdx).Top = Artikel(iIdx).Top
    Restzeit(iIdx).BackColor = &H8000000F
    Restzeit(iIdx).Visible = True
    
    Load Status(iIdx)
    Status(iIdx).Top = Artikel(iIdx).Top
    Status(iIdx).BackColor = &H8000000F
    Status(iIdx).Visible = True
    
    Load Gebot(iIdx)
    Gebot(iIdx).Top = Artikel(iIdx).Top
    Gebot(iIdx).Visible = True
    Gebot(iIdx).FontBold = True
        
    Load Bietgruppe(iIdx)
    Bietgruppe(iIdx).Top = Artikel(iIdx).Top
    Bietgruppe(iIdx).Visible = True
    
    Load Waehrung(iIdx)
    Waehrung(iIdx).Top = Artikel(iIdx).Top
    Waehrung(iIdx).Visible = True
    
    Load Ecke(iIdx)
    Ecke(iIdx).Top = Titel(iIdx).Top + Titel(iIdx).Height - Ecke(iIdx).Height
    Ecke(iIdx).Left = Titel(iIdx).Left + Titel(iIdx).Width - Ecke(iIdx).Width + 10
    Ecke(iIdx).Visible = False
    Ecke(iIdx).ZOrder vbBringToFront '0
    
    'Inhalte löschen
    Artikel(iIdx).Text = ""
    EndeZeit(iIdx).Caption = ""
    Gebot(iIdx).Text = ""
    Bietgruppe(iIdx).Text = ""
    EndeZeit(iIdx).Caption = ""
    Titel(iIdx).Caption = ""
    Preis(iIdx).Caption = ""
    Restzeit(iIdx).Caption = ""
    
    'Sprachen ..
    Artikel(iIdx).ToolTipText = gsarrLangTxt(244)
    EndeZeit(iIdx).ToolTipText = gsarrLangTxt(31) & " " & gsarrLangTxt(60)
    Gebot(iIdx).ToolTipText = gsarrLangTxt(242)
    Bietgruppe(iIdx).ToolTipText = gsarrLangTxt(243)
    EndeZeit(iIdx).ToolTipText = gsarrLangTxt(245)
    Titel(iIdx).ToolTipText = gsarrLangTxt(246)
    Preis(iIdx).ToolTipText = gsarrLangTxt(247)
    Restzeit(iIdx).ToolTipText = gsarrLangTxt(248)
End If 'idx > 0
End Sub

Private Sub SetFormSize()

VScroll1.Height = Artikel(giMaxRow).Height + Artikel(giMaxRow).Top - VScroll1.Top

Zusatzfeld.Top = Artikel(giMaxRow).Height + Artikel(giMaxRow).Top + 100

Me.Height = Zusatzfeld.Top + Zusatzfeld.Height + StatusBar1.Height + 50

End Sub

Public Sub ScaleToToolbar(bIsOn As Boolean)
Dim i As Integer
Dim fOffset As Double
On Error Resume Next

  fOffset = Toolbar1.Height + 100
  If bIsOn Then fOffset = -fOffset
  
  For i = 0 To Me.Controls.Count - 1
    With Me.Controls(i)
        If Not TypeOf Me.Controls(i).Container Is ToolBar Then
            If TypeOf Me.Controls(i) Is Line Then
              .Y1 = .Y1 - fOffset
              .Y2 = .Y2 - fOffset
            Else
              .Top = .Top - fOffset
            End If
        End If
    End With
  Next i
End Sub
Private Sub InitToolbar()
    Dim i As Integer
    Dim fOffset As Single
    
    On Error Resume Next  'MD-Marker , entfernen birgt (noch) unerwünschten Seiteneffekt
    
    If gbShowToolbar Then
        
        With Toolbar1
            .Tag = "out"
            
            Call SetToolbarImage(.Buttons(1), 14)
            Call SetToolbarImage(.Buttons(2), 15)
            
            Call SetToolbarImage(.Buttons(4), 8)
            Call SetToolbarImage(.Buttons(5), 6)
            Call SetToolbarImage(.Buttons(6), 11)
            Call SetToolbarImage(.Buttons(7), 12)
            Call SetToolbarImage(.Buttons(8), 13)
            
            Call SetToolbarImage(.Buttons(10), 4)
            Call SetToolbarImage(.Buttons(11), 2)
            Call SetToolbarImage(.Buttons(12), 1)
            
            Call SetToolbarImage(.Buttons(14), 5)
            Call SetToolbarImage(.Buttons(15), 7)
        End With
        
        Call SwapToolbar("in")
        
        DoEvents ' Sonst klappt Toolbar1.Height nicht :( lg 06.08.03
        fOffset = Toolbar1.Height + 100
        
    End If
    
    For i = 0 To Me.Controls.Count - 1
        With Me.Controls(i)
            If Not TypeOf Me.Controls(i).Container Is ToolBar Then
                If TypeOf Me.Controls(i) Is Line Then
                    .Y1 = .Y1 + fOffset
                    .Y2 = .Y2 + fOffset
                Else
                    .Top = .Top + fOffset
                End If
            End If
        End With
    Next i
End Sub


Private Sub Waehrung_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Do_OLEDragDrop(Index + VScroll1.Value, Data, Effect, Button, Shift, X, Y)
    
End Sub

Private Sub WakeupTimer_Timer()
    
    Static lWakeupCnt As Long
    Dim bWakeup As Boolean
    
    If giSuspendState = 1 Then
        If gdatFallAsleepDate + myTimeSerial(0, 1, 0) < Now Then giSuspendState = 2
    End If
    
    If giSuspendState = 2 Then
        lWakeupCnt = lWakeupCnt + 1
        If lWakeupCnt >= giSleepAfterWakeup Then bWakeup = True
    End If
    
    If bWakeup Then
        giSuspendState = 0
        lWakeupCnt = 0
        Dir App.Path
        Dir gsAppDataPath
    End If
    
End Sub

Private Sub ZeilenUnloadTimer_Timer()
    
    Dim i As Integer
    
    If giSuspendState = 0 Then
    
        ZeilenUnloadTimer.Enabled = False
        
        For i = 1 To giMaxRow
            EntferneZeile (i)
        Next
        Gebot(0).Enabled = False
        Gebot(0).Enabled = True
        Gebot(0).FontBold = True
        
    End If 'giSuspendState = 0
    
    mbIsUnloadingZeilen = False
    If mbRestoreQueued Then
      mbRestoreQueued = False
      mbIsRestoring = True
      FromTaskbar
      mbIsRestoring = False
    End If
    
End Sub

Private Sub Zusatzfeld_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    
    If KeyCode = vbKeyReturn Then 'Enter , 13
        Call Zusatzfeld_LostFocus
    End If
    
End Sub

Private Sub Zusatzfeld_LostFocus()

    On Error Resume Next
    Dim sTmp As String
    
    'prüfen ob ok
    sTmp = Trim(Zusatzfeld.Text)
    
    If sTmp = " " Or Val(sTmp) = 0 Then
        sTmp = ""
        Zusatzfeld.Text = sTmp
    End If
    
    If sTmp <> "" Then
        Call AddArtikel(sTmp)
    End If
    
    Zusatzfeld.Text = ""
    
End Sub

Private Sub Zusatzfeld_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Do_OLEDragDrop 0, Data, Effect, Button, Shift, X, Y

End Sub

Public Sub SwitchToolbar(iToolbarSize As Integer, bIsVisible As Boolean)

If gbShowToolbar Then
    'erstmal Toolbar abschalten:
    Toolbar1.Visible = False
    ScaleToToolbar False
    gbShowToolbar = False
    SwapToolbar "out"
End If

If bIsVisible Then
  giToolbarSize = iToolbarSize
   gbShowToolbar = True
   InitToolbar
    'ScaleToToolbar True
   Toolbar1.Visible = True
End If

End Sub

Public Sub ScaleRelativ(fWidth As Double)
Dim fFactor As Double
Dim fNewHeight As Double

fFactor = miStartHeight / miStartWidth
fNewHeight = fWidth * fFactor
Me.Height = fNewHeight
Me.Width = fWidth

End Sub

Private Sub SetLanguage()
'
' Übernahme der Sprachtexte .. reine Fleissarbeit :-)
'
Dim i As Integer

'Menüeinträge
mnuSave.Caption = gsarrLangTxt(30)
mnuSaveArtikel.Caption = gsarrLangTxt(31)
mnuSaveSettings.Caption = gsarrLangTxt(32)
mnuSaveAll.Caption = gsarrLangTxt(33)
mnuLesen.Caption = gsarrLangTxt(34)
mnuReadArtikel.Caption = gsarrLangTxt(35)
mnuActions.Caption = gsarrLangTxt(36)
mnuUpdateArtikel.Caption = gsarrLangTxt(37)
mnuMyEbay.Caption = gsarrLangTxt(38)
mnuCleanupArtikel.Caption = gsarrLangTxt(59)
mnuCleanupArtikel2.Caption = gsarrLangTxt(751)
mnuDeleteArtikel.Caption = gsarrLangTxt(754)
mnuSync.Caption = gsarrLangTxt(39)
mnuPasswd.Caption = gsarrLangTxt(40)
mnuAuto.Caption = gsarrLangTxt(41)
mnuWindow.Caption = gsarrLangTxt(42)
mnuSettings.Caption = gsarrLangTxt(32) & "..."
mnuArtikel.Caption = gsarrLangTxt(223)
mnuAbout.Caption = gsarrLangTxt(52) & "..."
mnuFile.Caption = gsarrLangTxt(53)
mnuInfo.Caption = gsarrLangTxt(54)
mnuBrowser.Caption = gsarrLangTxt(55)
mnuBrowse.Caption = gsarrLangTxt(55)
mnuCurrUpdate.Caption = gsarrLangTxt(65)
mnuHomepage.Caption = gsarrLangTxt(213)
mnuReleasenotes.Caption = gsarrLangTxt(437)

mnuVersion.Caption = gsarrLangTxt(43)
mnuHelp.Caption = gsarrLangTxt(44)
mnuExit.Caption = gsarrLangTxt(45)

mnuComment.Caption = gsarrLangTxt(48)
mnuBid.Caption = gsarrLangTxt(49) ' alternativ auf 732 bei Sofort-Kauf
mnuAkt.Caption = gsarrLangTxt(50)
mnuDel.Caption = gsarrLangTxt(51)
mnuAccount.Caption = gsarrLangTxt(728)
mnuProductSearch.Caption = gsarrLangTxt(729)
mnuTools.Caption = gsarrLangTxt(730)
mnuSendItemTo.Caption = gsarrLangTxt(731)
mnuEditShipping.Caption = gsarrLangTxt(748)
mnuSearch.Caption = gsarrLangTxt(744) & "..."
mnuSearchContinue.Caption = gsarrLangTxt(745)
'Buttons
NewArtikel.Caption = gsarrLangTxt(31)
SortEnde.Caption = gsarrLangTxt(60)

'Label
Label7.Caption = gsarrLangTxt(66)
Label9.Caption = gsarrLangTxt(67)
Label10.Caption = gsarrLangTxt(68)
Label2.Caption = gsarrLangTxt(69)
Label11.Caption = gsarrLangTxt(70)
mnuLanguage.Caption = gsarrLangTxt(230)
        
Label7.ToolTipText = gsarrLangTxt(247)
Label9.ToolTipText = gsarrLangTxt(248)
Label10.ToolTipText = gsarrLangTxt(242)
Label2.ToolTipText = gsarrLangTxt(243)
Label11.ToolTipText = gsarrLangTxt(264)
        
'ToolTips
Toolbar1.Buttons(1).ToolTipText = gsarrLangTxt(34)
Toolbar1.Buttons(2).ToolTipText = gsarrLangTxt(30)
Toolbar1.Buttons(4).ToolTipText = gsarrLangTxt(37)
Toolbar1.Buttons(5).ToolTipText = gsarrLangTxt(38)
Toolbar1.Buttons(6).ToolTipText = gsarrLangTxt(39)
Toolbar1.Buttons(7).ToolTipText = gsarrLangTxt(40)
Toolbar1.Buttons(8).ToolTipText = gsarrLangTxt(41)
Toolbar1.Buttons(10).ToolTipText = gsarrLangTxt(32)
Toolbar1.Buttons(11).ToolTipText = Replace(gsarrLangTxt(603), "%SRV%", gsMainUrl)
Toolbar1.Buttons(12).ToolTipText = gsarrLangTxt(604)
Toolbar1.Buttons(14).ToolTipText = gsarrLangTxt(52)
Toolbar1.Buttons(15).ToolTipText = gsarrLangTxt(605)

NewArtikel.ToolTipText = gsarrLangTxt(240)
SortEnde.ToolTipText = gsarrLangTxt(241)


For i = 0 To giMaxRow
    Artikel(i).ToolTipText = gsarrLangTxt(244)
    EndeZeit(i).ToolTipText = gsarrLangTxt(31) & " " & gsarrLangTxt(60)
    Gebot(i).ToolTipText = gsarrLangTxt(242)
    Bietgruppe(i).ToolTipText = gsarrLangTxt(243)
    EndeZeit(i).ToolTipText = gsarrLangTxt(245)
    Titel(i).ToolTipText = gsarrLangTxt(246)
    Preis(i).ToolTipText = gsarrLangTxt(247)
    Restzeit(i).ToolTipText = gsarrLangTxt(248)
    Status(i).ToolTipText = gsarrLangTxt(264)
Next i

CheckAutoMode 'lg 04.05.03

End Sub

Private Sub SetSupportedLanguages()

  Dim vntLanguages As Variant
  Dim vntlanguage As Variant
  Dim i As Integer
  
  vntLanguages = Split(GetSupportedLanguage(), ",")
  
  For Each vntlanguage In vntLanguages
    If vntlanguage <> vbNullString Then
      i = i + 1
      mnul(i).Caption = vntlanguage
      mnul(i).Enabled = True
      mnul(i).Visible = True
      If vntlanguage = gsAktLanguage Then mnul(i).Checked = True
    End If
  Next

End Sub

Private Sub CheckIniVersion()

  Call CheckVersionOfLanguageFile
  Call CheckVersionOfKeywordsFile

End Sub

Private Sub ShowPopupMenu_TrayIcon()

  Dim lRet As Long
  Dim myMenu As clsPopupMenu
  
  Set myMenu = New clsPopupMenu
  
  myMenu.hWnd = Me.hWnd
  'die CStr's müssen sein, ohne gehts nicht:
  lRet = myMenu.Popup(1, CStr(gsarrLangTxt(46)), "-", CStr(gsarrLangTxt(57)), "-", CStr(gsarrLangTxt(52)) & " BOM", "-", IIf(gbAutoMode, "CHECKED_", "") & gsarrLangTxt(58), "-", CStr(gsarrLangTxt(45)))
  
  Select Case lRet
    Case 1: mnuMax_Click
    Case 3: mnuArtikel_Click
    Case 5: mnuAbout2_Click
    Case 7: gbAutoMode = Not gbAutoMode: CheckAutoMode
    Case 9: mnuExit2_Click
  End Select
  
  Set myMenu = Nothing

End Sub

Private Function IstWasZuTun(sWann As String, Optional iAnzahl As Integer, Optional sGruppe As String = "") As Boolean
    
    Dim lAktRow As Long
    Dim fNaechstAuktion As Double
    Dim sGruppeTmp As String
    
    iAnzahl = 0
    IstWasZuTun = False
    fNaechstAuktion = 999
    For lAktRow = 1 To giAktAnzArtikel
    
        Call ZeitPrüfung(lAktRow, True)
        
        With gtarrArtikelArray(lAktRow)
            If .Gebot > 0 And GetRestzeitFromItem(lAktRow) > 0 And Not .Titel = "" And .Status <= [asNixLos] Then
                sGruppeTmp = sGruppe
                If sGruppe = "" Then sGruppeTmp = .Gruppe
                sGruppeTmp = GetGruppe(sGruppeTmp)
                If GetGruppe(.Gruppe) = sGruppeTmp Then
                    iAnzahl = iAnzahl + 1
                    If GetRestzeitFromItem(lAktRow) < fNaechstAuktion Then
                        fNaechstAuktion = GetRestzeitFromItem(lAktRow)
                    End If
                End If
            End If
        End With
    Next lAktRow
    
    If fNaechstAuktion < 999 Then
        IstWasZuTun = True
        sWann = TimeLeft2String(fNaechstAuktion) 'lg 31.07.03
    End If
    
End Function

Public Sub SwapToolbar(SwapInOrOut As String)

  Dim i As Long
  
  If LCase(SwapInOrOut) = "out" And Toolbar1.Tag = "in" Then
  
    For i = 1 To Toolbar1.Buttons.Count
      Toolbar1.Buttons(i).Caption = Toolbar1.Buttons(i).Image
    Next i
    Toolbar1.ImageList = Nothing
    Toolbar1.Tag = "out"
    ImageList1.ListImages.Clear
    
  ElseIf LCase(SwapInOrOut) = "in" And Toolbar1.Tag = "out" Then
  
    If giToolbarSize = 1 Then
      ImageList1.ImageHeight = 32
      ImageList1.ImageWidth = 32
    Else
      ImageList1.ImageHeight = 16
      ImageList1.ImageWidth = 16
    End If
       
    For i = glMINRESICONS To glMAXRESICONS
      ImageList1.ListImages.Add i Mod 100, "", MyLoadResPicture(i, ImageList1.ImageHeight)
    Next i
  
    Set Toolbar1.ImageList = ImageList1
    
    For i = 1 To Toolbar1.Buttons.Count
      Toolbar1.Buttons(i).Image = Val(Toolbar1.Buttons(i).Caption)
      Toolbar1.Buttons(i).Caption = ""
    Next i
    Toolbar1.Tag = "in"
  
  End If

End Sub

Public Sub SetToolbarImage(Button As Object, Image As Integer)

  If Toolbar1.Tag = "in" Then
    Button.Image = Image
  Else
    Button.Caption = Image
  End If

End Sub

Public Sub SwapZeilen(sSwapInOrOut As String)

  Dim i As Integer
  
  If LCase(sSwapInOrOut) = "out" And Artikel(0).Tag = "in" Then
  
    On Error Resume Next
    ' try to trigger a lost-focus-event before minimizing
    If GebotTimer.Enabled Then GebotTimer_Timer
    If Gruppe_Timer.Enabled Then Gruppe_Timer_Timer
    On Error GoTo 0
    
    Artikel(0).Tag = "out"
    ZeilenUnloadTimer.Enabled = True
    mbIsUnloadingZeilen = True
  
  ElseIf LCase(sSwapInOrOut) = "in" And Artikel(0).Tag = "out" Then

    For i = 1 To giMaxRow
        GeneriereZeile (i)
    Next
    Artikel(0).Tag = "in"
    ArtikelArrayToScreen VScroll1.Value

  End If

End Sub

Public Sub PanelText(sb As StatusBar, Index As Long, aText As String, Optional Resetable As Boolean = False, _
      Optional bgColor As Long = &H8000000F, Optional fgColor As Long = &H80000012, _
      Optional FixText As Boolean = False)
  
  With picPanel
    Set Me.Font = sb.Font
    sb.Height = Me.TextHeight("IygX") + 85 'nur ein Quatsch-String
    Set .Font = sb.Font
    .Move 0, 0, 10000, Me.TextHeight("IygX") + 85
    .BackColor = bgColor
    .Cls
    .ForeColor = fgColor
    picPanel.Print " " & aText
    sb.Panels(Index).Text = aText
    Set sb.Panels(Index).Picture = .Image
    sb.Panels(Index).Tag = bgColor
  End With
  
  gdatarrPanelTimes(Index) = IIf(Resetable, MyNow, 0)
  If FixText Then
    gsarrPanelFixText(Index) = aText
    glarrPanelFixBackColor(Index) = bgColor
    glarrPanelFixForeColor(Index) = fgColor
  End If
  
End Sub

Public Sub PanelRepaint()

  On Error Resume Next
  Dim sb As StatusBar
  Dim X As Long
  
  Set sb = StatusBar1
  For X = 1 To sb.Panels.Count
    PanelText sb, X, sb.Panels(X).Text, IIf(gdatarrPanelTimes(X) <> 0, True, False), sb.Panels(X).Tag
  Next X
  
End Sub

Private Function Zeitsync() As String
    
    Dim fTimeSyncStart As Double
    Dim fTimeSyncEnd As Double
    Dim fTimeSyncBefore As Double
    Dim fTimeSyncAfter As Double
    Dim fLap As Double
    
    Call PanelText(StatusBar1, 2, gsarrLangTxt(126) & "...")
    
    'wir ziehen die Dauer des Timesync von der Differenz Uhrzeit vorher/nachher ab.
    fTimeSyncBefore = Timer()
    fTimeSyncStart = GetSystemUptime()
    
    If giUseNtp Then
        Zeitsync = GetINetTime()
    Else
        Zeitsync = sync_ebaytime()
    End If
    
    fTimeSyncAfter = Timer()
    fTimeSyncEnd = GetSystemUptime()
    fLap = fTimeSyncAfter - fTimeSyncBefore - (fTimeSyncEnd - fTimeSyncStart) + gfTimeDeviation
    
    If Zeitsync <> "" Then
        Call PanelText(StatusBar1, 2, gsarrLangTxt(74), True, vbGreen)
        If (fLap > 1 Or giDebugLevel > 1) Then
            Call DebugPrint("Zeitsync Differenz " & CStr(CInt(fLap)) & " Sekunden")
        End If
    Else
        Call PanelText(StatusBar1, 2, gsarrLangTxt(75), False, vbRed)
        Call DebugPrint("Zeitsync nicht erfolgreich")
    End If
    
End Function

Public Sub PanelReset()

  Dim X As Long
  For X = LBound(gdatarrPanelTimes) To UBound(gdatarrPanelTimes)
    If gdatarrPanelTimes(X) <> 0 Then
      If gsarrPanelFixText(X) = "" Then glarrPanelFixForeColor(X) = &H80000012: glarrPanelFixBackColor(X) = &H8000000F
      If DateDiff("s", gdatarrPanelTimes(X), MyNow) >= glCleanStatusTime Then PanelText StatusBar1, X, gsarrPanelFixText(X), False, glarrPanelFixBackColor(X), glarrPanelFixForeColor(X)
    End If
  Next

End Sub

Private Sub ShowContextMenu(Index As Integer)

  If Artikel(Index) <> "" Then
    If Toolbar1.Buttons(4).Image <> 3 Then 'nicht wenn Artikelupdate läuft
      miGlobArtikel = Index + VScroll1.Value
      If miGlobArtikel <= UBound(gtarrArtikelArray) Then
      
        If StatusIstBuyItNowStatus(gtarrArtikelArray(miGlobArtikel).Status) Then
          mnuBid.Caption = gsarrLangTxt(732)
        Else
          mnuBid.Caption = gsarrLangTxt(49)
        End If
        
        If gtarrArtikelArray(miGlobArtikel).Status = [asAdvertisement] Or _
           GetRestzeitFromItem(miGlobArtikel) = 0 Then
          mnuBid.Enabled = False
        Else
          mnuBid.Enabled = True
        End If
      
        Me.PopupMenu mnuMouse
      End If
    End If
  End If
End Sub
        
Private Function AnzValidArtikel(ByVal bNurAnzahl As Boolean) As Integer
'Artikelanzahl mit Restzeit>0 oder nächster anstehender
Dim i As Integer
Dim iTmp As Integer
Dim iZahl As Integer
Dim fDiffAn As Double
Dim tarrEZeit() As udtEzeitPos

i = 1

If bNurAnzahl Then
    Do
        If gtarrArtikelArray(i).EndeZeit - MyNow > 0 Then
            iTmp = iTmp + 1
        End If
        i = i + 1
    Loop Until i = giAktAnzArtikel + 1
    AnzValidArtikel = iTmp
    
Else
    
    ReDim tarrEZeit(0 To giAktAnzArtikel) As udtEzeitPos
    Do
        fDiffAn = gtarrArtikelArray(i).EndeZeit - MyNow()
        If fDiffAn > 0 Then
            tarrEZeit(i - iZahl).ts_Ezeit = gtarrArtikelArray(i).EndeZeit
            tarrEZeit(i - iZahl).ts_AnzPos = i
        Else
            iZahl = iZahl + 1
        End If
        i = i + 1
    Loop Until i = giAktAnzArtikel + 1
    
    ReDim Preserve tarrEZeit(0 To (UBound(tarrEZeit()) - iZahl)) As udtEzeitPos
  
    If tarrEZeit(1).ts_Ezeit < tarrEZeit(UBound(tarrEZeit())).ts_Ezeit Then
        AnzValidArtikel = tarrEZeit(1).ts_AnzPos
    Else
        AnzValidArtikel = tarrEZeit(UBound(tarrEZeit())).ts_AnzPos
    End If
End If
Erase tarrEZeit()
End Function

Public Sub Upd_Art(ByVal iAktRow As Integer, Optional ByVal sTxt As String, Optional ByVal bWait As Boolean = True)
'sh 30.10.03 aktualisierung sichtbar
On Error GoTo errhdl
Dim FM As String, eNR As Integer, iVScrollTmp As Integer
Dim datEndezeitOld As Date
Dim sItem As String

datEndezeitOld = gtarrArtikelArray(iAktRow).EndeZeit

iVScrollTmp = VScroll1.Value

If iAktRow - VScroll1.Value >= 0 And iAktRow - VScroll1.Value <= giMaxRow And _
  Me.WindowState <> vbMinimized Then
    Titel(iAktRow - VScroll1.Value).BackColor = vbYellow
End If

gtarrArtikelArray(iAktRow).UpdateInProgressSince = MyNow

sItem = gtarrArtikelArray(iAktRow).Artikel
Update_Artikel iAktRow, sTxt, bWait
iAktRow = ItemToIndex(sItem)
   
If iAktRow > 0 Then
   
  If iAktRow - iVScrollTmp >= 0 And iAktRow - iVScrollTmp <= giMaxRow Then
    If Me.WindowState <> vbMinimized Then
      If bWait Then
        Titel(iAktRow - iVScrollTmp).BackColor = &H8000000F
      End If
    End If
  End If
  
  If datEndezeitOld <> gtarrArtikelArray(iAktRow).EndeZeit Then Sortiere
End If

Exit Sub

errhdl:
FM = Err.Description
eNR = Err.Number
Err.Clear
If eNR <> 0 And eNR <> 20 Then
  DebugPrint "Upd_Art Fehler: " & eNR & "(" & FM & ") Arow: " & CStr(iAktRow) & " Maxrow: " & CStr(giMaxRow) & " vscroll:" & CStr(VScroll1.Value)
  Resume Next
End If
End Sub

Private Sub AddTrayIcon()

With mtTrayIcon
    If Not .hWnd Then
        .cbSize = Len(mtTrayIcon)
        .hWnd = Me.Picture1.hWnd
        .uId = 1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_MOUSEMOVE
        .hIcon = frmDummy.Icon
        
        Call Shell_NotifyIcon(NIM_ADD, mtTrayIcon)
    End If
End With

End Sub

Public Sub RemoveTrayIcon()
  If mtTrayIcon.hWnd Then Shell_NotifyIcon NIM_DELETE, mtTrayIcon
End Sub

Private Sub WatchArtikelZeilen()

  If Me.WindowState = vbMinimized Then
    If Artikel(0).Tag = "in" And Not mbIsMinimizing Then Form_Resize
  Else
    If Artikel(0).Tag = "out" And Not mbIsRestoring Then Form_Resize
  End If

End Sub

Private Sub WatchTrayIcon()
  On Error Resume Next
  
  Static lCounter As Long
  Static lLastTrayHwnd As Long
  Dim lCurrentTrayHwnd As Long
  
  lCounter = lCounter + 1
  If lCounter > 10 Then
    lCounter = 0

    lCurrentTrayHwnd = GetTaskBarProps("HANDLE")
    If lCurrentTrayHwnd <> lLastTrayHwnd Then
      
      lLastTrayHwnd = lCurrentTrayHwnd
      
      If mtTrayIcon.hWnd Then
        If Shell_NotifyIcon(NIM_MODIFY, mtTrayIcon) = False Then
            Call RemoveTrayIcon
            Call AddTrayIcon
        End If
      End If
    End If
    
  End If
End Sub

Private Sub ShowSplashOnce()
'zeigt frmAbout wenigstens 1mal an, auch wenn sonst
'der Splash nicht angezeigt werden soll
'wird in Form.Load aufgerufen
Dim lFile As String
Dim tmpstr As String

lFile = gsAppDataPath & "\Settings.ini"

If INIGetValue(lFile, "Diverses", "ShowSplashOnce", tmpstr) = 0 Then tmpstr = "0"

'If Val(tmpstr) < 1 Then
'    tmp = INISetValue(lFile, "Diverses", "ShowSplashOnce", 1)
'    frmAbout.Timer1.Interval = 10000 '10 Sekunden - nur beim ersten Mal
'    frmAbout.SetShowSplashOnly
'    frmAbout.ShowAWhile
'    frmAbout.Show vbModal, Me
'    Exit Sub
'End If

If gbShowSplash And giStartupSize <> vbMinimized Then
    'Zuerst das entsprechende Property setzen , dann
    frmAbout.SetShowSplashOnly
    'frmAbout.SetSpendeActiv
    Load frmAbout ' und anschließend anzeigen
    frmAbout.Show
    SetForeground True, frmAbout.hWnd
End If

End Sub

Private Sub DoTheJob(sJobfile As String)
    
  Dim sTmp As String
  
  Dim sUsername As String
  Dim sPassword As String
  Dim sItemCode As String
  Dim sBidPrice As String
  Dim bIsBuyItNow As Boolean
  Dim datEndTime As Date
  Dim sStatus As String
  Dim bResult As Boolean
  Dim sResultFile As String
  
  AutoSave.Enabled = False
  ReDim gtarrArtikelArray(1 To 1) As udtArtikelZeile
  giAktAnzArtikel = 1
  
  If InStr(1, sJobfile, "\") <= 0 Then sJobfile = ".\" & sJobfile
  
  If Dir(sJobfile) = "" Then
    Call DebugPrint("Job file not found, exiting! (" & sJobfile & ")")
    Exit Sub
  End If
  Call DebugPrint("Parsing job file")
  
  If INIGetValue(sJobfile, "RESULT", "STATUS", sTmp) <> 0 Then sStatus = Trim(sTmp)
  If sStatus > "" Then Call DebugPrint("Job is already done, exiting!"): Exit Sub
  
  If INIGetValue(sJobfile, "JOB", "USERNAME", sTmp) <> 0 Then sUsername = Trim(sTmp)
  If INIGetValue(sJobfile, "JOB", "PASSWORD", sTmp) <> 0 Then sPassword = Trim(sTmp)
  If INIGetValue(sJobfile, "JOB", "ITEMCODE", sTmp) <> 0 Then sItemCode = Trim(sTmp)
  If INIGetValue(sJobfile, "JOB", "BIDPRICE", sTmp) <> 0 Then sBidPrice = Trim(sTmp)
  If INIGetValue(sJobfile, "JOB", "BUYITNOW", sTmp) <> 0 Then bIsBuyItNow = CBool(sTmp)
  If INIGetValue(sJobfile, "JOB", "ENDTIME", sTmp) <> 0 Then datEndTime = CDate(sTmp)
  If INIGetValue(sJobfile, "JOB", "TIMEDIV", sTmp) <> 0 Then gfTimeDeviation = CDbl(sTmp)
  
  If sUsername = "" Then sUsername = gsUser: sPassword = gsPass
  If sPassword = "" Then
    If UsrAccToIndex(sUsername) > 0 Then sPassword = gtarrUserArray(UsrAccToIndex(sUsername)).UaPass
  End If
  
  If sUsername = "" Then Call DebugPrint("Username missing, exiting!"): sStatus = "ERROR_USERNAME_MISSING"
  If sPassword = "" Then Call DebugPrint("Password missing, exiting!"): sStatus = "ERROR_PASSWORD_MISSING"
  If sItemCode = "" Then Call DebugPrint("Itemcode missing, exiting!"): sStatus = "ERROR_ITEMCODE_MISSING"
  If sBidPrice = "" Then Call DebugPrint("Bidprice missing, exiting!"): sStatus = "ERROR_BIDPRICE_MISSING"

  If sStatus = "" Then
    Call DebugPrint("Bidding " & sBidPrice & " on item " & sItemCode & " with user " & sUsername & " and password " & String(Len(sPassword), "*") & "")
    giDefaultUser = 0
    gtarrArtikelArray(1).Artikel = sItemCode
    gtarrArtikelArray(1).EndeZeit = datEndTime
    gtarrArtikelArray(1).TimeZone = GetUTCOffset
    bResult = Bieten(sItemCode, sBidPrice, sUsername, sPassword, 1, IIf(datEndTime > 0, False, True), False, bIsBuyItNow)
    If bResult Then
      sStatus = "SUCCESS"
    ElseIf miErrStatus = [asUeberboten] Then
      sStatus = "OUTBID"
    Else
      sStatus = "ERROR_GENERAL_ERROR"
    End If
    Call DebugPrint("Result: " & sStatus)
  End If
  
  Call DebugPrint("Writing result")
  Call INISetValue(sJobfile, "RESULT", "STATUS", sStatus)
  
  Call DebugPrint("Renaming job file")
  
  sResultFile = Left(sJobfile, Len(sJobfile) - 3) & "result"
  If Dir(sResultFile) > "" Then Kill sResultFile
  Name sJobfile As sResultFile
   
  Call DebugPrint("All done, exiting!")
   
End Sub

Public Sub ReadLocalServStrings()
Dim sPath As String

On Error Resume Next
ReDim gsarrServerStrArr(0)

sPath = Dir(App.Path & "\Serverstrings_*.ini")
Do While Len(sPath) > 0
  'DebugPrint "Dateiname: " & sPath
  If Not LCase(getFile(sPath)) Like "*.new.*" Then
    ReDim Preserve gsarrServerStrArr(UBound(gsarrServerStrArr) + 1)
    gsarrServerStrArr(UBound(gsarrServerStrArr)) = sPath
    'DebugPrint "ServerStrArr-" & UBound(gsarrServerStrArr) & ": " & sPath
  End If
  sPath = Dir
Loop

End Sub

Public Sub CheckSofortkaufArtikel(Optional ByVal sGruppe As String = "")
        
    Dim i As Integer
    Dim iAnzahl As Integer
    Dim sGruppeTmp As String
    
    
    If gbBuyItNow Then
        If gbAutoMode Then
            If Not mbAlreadyRunning Then
                mbAlreadyRunning = True
                
                For i = 1 To giAktAnzArtikel
                    With gtarrArtikelArray(i)
                        sGruppeTmp = sGruppe
                        If sGruppeTmp = "" Then sGruppeTmp = .Gruppe
                        If .Status = [asBuyOnly] And GetGruppe(.Gruppe) = GetGruppe(sGruppeTmp) And .Gruppe > "" And GetRestzeitFromItem(i) > 0 Then
                            Call IstWasZuTun("", iAnzahl, .Gruppe)
                            If iAnzahl < GetAnzahlVonGruppe(.Gruppe) Then ' es gibt weniger Artikel zum Bebieten als wir brauchen
                                If MsgBoxEx(gsarrLangTxt(732) & ": " & .Artikel & vbCrLf & .Titel & " ?", gsarrLangTxt(408) & " [10]*-" & gsarrLangTxt(409) & "-") = 1 Then
                                    .Status = [asBuyOnlyBuyItNow]
                                    .LastChangedId = GetChangeID()
                                    Call Timer1_Timer
                                Else
                                    gbAutoMode = False
                                    Call CheckAutoMode
                                End If
                            End If
                        End If
                    End With
                Next i
                
                mbAlreadyRunning = False
            End If
        End If
    End If
End Sub

Private Sub CheckAlleBietgruppen()
    
    Dim i As Integer
    Dim sGruppeTmp As String
    Dim colTmp As Collection
    
    Set colTmp = New Collection
    
    On Error Resume Next
    
    For i = 1 To giAktAnzArtikel
        sGruppeTmp = GetGruppe(gtarrArtikelArray(i).Gruppe)
        Call colTmp.Add(sGruppeTmp, sGruppeTmp)
    Next i
    
    Do While colTmp.Count > 0
        Call CheckBietgruppe(colTmp.Item(1))
        Call colTmp.Remove(1)
    Loop
    
    Call CheckSofortkaufArtikel
    
    Set colTmp = Nothing
End Sub

Private Function GetGruppe(ByVal sGruppeMitAnzahl As String) As String
  
  Dim iPos As Integer
 
  GetGruppe = sGruppeMitAnzahl
  iPos = InStr(1, GetGruppe, ";")
  If iPos > 0 Then GetGruppe = Left(GetGruppe, iPos - 1)
  GetGruppe = Trim(GetGruppe)
  iPos = InStr(1, GetGruppe, ":")
  If iPos > 0 Then
    GetGruppe = Trim(Mid(GetGruppe, iPos + 1))
  End If
  
End Function

Private Function GetMasterGruppe(ByVal sGruppeMitAnzahl As String) As String
  
  Dim iPos As Integer
 
  GetMasterGruppe = sGruppeMitAnzahl
  iPos = InStr(1, GetMasterGruppe, ";")
  If iPos > 0 Then GetMasterGruppe = Left(GetMasterGruppe, iPos - 1)
  GetMasterGruppe = Trim(GetMasterGruppe)
  iPos = InStr(1, GetMasterGruppe, ":")
  If iPos > 0 Then
    GetMasterGruppe = Trim(Mid(GetMasterGruppe, 1, iPos - 1))
  End If
  
End Function

Private Function GetAnzahlVonGruppe(ByVal sGruppeMitAnzahl As String) As Integer
  
  Dim lPos As Integer
  GetAnzahlVonGruppe = 1
 
  lPos = InStr(1, sGruppeMitAnzahl, ";")
  If lPos > 0 Then
    GetAnzahlVonGruppe = Val(Mid(sGruppeMitAnzahl, lPos + 1))
  End If
  
End Function

Public Sub CheckBietgruppe(ByVal sGruppe As String)
    
  Dim i As Integer
  Dim j As Integer
  Dim sVorzeichen As String
  Dim bGruppeAktiv As Boolean
  Dim bGruppeVerloren As Boolean
  Dim bVergleichwert As Boolean
  Dim sUntergruppe As String
  Dim sGruppeTmp As String
  Dim iAnzahl As Integer
  Dim sHauptgruppe As String
  
  bGruppeVerloren = True ' erstmal annehmen, dass alles zu spät ist

  sHauptgruppe = GetGruppe(sGruppe)
  
  ' Ist min. 1 Hauptartikel OK ?
  For i = 1 To giAktAnzArtikel
    sGruppeTmp = GetGruppe(gtarrArtikelArray(i).Gruppe)
    If sGruppeTmp = sHauptgruppe Then
      If gtarrArtikelArray(i).Status = [asOK] Then
        bGruppeAktiv = True
        bGruppeVerloren = False
        Exit For
      End If
      If (gtarrArtikelArray(i).Status = [asHoldGroup] And gtarrArtikelArray(i).Gebot > 0) Or _
         (gtarrArtikelArray(i).Status = [asBuyOnlyOnHold] And gbBuyItNow) Then  ' da geht noch was, der wird vielleicht gleich noch gekauft
        bGruppeVerloren = False
        Exit For
      End If
    End If
  Next i
  
  If Not bGruppeAktiv Then ' kein Artikel gewonnen, mal schaun ob noch Möglichkeiten kommen:
    Call IstWasZuTun("", iAnzahl, sHauptgruppe)
    If iAnzahl > 0 Then bGruppeVerloren = False ' da geht noch was...
  End If
  
    
  ' Alle Unterartikel prüfen und Hold setzten/entfernen
  For i = 1 To giAktAnzArtikel
    sUntergruppe = GetMasterGruppe(gtarrArtikelArray(i).Gruppe)
    
    For j = 1 To 2
      If j = 1 Then sVorzeichen = "+": bVergleichwert = bGruppeAktiv
      If j = 2 Then sVorzeichen = "-": bVergleichwert = bGruppeVerloren

      If sUntergruppe = sVorzeichen & sHauptgruppe Then
        If gtarrArtikelArray(i).Status <= [asNixLos] And Not bVergleichwert Then
          gtarrArtikelArray(i).Status = [asHoldGroup]
          gtarrArtikelArray(i).LastChangedId = GetChangeID()
          Call DebugPrint("Hold Art: " & gtarrArtikelArray(i).Titel & "(" & gtarrArtikelArray(i).Artikel & ")")
        ElseIf gtarrArtikelArray(i).Status = [asBuyOnly] And Not bVergleichwert Then
          gtarrArtikelArray(i).Status = [asBuyOnlyOnHold]
          gtarrArtikelArray(i).LastChangedId = GetChangeID()
          Call DebugPrint("Hold Art: " & gtarrArtikelArray(i).Titel & "(" & gtarrArtikelArray(i).Artikel & ")")
        End If
        If gtarrArtikelArray(i).Status = [asHoldGroup] And bVergleichwert Then
          gtarrArtikelArray(i).Status = [asNixLos]
          gtarrArtikelArray(i).LastChangedId = GetChangeID()
          Call DebugPrint("Unhold Art: " & gtarrArtikelArray(i).Titel & "(" & gtarrArtikelArray(i).Artikel & ")")
        ElseIf gtarrArtikelArray(i).Status = [asBuyOnlyOnHold] And bVergleichwert Then
          gtarrArtikelArray(i).Status = [asBuyOnly]
          gtarrArtikelArray(i).LastChangedId = GetChangeID()
          Call DebugPrint("Unhold Art: " & gtarrArtikelArray(i).Titel & "(" & gtarrArtikelArray(i).Artikel & ")")
        End If
      ElseIf Left(gtarrArtikelArray(i).Gruppe, 1) <> "+" And Left(gtarrArtikelArray(i).Gruppe, 1) <> "-" And gtarrArtikelArray(i).Status = [asHoldGroup] Then
        gtarrArtikelArray(i).Status = [asNixLos]
        gtarrArtikelArray(i).LastChangedId = GetChangeID()
        Call DebugPrint("Unhold Art: " & gtarrArtikelArray(i).Titel & "(" & gtarrArtikelArray(i).Artikel & ")")
      ElseIf Left(gtarrArtikelArray(i).Gruppe, 1) <> "+" And Left(gtarrArtikelArray(i).Gruppe, 1) <> "-" And gtarrArtikelArray(i).Status = [asBuyOnlyOnHold] Then
        gtarrArtikelArray(i).Status = [asBuyOnly]
        gtarrArtikelArray(i).LastChangedId = GetChangeID()
        Call DebugPrint("Unhold Art: " & gtarrArtikelArray(i).Titel & "(" & gtarrArtikelArray(i).Artikel & ")")
      End If
    
    Next j
    
  Next i
  
  If sGruppe Like "+*" Then Call CheckBietgruppe(Mid(GetMasterGruppe(sGruppe), 2))
  If sGruppe Like "-*" Then Call CheckBietgruppe(Mid(GetMasterGruppe(sGruppe), 2))
  
  Call ArtikelArrayToScreen(VScroll1.Value)
  
End Sub

Sub SendenAn(iIdx As Integer)

  Dim sTmp As String
  Dim oRC4 As clsRC4
  
  sTmp = InputBox(gsarrLangTxt(731), App.Title, gsSendItemTo)
  If sTmp > "" Then
    gsSendItemTo = Trim(sTmp)
    sTmp = "B-O-M: ADDMOD ART " & gtarrArtikelArray(iIdx).Artikel & " EUR " & gtarrArtikelArray(iIdx).Gebot
    If gtarrArtikelArray(iIdx).Gruppe > "" Then sTmp = sTmp & " GRUPPE " & gtarrArtikelArray(iIdx).Gruppe
    If gtarrArtikelArray(iIdx).UserAccount > "" Then sTmp = sTmp & " ACCOUNT " & gtarrArtikelArray(iIdx).UserAccount
    
    If gbSendItemEncrypted Then
      
      Set oRC4 = New clsRC4
      sTmp = Replace(oRC4.EncryptString(sTmp, gsPass, True), vbCrLf, "")
      Set oRC4 = Nothing
    End If
    
    sTmp = "To: " & gsSendItemTo & vbCrLf & "Subject: " & sTmp
    
    InsertMailBuff sTmp & vbCrLf
    If StatusIstBuyItNowStatus(gtarrArtikelArray(iIdx).Status) Then
      gtarrArtikelArray(iIdx).Status = [asBuyOnlyDelegated]
    Else
      gtarrArtikelArray(iIdx).Status = [asDelegatedBom]
    End If
    gtarrArtikelArray(iIdx).LastChangedId = GetChangeID()
    SaveSetting "Diverses", "SendItemTo", gsSendItemTo
    ArtikelArrayToScreen VScroll1.Value
    CheckBietgruppe gtarrArtikelArray(iIdx).Gruppe
    CheckSofortkaufArtikel
  End If

End Sub

Public Function StatusIstBuyItNowStatus(lStatus As Long) As Boolean
    
    If lStatus = [asBuyOnly] Or lStatus = [asBuyOnlyBuyItNow] _
        Or lStatus = [asBuyOnlyCanceled] Or lStatus = [asBuyOnlyOnHold] _
        Or lStatus = [asBuyOnlyDelegated] Then StatusIstBuyItNowStatus = True
                
End Function

Private Function GetGebotColorFromItem(ByVal iIdx As Integer) As Long
    
    GetGebotColorFromItem = vbWindowBackground
    With gtarrArtikelArray(iIdx)
        'If .Artikel = "220115309795" Then Stop
        If .AnzGebote = 0 Then GetGebotColorFromItem = vbGreen
        If .Gebot > 0 Then
            If .MinGebot > .Gebot Then GetGebotColorFromItem = vbYellow
        End If
        If StatusIstBuyItNowStatus(.Status) Then GetGebotColorFromItem = RGB(200, 200, 200)
        If .Status = [asAdvertisement] Then GetGebotColorFromItem = RGB(200, 200, 200)
    End With
    
End Function

Private Function GetRestzeitFromItem(ByVal iIdx As Integer) As Double

  GetRestzeitFromItem = gtarrArtikelArray(iIdx).EndeZeit - MyNow
  If GetRestzeitFromItem < 0 Then GetRestzeitFromItem = 0

End Function

Private Sub Titel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      
    If gbBrowseInline Then
        
        If giInlineBrowserModifierKey > 0 And giInlineBrowserModifierKey <> Shift Then
            'Call DebugPrint("BrowseInline : " & CStr(gbBrowseInline & "  ModifierKey :   " & CStr(giInlineBrowserModifierKey) & "  Shift :  " & CStr(Shift)) & "  frmInfo-Sichtbar :  Falsch")
            Call HideInfo
        Else
        
            If miMouseIndex <> Index Then
                'Call DebugPrint("BrowseInline : " & CStr(gbBrowseInline & "  ModifierKey :   " & CStr(giInlineBrowserModifierKey) & "  Shift :  " & CStr(Shift)) & "  frmInfo-Sichtbar :  Wahr")
                If ShowInfo.Enabled Then
                    ShowInfo.Enabled = False
                End If
            End If
            
            miMouseIndex = Index
            
            If Not ShowInfo.Enabled Then
                ShowInfo.Interval = giInlineBrowserDelay / 2
                ShowInfo.Enabled = True
            End If
            
        End If 'giInlineBrowserModifierKey > 0 And
    End If 'gbBrowseInline
End Sub

Private Sub EndeZeit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call HideInfo
End Sub

Public Sub HideInfo()

  miMouseIndex = -1
  miShowIndex = -1
  miTmpIndex = -1
  frmInfo.Hide

End Sub

Private Sub ShowInfo_Timer()
    
    On Error Resume Next
    Dim iTmp As Integer
    Dim sTmp, sBuffer As String
    Dim sKommando As String
    
    If giSuspendState = 0 Then
    
        ShowInfo.Enabled = False
        'Debug.Print "MouseIndex: " & miMouseIndex, "TmpIndex: " & miTmpIndex
        
        If miMouseIndex <> miTmpIndex Then
            'verzoegerung
            iTmp = miMouseIndex
            miMouseIndex = miTmpIndex
            miTmpIndex = iTmp
            'Exit Sub
        Else
        
            If miMouseIndex >= 0 Then
                If miMouseIndex <> miShowIndex And Len(Artikel(miMouseIndex).Text) > 3 Then
                    miShowIndex = miMouseIndex
                    sBuffer = ""
                    
                    sTmp = gsCmdViewItem
                    sTmp = Replace(sTmp, "[Item]", Artikel(miShowIndex).Text)
                    
                    sKommando = sBuffer & sTmp
                    
                    gsGlobalUrl = "http://" & gsScript4 & gsScriptCommand4 & sKommando
                    frmInfo.Show
                    
                    If Not gsGlobalUrl = frmInfo.WebBrowser1.LocationURL Then
                        frmInfo.WebBrowser1.Navigate gsGlobalUrl
                    End If
                End If
            Else
                If miShowIndex >= 0 Then
                    miShowIndex = -1
                    frmInfo.Hide
                    ShowInfo.Enabled = False
                End If
            End If
        
        End If 'miMouseIndex <> miTmpIndex
        
    End If 'giSuspendState = 0
    
End Sub

Public Sub SetFocusRect(ByVal iIdx As Integer)

  gbWarSchonWach = True

  On Error GoTo ERR_HANDLER
  Static iLastIdx As Integer
  Static oLastLineAbove As Object
  Static oLastLineBelow As Object
  Dim oLineAbove As Object
  Dim oLineBelow As Object

  
  If VersandkostenEdit.Tag = "-" Then Exit Sub    ' brauchen wir, weil sonst bei Klick auf ein nicht-Edit-Feld einer anderen Zeile der Fokus innerhalb der vorherigen Zeile bleibt
  If iIdx = -1 Then iIdx = iLastIdx: iLastIdx = -1
  If iIdx = -2 Then Fokus.Visible = False: Exit Sub
  If iIdx = iLastIdx Then Exit Sub
  iLastIdx = iIdx
  
  On Error Resume Next
  oLastLineAbove.Visible = True
  oLastLineBelow.Visible = True
  On Error GoTo ERR_HANDLER
  
  If gbShowFocusRect And iIdx >= 0 Then
    Fokus.Visible = True
    Fokus.BorderColor = glFocusRectColor
    Set oLineBelow = Line1(iIdx)
    If iIdx > 0 Then
      Set oLineAbove = Line1(iIdx - 1)
    Else
      Set oLineAbove = Line12
    End If
    Fokus.Move Artikel(iIdx).Left - 2 * Screen.TwipsPerPixelX, oLineAbove.Y1, Status(iIdx).Left + Status(iIdx).Width + 2 * Screen.TwipsPerPixelX - (Artikel(iIdx).Left - 2 * Screen.TwipsPerPixelX), oLineBelow.Y2 - oLineAbove.Y1 + 20
    oLineAbove.Visible = False
    oLineBelow.Visible = False
    Set oLastLineAbove = oLineAbove
    Set oLastLineBelow = oLineBelow
  Else
ERR_HANDLER:
    Fokus.Visible = False
  End If

End Sub

Private Sub CallExtCmd(iIdx As Integer, sExtCmd As String)
    
    Dim sCmd As String
    Dim vntKeyname As Variant
    Dim vntKeywert As Variant
    Dim i As Integer
    
    sCmd = sExtCmd
    With gtarrArtikelArray(iIdx)
        vntKeyname = Array("url", "seller", "item", "highbidder", "title", "location", "price", "currency", "group", "comment", "bid", "endtime", "bidcount", "minbid", "timeleft", "timenext", "status", "timediv")
        vntKeywert = Array("http://" & gsScript4 & gsScriptCommand4 & gsCmdViewItem, .Verkaeufer, .Artikel, .Bieter, .Titel, .Standort, Format(.AktPreis, "###,##0.00"), .WE, .Gruppe, .Kommentar, Format(.Gebot, "###,##0.00"), Date2Str(.EndeZeit), .AnzGebote, Format(.MinGebot, "###,##0.00"), Abs(Date2UnixDate(.EndeZeit) - Date2UnixDate(MyNow)), Date2UnixDate(MyNow + gfRestzeitZaehler) - Date2UnixDate(MyNow), IIf(.Status = [asOK], 1, 0), gfTimeDeviation)
    End With
    
    For i = LBound(vntKeywert) To UBound(vntKeywert)
        sCmd = Replace(sCmd, "[" & vntKeyname(i) & "]", vntKeywert(i), , , vbTextCompare)
    Next i
    
    On Error Resume Next
    If sCmd > "" Then Call Shell(sCmd, giExtCmdWindowStyle)
    On Error GoTo 0
    
End Sub

Private Function GetSecurityToken(sTxtIn) As String
    
    'MD-Marker Neu 20090410
    frmSecurityToken.SecurityToken = sTxtIn
    Load frmSecurityToken
    frmSecurityToken.Show vbModal, Me
    '...
    GetSecurityToken = frmSecurityToken.SecurityToken
    Unload frmSecurityToken
    
End Function
