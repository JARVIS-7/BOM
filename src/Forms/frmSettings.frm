VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Biet-O-Matic Einstellungen"
   ClientHeight    =   6930
   ClientLeft      =   1800
   ClientTop       =   1605
   ClientWidth     =   9495
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "frmSettings"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7407.322
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   9678.036
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnVerwerfen 
      Caption         =   "Zur¸cksetzen"
      Height          =   375
      Left            =   6480
      TabIndex        =   152
      ToolTipText     =   "Einstellungen erweitern / reduzieren"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton btnSpeichern 
      Caption         =   "Speichern"
      Height          =   375
      Left            =   3360
      TabIndex        =   150
      ToolTipText     =   "Einstellungen ¸bernehmen und speichern"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton btnAbbruch 
      Cancel          =   -1  'True
      Caption         =   "Ende"
      Height          =   375
      Left            =   8040
      TabIndex        =   153
      ToolTipText     =   "Einstellungsdialog schliesen"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton btnUebernehmen 
      Caption         =   "‹bernehmen"
      Height          =   375
      Left            =   4920
      TabIndex        =   151
      ToolTipText     =   "Einstellungen ¸bernehmen"
      Top             =   6360
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5685
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10028
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Bieten"
      TabPicture(0)   =   "frmSettings.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Verbindung"
      TabPicture(1)   =   "frmSettings.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(1)=   "Frame12"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Automatik"
      TabPicture(2)   =   "frmSettings.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "Frame19"
      Tab(2).Control(3)=   "Frame21"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Mail"
      TabPicture(3)   =   "frmSettings.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame13"
      Tab(3).Control(1)=   "Frame14"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Server"
      TabPicture(4)   =   "frmSettings.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label58"
      Tab(4).Control(1)=   "Label54"
      Tab(4).Control(2)=   "Label53"
      Tab(4).Control(3)=   "Label19"
      Tab(4).Control(4)=   "Label34"
      Tab(4).Control(5)=   "Label18"
      Tab(4).Control(6)=   "Label17"
      Tab(4).Control(7)=   "Label16"
      Tab(4).Control(8)=   "txtServer1"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "txtServer2"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "txtServer3"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "txtMainUrl"
      Tab(4).Control(12)=   "txtServer4"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "txtServer5"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "cboServerStrings"
      Tab(4).ControlCount=   15
      TabCaption(5)   =   "Anzeige"
      TabPicture(5)   =   "frmSettings.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame4"
      Tab(5).Control(1)=   "Frame2"
      Tab(5).Control(2)=   "Frame1"
      Tab(5).Control(3)=   "Frame15"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Zeitsync"
      TabPicture(6)   =   "frmSettings.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame9"
      Tab(6).Control(1)=   "Frame6"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "ODBC"
      TabPicture(7)   =   "frmSettings.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame18"
      Tab(7).ControlCount=   1
      Begin VB.ComboBox cboServerStrings 
         Height          =   315
         Left            =   -74400
         Style           =   2  'Dropdown-Liste
         TabIndex        =   98
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtServer5 
         Height          =   285
         Left            =   -70590
         TabIndex        =   227
         TabStop         =   0   'False
         Text            =   "cgi.x.de"
         Top             =   4830
         Width           =   3405
      End
      Begin VB.TextBox txtServer4 
         Height          =   285
         Left            =   -70590
         TabIndex        =   225
         TabStop         =   0   'False
         Text            =   "cgi.x.de"
         Top             =   4230
         Width           =   3405
      End
      Begin VB.Frame Frame21 
         Caption         =   "Aktualisierungsoptionen"
         Height          =   2080
         Left            =   -74700
         TabIndex        =   217
         Top             =   3480
         Width           =   7995
         Begin VB.CheckBox chkMultiAkt 
            Caption         =   "gleichzeitiges Aktualisieren"
            Height          =   195
            Left            =   5400
            TabIndex        =   66
            ToolTipText     =   "Die Artikel werden im Zyklus x sec je 1 Artikel aktualisiert"
            Top             =   240
            Width           =   2505
         End
         Begin VB.TextBox txtReLogin 
            Alignment       =   2  'Zentriert
            Height          =   285
            Left            =   600
            TabIndex        =   74
            Text            =   "Text1"
            Top             =   1710
            Width           =   375
         End
         Begin VB.CheckBox chkReLogin 
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   1730
            Width           =   255
         End
         Begin VB.TextBox txtReloadTimes 
            Alignment       =   2  'Zentriert
            Height          =   285
            Left            =   240
            TabIndex        =   72
            Text            =   "3"
            Top             =   1350
            Width           =   375
         End
         Begin VB.CheckBox chkAktualisieren 
            Caption         =   "bis Auktionsende"
            Height          =   195
            Left            =   260
            TabIndex        =   63
            ToolTipText     =   "Die Artikel werden im Zyklus x sec je 1 Artikel aktualisiert"
            Top             =   285
            Width           =   1665
         End
         Begin VB.TextBox txtAktCycle 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   300
            Left            =   4020
            TabIndex        =   65
            Text            =   "60"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chkAktualXvor 
            Caption         =   "bis"
            Enabled         =   0   'False
            Height          =   195
            Left            =   260
            TabIndex        =   67
            ToolTipText     =   "Bis X-min vor Auktionsende wird alle X-min jeweils 1 Artikel aktualisiert, danach zyklisch alle x-sec"
            Top             =   645
            Width           =   735
         End
         Begin VB.TextBox txtAktXminvor 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            MaxLength       =   2
            TabIndex        =   68
            Text            =   "3"
            ToolTipText     =   "Minimum 3 Minuten"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtAktXminvorCycle 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   300
            Left            =   4020
            MaxLength       =   2
            TabIndex        =   69
            Text            =   "10"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cboArtAktOptions 
            Height          =   315
            ItemData        =   "frmSettings.frx":00EC
            Left            =   1920
            List            =   "frmSettings.frx":00F9
            Style           =   2  'Dropdown-Liste
            TabIndex        =   70
            ToolTipText     =   "Auswahloptionen"
            Top             =   960
            Width           =   2040
         End
         Begin VB.TextBox txtArtAktOptionsValue 
            Alignment       =   1  'Rechts
            Height          =   300
            Left            =   4020
            TabIndex        =   71
            Text            =   "0"
            Top             =   975
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cboAktualisierenOpt 
            Height          =   315
            ItemData        =   "frmSettings.frx":0131
            Left            =   1920
            List            =   "frmSettings.frx":013B
            Style           =   2  'Dropdown-Liste
            TabIndex        =   64
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblReLogin 
            Caption         =   "Minuten vor n‰chster Auktion automatisch einloggen"
            Height          =   255
            Left            =   1080
            TabIndex        =   232
            Top             =   1750
            Width           =   6615
         End
         Begin VB.Label lblReloadDescription 
            Caption         =   "Aktualisierungsversuche"
            Height          =   255
            Left            =   720
            TabIndex        =   231
            Top             =   1395
            Width           =   7215
         End
         Begin VB.Label Label51 
            Caption         =   "min"
            Height          =   195
            Left            =   5280
            TabIndex        =   224
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label24 
            Caption         =   "sec "
            Height          =   195
            Left            =   4665
            TabIndex        =   223
            Top             =   285
            Width           =   795
         End
         Begin VB.Label Label47 
            Caption         =   "min vor Auktionsende alle "
            Enabled         =   0   'False
            Height          =   195
            Left            =   1905
            TabIndex        =   222
            Top             =   660
            Width           =   2055
         End
         Begin VB.Label Label48 
            Caption         =   "min"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4665
            TabIndex        =   221
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label49 
            Caption         =   "Artikelauswahl:"
            Height          =   195
            Left            =   525
            TabIndex        =   220
            Top             =   1005
            Width           =   1335
         End
         Begin VB.Label Label50 
            Caption         =   "Artikel"
            Height          =   195
            Left            =   4665
            TabIndex        =   219
            Top             =   1020
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label52 
            Alignment       =   2  'Zentriert
            Caption         =   "alle"
            Height          =   255
            Left            =   3480
            TabIndex        =   218
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "POP"
         Height          =   4695
         Left            =   -74640
         TabIndex        =   209
         Top             =   840
         Width           =   3855
         Begin VB.CommandButton btnTestPop 
            Caption         =   "POP Test"
            Height          =   455
            Left            =   2520
            TabIndex        =   85
            ToolTipText     =   "POP- und SMTP- Einstellungen testen"
            Top             =   4080
            Width           =   1110
         End
         Begin VB.CheckBox chkPopEncryptedOnly 
            Caption         =   "Nur passwortverschl. Mails akzeptieren"
            Height          =   195
            Left            =   240
            TabIndex        =   83
            ToolTipText     =   "SMTP Authentifizierung benutzen"
            Top             =   3240
            Width           =   3495
         End
         Begin VB.CheckBox chkPopUseSSL 
            Caption         =   "SSL"
            Height          =   195
            Left            =   240
            TabIndex        =   84
            Top             =   3540
            Width           =   3495
         End
         Begin VB.TextBox txtPopPort 
            Height          =   315
            Left            =   2640
            TabIndex        =   76
            Top             =   540
            Width           =   915
         End
         Begin VB.TextBox txtAbsender 
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   240
            TabIndex        =   79
            ToolTipText     =   "Erlaubte Absender; leer: alle, sonst Mailadressen durch Semikolon getrennt"
            Top             =   2250
            Width           =   3285
         End
         Begin VB.CheckBox chkUsePop 
            Caption         =   "POP- Zugriff alle"
            Height          =   300
            Left            =   240
            TabIndex        =   81
            ToolTipText     =   "Lesen von Update- Befehlen aus einem POP- Mailserver"
            Top             =   2920
            Width           =   1500
         End
         Begin VB.TextBox txtPopServer 
            Height          =   315
            Left            =   240
            TabIndex        =   75
            ToolTipText     =   "Name des Mailservers"
            Top             =   540
            Width           =   2175
         End
         Begin VB.TextBox txtPopUser 
            Height          =   315
            Left            =   240
            TabIndex        =   77
            Top             =   1110
            Width           =   3315
         End
         Begin VB.TextBox txtPopPass 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   78
            Top             =   1680
            Width           =   3315
         End
         Begin VB.TextBox txtPopTimeOut 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   1800
            TabIndex        =   80
            Text            =   "10"
            ToolTipText     =   "Wie lange soll auf Antwort des POP- Servers gewartet werden"
            Top             =   2610
            Width           =   495
         End
         Begin VB.TextBox txtPopZykl 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   300
            Left            =   1800
            TabIndex        =   82
            Text            =   "60"
            ToolTipText     =   "Wie oft soll der POP- Zugriff durchgef¸hrt werden"
            Top             =   2910
            Width           =   495
         End
         Begin VB.Label Label55 
            Caption         =   "Port"
            Height          =   225
            Left            =   2640
            TabIndex        =   229
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label10 
            Caption         =   "POP- Timeout "
            Height          =   240
            Left            =   240
            TabIndex        =   216
            Top             =   2655
            Width           =   1500
         End
         Begin VB.Label Label13 
            Caption         =   "Akzeptierte Absender"
            Height          =   255
            Left            =   240
            TabIndex        =   215
            Top             =   2025
            Width           =   3240
         End
         Begin VB.Label Label7 
            Caption         =   "POP- Server"
            Height          =   225
            Left            =   240
            TabIndex        =   214
            Top             =   315
            Width           =   2160
         End
         Begin VB.Label Label6 
            Caption         =   "Minuten"
            Height          =   285
            Left            =   2520
            TabIndex        =   213
            Top             =   2955
            Width           =   1200
         End
         Begin VB.Label Label8 
            Caption         =   "Username"
            Height          =   180
            Left            =   240
            TabIndex        =   212
            Top             =   885
            Width           =   3240
         End
         Begin VB.Label Label9 
            Caption         =   "Passwort"
            Height          =   180
            Left            =   240
            TabIndex        =   211
            Top             =   1455
            Width           =   3360
         End
         Begin VB.Label Label11 
            Caption         =   "Sec."
            Height          =   240
            Left            =   2505
            TabIndex        =   210
            Top             =   2655
            Width           =   1185
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Ende"
         Height          =   1300
         Left            =   -71400
         TabIndex        =   208
         Top             =   700
         Width           =   4695
         Begin VB.CheckBox chkEndWin 
            Caption         =   "Rechner nach Auktionsende herunterfahren"
            Height          =   285
            Left            =   260
            TabIndex        =   53
            ToolTipText     =   "Nach Ende aller Auktionen wird der Rechner heruntergefahren"
            Top             =   720
            Width           =   4320
         End
         Begin VB.CheckBox chkDoServer 
            Caption         =   "Nur explizit beenden"
            Height          =   285
            Left            =   260
            TabIndex        =   51
            ToolTipText     =   "wird nur bei Anwahl ""Programmende"" beendet"
            Top             =   240
            Width           =   4320
         End
         Begin VB.CheckBox chkBeendenNachAuktion 
            Caption         =   "BOM nach Auktionsende beenden"
            Height          =   285
            Left            =   260
            TabIndex        =   54
            ToolTipText     =   "Nach Ende aller Auktionen wird BOM beendet"
            Top             =   960
            Width           =   4320
         End
         Begin VB.CheckBox chkWarnenBeimBeenden 
            Caption         =   "Hinweis auf anstehende Auktionen beim Beenden"
            Height          =   285
            Left            =   260
            TabIndex        =   52
            ToolTipText     =   "Beim Beenden wird auf anstehende Auktionen hingewiesen"
            Top             =   480
            Width           =   4320
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'Kein
         Height          =   3855
         Left            =   -74640
         TabIndex        =   198
         Top             =   960
         Width           =   6375
         Begin VB.TextBox txtOdbcDB 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   145
            Top             =   750
            Width           =   3435
         End
         Begin VB.CheckBox chkUsesOdbc 
            Caption         =   "ODBC benutzen "
            Height          =   195
            Left            =   0
            TabIndex        =   143
            Top             =   0
            Width           =   4425
         End
         Begin VB.TextBox txtOdbcProvider 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   144
            Top             =   375
            Width           =   3435
         End
         Begin VB.TextBox txtOdbcUser 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   146
            Top             =   1125
            Width           =   1695
         End
         Begin VB.TextBox txtOdbcPass 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   147
            Top             =   1500
            Width           =   1710
         End
         Begin VB.TextBox txtOdbcZyklus 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   148
            Top             =   1875
            Width           =   645
         End
         Begin VB.CommandButton btnOdbcConnect 
            Caption         =   "Verbindung init"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1485
            TabIndex        =   149
            Top             =   2835
            Width           =   1275
         End
         Begin VB.Label Label36 
            Caption         =   "Provider-String"
            Height          =   270
            Left            =   270
            TabIndex        =   204
            Top             =   420
            Width           =   1470
         End
         Begin VB.Label Label37 
            Caption         =   "Datenbank"
            Height          =   270
            Left            =   270
            TabIndex        =   203
            Top             =   795
            Width           =   1395
         End
         Begin VB.Label Label38 
            Caption         =   "User"
            Height          =   270
            Left            =   270
            TabIndex        =   202
            Top             =   1170
            Width           =   1350
         End
         Begin VB.Label Label39 
            Caption         =   "Passwort"
            Height          =   270
            Left            =   270
            TabIndex        =   201
            Top             =   1545
            Width           =   1470
         End
         Begin VB.Label Label40 
            Caption         =   "Abfragezyklus "
            Height          =   270
            Left            =   270
            TabIndex        =   200
            Top             =   1920
            Width           =   1350
         End
         Begin VB.Label Label41 
            Caption         =   "Minuten"
            Height          =   270
            Left            =   2640
            TabIndex        =   199
            Top             =   1920
            Width           =   840
         End
      End
      Begin VB.Frame Frame16 
         BorderStyle     =   0  'Kein
         Height          =   1455
         Left            =   6720
         TabIndex        =   197
         Top             =   1560
         Width           =   1695
         Begin VB.CommandButton btnSpeedCheck 
            Caption         =   "Speed Test"
            Height          =   525
            Left            =   0
            TabIndex        =   23
            ToolTipText     =   "Test der Verbindungsgeschwindigkeit f¸r 10 Gebote gemittelt"
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton btnLogInTest 
            Caption         =   "Login Test"
            Height          =   525
            Left            =   0
            TabIndex        =   22
            ToolTipText     =   "User- und Passwortpr¸fung"
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Browser und Proxy"
         Height          =   4545
         Left            =   -70200
         TabIndex        =   189
         Top             =   840
         Width           =   3600
         Begin VB.CheckBox chkUseCurl 
            Caption         =   "Curl benutzen"
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   1120
            Width           =   3255
         End
         Begin VB.CheckBox chkUseIECookies 
            Caption         =   "IE-Cookies benutzen"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   1400
            Width           =   3255
         End
         Begin VB.OptionButton optConnectDirect 
            Caption         =   "kein Proxy / direkte Verbindung"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            ToolTipText     =   "Es wird kein Proxy benutzt, sondern diretk verbunden"
            Top             =   1680
            Width           =   3255
         End
         Begin VB.CheckBox chkUseProxyAuth 
            Caption         =   "Beim Proxy anmelden mit:"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            ToolTipText     =   "Wenn der Proxy eine Authentifizierung verlangt"
            Top             =   3400
            Width           =   3135
         End
         Begin VB.TextBox txtProxyPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1800
            PasswordChar    =   "*"
            TabIndex        =   46
            Top             =   4080
            Width           =   1515
         End
         Begin VB.TextBox txtProxyUser 
            Height          =   285
            Left            =   1800
            TabIndex        =   45
            Top             =   3690
            Width           =   1515
         End
         Begin VB.TextBox txtProxyName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   42
            Text            =   "Proxy-Name"
            Top             =   2610
            Width           =   1515
         End
         Begin VB.TextBox txtProxyPort 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   43
            Text            =   "80"
            Top             =   3000
            Width           =   510
         End
         Begin VB.TextBox txtBrowserString 
            Height          =   495
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   36
            ToolTipText     =   "Der Identifikationsstring; beliebig "
            Top             =   550
            Width           =   3045
         End
         Begin VB.OptionButton optConnectDefault 
            Caption         =   "Proxy-Einstellungen wie im IE benutzen"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            ToolTipText     =   "Es werden die Proxy-Einstellungen aus dem Internet-Explorer ¸bernommen"
            Top             =   1980
            Width           =   3255
         End
         Begin VB.OptionButton optUseProxy 
            Caption         =   "Alternativen Proxy benutzen:"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            ToolTipText     =   "Hier kann ein alternativer Proxy angegeben werden"
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label44 
            Caption         =   "Benutzer"
            Height          =   195
            Left            =   480
            TabIndex        =   194
            Top             =   3735
            Width           =   1305
         End
         Begin VB.Label Label43 
            Caption         =   "Passwort"
            Height          =   240
            Left            =   480
            TabIndex        =   193
            Top             =   4095
            Width           =   1395
         End
         Begin VB.Label Label20 
            Caption         =   "Proxy- Port:"
            Height          =   240
            Left            =   480
            TabIndex        =   192
            Top             =   3015
            Width           =   1395
         End
         Begin VB.Label Label21 
            Caption         =   "Proxy- Adresse:"
            Height          =   195
            Left            =   480
            TabIndex        =   191
            Top             =   2655
            Width           =   1305
         End
         Begin VB.Label Label23 
            Caption         =   "Browser- ID"
            Height          =   180
            Left            =   240
            TabIndex        =   190
            Top             =   300
            Width           =   3195
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Zugang zum Internet"
         Height          =   4545
         Left            =   -74760
         TabIndex        =   184
         Top             =   840
         Width           =   4335
         Begin VB.TextBox txtModemVorlauf 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   840
            TabIndex        =   32
            Text            =   "5"
            ToolTipText     =   "Wann wird das Modem verbunden?"
            Top             =   2715
            Width           =   375
         End
         Begin VB.TextBox txtCheckForUpdateInterval 
            Height          =   285
            Left            =   2160
            TabIndex        =   28
            Top             =   1320
            Width           =   555
         End
         Begin VB.CheckBox chkCheckForUpdateBeta 
            Caption         =   "auch Beta-Versionen melden"
            Height          =   195
            Left            =   585
            TabIndex        =   29
            Top             =   1680
            Width           =   3615
         End
         Begin VB.CheckBox chkAutoUpdateCurrencies 
            Caption         =   "W‰hrungskurse automatisch aktualisieren"
            Enabled         =   0   'False
            Height          =   195
            Left            =   585
            TabIndex        =   30
            ToolTipText     =   "Bei jedem Start W‰hrungskurse automatisch aktualisieren"
            Top             =   2040
            Width           =   3615
         End
         Begin VB.CommandButton btnDfueLesen 
            Caption         =   "Lesen"
            Height          =   225
            Left            =   2160
            TabIndex        =   34
            ToolTipText     =   "Einlesen der vorhandenen DFUE- Verbindungen"
            Top             =   3120
            Width           =   1065
         End
         Begin VB.ListBox lstDfue 
            Height          =   450
            Left            =   840
            TabIndex        =   33
            Top             =   3360
            Width           =   2445
         End
         Begin VB.OptionButton optModem 
            Caption         =   "Modem- Einwahlbetrieb"
            Height          =   390
            Left            =   210
            TabIndex        =   31
            ToolTipText     =   "Anwahl ¸ber Modem. ACHTUNG: Automatisches Verbinden muss beim Modem eingestellt sein."
            Top             =   2325
            Width           =   4005
         End
         Begin VB.OptionButton optLan 
            Caption         =   "LAN- oder FLAT- Betrieb"
            Height          =   360
            Left            =   210
            TabIndex        =   24
            ToolTipText     =   "Dauernder Anschluss ans Internet"
            Top             =   300
            Width           =   3225
         End
         Begin VB.TextBox txtPreConnect 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Left            =   870
            TabIndex        =   26
            Text            =   "2"
            ToolTipText     =   "f¸r DSL; kurz vor den Auktionen die Verbindung erˆffnen"
            Top             =   660
            Width           =   345
         End
         Begin VB.CheckBox chkUsePre 
            Height          =   300
            Left            =   585
            TabIndex        =   25
            Top             =   660
            Width           =   225
         End
         Begin VB.CheckBox chkCheckForUpdate 
            Caption         =   "automatisch auf BOM- Updates pr¸fen"
            Enabled         =   0   'False
            Height          =   195
            Left            =   585
            TabIndex        =   27
            ToolTipText     =   "Bei jedem Start auf verf¸gbare Updates pr¸fen"
            Top             =   1080
            Width           =   3615
         End
         Begin VB.CheckBox chkTestConnect 
            Caption         =   "Provider lenkt um"
            Height          =   195
            Left            =   720
            TabIndex        =   35
            ToolTipText     =   "Wenn der Provider zuerst seine HP aufschaltet"
            Top             =   4080
            Width           =   3495
         End
         Begin VB.Label Label60 
            Caption         =   "Stunden"
            Height          =   195
            Left            =   2880
            TabIndex        =   241
            Top             =   1365
            Width           =   1305
         End
         Begin VB.Label Label59 
            Caption         =   "Intervall:"
            Height          =   195
            Left            =   870
            TabIndex        =   240
            Top             =   1365
            Width           =   1305
         End
         Begin VB.Label Label14 
            Caption         =   "Min vor Auktion Verbindung pr¸fen"
            Height          =   255
            Left            =   1380
            TabIndex        =   187
            Top             =   720
            Width           =   2835
         End
         Begin VB.Label label5 
            Caption         =   "Minuten vor Auktion Online gehen"
            Height          =   285
            Left            =   1380
            TabIndex        =   186
            Top             =   2745
            Width           =   2880
         End
         Begin VB.Label Label15 
            Caption         =   "Verbinden mit:"
            Height          =   210
            Left            =   855
            TabIndex        =   185
            Top             =   3090
            Width           =   2640
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Verschiedenes"
         Height          =   4575
         Left            =   -74760
         TabIndex        =   183
         Top             =   840
         Width           =   4260
         Begin VB.CheckBox chkNewItemWindowOpenOnStartup 
            Caption         =   "Artikelfenster beim Start ˆffnen"
            Height          =   195
            Left            =   300
            TabIndex        =   112
            ToolTipText     =   "Das Artikelfenster wird beim Programmstart geˆffnet."
            Top             =   2400
            Width           =   3645
         End
         Begin VB.ComboBox cboIconSet 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2115
            TabIndex        =   106
            Text            =   "f_IconSet"
            ToolTipText     =   "Icon-Satz laden"
            Top             =   840
            Width           =   1950
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'Kein
            Height          =   305
            Left            =   3200
            TabIndex        =   239
            Top             =   3700
            Width           =   735
            Begin VB.Shape Shape1 
               BorderWidth     =   2
               Height          =   180
               Left            =   50
               Shape           =   4  'Gerundetes Rechteck
               Top             =   120
               Width           =   675
            End
         End
         Begin VB.CheckBox chkShowWeekday 
            Caption         =   "Wochentag anzeigen"
            Height          =   195
            Left            =   300
            TabIndex        =   118
            ToolTipText     =   "Positive Statusmeldungen werden nach der vorgegebenen Zeit ausgeblendet"
            Top             =   3600
            Width           =   3585
         End
         Begin VB.CheckBox chkShowFocusRect 
            Caption         =   "Zeilencursor anzeigen, Farbe"
            Height          =   195
            Left            =   300
            TabIndex        =   119
            ToolTipText     =   "Positive Statusmeldungen werden nach der vorgegebenen Zeit ausgeblendet"
            Top             =   3840
            Width           =   2865
         End
         Begin VB.CheckBox chkShowShippingCosts 
            Caption         =   "Versandkosten anzeigen"
            Height          =   195
            Left            =   300
            TabIndex        =   117
            ToolTipText     =   "Positive Statusmeldungen werden nach der vorgegebenen Zeit ausgeblendet"
            Top             =   3360
            Width           =   3585
         End
         Begin VB.CheckBox chkCleanStatus 
            Caption         =   "Status ausblenden nach"
            Height          =   195
            Left            =   300
            TabIndex        =   115
            ToolTipText     =   "Positive Statusmeldungen werden nach der vorgegebenen Zeit ausgeblendet"
            Top             =   3120
            Width           =   2265
         End
         Begin VB.TextBox txtCleanStatusTime 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Left            =   2595
            TabIndex        =   116
            Text            =   "3"
            ToolTipText     =   "Zeit bis zum Ausblenden der Statusmeldung"
            Top             =   3120
            Width           =   495
         End
         Begin VB.CheckBox chkShowTitleTimeLeft 
            Caption         =   "Restzeit in der Titelleiste anzeigen"
            Height          =   195
            Left            =   300
            TabIndex        =   113
            Top             =   2640
            Width           =   3645
         End
         Begin VB.CheckBox chkShowTitleDateTime 
            Caption         =   "Datum und Uhrzeit in der Titelleiste anzeigen"
            Height          =   195
            Left            =   300
            TabIndex        =   114
            Top             =   2880
            Width           =   3645
         End
         Begin VB.CheckBox chkNewItemWindowAlwaysOnTop 
            Caption         =   "Artikelfenster mit 'immer sichtbar' ˆffnen"
            Height          =   195
            Left            =   300
            TabIndex        =   111
            ToolTipText     =   "Wenn minimiert wird BOM als Symbol im Tray dargestellt."
            Top             =   2160
            Width           =   3645
         End
         Begin VB.TextBox txtMaxArtikel 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   3480
            TabIndex        =   103
            Text            =   "12"
            Top             =   250
            Width           =   570
         End
         Begin VB.ComboBox cboToolbarSize 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2115
            Style           =   2  'Dropdown-Liste
            TabIndex        =   105
            ToolTipText     =   "grosse oder kleine Icons anzeigen"
            Top             =   555
            Width           =   1950
         End
         Begin VB.CheckBox chkShowToolbar 
            Caption         =   "ToolBar anzeigen"
            Height          =   210
            Left            =   300
            TabIndex        =   104
            ToolTipText     =   "Zeigt die Symbolleiste an"
            Top             =   620
            Width           =   1740
         End
         Begin VB.CheckBox chkOperaField 
            Caption         =   "zus‰tzliches Eingabefeld"
            Height          =   195
            Left            =   300
            TabIndex        =   107
            ToolTipText     =   "F¸r nicht- D&D- f‰hige Browser "
            Top             =   1215
            Width           =   3600
         End
         Begin VB.CheckBox chkUseWheel 
            Caption         =   "Wheelmouse benutzen"
            Height          =   195
            Left            =   300
            TabIndex        =   108
            ToolTipText     =   "Bei Anschluss einer Scrollmouse kann durch die Artikelliste gescrollt werden"
            Top             =   1455
            Width           =   3645
         End
         Begin VB.CheckBox chkMinToTray 
            Caption         =   "in den Tray minimieren"
            Height          =   195
            Left            =   300
            TabIndex        =   109
            ToolTipText     =   "Wenn minimiert wird BOM als Symbol im Tray dargestellt."
            Top             =   1680
            Width           =   3645
         End
         Begin VB.CheckBox chkNewItemWindowKeepsValues 
            Caption         =   "Artikelfenster beh‰lt erweiterte Eingaben"
            Height          =   195
            Left            =   300
            TabIndex        =   110
            Top             =   1920
            Width           =   3645
         End
         Begin VB.TextBox txtFocusRectColor 
            Height          =   285
            Left            =   2880
            MaxLength       =   6
            TabIndex        =   120
            Top             =   3840
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label61 
            Caption         =   "Icon-Satz laden"
            Height          =   210
            Left            =   600
            TabIndex        =   243
            Top             =   900
            Width           =   1515
         End
         Begin VB.Label Label46 
            Caption         =   " sec "
            Height          =   195
            Left            =   3195
            TabIndex        =   205
            Top             =   3150
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Anzahl der dargestellten Artikelzeilen"
            Height          =   210
            Left            =   300
            TabIndex        =   188
            Top             =   300
            Width           =   3075
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "SMTP"
         Height          =   4695
         Left            =   -70680
         TabIndex        =   179
         Top             =   840
         Width           =   3975
         Begin VB.CheckBox chkSendLow 
            Caption         =   "Mail wenn Gebot zu niedrig"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            ToolTipText     =   "Eine Mail versenden, wenn eigenes Gebot vor Ende zu niedrig ist"
            Top             =   3220
            Width           =   3615
         End
         Begin VB.CommandButton btnTestSmtp 
            Caption         =   "SMTP Test"
            Height          =   455
            Left            =   2640
            TabIndex        =   97
            ToolTipText     =   "POP- und SMTP- Einstellungen testen"
            Top             =   4080
            Width           =   1110
         End
         Begin VB.CheckBox chkSendTestMail 
            Caption         =   "Testmail schicken"
            Height          =   195
            Left            =   240
            TabIndex        =   96
            ToolTipText     =   "SMTP Authentifizierung benutzen"
            Top             =   4245
            Width           =   2295
         End
         Begin VB.CheckBox chkSmtpUseSSL 
            Caption         =   "SSL"
            Height          =   195
            Left            =   240
            TabIndex        =   95
            Top             =   3860
            Width           =   3615
         End
         Begin VB.TextBox txtSmtpPort 
            Height          =   315
            Left            =   2745
            TabIndex        =   87
            Top             =   540
            Width           =   915
         End
         Begin VB.TextBox txtSendOkFromRealname 
            Height          =   285
            Left            =   240
            TabIndex        =   90
            ToolTipText     =   "Name des Mailabsenders (z.B. Otto Mustermann)"
            Top             =   2250
            Width           =   3420
         End
         Begin VB.CheckBox chkSendNok 
            Caption         =   "Auktionsende- Mail wenn ¸berboten"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            ToolTipText     =   "Eine Mail bei erfolgreichem Gebot versenden"
            Top             =   2920
            Width           =   3615
         End
         Begin VB.TextBox txtSmtpServer 
            Height          =   315
            Left            =   240
            TabIndex        =   86
            ToolTipText     =   "Name des SMTP-Servers"
            Top             =   540
            Width           =   2280
         End
         Begin VB.CheckBox chkUseAuth 
            Caption         =   "SMTP Auth benutzen"
            Height          =   195
            Left            =   240
            TabIndex        =   94
            ToolTipText     =   "SMTP Authentifizierung benutzen"
            Top             =   3560
            Width           =   3615
         End
         Begin VB.TextBox txtSendOkFrom 
            Height          =   285
            Left            =   240
            TabIndex        =   89
            ToolTipText     =   "Adresse des Mailabsenders (z.B. Otto@Otto.com)"
            Top             =   1680
            Width           =   3420
         End
         Begin VB.TextBox txtSendOkTo 
            Height          =   285
            Left            =   240
            TabIndex        =   88
            ToolTipText     =   "Adresse des Mailemp‰ngers (z.B. Otto@Otto.com)"
            Top             =   1110
            Width           =   3420
         End
         Begin VB.CheckBox chkSendOk 
            Caption         =   "Auktionsende- Mail schicken bei Erfolg"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            ToolTipText     =   "Eine Mail bei erfolgreichem Gebot versenden"
            Top             =   2640
            Width           =   3615
         End
         Begin VB.Label Label56 
            Caption         =   "Port"
            Height          =   225
            Left            =   2745
            TabIndex        =   230
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label42 
            Caption         =   "Absendername"
            Height          =   225
            Left            =   240
            TabIndex        =   195
            Top             =   2025
            Width           =   3450
         End
         Begin VB.Label Label12 
            Caption         =   "SMTP- Server"
            Height          =   210
            Left            =   240
            TabIndex        =   182
            Top             =   315
            Width           =   2325
         End
         Begin VB.Label Label27 
            Caption         =   "Empf‰ngeradresse"
            Height          =   225
            Left            =   240
            TabIndex        =   181
            Top             =   885
            Width           =   3330
         End
         Begin VB.Label Label32 
            Caption         =   "Absenderadresse"
            Height          =   225
            Left            =   240
            TabIndex        =   180
            Top             =   1455
            Width           =   3450
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Benutzerdaten"
         Height          =   4800
         Left            =   240
         TabIndex        =   172
         Top             =   720
         Width           =   6255
         Begin VB.TextBox txtVorlaufSnipe 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   2760
            TabIndex        =   8
            Text            =   "3"
            ToolTipText     =   "Wann soll geboten werden?"
            Top             =   2205
            Width           =   420
         End
         Begin VB.TextBox txtVorlauf 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   2040
            TabIndex        =   7
            Text            =   "15"
            ToolTipText     =   "Wann soll geboten werden?"
            Top             =   2205
            Width           =   420
         End
         Begin VB.CheckBox chkQuietAfterManBid 
            Caption         =   "Keine Meldung nach manuellem Bieten"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   3120
            Width           =   5415
         End
         Begin VB.CheckBox chkBuyItNow 
            Caption         =   "Sofort-Kaufen aktivieren"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   2880
            Width           =   5415
         End
         Begin VB.CheckBox chkUseSecurityTokenNeuEdit 
            Height          =   195
            Left            =   3720
            TabIndex        =   236
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkUseSecurityToken 
            Caption         =   "Sicherheitsschluessel benutzen"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1440
            TabIndex        =   235
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtSoundOnBidFail 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "das Passwort"
            Top             =   4320
            Width           =   4695
         End
         Begin VB.CommandButton btnTestPlaySoundFail 
            Height          =   285
            Left            =   5400
            Picture         =   "frmSettings.frx":014F
            Style           =   1  'Grafisch
            TabIndex        =   21
            Top             =   4320
            Width           =   285
         End
         Begin VB.CommandButton btnBrowseSoundFail 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5085
            Style           =   1  'Grafisch
            TabIndex        =   20
            Top             =   4320
            Width           =   285
         End
         Begin VB.TextBox txtSoundOnBidSuccess 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   "das Passwort"
            Top             =   3960
            Width           =   4695
         End
         Begin VB.CommandButton btnTestPlaySoundSuccess 
            Height          =   285
            Left            =   5400
            Picture         =   "frmSettings.frx":0299
            Style           =   1  'Grafisch
            TabIndex        =   18
            Top             =   3960
            Width           =   285
         End
         Begin VB.CommandButton btnBrowseSoundSuccess 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5085
            Style           =   1  'Grafisch
            TabIndex        =   17
            Top             =   3960
            Width           =   285
         End
         Begin VB.TextBox txtUsersNeuEdit 
            Height          =   285
            Left            =   4560
            TabIndex        =   207
            Top             =   1365
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtPassNeuEdit 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   4080
            PasswordChar    =   "*"
            TabIndex        =   206
            Top             =   1365
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton btnAddUser 
            Caption         =   "Neu"
            Height          =   300
            Left            =   5080
            TabIndex        =   3
            Top             =   705
            Width           =   1000
         End
         Begin VB.CommandButton btnDelUser 
            Caption         =   "Lˆschen"
            Height          =   300
            Left            =   5080
            TabIndex        =   6
            Top             =   1740
            Width           =   1000
         End
         Begin VB.CommandButton btnEditUser 
            Caption         =   "Edit"
            Height          =   300
            Left            =   5080
            TabIndex        =   4
            Top             =   1055
            Width           =   1000
         End
         Begin VB.CommandButton btnUserCancel 
            Caption         =   "Abbrechen"
            Enabled         =   0   'False
            Height          =   300
            Left            =   5080
            TabIndex        =   5
            Top             =   1400
            Width           =   1000
         End
         Begin VB.ComboBox cboUsers 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown-Liste
            TabIndex        =   1
            Top             =   840
            Width           =   3510
         End
         Begin VB.TextBox txtPass1 
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1440
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   1240
            Width           =   3510
         End
         Begin VB.CommandButton btnBrowseSound 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5080
            Style           =   1  'Grafisch
            TabIndex        =   14
            Top             =   3600
            Width           =   285
         End
         Begin VB.CommandButton btnTestPlaySound 
            Height          =   285
            Left            =   5400
            Picture         =   "frmSettings.frx":03E3
            Style           =   1  'Grafisch
            TabIndex        =   15
            Top             =   3600
            Width           =   285
         End
         Begin VB.TextBox txtSoundOnBid 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            TabIndex        =   13
            ToolTipText     =   "das Passwort"
            Top             =   3600
            Width           =   4695
         End
         Begin VB.CheckBox chkPlaySoundOnBid 
            Caption         =   "Beim Bieten diese Wave- Dateien abspielen"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   3360
            Width           =   5415
         End
         Begin VB.TextBox txtTestArtikel 
            Height          =   285
            Left            =   3360
            TabIndex        =   9
            ToolTipText     =   "Hier die Test- Artikelnummer f¸r SpeedTest eingeben."
            Top             =   2520
            Visible         =   0   'False
            Width           =   1570
         End
         Begin VB.Label Label57 
            Caption         =   "/"
            Height          =   255
            Left            =   2550
            TabIndex        =   233
            Top             =   2235
            Width           =   255
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            X1              =   0
            X2              =   6240
            Y1              =   2080
            Y2              =   2080
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   0
            X2              =   6240
            Y1              =   650
            Y2              =   650
         End
         Begin VB.Label Label4 
            Caption         =   "sec. vor Ablauf der Auktion"
            Height          =   255
            Left            =   3480
            TabIndex        =   178
            Top             =   2235
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Vorbereiten / Bieten"
            Height          =   255
            Left            =   270
            TabIndex        =   177
            Top             =   2235
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Passwort"
            Height          =   255
            Left            =   255
            TabIndex        =   176
            Top             =   1290
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Username"
            Height          =   255
            Left            =   240
            TabIndex        =   175
            Top             =   890
            Width           =   1170
         End
         Begin VB.Label Label28 
            Caption         =   "Artikelnummer f¸r Speed- Test:"
            Height          =   270
            Left            =   270
            TabIndex        =   174
            Top             =   2575
            Visible         =   0   'False
            Width           =   3060
         End
         Begin VB.Label Label35 
            Caption         =   "hier steht das Auktionshaus ;-)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   173
            Top             =   360
            Width           =   3800
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Browser"
         Height          =   1695
         Left            =   -70320
         TabIndex        =   171
         Top             =   3720
         Width           =   3615
         Begin VB.ComboBox cboInlineBrowserModifierKey 
            Height          =   315
            Left            =   2040
            Sorted          =   -1  'True
            Style           =   2  'Dropdown-Liste
            TabIndex        =   131
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtInlineBrowserDelay 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   130
            Text            =   "3"
            ToolTipText     =   "Zeit bis zum Einblenden des Inline-Browsers"
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox chkUseInline 
            Caption         =   "Inline-Browser nach "
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkUseNewWin 
            Caption         =   "neues Fenster"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2040
            TabIndex        =   128
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optUseExt 
            Caption         =   "Externer Browser "
            Height          =   255
            Left            =   120
            TabIndex        =   127
            ToolTipText     =   "Externen Browser verwenden"
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "mit Zusatztaste"
            Height          =   225
            Left            =   360
            TabIndex        =   242
            Top             =   1240
            Width           =   1545
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   " ms"
            Height          =   195
            Left            =   2520
            TabIndex        =   237
            Top             =   800
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Startgrˆsse"
         Height          =   1335
         Left            =   -70320
         TabIndex        =   170
         Top             =   2280
         Width           =   3615
         Begin VB.OptionButton optStartNormal 
            Caption         =   "als Fenster"
            Height          =   195
            Left            =   270
            TabIndex        =   125
            ToolTipText     =   "Programmstart als Fenster"
            Top             =   630
            Width           =   3030
         End
         Begin VB.OptionButton optStartMin 
            Caption         =   "minimiert im Taskbar"
            Height          =   195
            Left            =   270
            TabIndex        =   126
            ToolTipText     =   "Programmstart als Icon"
            Top             =   960
            Width           =   2985
         End
         Begin VB.OptionButton optStartMax 
            Caption         =   "als Vollbildschirm"
            Height          =   195
            Left            =   270
            TabIndex        =   124
            ToolTipText     =   "Programmstart als Vollbild"
            Top             =   300
            Width           =   2985
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Zeitabgleich"
         Height          =   1680
         Left            =   -74520
         TabIndex        =   167
         Top             =   3600
         Width           =   5085
         Begin VB.TextBox txtTimeSyncIntervall 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   300
            Left            =   2160
            TabIndex        =   141
            Text            =   "60"
            ToolTipText     =   "Wie oft soll der Zeitsync durchgef¸hrt werden"
            Top             =   970
            Width           =   435
         End
         Begin VB.CheckBox chkRepeatEvery 
            Caption         =   "wiederholt alle"
            Height          =   255
            Left            =   525
            TabIndex        =   140
            ToolTipText     =   "Synchronisation wiederholt im angegebenen Intervall"
            Top             =   1000
            Width           =   1635
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "t‰glich einmal"
            Height          =   255
            Left            =   525
            TabIndex        =   139
            ToolTipText     =   "Synchronisation bei erstmaligem Einschalten des AutoModus und nachts um 02:05"
            Top             =   760
            Width           =   4395
         End
         Begin VB.CheckBox chkPreGeb 
            Caption         =   "vor jeder Auktion"
            Height          =   255
            Left            =   525
            TabIndex        =   138
            ToolTipText     =   "Synchronisation kurz vor jeder Auktion "
            Top             =   520
            Width           =   4395
         End
         Begin VB.CheckBox chkStart 
            Caption         =   "beim Start"
            Height          =   255
            Left            =   525
            TabIndex        =   137
            ToolTipText     =   "Synchronisation beim Programmstart"
            Top             =   280
            Width           =   4395
         End
         Begin VB.CheckBox chkKeinHinweisNachZeitsync 
            Caption         =   "kein Hinweis nach Zeitabgleich"
            Height          =   255
            Left            =   525
            TabIndex        =   142
            ToolTipText     =   "Die Meldung nach dem Zeitabgleich wird unterdr¸ckt"
            Top             =   1240
            Width           =   4395
         End
         Begin VB.Label Label45 
            Caption         =   "Minuten"
            Height          =   285
            Left            =   2745
            TabIndex        =   196
            Top             =   1005
            Width           =   1080
         End
      End
      Begin VB.TextBox txtMainUrl 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74430
         TabIndex        =   99
         Top             =   3630
         Width           =   3405
      End
      Begin VB.Frame Frame5 
         Caption         =   "Artikel aktualisieren"
         Height          =   1370
         Left            =   -74700
         TabIndex        =   160
         Top             =   2050
         Width           =   8000
         Begin VB.CheckBox chkEditShippingOnClick 
            Caption         =   "Versandkosten beim Anklicken bearbeiten"
            Height          =   195
            Left            =   3550
            TabIndex        =   62
            Top             =   1080
            Width           =   4335
         End
         Begin VB.CheckBox chkAutoWarnNoBid 
            Caption         =   "Im Automatikmodus warnen wenn kein Gebot"
            Height          =   195
            Left            =   3550
            TabIndex        =   61
            Top             =   820
            Width           =   4335
         End
         Begin VB.CheckBox chkAutoAktNext 
            Caption         =   "Nur n‰chster anstehender Artikel"
            Height          =   195
            Left            =   3550
            TabIndex        =   59
            ToolTipText     =   "Nur der n‰chste anstehende Artikel wird beim Start aktialisiert"
            Top             =   300
            Width           =   4335
         End
         Begin VB.CheckBox chkPostGebAktualisieren2 
            Caption         =   "auch wenn bereits ¸berboten"
            Height          =   195
            Left            =   255
            TabIndex        =   58
            ToolTipText     =   "nach Ende der Auktion wird nochmals aktualisiert und gepr¸ft, ob die Auktion gewonnen wurde"
            Top             =   1080
            Width           =   3150
         End
         Begin VB.CheckBox chkAutoAktualisieren 
            Caption         =   "beim Start"
            Height          =   195
            Left            =   255
            TabIndex        =   55
            ToolTipText     =   "Nach dem Start alle Artikel aktualisieren"
            Top             =   300
            Width           =   3150
         End
         Begin VB.CheckBox chkAutoSave 
            Caption         =   "Artikellisten automatisch speicherrn"
            Height          =   195
            Left            =   3550
            TabIndex        =   60
            ToolTipText     =   "Alle 10 Minuten die Artikelliste speichern"
            Top             =   560
            Width           =   4350
         End
         Begin VB.CheckBox chkPostGebAktualisieren 
            Caption         =   "nach Auktionsende"
            Height          =   195
            Left            =   255
            TabIndex        =   57
            ToolTipText     =   "nach Ende der Auktion wird nochmals aktualisiert und gepr¸ft, ob die Auktion gewonnen wurde"
            Top             =   820
            Width           =   3105
         End
         Begin VB.CheckBox chkPostManBidAktualisieren 
            Caption         =   "nach manuellem Bieten"
            Height          =   195
            Left            =   255
            TabIndex        =   56
            Top             =   560
            Width           =   3105
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Font"
         Height          =   1335
         Left            =   -70320
         TabIndex        =   159
         Top             =   840
         Width           =   3615
         Begin VB.CommandButton btnTesten 
            Caption         =   "Anzeigeeinstellungen Testen"
            Height          =   495
            Left            =   1920
            TabIndex        =   238
            ToolTipText     =   "Test der Einstellung f¸r Font und Skalierung (nur ungef‰hr!)"
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtFontSize 
            Alignment       =   1  'Rechts
            Height          =   300
            Left            =   1350
            TabIndex        =   122
            Text            =   "8"
            ToolTipText     =   "Schriftgrˆsse, Default = 8"
            Top             =   600
            Width           =   465
         End
         Begin VB.TextBox txtFieldHeight 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   1350
            TabIndex        =   123
            Text            =   "450"
            ToolTipText     =   "Hˆhe der Artikelzeilen; Default = 450"
            Top             =   930
            Width           =   465
         End
         Begin VB.ComboBox cboFonts 
            Height          =   315
            Left            =   1350
            Sorted          =   -1  'True
            TabIndex        =   121
            ToolTipText     =   "den gew¸nschten Schriftfont ausw‰hlen"
            Top             =   270
            Width           =   2130
         End
         Begin VB.Label Label33 
            Caption         =   "Fontname"
            Height          =   270
            Left            =   240
            TabIndex        =   154
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "Feldhˆhe"
            Height          =   210
            Left            =   240
            TabIndex        =   162
            Top             =   990
            Width           =   990
         End
         Begin VB.Label Label25 
            Caption         =   "Fontgrˆsse "
            Height          =   255
            Left            =   240
            TabIndex        =   155
            Top             =   660
            Width           =   1020
         End
      End
      Begin VB.TextBox txtServer3 
         Height          =   285
         Left            =   -70590
         TabIndex        =   102
         TabStop         =   0   'False
         Text            =   "cgi.x.de"
         Top             =   3630
         Width           =   3405
      End
      Begin VB.TextBox txtServer2 
         Height          =   285
         Left            =   -74430
         TabIndex        =   101
         TabStop         =   0   'False
         Text            =   "cgi.x.de"
         Top             =   4830
         Width           =   3405
      End
      Begin VB.TextBox txtServer1 
         Height          =   285
         Left            =   -74430
         TabIndex        =   100
         TabStop         =   0   'False
         Text            =   "cgi.x.de"
         Top             =   4230
         Width           =   3405
      End
      Begin VB.Frame Frame7 
         Caption         =   "Start"
         Height          =   1300
         Left            =   -74700
         TabIndex        =   161
         Top             =   700
         Width           =   3165
         Begin VB.CheckBox chkShowSplash 
            Caption         =   "Splash beim Start"
            Height          =   210
            Left            =   260
            TabIndex        =   47
            ToolTipText     =   """About""- Fenster beim Start zeigen"
            Top             =   280
            Width           =   2790
         End
         Begin VB.CheckBox chkAutoStart 
            Caption         =   "Automatikmodus beim Start"
            Height          =   285
            Left            =   260
            TabIndex        =   49
            ToolTipText     =   "Nach dem Start die gespeicherten Artikel bebieten"
            Top             =   720
            Width           =   2790
         End
         Begin VB.CheckBox chkStartPass 
            Caption         =   "Passwortabfrage beim Start"
            Height          =   285
            Left            =   260
            TabIndex        =   48
            ToolTipText     =   "Bei Start das Passwort pr¸fen (= gespeichertes User- Passwort)"
            Top             =   480
            Width           =   2790
         End
         Begin VB.CheckBox chkAutoLogin 
            Caption         =   "automatisches Login beim Start"
            Height          =   285
            Left            =   260
            TabIndex        =   50
            ToolTipText     =   "Nach dem Start einloggen"
            Top             =   960
            Width           =   2790
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "TimeSync"
         Height          =   2580
         Left            =   -74520
         TabIndex        =   163
         Top             =   840
         Width           =   5085
         Begin VB.OptionButton optUseSntp 
            Caption         =   "Zeitsynchronisation ¸ber SNTP- Protokoll (Internetzeit)"
            Height          =   195
            Left            =   525
            TabIndex        =   134
            ToolTipText     =   "zur Synchronisation wird die normierte Internetzeit von einem Zeitserver gelesen"
            Top             =   990
            Width           =   4440
         End
         Begin VB.CommandButton btnTestNtp 
            Caption         =   "Zeit Test"
            Height          =   375
            Left            =   3520
            TabIndex        =   136
            ToolTipText     =   "Klick, um den Test durchzuf¸hren; im Erfolgsfall wird die Rechnerzeit gesetzt"
            Top             =   1800
            Width           =   1140
         End
         Begin VB.TextBox txtNtpServer 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1890
            TabIndex        =   135
            ToolTipText     =   "Der Zeitserver als Name"
            Top             =   1320
            Width           =   2790
         End
         Begin VB.OptionButton optUseHTML 
            Caption         =   "Zeitsynchronisation ¸ber HTML "
            Height          =   225
            Left            =   525
            TabIndex        =   132
            ToolTipText     =   "Zur Synchronisation wird die Seite ""aktuelle Zeit"" genutzt; relativ ungenau"
            Top             =   330
            Width           =   4500
         End
         Begin VB.OptionButton optUseTime 
            Caption         =   "Zeitsynchronisation ¸ber TIME- Protokoll (Internetzeit)"
            Height          =   195
            Left            =   525
            TabIndex        =   133
            ToolTipText     =   "zur Synchronisation wird die normierte Internetzeit von einem Zeitserver gelesen"
            Top             =   660
            Width           =   4440
         End
         Begin VB.Label lblTimeDiff 
            Height          =   270
            Left            =   540
            TabIndex        =   168
            Top             =   2145
            Width           =   2475
         End
         Begin VB.Label lblShowTime 
            Caption         =   "gelesene Zeit:"
            Height          =   195
            Left            =   540
            TabIndex        =   166
            Top             =   1920
            Width           =   3450
         End
         Begin VB.Label Label29 
            Caption         =   "Zeit- Server:"
            Height          =   225
            Left            =   795
            TabIndex        =   164
            Top             =   1350
            Width           =   1080
         End
      End
      Begin VB.Label Label58 
         Caption         =   "ServerStrings:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   234
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label54 
         Caption         =   "Server zum Anmelden"
         Height          =   180
         Left            =   -70560
         TabIndex        =   228
         Top             =   4605
         Width           =   3405
      End
      Begin VB.Label Label53 
         Caption         =   "Server f¸r Artikelupdates"
         Height          =   180
         Left            =   -70560
         TabIndex        =   226
         Top             =   4005
         Width           =   3405
      End
      Begin VB.Label Label19 
         Caption         =   "Die Eintr‰ge stehen in ServerSettings.ini und m¸ssen da ge‰ndert werden!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -74400
         TabIndex        =   169
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label34 
         Caption         =   "Webadresse: "
         Height          =   180
         Left            =   -74400
         TabIndex        =   165
         Top             =   3405
         Width           =   3285
      End
      Begin VB.Label Label18 
         Caption         =   "Server zum Bieten"
         Height          =   180
         Left            =   -70560
         TabIndex        =   158
         Top             =   3405
         Width           =   3405
      End
      Begin VB.Label Label17 
         Caption         =   "Server f¸r eBay-Zeit / Servicestatus"
         Height          =   180
         Left            =   -74400
         TabIndex        =   157
         Top             =   4605
         Width           =   3405
      End
      Begin VB.Label Label16 
         Caption         =   "Server f¸r ""beobachtete Artikel"""
         Height          =   180
         Left            =   -74400
         TabIndex        =   156
         Top             =   4005
         Width           =   3405
      End
   End
End
Attribute VB_Name = "frmSettings"
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
'
Private mbChangeFlag As Boolean
Private msUserEditTmp As String

Private Sub btnAbbruch_Click()
  Unload Me
End Sub

Private Sub btnAddUser_Click()

If btnAddUser.Caption = gsarrLangTxt(717) Then
  cboUsers.Visible = False
  txtPass1.Visible = False
  chkUseSecurityToken.Visible = False
  txtUsersNeuEdit.Visible = True
  txtUsersNeuEdit.Text = ""
  txtUsersNeuEdit.TabIndex = 1
  txtPassNeuEdit.Visible = True
  txtPassNeuEdit.Text = ""
  txtPassNeuEdit.TabIndex = 2
  chkUseSecurityTokenNeuEdit.Visible = True
  chkUseSecurityTokenNeuEdit.Value = vbUnchecked
  chkUseSecurityTokenNeuEdit.TabIndex = 3
  btnAddUser.Caption = "OK"
  btnEditUser.Enabled = False
  btnUserCancel.Enabled = True
  btnDelUser.Enabled = False
  txtUsersNeuEdit.SetFocus
Else

  If txtUsersNeuEdit.Text = "" Or txtPassNeuEdit.Text = "" Then
    MsgBox gsarrLangTxt(721) & vbCrLf & gsarrLangTxt(722), vbInformation, gsarrLangTxt(723)
      If txtUsersNeuEdit.Text = "" Then
        txtUsersNeuEdit.SetFocus
      Else
        txtPassNeuEdit.SetFocus
      End If
  Else
    If UsrAccToIndex(txtUsersNeuEdit.Text) > 0 Then
      MsgBox gsarrLangTxt(733), vbExclamation
      Exit Sub
    Else
      cboUsers.AddItem txtUsersNeuEdit.Text
      ReDim Preserve gtarrUserArray(cboUsers.ListCount) As udtUserPass
      gtarrUserArray(cboUsers.ListCount).UaUser = Trim(txtUsersNeuEdit.Text)
      gtarrUserArray(cboUsers.ListCount).UaPass = txtPassNeuEdit.Text
      gtarrUserArray(cboUsers.ListCount).UaToken = IIf(chkUseSecurityTokenNeuEdit.Value = vbChecked, True, False)
      giUserAnzahl = cboUsers.ListCount
      txtUsersNeuEdit.Visible = False
      txtPassNeuEdit.Visible = False
      chkUseSecurityTokenNeuEdit.Visible = False
      cboUsers.Visible = True
      txtPass1.Visible = True
      chkUseSecurityToken.Visible = True
      cboUsers.TabIndex = 1
      txtPass1.TabIndex = 2
      chkUseSecurityToken.TabIndex = 3
      btnAddUser.Caption = gsarrLangTxt(717)
      btnEditUser.Enabled = True
      btnUserCancel.Enabled = False
      btnDelUser.Enabled = True
      cboUsers.ListIndex = cboUsers.ListCount - 1
      Call cboUsers_Click
      mbChangeFlag = True
      gbEmpUserEnd = False
    End If
  End If
End If
End Sub

Private Sub btnBrowseSound_Click()

  Call BrowseSounds(txtSoundOnBid)

End Sub

Private Sub btnBrowseSoundFail_Click()

  BrowseSounds txtSoundOnBidFail

End Sub

Private Sub btnBrowseSoundSuccess_Click()

  BrowseSounds txtSoundOnBidSuccess

End Sub

Private Sub cboServerStrings_Click()
  Dim sTmp As String
  sTmp = gsServerStringsFile
  gsServerStringsFile = cboServerStrings.List(cboServerStrings.ListIndex)

  If gsServerStringsFile <> sTmp Then
    modKeywords.ReadAllKeywords
  
    txtMainUrl = gsMainUrl
  
    txtServer1.Text = gsScript1
    txtServer2.Text = gsScript2
    txtServer3.Text = gsScript3
    txtServer4.Text = gsScript4
    txtServer5.Text = gsScript5
    mbChangeFlag = True
      
    gsServerStringsFile = sTmp
    modKeywords.ReadAllKeywords
  End If
End Sub

Private Sub cboServerStrings_GotFocus()
  SSTab1.Tab = 4
End Sub

Private Sub chkNewItemWindowOpenOnStartup_Click()
  mbChangeFlag = True
End Sub

Private Sub chkPopEncryptedOnly_Click()
  mbChangeFlag = True
End Sub

Private Sub btnLogInTest_Click()

On Error GoTo errhdl
 
 If giUserAnzahl > 0 Then
   If btnEditUser.Caption = gsarrLangTxt(718) And btnAddUser.Caption = gsarrLangTxt(717) Then
  
  'gsUser = f_user
  'pass = f_pass
  gsUser = gtarrUserArray(cboUsers.ListIndex + 1).UaUser
  gsPass = gtarrUserArray(cboUsers.ListIndex + 1).UaPass
  gbUseSecurityToken = gtarrUserArray(cboUsers.ListIndex + 1).UaToken
  gsEbayLocalPass = ""
  frmHaupt.LogIn
  If gsEbayLocalPass > "" Then 'frmHaupt.Toolbar1.Buttons(7).Image = 16
    MsgBox gsarrLangTxt(100), vbInformation
  Else
    MsgBox gsarrLangTxt(101), vbCritical
  End If
   Else
     MsgBox gsarrLangTxt(721), vbInformation + vbOKOnly
   End If
 Else
   MsgBox gsarrLangTxt(2), vbInformation + vbOKOnly
 End If
 
errhdl:
 If Err.Number <> 0 And Err.Number <> 20 Then
   DebugPrint "err: " & Err.Number & " " & Err.Description
   Err.Clear
   Exit Sub
 
 End If

End Sub

Private Sub btnDelUser_Click()
Dim i As Integer
Dim iBetrifft As Integer
Dim iTmp As Integer

If giUserAnzahl > 0 Then
  iTmp = cboUsers.ListIndex + 1
  i = 1
  If giAktAnzArtikel > 0 Then
    Do While i < giAktAnzArtikel + 1
      If gtarrArtikelArray(i).UserAccount = gtarrUserArray(iTmp).UaUser Then
        iBetrifft = iBetrifft + 1
      End If
      i = i + 1
    Loop
    If iBetrifft > 0 Then
      i = 0
      i = MsgBox(gsarrLangTxt(724) & CStr(iBetrifft) & " " & gsarrLangTxt(725) & "." & vbCrLf & _
                 gsarrLangTxt(726) & vbCrLf & vbCrLf & gsarrLangTxt(727), vbExclamation + vbOKCancel)
      If i = vbOK Then
        i = 1
        Do While i < giAktAnzArtikel + 1
          If gtarrArtikelArray(i).UserAccount = gtarrUserArray(iTmp).UaUser Then
            gtarrArtikelArray(i).UserAccount = ""
          End If
          i = i + 1
        Loop
        i = 1
      End If
    Else
      i = 1
    End If
  Else
    i = 1
  End If
End If

If i = 1 Then
  'User Lˆschen
  DeleteValue HKEY_CURRENT_USER, "Software\Biet-O-Matic\", gtarrUserArray(iTmp).UaUser
  giUserAnzahl = giUserAnzahl - 1
  For i = iTmp To giUserAnzahl
    gtarrUserArray(i) = gtarrUserArray(i + 1)
  Next i
  ReDim Preserve gtarrUserArray(giUserAnzahl)
  cboUsers.RemoveItem (cboUsers.ListIndex)

  If cboUsers.ListCount > 0 Then
    If iTmp < giDefaultUser Then
      giDefaultUser = giDefaultUser - 1
      cboUsers.ListIndex = giDefaultUser - 1
    ElseIf iTmp = giDefaultUser Then '<--
      giDefaultUser = 1
      cboUsers.ListIndex = giDefaultUser - 1
    Else
      cboUsers.ListIndex = giDefaultUser - 1
    End If
  Else
    giDefaultUser = 0
    btnDelUser.Enabled = False
    btnEditUser.Enabled = False
    ReDim gtarrUserArray(0)
    gsUser = ""
    gsPass = ""
    gbEmpUserEnd = True
  End If

  cboUsers.Refresh
  Call cboUsers_Click
  mbChangeFlag = True
End If
End Sub

Private Sub lstDfue_Click()
  mbChangeFlag = True
End Sub

Private Sub btnDfueLesen_Click()
  
  Dim i As Integer
  
  GetDFUEList
  mbChangeFlag = True
  
  'mal sehen, ob wir die verbindung wiederfinden ..
  
  If lstDfue.ListCount > 0 Then
    For i = 0 To lstDfue.ListCount - 1
      If lstDfue.List(i) = gsConnectName Then
        lstDfue.ListIndex = i
      End If
    Next
  End If

End Sub

Private Sub btnEditUser_Click()
Dim i As Integer

If btnEditUser.Caption = gsarrLangTxt(718) Then
  cboUsers.Visible = False
  txtPass1.Visible = False
  chkUseSecurityToken.Visible = False
  txtUsersNeuEdit.Visible = True
  txtUsersNeuEdit.TabIndex = 1
  txtPassNeuEdit.Visible = True
  txtPassNeuEdit.TabIndex = 2
  chkUseSecurityTokenNeuEdit.Visible = True
  chkUseSecurityTokenNeuEdit.TabIndex = 3
  txtUsersNeuEdit.Text = gtarrUserArray(cboUsers.ListIndex + 1).UaUser
  msUserEditTmp = gtarrUserArray(cboUsers.ListIndex + 1).UaUser
  txtPassNeuEdit.Text = gtarrUserArray(cboUsers.ListIndex + 1).UaPass
  chkUseSecurityTokenNeuEdit.Value = IIf(gtarrUserArray(cboUsers.ListIndex + 1).UaToken, vbChecked, vbUnchecked)
  btnEditUser.Caption = "OK"
  btnAddUser.Enabled = False
  btnUserCancel.Enabled = True
  btnDelUser.Enabled = False
  txtUsersNeuEdit.SetFocus
Else

  If txtUsersNeuEdit.Text = "" Or txtPassNeuEdit = "" Then
    MsgBox gsarrLangTxt(721) & vbCrLf & gsarrLangTxt(722), vbInformation, gsarrLangTxt(723)
      If txtUsersNeuEdit.Text = "" Then
        txtUsersNeuEdit.SetFocus
      Else
        txtPassNeuEdit.SetFocus
      End If
  Else
    cboUsers.List(cboUsers.ListIndex) = Trim(txtUsersNeuEdit.Text)
    gtarrUserArray(cboUsers.ListIndex + 1).UaUser = Trim(txtUsersNeuEdit.Text)
    gtarrUserArray(cboUsers.ListIndex + 1).UaPass = txtPassNeuEdit.Text
    gtarrUserArray(cboUsers.ListIndex + 1).UaToken = IIf(chkUseSecurityTokenNeuEdit.Value = vbChecked, True, False)
    i = 1
    If giAktAnzArtikel > 0 Then
      Do While i < giAktAnzArtikel + 1
        If gtarrArtikelArray(i).UserAccount = msUserEditTmp Then
          gtarrArtikelArray(i).UserAccount = gtarrUserArray(cboUsers.ListIndex + 1).UaUser
        End If
        i = i + 1
      Loop
    End If
    Call cboUsers_Click
    txtUsersNeuEdit.Text = ""
    txtPassNeuEdit.Text = ""
    chkUseSecurityTokenNeuEdit.Value = vbUnchecked
    txtUsersNeuEdit.Visible = False
    txtPassNeuEdit.Visible = False
    chkUseSecurityTokenNeuEdit.Visible = False
    cboUsers.Visible = True
    txtPass1.Visible = True
    chkUseSecurityToken.Visible = True
    cboUsers.TabIndex = 1
    txtPass1.TabIndex = 2
    chkUseSecurityToken.TabIndex = 3
    btnEditUser.Caption = gsarrLangTxt(718)
    btnAddUser.Enabled = True
    btnUserCancel.Enabled = False
    btnDelUser.Enabled = True
    mbChangeFlag = True
  End If
End If
End Sub

Private Sub optStartMin_Click()
mbChangeFlag = True
End Sub

Private Sub txtAbsender_Change()
  mbChangeFlag = True
End Sub

Private Sub txtAktCycle_Change()
  mbChangeFlag = True
End Sub

Private Sub chkAktualisieren_GotFocus()
  SSTab1.Tab = 2
End Sub

Private Sub cboAktualisierenOpt_Click()
mbChangeFlag = True
End Sub

Private Sub chkAktualXvor_Click()
  txtAktXminvor.Enabled = chkAktualXvor > 0
  txtAktXminvorCycle.Enabled = chkAktualXvor > 0
  Label47.Enabled = chkAktualXvor > 0
  Label48.Enabled = chkAktualXvor > 0
  Label49.Enabled = chkAktualXvor > 0
  cboArtAktOptions.Enabled = chkAktualXvor > 0
  If cboArtAktOptions.ListIndex <> 0 Then
    txtArtAktOptionsValue.Enabled = chkAktualXvor > 0
      Label50.Enabled = chkAktualXvor > 0
      Label51.Enabled = chkAktualXvor > 0
  End If
  mbChangeFlag = True
End Sub

Private Sub chkAktualisieren_Click()
If chkAktualisieren = False Then
  txtAktCycle.Enabled = False
  chkAktualXvor = False
  chkAktualXvor.Enabled = False
  txtAktXminvor.Enabled = False
  txtAktXminvorCycle.Enabled = False
  cboAktualisierenOpt.Enabled = False
  Label24.Enabled = False
  Label47.Enabled = False
  Label48.Enabled = False
  Label52.Enabled = False
Else
  txtAktCycle.Enabled = True
  chkAktualXvor.Enabled = True
  cboAktualisierenOpt.Enabled = True
  Label24.Enabled = True
  Label52.Enabled = True
End If
mbChangeFlag = True
End Sub

Private Sub chkAktualXvor_GotFocus()
  SSTab1.Tab = 2
End Sub

Private Sub txtAktXminvor_Change()
  mbChangeFlag = True
End Sub

Private Sub txtAktXminvor_Validate(Cancel As Boolean)
If Not IsNumeric(txtAktXminvor.Text) Then
  Cancel = True
ElseIf txtAktXminvor < 1 Then
  Cancel = True
End If
End Sub

Private Sub txtAktXminvorCycle_Change()
  mbChangeFlag = True
End Sub

Private Sub txtAktXminvorCycle_Validate(Cancel As Boolean)
If Not IsNumeric(txtAktXminvorCycle.Text) Then
  Cancel = True
ElseIf txtAktXminvorCycle < 1 Then
    Cancel = True
End If
End Sub

Private Sub cboArtAktOptions_Click()
Select Case cboArtAktOptions.ListIndex
  Case 0, 3
    txtArtAktOptionsValue.Text = "0"
    txtArtAktOptionsValue.Visible = False
    Label50.Visible = False
    Label51.Visible = False
  Case 1
    txtArtAktOptionsValue.Visible = True
    txtArtAktOptionsValue.Text = "1"
    Label50.Visible = True
    Label51.Visible = False
  Case 2
    txtArtAktOptionsValue.Visible = True
    txtArtAktOptionsValue.Text = "5"
    Label50.Visible = False
    Label51.Visible = True
    With Label50
      Label51.Move .Left, .Top, .Width, .Height
    End With
End Select
mbChangeFlag = True
End Sub

Private Sub cboArtAktOptions_GotFocus()
  SSTab1.Tab = 2
End Sub

Private Sub txtArtAktOptionsValue_Change()
  mbChangeFlag = True
End Sub

Private Sub txtArtAktOptionsValue_GotFocus()
  SSTab1.Tab = 2
End Sub

Private Sub txtArtAktOptionsValue_Validate(Cancel As Boolean)
If Not IsNumeric(txtArtAktOptionsValue.Text) Then
  Cancel = True
Else
  Select Case cboArtAktOptions.ListIndex
    Case 1
      If txtArtAktOptionsValue < 1 Then
        Cancel = True
      End If
    Case 2
      If txtArtAktOptionsValue < 5 Then
        Cancel = True
      End If
  End Select
End If
End Sub

Private Sub chkAutoAktNext_Click()
  mbChangeFlag = True
End Sub

Private Sub chkAutoAktualisieren_Click()
  If chkAutoAktNext.Enabled = True Then
    chkAutoAktNext.Enabled = False
    chkAutoAktNext = 0
  Else
    chkAutoAktNext.Enabled = True
  End If
  mbChangeFlag = True
End Sub

Private Sub chkAutoLogin_Click()
  mbChangeFlag = True
End Sub

Private Sub chkAutoSave_Click()
  mbChangeFlag = True
End Sub

Private Sub chkAutoStart_Click()
  mbChangeFlag = True
End Sub

Private Sub chkAutoUpdateCurrencies_Click()
  mbChangeFlag = True
End Sub

Private Sub chkAutoWarnNoBid_Click()
mbChangeFlag = True
End Sub

Private Sub chkBeendenNachAuktion_Click()
  mbChangeFlag = True
End Sub

Private Sub txtBrowserString_Change()
  mbChangeFlag = True
End Sub

Private Sub chkBuyItNow_Click()
  If gbBuyItNow = False And chkBuyItNow.Value = vbChecked Then
    If MsgBox(gsarrLangTxt(543), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
      mbChangeFlag = True
    Else
      chkBuyItNow.Value = vbUnchecked
    End If
  Else
    If gbBuyItNow Then mbChangeFlag = True
  End If
End Sub

Private Sub chkCheckForUpdate_Click()
  chkCheckForUpdateBeta.Enabled = chkCheckForUpdate.Value <> 0
  txtCheckForUpdateInterval.Enabled = chkCheckForUpdate.Value <> 0
  mbChangeFlag = True
End Sub

Private Sub chkCheckForUpdateBeta_Click()
  mbChangeFlag = True
End Sub

Private Sub txtCheckForUpdateInterval_Change()
  mbChangeFlag = True
End Sub

Private Sub chkCleanStatus_Click()
  txtCleanStatusTime.Enabled = chkCleanStatus <> 0
  mbChangeFlag = True
End Sub

Private Sub txtCleanStatusTime_Change()
  mbChangeFlag = True
End Sub

Private Sub optConnectDefault_Click()
  chkUseProxyAuth.Enabled = (optConnectDirect = 0)
  optUseProxy_Click
  mbChangeFlag = True
End Sub

Private Sub optConnectDirect_Click()
  chkUseProxyAuth.Enabled = (optConnectDirect = 0)
  optUseProxy_Click
  mbChangeFlag = True
End Sub

Private Sub optConnectDirect_GotFocus()
  SSTab1.Tab = 1
End Sub

Private Sub chkDay_Click()
  mbChangeFlag = True
End Sub

Private Sub chkDoServer_Click()
  mbChangeFlag = True
End Sub

Private Sub chkEditShippingOnClick_Click()
mbChangeFlag = True
End Sub

Private Sub chkEndWin_Click()
  mbChangeFlag = True
End Sub

Private Sub txtFieldHeight_Change()
  mbChangeFlag = True
End Sub

Private Sub txtFieldHeight_LostFocus()
  If Val(txtFieldHeight.Text) <= 0 Then txtFieldHeight.Text = "440"
End Sub

Private Sub txtFocusRectColor_Change()
  mbChangeFlag = True
  Shape1.BorderColor = GetColorFromRgbHex(txtFocusRectColor.Text)
End Sub

Private Sub cboFonts_Change()
  mbChangeFlag = True
End Sub

Private Sub cboFonts_Click()
  mbChangeFlag = True
End Sub

Private Sub cboFonts_GotFocus()
  Dim i As Integer
  If Not gbNoEnumFonts Then
    If cboFonts.ListCount = 0 Then
      For i = 0 To Screen.FontCount - 1
        cboFonts.AddItem Screen.Fonts(i)
      Next i
      cboFonts = gsGlobFontName
    End If
  End If
End Sub

Private Sub txtFontSize_Change()
  mbChangeFlag = True
End Sub

Private Sub cboIconSet_Change()
  mbChangeFlag = True
End Sub

Private Sub cboIconSet_Click()
  mbChangeFlag = True
End Sub

Private Sub cboIconSet_Validate(Cancel As Boolean)
  Dim i As Integer
  
  If cboIconSet.ListIndex < 0 Then
    For i = 0 To cboIconSet.ListCount - 1
      If cboIconSet.Text = cboIconSet.List(i) Then cboIconSet.ListIndex = i
    Next
  End If
  If cboIconSet.ListIndex < 0 Then cboIconSet.Text = ""

End Sub

Private Sub txtInlineBrowserDelay_Change()
  mbChangeFlag = True
End Sub

Private Sub cboInlineBrowserModifierKey_Click()
  mbChangeFlag = True
End Sub

Private Sub cboInlineBrowserModifierKey_GotFocus()
  SSTab1.Tab = 5
End Sub

Private Sub chkKeinHinweisNachZeitsync_Click()
  mbChangeFlag = True
End Sub

Private Sub chkKeinHinweisNachZeitsync_GotFocus()
  SSTab1.Tab = 6
End Sub

Private Sub txtMainUrl_GotFocus()
SSTab1.Tab = 4
End Sub

Private Sub txtMaxArtikel_Change()
  mbChangeFlag = True
End Sub

Private Sub txtMaxArtikel_LostFocus()
  If Val(txtMaxArtikel.Text) <= 0 Then txtMaxArtikel.Text = "13"
End Sub

Private Sub chkMinToTray_Click()
  mbChangeFlag = True
End Sub

Private Sub txtModemVorlauf_Change()
  mbChangeFlag = True
End Sub

Private Sub chkMultiAkt_Click()
  mbChangeFlag = True
End Sub

Private Sub chkNewItemWindowAlwaysOnTop_Click()
  mbChangeFlag = True
End Sub

Private Sub chkNewItemWindowKeepsValues_Click()
  mbChangeFlag = True
End Sub

Private Sub txtNtpServer_Change()
  mbChangeFlag = True
End Sub

Private Sub txtNtpServer_LostFocus()
  If optUseSntp Then txtNtpServer.Text = Trim(GetServerFromServer(txtNtpServer.Text))
End Sub

Private Sub txtOdbcDB_Change()
  mbChangeFlag = True
End Sub

Private Sub txtOdbcPass_Change()
  mbChangeFlag = True
End Sub

Private Sub txtOdbcProvider_Change()
  mbChangeFlag = True
End Sub

Private Sub txtOdbcUser_Change()
  mbChangeFlag = True
End Sub

Private Sub txtOdbcZyklus_Change()
  mbChangeFlag = True
End Sub

Private Sub btnOdbcConnect_Click()

  Dim sTmp As String
  
  If mbChangeFlag Then
    If Not vbYes = MsgBox("Dieser Test erfordert die ‹bernahme der Einstellungen." & vbCrLf & "Sollen die Einstellungen jetzt ¸bernommen und der Test durchgef¸hrt werden ?", vbYesNo + vbQuestion) Then
        Exit Sub
    End If
  End If
  
  Call btnUebernehmen_Click
  Call ODBC_ResetConnection
  Call ODBC_Connect
  
  If ODBC_Check Then
    sTmp = gsarrLangTxt(102)
    frmHaupt.ODBC_Timer.Enabled = gbUsesOdbc
    If gbUsesOdbc Then
        sTmp = sTmp & vbCrLf & Replace(gsarrLangTxt(103), "%MIN%", giOdbcZyklus)
    End If
    MsgBox sTmp
  Else
    frmHaupt.ODBC_Timer.Enabled = False
  End If

End Sub

Private Sub btnOdbcConnect_GotFocus()
  SSTab1.Tab = 7
End Sub

Private Sub chkOperaField_Click()
  mbChangeFlag = True
End Sub

Private Sub chkPlaySoundOnBid_Click()
  mbChangeFlag = True
  txtSoundOnBid.Enabled = (chkPlaySoundOnBid = 1)
  btnTestPlaySound.Enabled = (chkPlaySoundOnBid = 1)
  btnBrowseSound.Enabled = (chkPlaySoundOnBid = 1)
  txtSoundOnBidSuccess.Enabled = (chkPlaySoundOnBid = 1)
  btnTestPlaySoundSuccess.Enabled = (chkPlaySoundOnBid = 1)
  btnBrowseSoundSuccess.Enabled = (chkPlaySoundOnBid = 1)
  txtSoundOnBidFail.Enabled = (chkPlaySoundOnBid = 1)
  btnTestPlaySoundFail.Enabled = (chkPlaySoundOnBid = 1)
  btnBrowseSoundFail.Enabled = (chkPlaySoundOnBid = 1)
End Sub

Private Sub txtPass1_Change()
    
    With chkStartPass
        .Enabled = CBool(Len(txtPass1.Text))
        If Not .Enabled Then .Value = vbUnchecked
    End With
    
End Sub

Private Sub txtPassNeuEdit_Change()
    
    With chkStartPass
        .Enabled = CBool(Len(txtPassNeuEdit.Text))
        If Not .Enabled Then .Value = vbUnchecked
    End With
    
End Sub


Private Sub txtPopPass_Change()
  frmHaupt.PanelText frmHaupt.StatusBar1, 2, ""
  mbChangeFlag = True
End Sub

Private Sub txtPopPort_Change()
  mbChangeFlag = True
End Sub

Private Sub txtPopServer_Change()
  frmHaupt.PanelText frmHaupt.StatusBar1, 2, ""
  mbChangeFlag = True
End Sub

Private Sub txtPopServer_GotFocus()
  SSTab1.Tab = 3
End Sub

Private Sub txtPopTimeOut_Change()
  mbChangeFlag = True
End Sub

Private Sub txtPopUser_Change()
  frmHaupt.PanelText frmHaupt.StatusBar1, 2, ""
  mbChangeFlag = True
End Sub

Private Sub chkPopUseSSL_Click()
  mbChangeFlag = True
End Sub

Private Sub txtPopZykl_Change()
  mbChangeFlag = True
End Sub

Private Sub chkPostGebAktualisieren_Click()
  mbChangeFlag = True
  chkPostGebAktualisieren2.Enabled = (chkPostGebAktualisieren = 1)
End Sub

Private Sub chkPostGebAktualisieren2_Click()
  mbChangeFlag = True
End Sub

Private Sub txtPreConnect_Change()
  mbChangeFlag = True
End Sub

Private Sub chkPreGeb_Click()
  mbChangeFlag = True
End Sub

Private Sub chkPostManBidAktualisieren_Click()
  mbChangeFlag = True
End Sub

Private Sub txtProxyName_Change()
  mbChangeFlag = True
End Sub

Private Sub txtProxyPass_Change()
  mbChangeFlag = True
End Sub

Private Sub txtProxyPass_GotFocus()
  SSTab1.Tab = 1
End Sub

Private Sub txtProxyPort_Change()
  mbChangeFlag = True
End Sub

Private Sub txtProxyUser_Change()
  mbChangeFlag = True
End Sub

Private Sub chkQuietAfterManBid_Click()
  mbChangeFlag = True
End Sub

Private Sub txtReloadTimes_Change()
    mbChangeFlag = True
End Sub

Private Sub txtReloadTimes_GotFocus()
  SSTab1.Tab = 2
End Sub

Private Sub txtReloadTimes_Validate(Cancel As Boolean)
    If IsNumeric(txtReloadTimes) Then
      If txtReloadTimes <= 10 And txtReloadTimes >= 0 Then Exit Sub
    End If
    MsgBox "0 <= X <= 10 !!!"
    txtReloadTimes = giReloadTimes
    Cancel = True
End Sub

Private Sub chkReLogin_Click()
  mbChangeFlag = True
End Sub

Private Sub chkRepeatEvery_Click()
  txtTimeSyncIntervall.Enabled = (chkRepeatEvery > 0)
  mbChangeFlag = True
End Sub

Private Sub chkSendLow_Click()
  mbChangeFlag = True
End Sub

Private Sub chkSendNok_Click()
  mbChangeFlag = True
End Sub

Private Sub chkSendOk_Click()
  mbChangeFlag = True
End Sub

Private Sub txtSendOkFrom_Change()
  mbChangeFlag = True
End Sub

Private Sub txtSendOkFromRealname_Change()
  mbChangeFlag = True
End Sub

Private Sub txtSendOkTo_Change()
  mbChangeFlag = True
End Sub

Private Sub txtServer3_GotFocus()
SSTab1.Tab = 4
End Sub

Private Sub chkShowFocusRect_Click()
  mbChangeFlag = True
  txtFocusRectColor.Enabled = (chkShowFocusRect = 1)
End Sub

Private Sub chkShowShippingCosts_Click()
  mbChangeFlag = True
End Sub

Private Sub chkShowSplash_Click()
  mbChangeFlag = True
End Sub

Private Sub chkShowSplash_GotFocus()
  SSTab1.Tab = 2
End Sub

Private Sub chkShowTitleDateTime_Click()
  mbChangeFlag = True
End Sub

Private Sub chkShowTitleTimeLeft_Click()
  mbChangeFlag = True
End Sub

Private Sub chkShowToolbar_Click()
  cboToolbarSize.Enabled = chkShowToolbar > 0
  mbChangeFlag = True
End Sub

Private Sub chkShowWeekday_Click()
  mbChangeFlag = True
End Sub

Private Sub txtSmtpPort_Change()
  mbChangeFlag = True
End Sub

Private Sub txtSmtpServer_Change()
  mbChangeFlag = True
End Sub

Private Sub chkSmtpUseSSL_Click()
  mbChangeFlag = True
End Sub

Private Sub txtSoundOnBid_Change()
  mbChangeFlag = True
End Sub

Private Sub txtSoundOnBidFail_Change()
  mbChangeFlag = True
End Sub

Private Sub txtSoundOnBidSuccess_Change()
  mbChangeFlag = True
End Sub

Private Sub chkStart_Click()
  mbChangeFlag = True
End Sub

Private Sub optStartMax_Click()
  mbChangeFlag = True
End Sub

Private Sub optStartNormal_Click()
  mbChangeFlag = True
End Sub
Private Sub chkStartPass_Click()
  mbChangeFlag = True
End Sub

Private Sub txtTestArtikel_Change()
  gsTestArtikel = txtTestArtikel
  mbChangeFlag = True
End Sub

Private Sub chkTestConnect_Click()
  mbChangeFlag = True
End Sub

Private Sub txtTimeSyncIntervall_Change()
  mbChangeFlag = True
End Sub

Private Sub cboToolbarSize_Click()
  mbChangeFlag = True
End Sub

Private Sub chkUseAuth_Click()
  mbChangeFlag = True
End Sub

Private Sub chkUseCurl_Click()
  mbChangeFlag = True
  chkUseIECookies.Enabled = False ' (chkUseCurl = 0)
  optConnectDefault.Enabled = (chkUseCurl = 0)
  chkMultiAkt.Enabled = (chkUseCurl <> 0)
  If chkUseIECookies.Enabled = False Then chkUseIECookies.Value = vbUnchecked
  If optConnectDefault.Enabled = False And optConnectDefault.Value = True Then optConnectDirect.Value = True
  If chkMultiAkt.Enabled = False Then chkMultiAkt.Value = vbUnchecked
End Sub

Private Sub optUseExt_Click()
  chkUseNewWin.Enabled = optUseExt.Value
  mbChangeFlag = True
End Sub

Private Sub optUseHTML_Click()

  optUseHTML.Value = True
  optUseTime.Value = False
  optUseSntp.Value = False
  txtNtpServer.Enabled = False
  mbChangeFlag = True
  
End Sub

Private Sub optUseHTML_GotFocus()
  SSTab1.Tab = 6
End Sub

Private Sub chkUseIECookies_Click()
  mbChangeFlag = True
End Sub

Private Sub chkUseInline_Click()
  mbChangeFlag = True
  txtInlineBrowserDelay.Enabled = chkUseInline
  cboInlineBrowserModifierKey.Enabled = chkUseInline
End Sub

Private Sub chkUseInline_GotFocus()
  SSTab1.Tab = 5
End Sub

Private Sub optUseTime_GotFocus()
  SSTab1.Tab = 6
End Sub

Private Sub optUseSntp_GotFocus()
  SSTab1.Tab = 6
End Sub

Private Sub chkUseNewWin_Click()
  mbChangeFlag = True
End Sub

Private Sub optUseTime_Click()

  optUseHTML.Value = False
  optUseTime.Value = True
  optUseSntp.Value = False
  txtNtpServer.Enabled = True
  mbChangeFlag = True
  
End Sub

Private Sub optUseSntp_Click()

  optUseHTML.Value = False
  optUseTime.Value = False
  optUseSntp.Value = True
  txtNtpServer.Enabled = True
  txtNtpServer.Text = Trim(GetServerFromServer(txtNtpServer.Text))
  mbChangeFlag = True
  
End Sub

Private Sub chkUsePop_Click()

  frmHaupt.PanelText frmHaupt.StatusBar1, 2, ""
  If chkUsePop.Value > 0 Then
    txtPopZykl.Enabled = True
    txtAbsender.Enabled = True
  Else
    frmHaupt.PanelText frmHaupt.StatusBar1, 2, " "
    txtPopZykl.Enabled = False
    txtAbsender.Enabled = False
  End If
  mbChangeFlag = True
  
End Sub


Private Sub chkUsePre_Click()
  mbChangeFlag = True
End Sub

Private Sub optUseProxy_Click()
  chkUseProxyAuth.Enabled = (optConnectDirect = 0)
  txtProxyName.Enabled = (optUseProxy <> 0)
  txtProxyPort.Enabled = (optUseProxy <> 0)
  txtProxyUser.Enabled = (chkUseProxyAuth <> 0) And chkUseProxyAuth.Enabled
  txtProxyPass.Enabled = (chkUseProxyAuth <> 0) And chkUseProxyAuth.Enabled
  mbChangeFlag = True
End Sub

Private Sub chkUseProxyAuth_Click()
  mbChangeFlag = True
  txtProxyUser.Enabled = (chkUseProxyAuth <> 0)
  txtProxyPass.Enabled = (chkUseProxyAuth <> 0)
End Sub

Private Sub chkUseProxyAuth_GotFocus()
  SSTab1.Tab = 1
End Sub

Private Sub cboUsers_GotFocus()
  SSTab1.Tab = 0
End Sub

Private Sub cboUsers_Click()
txtPass1.Text = gtarrUserArray(cboUsers.ListIndex + 1).UaPass
chkUseSecurityToken.Value = IIf(gtarrUserArray(cboUsers.ListIndex + 1).UaToken, vbChecked, vbUnchecked)
mbChangeFlag = True
End Sub

Private Sub txtUsersNeuEdit_LostFocus()
  txtUsersNeuEdit.Text = Trim(LCase(txtUsersNeuEdit.Text))
End Sub

Private Sub chkUsesOdbc_Click()

  If chkUsesOdbc <> 0 Then
    txtOdbcDB.Enabled = True
    txtOdbcPass.Enabled = True
    txtOdbcProvider.Enabled = True
    txtOdbcUser.Enabled = True
    txtOdbcZyklus.Enabled = True
    btnOdbcConnect.Enabled = True
  Else
    txtOdbcDB.Enabled = False
    txtOdbcPass.Enabled = False
    txtOdbcProvider.Enabled = False
    txtOdbcUser.Enabled = False
    txtOdbcZyklus.Enabled = False
    btnOdbcConnect.Enabled = False
  End If
  mbChangeFlag = True
  
End Sub

Private Sub chkUsesOdbc_GotFocus()
  SSTab1.Tab = 7
End Sub

Private Sub chkUseWheel_Click()
  mbChangeFlag = True
End Sub

Private Sub txtVorlauf_Change()
  mbChangeFlag = True
End Sub

Private Sub txtVorlauf_Validate(Cancel As Boolean)
  On Error GoTo ERROR_HANDLER
  If CInt(txtVorlauf.Text) <= CDbl(txtVorlaufSnipe.Text) Then
    txtVorlaufSnipe = CInt(txtVorlauf.Text) - 1
  End If
ERROR_HANDLER:
End Sub

Private Sub txtVorlaufSnipe_Change()
  mbChangeFlag = True
End Sub

Private Sub txtVorlaufSnipe_Validate(Cancel As Boolean)
  If txtVorlaufSnipe = "" Then txtVorlaufSnipe = "0"
  
  On Error GoTo ERROR_HANDLER
  If CInt(txtVorlauf.Text) <= CDbl(txtVorlaufSnipe.Text) Then
    txtVorlauf = CInt(CDbl(txtVorlaufSnipe.Text) + 1)
  End If
ERROR_HANDLER:

End Sub

Private Sub chkWarnenBeimBeenden_Click()
  mbChangeFlag = True
End Sub

Private Sub Form_Load()

  Dim i As Integer
  
  On Error Resume Next
  
  Me.Icon = MyLoadResPicture(104, 16)
  Call SendMessage(Me.hWnd, WM_SETICON, 0, Me.Icon)

  gbSettingsIsUp = True
  
  Call SetLanguage
  Label35.Caption = gsarrLangTxt(319) & gsAuctionHome
  
  'Bieten
  'f_user = gsUser
  'f_pass = gsPass
    
  cboUsers.Clear
  If giUserAnzahl > 0 Then
    For i = 1 To giUserAnzahl
    cboUsers.AddItem gtarrUserArray(i).UaUser
      If giUserAnzahl = 1 Then
        cboUsers.ListIndex = 0
        giDefaultUser = 1
      Else
        If giDefaultUser = i Then cboUsers.ListIndex = i - 1
      End If
    Next i
  txtPass1.Text = gtarrUserArray(cboUsers.ListIndex + 1).UaPass
  chkUseSecurityToken.Value = IIf(gtarrUserArray(cboUsers.ListIndex + 1).UaToken, vbChecked, vbUnchecked)
  Else
    btnEditUser.Enabled = False
    btnDelUser.Enabled = False
  End If
  
  With cboUsers
    txtUsersNeuEdit.Move .Left, .Top, .Width, .Height
  End With
  
  With txtPass1
    txtPassNeuEdit.Move .Left, .Top, .Width, .Height
  End With
  
  With chkUseSecurityToken
    chkUseSecurityTokenNeuEdit.Move .Left, .Top, .Width, .Height
  End With
  
  If gbPlaySoundOnBid Then chkPlaySoundOnBid.Value = 1 Else chkPlaySoundOnBid.Value = 0
  txtSoundOnBid.Text = gsSoundOnBid
  txtSoundOnBidSuccess.Text = gsSoundOnBidSuccess
  txtSoundOnBidFail.Text = gsSoundOnBidFail
  txtSoundOnBid.Enabled = gbPlaySoundOnBid
  txtSoundOnBidSuccess.Enabled = gbPlaySoundOnBid
  txtSoundOnBidFail.Enabled = gbPlaySoundOnBid
  btnTestPlaySound.Enabled = gbPlaySoundOnBid
  btnTestPlaySoundSuccess.Enabled = gbPlaySoundOnBid
  btnTestPlaySoundFail.Enabled = gbPlaySoundOnBid
  btnBrowseSound.Enabled = gbPlaySoundOnBid
  btnBrowseSoundSuccess.Enabled = gbPlaySoundOnBid
  btnBrowseSoundFail.Enabled = gbPlaySoundOnBid
  
  txtBrowserString.Text = gsBrowserIdString
  If gbUseIECookies Then chkUseIECookies.Value = 1 Else chkUseIECookies.Value = 0
  If gbUseCurl Then chkUseCurl.Value = 1 Else chkUseCurl.Value = 0
  
  txtVorlauf.Text = glVorlaufGebot
  txtVorlaufSnipe.Text = gfVorlaufSnipe
  
  'Verbindung
  optModem.Value = gbUsesModem
  optLan.Value = Not gbUsesModem
  
  If gbUsesModem Then optModem_Click Else optLan_Click
  
  If gbCheckForUpdate Then chkCheckForUpdate.Value = 1 Else chkCheckForUpdate.Value = 0
  chkCheckForUpdate.Enabled = Not gbUsesModem
  If Not chkCheckForUpdate.Enabled Then chkCheckForUpdate.Value = vbUnchecked
  txtCheckForUpdateInterval.Text = glCheckForUpdateInterval
  txtCheckForUpdateInterval.Enabled = chkCheckForUpdate.Value <> 0
  
  If gbCheckForUpdateBeta Then chkCheckForUpdateBeta.Value = 1 Else chkCheckForUpdateBeta.Value = 0
  
  If gbAutoUpdateCurrencies Then chkAutoUpdateCurrencies.Value = 1 Else chkAutoUpdateCurrencies.Value = 0
  
  txtPreConnect.Text = CStr(giVorlaufLan)
  If giVorlaufLan > 0 Then chkUsePre.Value = 1 Else chkUsePre.Value = 0
  
  If gbUseProxy Then
    optUseProxy.Value = 1
  ElseIf gbUseDirectConnect Then
    optConnectDirect.Value = 1
  Else
    optConnectDefault.Value = 1
  End If
  
  If gbUseProxyAuthentication Then chkUseProxyAuth.Value = 1 Else chkUseProxyAuth.Value = 0
  
  txtProxyName.Text = gsProxyName
  txtProxyPort.Text = giProxyPort
  txtProxyUser.Text = gsProxyUser
  txtProxyPass.Text = gsProxyPass
  
  chkUseCurl.Enabled = TestForCurl()
  chkUseIECookies.Enabled = False ' CBool(chkUseCurl.Value = 0)
  optConnectDefault.Enabled = CBool(chkUseCurl.Value = 0)
  chkMultiAkt.Enabled = CBool(chkUseCurl.Value <> 0)
  If chkUseIECookies.Enabled = False Then chkUseIECookies.Value = vbUnchecked
  If optConnectDefault.Enabled = False And optConnectDefault.Value = True Then optConnectDirect.Value = True
  'If f_multiakt.Enabled = False Then f_multiakt.Value = vbUnchecked ' -> ... And UseCurl
  
  chkUseProxyAuth.Enabled = CBool(optConnectDirect.Value = 0)
  txtProxyName.Enabled = CBool(optUseProxy.Value <> 0)
  txtProxyPort.Enabled = CBool(optUseProxy.Value <> 0)
  txtProxyUser.Enabled = (chkUseProxyAuth.Value <> 0) And chkUseProxyAuth.Enabled
  txtProxyPass.Enabled = (chkUseProxyAuth.Value <> 0) And chkUseProxyAuth.Enabled
  
  
  txtModemVorlauf.Text = glVorlaufModem
  
  lstDfue.Clear
  lstDfue.AddItem gsConnectName
  lstDfue.ListIndex = 0
  If gbTestConnect Then chkTestConnect.Value = 1 Else chkTestConnect.Value = 0
  
  cboFonts.Text = gsGlobFontName
  
  'Automatik
  chkAutoLogin.Enabled = Not gbUsesModem
  chkUseInline.Enabled = Not gbUsesModem
  chkStartPass.Enabled = CBool(Len(gsPass)) 'MD-Marker , Neu 20090410
  If gbPassAtStart Then chkStartPass.Value = 1 Else chkStartPass.Value = 0
  If gbAutoStart Then chkAutoStart.Value = 1 Else chkAutoStart = 0
  If gbAutoLogin And Not gbUsesModem Then chkAutoLogin.Value = 1 Else chkAutoLogin.Value = 0
  If gbTrayAction Then chkDoServer.Value = 1 Else chkDoServer.Value = 0
  If gbFileWinShutdown Then chkEndWin.Value = 1 Else chkEndWin.Value = 0
  If gbShowSplash Then chkShowSplash.Value = 1 Else chkShowSplash.Value = 0
  If frmHaupt.AutoSave.Enabled Then chkAutoSave.Value = 1 Else chkAutoSave.Value = 0
  
  If gbConcurrentUpdates And gbUseCurl Then chkMultiAkt.Value = 1 Else chkMultiAkt.Value = 0
  If gbUpdateAfterManualBid Then chkPostManBidAktualisieren.Value = 1 Else chkPostManBidAktualisieren.Value = 0
  If gbQuietAfterManualBid Then chkQuietAfterManBid.Value = 1 Else chkQuietAfterManBid.Value = 0
  If gbGeboteAktualisieren Then chkAktualisieren.Value = 1 Else chkAktualisieren.Value = 0
  txtAktCycle.Enabled = chkAktualisieren.Value > 0
  cboAktualisierenOpt.Enabled = chkAktualisieren.Value > 0
  Label24.Enabled = chkAktualisieren.Value > 0
  Label52.Enabled = chkAktualisieren.Value > 0
  cboAktualisierenOpt.ListIndex = giAktualisierenOpt
  txtAktCycle.Text = CStr(giArtikelRefreshCycle)
    
  If gbAktualisierenXvor Then
    chkAktualXvor.Value = 1
    txtAktXminvor.Enabled = True
    txtAktXminvor.Text = CStr(giAktXminvor)
    txtAktXminvorCycle.Enabled = True
    txtAktXminvorCycle.Text = CStr(giAktXminvorCycle)
    Label49.Enabled = True
    cboArtAktOptions.Enabled = True
    cboArtAktOptions.ListIndex = giArtAktOptions
    txtArtAktOptionsValue.Text = CStr(giArtAktOptionsValue)
    Select Case giArtAktOptions
      Case 0 ', 3
        txtArtAktOptionsValue.Visible = False
        Label50.Visible = False
        Label51.Visible = False
      Case 1
        txtArtAktOptionsValue.Visible = True
        Label50.Visible = True
        Label51.Visible = False
      Case 2
        txtArtAktOptionsValue.Visible = True
        Label50.Visible = False
        Label51.Visible = True
        With Label50
          Label51.Move .Left, .Top, .Width, .Height
        End With
      End Select
  Else
    chkAktualXvor.Value = 0
    Label49.Enabled = False
    cboArtAktOptions.ListIndex = 0
    cboArtAktOptions.Enabled = False
  End If
  
  If gbArtikelRefreshPost Then chkPostGebAktualisieren.Value = 1 Else chkPostGebAktualisieren.Value = 0
  chkPostGebAktualisieren2.Enabled = CBool(chkPostGebAktualisieren.Value = 1)
  If gbArtikelRefreshPost2 Then chkPostGebAktualisieren2.Value = 1 Else chkPostGebAktualisieren2.Value = 0
  
  If gbAutoAktualisieren Then chkAutoAktualisieren.Value = 1 Else chkAutoAktualisieren.Value = 0
  chkAutoAktNext.Enabled = CBool(chkAutoAktualisieren.Value <> 0)
  If gbAutoAktualisierenNext Then chkAutoAktNext.Value = 1 Else chkAutoAktNext.Value = 0
  If gbAutoWarnNoBid Then chkAutoWarnNoBid.Value = 1 Else chkAutoWarnNoBid.Value = 0
  If gbEditShippingOnClick Then chkEditShippingOnClick.Value = 1 Else gbEditShippingOnClick = 0
 
  If gbWarnenBeimBeenden Then chkWarnenBeimBeenden.Value = 1 Else gbWarnenBeimBeenden = 0
  If gbBeendenNachAuktion Then chkBeendenNachAuktion.Value = 1 Else chkBeendenNachAuktion = 0
  
  txtReloadTimes.Text = CStr(giReloadTimes)
  
  txtReLogin.Text = CStr(giReLogin)
  If giReLogin > 0 Then chkReLogin.Value = 1 Else chkReLogin.Value = 0
  
  'Timesync
  '0: kein, 1: einmal, 2: vor Gebot, 4: beim Start, 8: regelm‰ﬂig alle x Minuten
  
   If (giUseTimeSync And 1) > 0 Then chkDay.Value = 1
   If (giUseTimeSync And 2) > 0 Then chkPreGeb.Value = 1
   If (giUseTimeSync And 4) > 0 Then chkStart.Value = 1
   If (giUseTimeSync And 8) > 0 Then chkRepeatEvery.Value = 1
  
   txtTimeSyncIntervall.Enabled = CBool(chkRepeatEvery.Value > 0)
   txtTimeSyncIntervall.Text = CStr(glTimeSyncIntervall)
  
  'POP- Server
  If gbUsePop Then chkUsePop.Value = 1 Else chkUsePop.Value = 0
  Call chkUsePop_Click
  txtPopZykl.Text = CStr(giPopZyklus)
  txtPopServer.Text = gsPopServer
  txtPopPort.Text = CStr(giPopPort)
  txtPopUser.Text = gsPopUser
  txtPopPass.Text = gsPopPass
  txtPopTimeOut.Text = CStr(giPopTimeOut)
  txtSmtpServer.Text = gsSmtpServer
  txtSmtpPort.Text = CStr(giSmtpPort)
  txtAbsender.Text = gsAbsender
  
  If gbUseSmtpAuth Then chkUseAuth.Value = 1 Else chkUseAuth.Value = 0
  If gbPopUseSSL Then chkPopUseSSL.Value = 1 Else chkPopUseSSL.Value = 0
  If gbSmtpUseSSL Then chkSmtpUseSSL.Value = 1 Else chkSmtpUseSSL.Value = 0
  If gbPopEncryptedOnly Then chkPopEncryptedOnly.Value = 1 Else gbPopEncryptedOnly = 0
  
  'Ebay- Server
  cboServerStrings.Clear
  For i = 1 To UBound(gsarrServerStrArr())
    cboServerStrings.AddItem gsarrServerStrArr(i)
    If gsarrServerStrArr(i) = gsServerStringsFile Then
      cboServerStrings.ListIndex = i - 1
    End If
  Next
  txtMainUrl.Text = gsMainUrl
  
  txtServer1.Enabled = False
  txtServer2.Enabled = False
  txtServer3.Enabled = False
  txtServer4.Enabled = False
  txtServer5.Enabled = False
  
  txtServer1.Text = gsScript1
  txtServer2.Text = gsScript2
  txtServer3.Text = gsScript3
  txtServer4.Text = gsScript4
  txtServer5.Text = gsScript5
  
  'Darstellung
  txtMaxArtikel.Text = CStr(giMaxRowSetting + 1)
  Select Case giStartupSize
  Case vbMaximized
    optStartMax.Value = True
  Case vbNormal
    optStartNormal.Value = True
  Case vbMinimized
    optStartMin.Value = True
  End Select
  
  'If gbUseIntBrowser Then optUseInt.Value = True Else optUseExt.Value = True
  optUseExt.Value = True 'MD-Marker 20090325 , interner browser nicht mehr vorhanden
  chkUseNewWin.Enabled = optUseExt.Value
  If gbBrowseInNewWindow Then chkUseNewWin.Value = 1 Else chkUseNewWin.Value = 0
  If gbBrowseInline And Not gbUsesModem Then chkUseInline.Value = 1 Else chkUseInline.Value = 0
  txtInlineBrowserDelay.Enabled = CBool(chkUseInline.Value)
  cboInlineBrowserModifierKey.Enabled = CBool(chkUseInline.Value)
  txtInlineBrowserDelay.Text = CStr(giInlineBrowserDelay)
  cboInlineBrowserModifierKey.Clear
  cboInlineBrowserModifierKey.AddItem gsarrLangTxt(478), 0
  cboInlineBrowserModifierKey.AddItem gsarrLangTxt(479), 1
  cboInlineBrowserModifierKey.AddItem gsarrLangTxt(480), 2
  cboInlineBrowserModifierKey.ListIndex = giInlineBrowserModifierKey
  
  If gbUseWheel Then chkUseWheel.Value = 1 Else chkUseWheel.Value = 0
  If gbShowToolbar Then chkShowToolbar.Value = 1 Else chkShowToolbar.Value = 0
  If gbOperaField Then chkOperaField.Value = 1 Else chkOperaField.Value = 0
  If gbNewItemWindowAlwaysOnTop Then chkNewItemWindowAlwaysOnTop.Value = 1 Else chkNewItemWindowAlwaysOnTop.Value = 0
  If gbNewItemWindowOpenOnStartup Then chkNewItemWindowOpenOnStartup.Value = 1 Else chkNewItemWindowOpenOnStartup = 0
  If gbNewItemWindowKeepsValues Then chkNewItemWindowKeepsValues.Value = 1 Else chkNewItemWindowKeepsValues.Value = 0
  If gbMinToTray Then chkMinToTray.Value = 1 Else chkMinToTray.Value = 0
  If gbKeinHinweisNachZeitsync Then chkKeinHinweisNachZeitsync.Value = 1 Else chkKeinHinweisNachZeitsync.Value = 0
  If gbShowTitleDateTime Then chkShowTitleDateTime.Value = 1 Else chkShowTitleDateTime.Value = 0
  If gbShowTitleTimeLeft Then chkShowTitleTimeLeft.Value = 1 Else chkShowTitleTimeLeft.Value = 0
  If gbCleanStatus Then chkCleanStatus.Value = 1 Else chkCleanStatus.Value = 0
  txtCleanStatusTime.Text = CStr(glCleanStatusTime)
  txtCleanStatusTime.Enabled = gbCleanStatus
  If gbShowShippingCosts Then chkShowShippingCosts.Value = 1 Else chkShowShippingCosts.Value = 0
  If gbShowWeekday Then chkShowWeekday.Value = 1 Else chkShowWeekday.Value = 0
  If gsSpecialDateFormat > "" Then chkShowWeekday.Value = 0: chkShowWeekday.Enabled = False
  If gbShowFocusRect Then chkShowFocusRect.Value = 1 Else chkShowFocusRect.Value = 0
  txtFocusRectColor.Enabled = CBool(chkShowFocusRect.Value = 1)
  txtFocusRectColor.Text = GetRgbHexFromColor(glFocusRectColor)
  
  cboToolbarSize.Clear
  cboToolbarSize.AddItem gsarrLangTxt(323), 0 'kleine Icons
  cboToolbarSize.AddItem gsarrLangTxt(324), 1 'grosse Icons
  cboToolbarSize.ListIndex = giToolbarSize
  cboToolbarSize.Enabled = CBool(chkShowToolbar.Value > 0)
  
  Call ReadIconSetNames
  cboIconSet.Clear
  For i = 1 To UBound(gsarrIconSet())
    cboIconSet.AddItem gsarrIconSet(i)
    If gsarrIconSet(i) = gsIconSet Then
      cboIconSet.ListIndex = i - 1
    End If
  Next
  cboIconSet.Enabled = IIf(UBound(gsarrIconSet()) < 1, False, True)
  
  'Fenster
  txtFieldHeight.Text = CStr(Abs(giDefaultHeight))
  txtFontSize.Text = CStr(Abs(giDefaultFontSize))
  
  'diverses
  If gbSendAuctionEnd Then chkSendOk.Value = 1 Else chkSendOk.Value = 0
  If gbSendAuctionEndNoSuccess Then chkSendNok.Value = 1 Else chkSendNok.Value = 0
  If gbSendIfLow Then chkSendLow.Value = 1 Else chkSendLow.Value = 0
  
  txtSendOkTo.Text = gsSendEndTo
  txtSendOkFrom.Text = gsSendEndFrom
  txtSendOkFromRealname.Text = gsSendEndFromRealname
  txtTestArtikel.Text = gsTestArtikel
  If gbBuyItNow Then chkBuyItNow.Value = 1 Else chkBuyItNow.Value = 0
  
  'NTP
  Select Case giUseNtp
    Case 0
      optUseHTML.Value = True
      Call optUseHTML_Click
    Case 1
      optUseTime.Value = True
      Call optUseTime_Click
    Case 2
      optUseSntp.Value = True
      Call optUseSntp_Click
  End Select
  
  txtNtpServer.Text = gsNtpServer
  
  
  'ODBC
  If gbUsesOdbc Then chkUsesOdbc.Value = 1 Else chkUsesOdbc.Value = 0
  txtOdbcZyklus.Text = CStr(giOdbcZyklus)
  txtOdbcProvider.Text = gsOdbcProvider
  txtOdbcDB.Text = gsOdbcDb
  txtOdbcUser.Text = gsOdbcUser
  txtOdbcPass.Text = gsOdbcPass
   
  chkPopUseSSL.Enabled = ShellTest(gsPopCmdSSL, vbHide)
  chkSmtpUseSSL.Enabled = ShellTest(gsSmtpCmdSSL, vbHide)
  If chkPopUseSSL.Enabled = False Then chkPopUseSSL.Value = vbUnchecked
  If chkSmtpUseSSL.Enabled = False Then chkSmtpUseSSL.Value = vbUnchecked
  
  SSTab1.Tab = 0
  
  mbChangeFlag = False
  
  Exit Sub
  
errhdl:
  MsgBox "Error frmSettings FormLoad: " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Dim lRet As VbMsgBoxResult
    
    lRet = vbNo
    
    If mbChangeFlag Then
        lRet = MsgBox(gsarrLangTxt(104), vbYesNoCancel Or vbQuestion)
    End If
    
    If lRet = vbCancel Then
        Cancel = 1
    Else
        gbSettingsIsUp = False
        If lRet = vbYes Then
            Call btnUebernehmen_Click
            Call SaveAllSettings
        End If
        If gbEmpUserEnd Then Call frmHaupt.AutoStatus_Click
    End If
    
End Sub

Private Sub Frame3_Click()
  
  If txtFocusRectColor.Enabled = False Then Exit Sub

  Dim CurrColor As Long
  CurrColor = GetColorFromRgbHex(txtFocusRectColor)
  If Dialog_Color(Me, CurrColor) Then txtFocusRectColor = GetRgbHexFromColor(CurrColor)
  DoEvents
  SSTab1.Tab = 5

End Sub

Private Sub optLan_Click()

  txtModemVorlauf.Enabled = False
  chkAktualisieren.Enabled = True
  txtPreConnect.Enabled = True
  chkCheckForUpdate.Enabled = True
  chkCheckForUpdate_Click
  chkAutoUpdateCurrencies.Enabled = True
  chkUsePre.Enabled = True
  btnDfueLesen.Enabled = False
  Label4.Enabled = True
  label5.Enabled = False
'  SSTab1.TabEnabled(3) = True
  chkUsePop.Enabled = True
  chkUseInline.Enabled = True

  mbChangeFlag = True
  
End Sub

Private Sub optLan_GotFocus()
  SSTab1.Tab = 1
End Sub

Private Sub optModem_Click()

  txtModemVorlauf.Enabled = True
  chkAktualisieren.Enabled = False
  txtPreConnect.Enabled = False
  chkCheckForUpdate.Enabled = False
  chkCheckForUpdate_Click
  chkAutoUpdateCurrencies.Enabled = False
  chkUsePre.Enabled = False
  Label4.Enabled = False
  label5.Enabled = True
  btnDfueLesen.Enabled = True
'  SSTab1.TabEnabled(3) = False
  chkUsePop = 0
  chkUsePop.Enabled = False
  chkUseInline.Enabled = False
  
  If lstDfue.ListCount = 0 Then
    GetDFUEList
  End If
  
  mbChangeFlag = True
  
End Sub

Private Sub optModem_GotFocus()
  SSTab1.Tab = 1
End Sub

Private Sub btnSpeedCheck_Click()
    
    On Error Resume Next
    Dim fLfz As Double
    Dim fVlfz As Integer
    Dim lRet As VbMsgBoxResult
    Dim sUsername As String
    Dim sPassword As String
    Dim bUseToken As Boolean

    gbStopTests = False
    
    If giUserAnzahl = 0 Then
        MsgBox gsarrLangTxt(2), vbInformation Or vbOKOnly
    Else
    
        If Not CheckInternetConnection Then
            Call frmHaupt.Ask_Online
        End If
        
        If CheckInternetConnection Then
            lRet = MsgBox(gsarrLangTxt(108), vbYesNo Or vbQuestion)
            If lRet = vbYes Then
                sUsername = gtarrUserArray(cboUsers.ListIndex + 1).UaUser
                sPassword = gtarrUserArray(cboUsers.ListIndex + 1).UaPass
                bUseToken = gtarrUserArray(cboUsers.ListIndex + 1).UaToken
                
                Screen.MousePointer = vbHourglass
                fLfz = frmHaupt.CheckSpeed(10, sUsername, sPassword, bUseToken)
                Screen.MousePointer = vbNormal
                fVlfz = fLfz + 2
                MsgBox gsarrLangTxt(109) & Format(fLfz, "#.##") & gsarrLangTxt(110) & vbCrLf & vbCrLf & gsarrLangTxt(111) _
                    & vbCrLf & vbCrLf & gsarrLangTxt(112) & fVlfz & " sec", vbInformation
                    
                If gbUsesModem And gbLastDialupWasManually Then Call frmHaupt.Ask_Offline
            End If 'lRet=vbYes
        Else
            MsgBox gsarrLangTxt(113), vbExclamation
        End If
    End If 'giUserAnzahl = 0
End Sub

Private Sub btnSpeedCheck_GotFocus()
  SSTab1.Tab = 0
End Sub

Private Sub btnSpeichern_Click()
  btnUebernehmen_Click
  SaveAllSettings
  mbChangeFlag = False
End Sub

Private Sub btnTestPop_Click()

  Dim bOk As Boolean
  Dim sTestOut As String

  On Error Resume Next

  If mbChangeFlag Then
    If Not vbYes = MsgBox("Dieser Test erfordert die ‹bernahme der Einstellungen." & vbCrLf & "Sollen die Einstellungen jetzt ¸bernommen und der Test durchgef¸hrt werden ?", vbYesNo + vbQuestion) Then Exit Sub
    btnUebernehmen_Click
  End If
  
  btnTestPop.Enabled = False
  
  btnLogInTest.Enabled = False
  Screen.MousePointer = vbHourglass
  
  gsStatusTxt = ""
  gbPopTestIsOk = False
  
  bOk = PopTest
  
  sTestOut = btnTestPop.Caption & ":" & vbCrLf & vbCrLf & gsStatusTxt & vbCrLf
  frmHaupt.tcpIn.Close
  
  Screen.MousePointer = vbNormal
  
  gbPopTestIsOk = bOk
  If bOk Then
    MsgBox gsarrLangTxt(106) & vbCrLf & vbCrLf & sTestOut, vbInformation
  Else
    MsgBox gsarrLangTxt(107) & vbCrLf & vbCrLf & sTestOut, vbCritical
  End If
  
  btnLogInTest.Enabled = True
  
  bOk = gbPopTestIsOk
  
  If bOk Then
    frmHaupt.PanelText frmHaupt.StatusBar1, 2, Replace(gsarrLangTxt(84), "%MIN%", giPopZyklus), True, vbGreen
  Else
    frmHaupt.PanelText frmHaupt.StatusBar1, 2, gsarrLangTxt(85), False, vbRed
  End If
  
  btnTestPop.Enabled = True

End Sub

Private Sub btnTestSmtp_Click()

  Dim bOk As Boolean
  Dim sTestOut As String

  On Error Resume Next

  If mbChangeFlag Then
    If Not vbYes = MsgBox("Dieser Test erfordert die ‹bernahme der Einstellungen." & vbCrLf & "Sollen die Einstellungen jetzt ¸bernommen und der Test durchgef¸hrt werden ?", vbYesNo + vbQuestion) Then Exit Sub
    btnUebernehmen_Click
  End If
  
  chkSendTestMail.SetFocus
  btnTestSmtp.Enabled = False
  
  If txtSendOkTo.Text <> "" And txtSendOkFrom.Text <> "" Then
    Screen.MousePointer = vbHourglass
    bOk = SendSMTP(txtSendOkFromRealname.Text & "<" & txtSendOkFrom.Text & ">", txtSendOkTo.Text, "Subject: Biet-O-Matic Testmail" & vbCrLf & "Testmail" & vbCrLf & vbCrLf, True, (chkSendTestMail = vbUnchecked))
    Screen.MousePointer = vbNormal
    sTestOut = btnTestSmtp.Caption & ":" & vbCrLf & vbCrLf & frmHaupt.SMTP_1.SMTPDebugOutput & vbCrLf

    If bOk Then
      MsgBox gsarrLangTxt(106) & vbCrLf & vbCrLf & sTestOut, vbInformation
    Else
      MsgBox gsarrLangTxt(107) & vbCrLf & vbCrLf & sTestOut, vbCritical
    End If
  Else
    MsgBox gsarrLangTxt(105), vbExclamation
  End If
  
  btnTestSmtp.Enabled = True
  btnTestSmtp.SetFocus

End Sub

Private Sub btnTestSmtp_GotFocus()
  SSTab1.Tab = 3
End Sub

Private Sub btnTesten_Click()

  frmHaupt.WindowState = vbNormal
  
  'Font setzen
  giDefaultHeight = Abs(Val(txtFieldHeight))
  giDefaultFontSize = Abs(Val(txtFontSize))
  gsGlobFontName = cboFonts.Text
  SetFont frmHaupt
  
  'Toolbar switchen
  frmHaupt.SwitchToolbar cboToolbarSize.ListIndex, chkShowToolbar
  giToolbarSize = cboToolbarSize.ListIndex
  gbShowToolbar = chkShowToolbar

End Sub

Private Sub btnTestNtp_Click()
    
    On Error Resume Next
    
    Dim sTmp As String
    Dim sOk As String
    Dim fLap As Double
    Dim datZeit As Date
    Dim fTimeSyncStart As Double
    Dim fTimeSyncEnd As Double
    Dim fTimeSyncBefore As Double
    Dim fTimeSyncAfter As Double
    Dim iUseNTPTmp As Integer
    Dim sNTPServerTmp As String
    
    
    If Not CheckInternetConnection Then
        frmHaupt.Ask_Online
        If Not IsOnline Then
            Exit Sub
        End If
    End If
    
    lblTimeDiff.Caption = ""
    lblShowTime.Caption = gsarrLangTxt(116)
    btnTestNtp.Enabled = False
    
    'lg 12.05.2003
    'wir ziehen die Dauer des Timesync von der Differenz Uhrzeit vorher/nachher ab.
    fTimeSyncBefore = Timer()
    fTimeSyncStart = GetSystemUptime()
    
    
    sNTPServerTmp = gsNtpServer
    iUseNTPTmp = giUseNtp
    
    If optUseHTML Then giUseNtp = 0
    If optUseTime Then giUseNtp = 1
    If optUseSntp Then giUseNtp = 2
    gsNtpServer = Trim(txtNtpServer)
    
    If giUseNtp Then
        sTmp = GetINetTime()
        
        'lg 12.05.2003
        fTimeSyncAfter = Timer()
        fTimeSyncEnd = GetSystemUptime()
        fLap = fTimeSyncAfter - fTimeSyncBefore - (fTimeSyncEnd - fTimeSyncStart) + gfTimeDeviation
        datZeit = myTimeSerial(0, 0, 1) * fLap 'lg 12.05.2003
        If sTmp = "" Then
            lblShowTime.Caption = gsarrLangTxt(117)
        Else
            lblShowTime.Caption = gsarrLangTxt(118) & Date2Str(MyNow) 'lg 12.05.2003
            lblTimeDiff.Caption = gsarrLangTxt(119) & Format(datZeit, "hh:mm:ss") 'lg 22.04.2004
        End If
    Else
        sOk = frmHaupt.sync_ebaytime
        
        'lg 12.05.2003
        fTimeSyncAfter = Timer()
        fTimeSyncEnd = GetSystemUptime()
        fLap = fTimeSyncAfter - fTimeSyncBefore - (fTimeSyncEnd - fTimeSyncStart)
        datZeit = myTimeSerial(0, 0, 1) * fLap 'lg 12.05.2003
        If sOk <> "" Then
            lblShowTime.Caption = gsarrLangTxt(118) & Date2Str(MyNow) 'lg 12.05.2003
            lblTimeDiff.Caption = gsarrLangTxt(119) & Format(datZeit, "hh:mm:ss") 'lg 22.04.2004
            Call DebugPrint("Zeitsync Differenz " & CStr(CInt(fLap)) & " Sekunden")
        Else
            lblShowTime.Caption = gsarrLangTxt(117)
            Call DebugPrint("Zeitsync nicht erfolgreich")
        End If
    End If
    
    gsNtpServer = sNTPServerTmp
    giUseNtp = iUseNTPTmp
    
    btnTestNtp.Enabled = True
    If gbUsesModem And gbLastDialupWasManually Then frmHaupt.Ask_Offline
    
End Sub

Private Sub btnTestPlaySound_Click()
  PlaySound txtSoundOnBid.Text
End Sub

Private Sub btnTestPlaySoundFail_Click()
  PlaySound txtSoundOnBidFail
End Sub

Private Sub btnTestPlaySoundSuccess_Click()
  PlaySound txtSoundOnBidSuccess.Text
End Sub

Private Sub txtReLogin_Change()
  mbChangeFlag = True
End Sub

Private Sub btnUebernehmen_Click()
  
  'On Error GoTo errhdl:
  Dim bChangeFlagSave As Boolean
  Dim obj As Object

  bChangeFlagSave = mbChangeFlag
  
  If Val(txtVorlauf.Text) < 2 Then
    MsgBox gsarrLangTxt(114)
    Exit Sub
  End If
  
  If Val(txtVorlaufSnipe.Text) < 0 Then
    txtVorlaufSnipe = 0
  End If
  
  'gsUser = Trim(f_user)
  'gsPass = f_pass

  giUserAnzahl = cboUsers.ListCount
  If giUserAnzahl > 0 Then
    giDefaultUser = cboUsers.ListIndex + 1
    If gsUser <> gtarrUserArray(giDefaultUser).UaUser Then gsEbayLocalPass = "" ' damit in MyEbay neu eingeloggt wird
    gsUser = gtarrUserArray(giDefaultUser).UaUser
    gsPass = gtarrUserArray(giDefaultUser).UaPass
    gbUseSecurityToken = gtarrUserArray(giDefaultUser).UaToken
  Else
    gsUser = ""
    gsPass = ""
  End If
  
  gbPlaySoundOnBid = CBool(chkPlaySoundOnBid.Value > 0)
  gsSoundOnBid = txtSoundOnBid.Text
  gsSoundOnBidSuccess = txtSoundOnBidSuccess.Text
  gsSoundOnBidFail = txtSoundOnBidFail.Text
  
  glVorlaufGebot = txtVorlauf.Text
  gfVorlaufSnipe = CDblSave(txtVorlaufSnipe.Text, gfVorlaufSnipe)
  'Vorlaufzeit wandeln
  gfVorlaufGebotTimeVal = myTimeSerial(0, 0, 1) * glVorlaufGebot 'lg 12.05.2003

  gbUsesModem = optModem.Value
  glVorlaufModem = Val(txtModemVorlauf.Text)
  If Not gbUsesModem Then
    gbGeboteAktualisieren = CBool(chkAktualisieren.Value > 0)
  Else
    gbGeboteAktualisieren = False
  End If
  
  gbAutoStart = CBool(chkAutoStart.Value > 0)
  gsConnectName = lstDfue.List(lstDfue.ListIndex)
  If gsConnectName = "" Then
    gsConnectName = "--"
  End If
  
  gsPopUser = txtPopUser.Text
  gsPopServer = txtPopServer.Text
  giPopPort = Val(txtPopPort.Text)
  gsPopPass = txtPopPass.Text
  giPopZyklus = Val(txtPopZykl)
  giPopTimeOut = Val(txtPopTimeOut.Text)
  frmHaupt.TimeoutTimer.Interval = 1000 ' 1 sec
  gsSmtpServer = txtSmtpServer.Text
  giSmtpPort = Val(txtSmtpPort.Text)
  gsAbsender = txtAbsender.Text
  gbUseSmtpAuth = chkUseAuth.Value
  gbPopUseSSL = chkPopUseSSL.Value
  gbSmtpUseSSL = chkSmtpUseSSL.Value
  gbPopEncryptedOnly = chkPopEncryptedOnly.Value
  
  gbUsePop = CBool(chkUsePop.Value > 0)
  chkUsePop_Click
  
  If chkUsePre.Value Then
      giVorlaufLan = Val(txtPreConnect.Text)
  Else
      giVorlaufLan = 0
  End If
  
  'automatisch vor Auktion einloggen (mae 050718)
  If chkReLogin Then
      giReLogin = Val(txtReLogin.Text)
  Else
      giReLogin = 0
  End If
  
  If chkCheckForUpdate.Value > 0 Then gbCheckForUpdate = True Else gbCheckForUpdate = False
  If chkCheckForUpdateBeta.Value > 0 Then gbCheckForUpdateBeta = True Else gbCheckForUpdateBeta = False
  glCheckForUpdateInterval = Val(txtCheckForUpdateInterval.Text)
  If glCheckForUpdateInterval = 0 Then glCheckForUpdateInterval = 1
  If chkAutoUpdateCurrencies.Value > 0 Then gbAutoUpdateCurrencies = True Else gbAutoUpdateCurrencies = False
  
  'gbUseIntBrowser = optUseInt.Value 'MD-Marker 20090325 , Ctrl wurde entfernt
  gbBrowseInNewWindow = chkUseNewWin.Value
  gbBrowseInline = chkUseInline.Value
  giInlineBrowserDelay = Val(txtInlineBrowserDelay.Text)
  giInlineBrowserModifierKey = cboInlineBrowserModifierKey.ListIndex
  gbFileWinShutdown = CBool(chkEndWin.Value = 1)
  
  'Startup- Size
  If optStartMax.Value Then giStartupSize = vbMaximized
  If optStartNormal.Value Then giStartupSize = vbNormal
  If optStartMin.Value Then giStartupSize = vbMinimized
  
  'Font
  gsGlobFontName = cboFonts.Text
  
  '1.7.1 Serververhalten
  gbTrayAction = CBool(chkDoServer.Value = 1)
  '1.7.3
  'Proxy
  gbUseProxy = CBool(optUseProxy.Value <> 0)
  gbUseDirectConnect = CBool(optConnectDirect.Value <> 0)
  gbUseProxyAuthentication = CBool(chkUseProxyAuth.Value <> 0)

  gsProxyName = txtProxyName.Text
  giProxyPort = Val(txtProxyPort.Text)
  gsProxyUser = txtProxyUser.Text
  gsProxyPass = txtProxyPass.Text
  
  ' 1.8.0
  'gbUseIECookies = CBool(chkUseIECookies.Value <> 0) ' wir wollen keine IE-Cookies mehr, das gibt nur Stress!
  gbUseCurl = CBool(chkUseCurl.Value <> 0)
  gsBrowserIdString = txtBrowserString.Text
  giMaxRowSetting = CInt(Val(txtMaxArtikel.Text)) - 1
  giDefaultHeight = Abs(Val(txtFieldHeight.Text))
  giDefaultFontSize = Abs(Val(txtFontSize.Text))
  gbPassAtStart = (chkStartPass.Value = 1)
  gbArtikelRefreshPost = (chkPostGebAktualisieren.Value = 1)
  gbArtikelRefreshPost2 = (chkPostGebAktualisieren2.Value = 1)
  giArtikelRefreshCycle = Val(txtAktCycle.Text)
  gbAutoLogin = CBool(chkAutoLogin.Value = 1)
  gbAutoAktualisieren = CBool(chkAutoAktualisieren.Value = 1)
  giAktualisierenOpt = cboAktualisierenOpt.ListIndex
  gbAutoWarnNoBid = CBool(chkAutoWarnNoBid.Value)
  gbEditShippingOnClick = CBool(chkEditShippingOnClick.Value)
  gbConcurrentUpdates = CBool(chkMultiAkt.Value)
  gbUpdateAfterManualBid = CBool(chkPostManBidAktualisieren.Value)
  gbQuietAfterManualBid = CBool(chkQuietAfterManBid.Value)
  
  'sh 25.10.03 nur n‰chste +  bis x min vor alle x min
  gbAutoAktualisierenNext = CBool(chkAutoAktNext.Value)
  gbAktualisierenXvor = CBool(chkAktualXvor.Value = 1)
  giAktXminvor = txtAktXminvor.Text
  giAktXminvorCycle = txtAktXminvorCycle
  giArtAktOptions = IIf(chkAktualXvor.Value = 1, cboArtAktOptions.ListIndex, 0)
  giArtAktOptionsValue = txtArtAktOptionsValue.Text
  gbWarnenBeimBeenden = CBool(chkWarnenBeimBeenden.Value = 1)
  gbBeendenNachAuktion = CBool(chkBeendenNachAuktion.Value = 1)
  giReloadTimes = CInt(txtReloadTimes.Text)
  
  frmHaupt.AutoSave.Enabled = (chkAutoSave = 1)
  
  'Timesync
  '0: kein, 1: einmal, 2: vor Gebot, 4: beim Start, 8: regelm‰ﬂig alle x Minuten
  giUseTimeSync = 0
  If chkDay.Value Then giUseTimeSync = giUseTimeSync Or 1
  If chkPreGeb.Value Then giUseTimeSync = giUseTimeSync Or 2
  If chkStart.Value Then giUseTimeSync = giUseTimeSync Or 4
  If chkRepeatEvery.Value Then giUseTimeSync = giUseTimeSync Or 8
  
  glTimeSyncIntervall = Val(txtTimeSyncIntervall.Text)
  
  gbUseWheel = CBool(chkUseWheel.Value = 1)
  
  If gbShowToolbar <> CBool(chkShowToolbar.Value = 1) Or _
     giToolbarSize <> cboToolbarSize.ListIndex Then
    frmHaupt.SwitchToolbar cboToolbarSize.ListIndex, chkShowToolbar
  End If
  gbShowToolbar = CBool(chkShowToolbar.Value = 1)
  giToolbarSize = cboToolbarSize.ListIndex
  
  gsIconSet = IIf(cboIconSet.ListIndex >= 0, cboIconSet.Text, "")
  
  gbSendAuctionEnd = CBool(chkSendOk.Value = 1)
  gbSendAuctionEndNoSuccess = CBool(chkSendNok.Value = 1)
  gbSendIfLow = CBool(chkSendLow.Value = 1)
  gsSendEndTo = txtSendOkTo.Text
  gsSendEndFrom = txtSendOkFrom.Text
  gsSendEndFromRealname = txtSendOkFromRealname.Text
  gsTestArtikel = txtTestArtikel.Text
  gbBuyItNow = CBool(chkBuyItNow.Value = 1)
  gbTestConnect = CBool(chkTestConnect.Value = 1)
  gbOperaField = CBool(chkOperaField.Value = 1)
  gbShowSplash = CBool(chkShowSplash.Value = 1)
  gbNewItemWindowKeepsValues = CBool(chkNewItemWindowKeepsValues.Value = 1)
  gbNewItemWindowAlwaysOnTop = CBool(chkNewItemWindowAlwaysOnTop.Value = 1)
  gbNewItemWindowOpenOnStartup = CBool(chkNewItemWindowOpenOnStartup.Value = 1)
  For Each obj In Forms
    If obj.Name = "frmNeuerArtikel" Then
      frmNeuerArtikel.Check2.Value = IIf(gbNewItemWindowKeepsValues, vbChecked, vbUnchecked)
      Exit For
    End If
  Next
  gbMinToTray = CBool(chkMinToTray.Value = 1)
  gbKeinHinweisNachZeitsync = CBool(chkKeinHinweisNachZeitsync.Value = 1)
  gbShowTitleDateTime = CBool(chkShowTitleDateTime.Value = 1)
  gbShowTitleTimeLeft = CBool(chkShowTitleTimeLeft.Value = 1)
  gbCleanStatus = CBool(chkCleanStatus.Value = 1)
  glCleanStatusTime = Val(txtCleanStatusTime.Text)
  gbShowShippingCosts = CBool(chkShowShippingCosts.Value = 1)
  gbShowWeekday = CBool(chkShowWeekday.Value = 1)
  gbShowFocusRect = CBool(chkShowFocusRect.Value = 1)
  glFocusRectColor = GetColorFromRgbHex(txtFocusRectColor.Text)
  
  'NTP
  If optUseHTML.Value Then giUseNtp = 0
  If optUseTime.Value Then giUseNtp = 1
  If optUseSntp.Value Then giUseNtp = 2
  gsNtpServer = Trim(txtNtpServer.Text)
  
  'ODBC
  gbUsesOdbc = CBool(chkUsesOdbc.Value = 1)
  giOdbcZyklus = Val(txtOdbcZyklus.Text)
  gsOdbcProvider = txtOdbcProvider.Text
  gsOdbcDb = txtOdbcDB.Text
  gsOdbcUser = txtOdbcUser.Text
  gsOdbcPass = txtOdbcPass.Text
  
  frmHaupt.Zusatzfeld.Visible = gbUsesOdbc Or gbOperaField 'lg 27.05.03
  frmHaupt.Zusatzfeld.Enabled = Not gbUsesOdbc
  If Not gbUsesOdbc Then
    frmHaupt.ODBC_Timer.Enabled = False
    frmHaupt.Zusatzfeld.BackColor = &H8000000F
    frmHaupt.Zusatzfeld.Text = ""
  End If
  
  gsServerStringsFile = cboServerStrings.List(cboServerStrings.ListIndex)
  modKeywords.ReadAllKeywords
  
  mbChangeFlag = bChangeFlagSave
  Exit Sub
  
errhdl:
  MsgBox gsarrLangTxt(115) & Err.Description
End Sub

Private Sub btnUserCancel_Click()
If btnAddUser.Enabled Then
  btnAddUser.Caption = gsarrLangTxt(717)
   btnEditUser.Enabled = (giUserAnzahl > 0)
   btnDelUser.Enabled = (giUserAnzahl > 0)
  btnUserCancel.Enabled = False
Else
  btnEditUser.Caption = gsarrLangTxt(718)
  btnAddUser.Enabled = True
  btnDelUser.Enabled = True
  btnUserCancel.Enabled = False
End If
  txtUsersNeuEdit.Text = ""
  txtUsersNeuEdit.Visible = False
  txtPassNeuEdit.Text = ""
  txtPassNeuEdit.Visible = False
  chkUseSecurityTokenNeuEdit.Value = vbUnchecked
  chkUseSecurityTokenNeuEdit.Visible = False
  cboUsers.Visible = True
  txtPass1.Visible = True
  chkUseSecurityToken.Visible = True
  cboUsers.TabIndex = 1
  txtPass1.TabIndex = 2
  chkUseSecurityToken.TabIndex = 3
End Sub

Private Sub btnVerwerfen_Click()

  Dim lTabSave As Long
  lTabSave = SSTab1.Tab
  
  Call ReadAllSettings
  Call Form_Load 'MD-Marker
  If giUserAnzahl > 0 Then btnDelUser.Enabled = True
  
  SSTab1.Tab = lTabSave

End Sub

Private Sub SetLanguage()

  Me.Caption = gsarrLangTxt(215) & " - " & gsarrLangTxt(32)
  SSTab1.TabCaption(0) = gsarrLangTxt(120)
  SSTab1.TabCaption(1) = gsarrLangTxt(121)
  SSTab1.TabCaption(2) = gsarrLangTxt(122)
  SSTab1.TabCaption(3) = gsarrLangTxt(123)
  SSTab1.TabCaption(4) = gsarrLangTxt(124)
  SSTab1.TabCaption(5) = gsarrLangTxt(125)
  SSTab1.TabCaption(6) = gsarrLangTxt(126)
  SSTab1.TabCaption(7) = gsarrLangTxt(127)
  
  Frame11.Caption = gsarrLangTxt(374)
  Frame12.Caption = gsarrLangTxt(375)
  Label1.Caption = gsarrLangTxt(360)
  Label2.Caption = gsarrLangTxt(130)
  Label3.Caption = gsarrLangTxt(131)
  Label4.Caption = gsarrLangTxt(132)
  Label28.Caption = gsarrLangTxt(133)
  
  Frame8.Caption = gsarrLangTxt(134)
  optLan.Caption = gsarrLangTxt(135)
  Label14.Caption = gsarrLangTxt(136)
  chkCheckForUpdate.Caption = gsarrLangTxt(137)
  chkCheckForUpdateBeta.Caption = gsarrLangTxt(464)
  chkAutoUpdateCurrencies.Caption = gsarrLangTxt(399)
  Label59.Caption = gsarrLangTxt(475)
  Label60.Caption = gsarrLangTxt(476)
  optModem.Caption = gsarrLangTxt(138)
  label5.Caption = gsarrLangTxt(139)
  Label15.Caption = gsarrLangTxt(140)
  chkTestConnect.Caption = gsarrLangTxt(141)
  chkStartPass.Caption = gsarrLangTxt(142)
  chkAutoStart.Caption = gsarrLangTxt(143)
  chkAutoLogin.Caption = gsarrLangTxt(144)
  chkDoServer.Caption = gsarrLangTxt(145)
  chkEndWin.Caption = gsarrLangTxt(146)
  chkAktualisieren.Caption = gsarrLangTxt(147)
  chkMultiAkt.Caption = gsarrLangTxt(470)
  chkPostManBidAktualisieren.Caption = gsarrLangTxt(471)
  chkQuietAfterManBid.Caption = gsarrLangTxt(472)
  Label24.Caption = gsarrLangTxt(361)
  chkPostGebAktualisieren.Caption = gsarrLangTxt(149)
  chkAutoSave.Caption = gsarrLangTxt(150)
  Frame7.Caption = gsarrLangTxt(151)
  Frame5.Caption = gsarrLangTxt(152)
  chkAutoAktualisieren.Caption = gsarrLangTxt(376)
  'sh
  chkAutoAktNext.Caption = gsarrLangTxt(700)
  chkAktualXvor.Caption = gsarrLangTxt(701)
  Label47.Caption = gsarrLangTxt(702)
  Label48.Caption = gsarrLangTxt(703)
  cboUsers.ToolTipText = gsarrLangTxt(716)
  btnAddUser.Caption = gsarrLangTxt(717)
  btnEditUser.Caption = gsarrLangTxt(718)
  btnUserCancel.Caption = gsarrLangTxt(719)
  btnDelUser.Caption = gsarrLangTxt(720)
  Label49.Caption = gsarrLangTxt(709)
  cboArtAktOptions.ToolTipText = gsarrLangTxt(711)
  cboArtAktOptions.List(0) = gsarrLangTxt(712)
  cboArtAktOptions.List(1) = gsarrLangTxt(713)
  cboArtAktOptions.List(2) = gsarrLangTxt(714)
  Label50.Caption = gsarrLangTxt(725)
  cboAktualisierenOpt.List(0) = gsarrLangTxt(738)
  cboAktualisierenOpt.List(1) = gsarrLangTxt(739)
  chkAutoWarnNoBid.Caption = gsarrLangTxt(741)
  chkEditShippingOnClick.Caption = gsarrLangTxt(710)
  
  Label52.Caption = gsarrLangTxt(737)
  Frame19.Caption = gsarrLangTxt(707)
  Frame21.Caption = gsarrLangTxt(715)
  
  chkPostGebAktualisieren2.Caption = gsarrLangTxt(377)
  chkWarnenBeimBeenden.Caption = gsarrLangTxt(389)
  chkBeendenNachAuktion.Caption = gsarrLangTxt(390)
  
  Frame13.Caption = gsarrLangTxt(378)
  Frame14.Caption = gsarrLangTxt(379)
  chkUsePop.Caption = gsarrLangTxt(153)
  chkUseAuth.Caption = gsarrLangTxt(362)
  chkPopUseSSL.Caption = gsarrLangTxt(462)
  chkSmtpUseSSL.Caption = gsarrLangTxt(462)
  chkPopEncryptedOnly.Caption = gsarrLangTxt(463)
  chkSendTestMail.Caption = gsarrLangTxt(380)
  Label6.Caption = gsarrLangTxt(363)
  Label7.Caption = gsarrLangTxt(364)
  Label8.Caption = gsarrLangTxt(360)
  Label10.Caption = gsarrLangTxt(365)
  Label11.Caption = gsarrLangTxt(361)
  Label12.Caption = gsarrLangTxt(366)
  Label9.Caption = gsarrLangTxt(130)
  Label13.Caption = gsarrLangTxt(154)
  Label19.Caption = gsarrLangTxt(155)
  Label34.Caption = gsarrLangTxt(156)
  Label16.Caption = gsarrLangTxt(157)
  Label17.Caption = gsarrLangTxt(158)
  Label18.Caption = gsarrLangTxt(159)
  Label53.Caption = gsarrLangTxt(456)
  Label54.Caption = gsarrLangTxt(457)
  Label55.Caption = gsarrLangTxt(459)
  Label56.Caption = gsarrLangTxt(459)
  
  Label22.Caption = gsarrLangTxt(160)
  Frame2.Caption = gsarrLangTxt(161)
  optStartMax.Caption = gsarrLangTxt(162)
  optStartMin.Caption = gsarrLangTxt(164)
  optStartNormal.Caption = gsarrLangTxt(163)
  'optUseInt.Caption = gsarrLangTxt(165) 'MD-Marker 20090325 , Ctrl wurde entfernt
  optUseExt.Caption = gsarrLangTxt(166)
  chkUseNewWin.Caption = gsarrLangTxt(455)
  chkUseInline.Caption = gsarrLangTxt(465)
  chkShowSplash.Caption = gsarrLangTxt(167)
  chkUseWheel.Caption = gsarrLangTxt(168)
  chkShowToolbar.Caption = gsarrLangTxt(169)
  chkOperaField.Caption = gsarrLangTxt(170)
  
  Frame1.Caption = gsarrLangTxt(382)
  Frame15.Caption = gsarrLangTxt(381)
  Frame4.Caption = gsarrLangTxt(369)
  
  Label25.Caption = gsarrLangTxt(173)
  Label26.Caption = gsarrLangTxt(174)
  Label33.Caption = gsarrLangTxt(370)
  chkMinToTray.Caption = gsarrLangTxt(383)
  chkNewItemWindowKeepsValues.Caption = gsarrLangTxt(384)
  chkNewItemWindowAlwaysOnTop.Caption = gsarrLangTxt(392)
  chkNewItemWindowOpenOnStartup.Caption = gsarrLangTxt(481)
  Label30.Caption = gsarrLangTxt(466)
  Label31.Caption = gsarrLangTxt(477)
  
  'Bieten II
  chkSendOk.Caption = gsarrLangTxt(175)
  chkSendNok.Caption = gsarrLangTxt(391)
  chkSendLow.Caption = gsarrLangTxt(468)
  Label23.Caption = gsarrLangTxt(371)
  Label27.Caption = gsarrLangTxt(176)
  Label32.Caption = gsarrLangTxt(177)
  Label42.Caption = gsarrLangTxt(396)
  optConnectDefault.Caption = gsarrLangTxt(178)
  optUseProxy.Caption = gsarrLangTxt(179)
  optConnectDirect.Caption = gsarrLangTxt(458)
  chkUseIECookies.Caption = gsarrLangTxt(461)
  chkUseCurl.Caption = gsarrLangTxt(467)
  Label21.Caption = gsarrLangTxt(180)
  Label20.Caption = gsarrLangTxt(181)
  chkUseProxyAuth.Caption = gsarrLangTxt(393)
  Label44.Caption = gsarrLangTxt(394)
  Label43.Caption = gsarrLangTxt(395)
  
  'TimeSync
  Frame9.Caption = gsarrLangTxt(182)
  Frame6.Caption = gsarrLangTxt(183)
  optUseHTML.Caption = gsarrLangTxt(184)
  optUseTime.Caption = gsarrLangTxt(185)
  optUseSntp.Caption = gsarrLangTxt(186)
  Label29.Caption = gsarrLangTxt(372)
  lblShowTime.Caption = gsarrLangTxt(187)
  chkDay.Caption = gsarrLangTxt(189)
  chkPreGeb.Caption = gsarrLangTxt(190)
  chkStart.Caption = gsarrLangTxt(385)
  chkRepeatEvery.Caption = gsarrLangTxt(398)
  Label45.Caption = gsarrLangTxt(363)
  chkKeinHinweisNachZeitsync.Caption = gsarrLangTxt(388)
  'odbc
  chkUsesOdbc.Caption = gsarrLangTxt(191)
  Label36.Caption = gsarrLangTxt(192)
  Label37.Caption = gsarrLangTxt(193)
  Label38.Caption = gsarrLangTxt(194)
  Label39.Caption = gsarrLangTxt(195)
  Label40.Caption = gsarrLangTxt(196)
  Label41.Caption = gsarrLangTxt(363)
  btnLogInTest.Caption = gsarrLangTxt(197)
  btnSpeedCheck.Caption = gsarrLangTxt(198)
  btnDfueLesen.Caption = gsarrLangTxt(199)
  btnTestPop.Caption = gsarrLangTxt(200)
  btnTesten.Caption = gsarrLangTxt(387)
  btnTestSmtp.Caption = gsarrLangTxt(201)
  btnTestNtp.Caption = gsarrLangTxt(386)
  btnOdbcConnect.Caption = gsarrLangTxt(202)
  btnSpeichern.Caption = gsarrLangTxt(203)
  btnUebernehmen.Caption = gsarrLangTxt(204)
  btnVerwerfen.Caption = gsarrLangTxt(205)
  btnAbbruch.Caption = gsarrLangTxt(207)
  chkPlaySoundOnBid.Caption = gsarrLangTxt(397)
  chkShowTitleDateTime.Caption = gsarrLangTxt(450)
  chkShowTitleTimeLeft.Caption = gsarrLangTxt(451)
  chkCleanStatus.Caption = gsarrLangTxt(452)
  Label46.Caption = gsarrLangTxt(361)
  chkUseSecurityToken.Caption = gsarrLangTxt(460)
  chkUseSecurityTokenNeuEdit.Caption = gsarrLangTxt(460)
  chkBuyItNow.Caption = gsarrLangTxt(732) & " (" & gsarrLangTxt(122) & ")"
  chkShowShippingCosts.Caption = gsarrLangTxt(469)
  chkShowWeekday.Caption = gsarrLangTxt(474)
  chkShowFocusRect.Caption = gsarrLangTxt(473)
  
  'neu:ToolTips
  txtPass1.ToolTipText = gsarrLangTxt(130)
  txtPassNeuEdit.ToolTipText = gsarrLangTxt(130)
  chkUseSecurityTokenNeuEdit.ToolTipText = gsarrLangTxt(520)
  chkPlaySoundOnBid.ToolTipText = gsarrLangTxt(348)
  txtSoundOnBid.ToolTipText = gsarrLangTxt(349)
  txtSoundOnBidSuccess.ToolTipText = gsarrLangTxt(511)
  txtSoundOnBidFail.ToolTipText = gsarrLangTxt(512)
  txtVorlauf.ToolTipText = gsarrLangTxt(270)
  txtVorlaufSnipe.ToolTipText = gsarrLangTxt(270)
  txtTestArtikel.ToolTipText = gsarrLangTxt(271)
  chkBuyItNow.ToolTipText = "" ' RTFM
  btnLogInTest.ToolTipText = gsarrLangTxt(272)
  btnSpeedCheck.ToolTipText = gsarrLangTxt(273)
  optLan.ToolTipText = gsarrLangTxt(274)
  optModem.ToolTipText = gsarrLangTxt(275)
  txtPreConnect.ToolTipText = gsarrLangTxt(276)
  chkCheckForUpdate.ToolTipText = gsarrLangTxt(277)
  chkAutoUpdateCurrencies.ToolTipText = gsarrLangTxt(510)
  txtModemVorlauf.ToolTipText = gsarrLangTxt(278)
  btnDfueLesen.ToolTipText = gsarrLangTxt(279)
  chkTestConnect.ToolTipText = gsarrLangTxt(280)
  chkStartPass.ToolTipText = gsarrLangTxt(281)
  chkAutoStart.ToolTipText = gsarrLangTxt(282)
  chkDoServer.ToolTipText = gsarrLangTxt(283)
  chkEndWin.ToolTipText = gsarrLangTxt(284)
  chkAktualisieren.ToolTipText = gsarrLangTxt(285)
  chkMultiAkt.ToolTipText = gsarrLangTxt(536)
  chkPostManBidAktualisieren.ToolTipText = gsarrLangTxt(537)
  chkQuietAfterManBid.ToolTipText = gsarrLangTxt(538)
  txtAktXminvor.ToolTipText = gsarrLangTxt(706) 'sh
  chkPostGebAktualisieren.ToolTipText = gsarrLangTxt(287)
  chkAutoSave.ToolTipText = gsarrLangTxt(288)
  chkUsePop.ToolTipText = gsarrLangTxt(289)
  txtPopServer.ToolTipText = gsarrLangTxt(290)
  txtAbsender.ToolTipText = gsarrLangTxt(291)
  chkShowSplash.ToolTipText = gsarrLangTxt(292)
  chkUseWheel.ToolTipText = gsarrLangTxt(293)
  chkShowToolbar.ToolTipText = gsarrLangTxt(294)
  txtMaxArtikel.ToolTipText = gsarrLangTxt(509)
  chkOperaField.ToolTipText = gsarrLangTxt(295)
'  f_small.ToolTipText = gsarrLangTxt(296)
'  f_high.ToolTipText = gsarrLangTxt(297)
'  f_special.ToolTipText = gsarrLangTxt(298)
  chkMinToTray.ToolTipText = gsarrLangTxt(339)
  chkNewItemWindowKeepsValues.ToolTipText = gsarrLangTxt(340)
  chkNewItemWindowAlwaysOnTop.ToolTipText = gsarrLangTxt(346)
  chkNewItemWindowOpenOnStartup.ToolTipText = gsarrLangTxt(542)
  cboFonts.ToolTipText = gsarrLangTxt(299)
  txtFontSize.ToolTipText = gsarrLangTxt(300)
  txtFieldHeight.ToolTipText = gsarrLangTxt(301)
  btnTesten.ToolTipText = gsarrLangTxt(302)
  chkSendOk.ToolTipText = gsarrLangTxt(303)
  chkSendNok.ToolTipText = gsarrLangTxt(345)
  chkSendLow.ToolTipText = gsarrLangTxt(534)
  txtSendOkTo.ToolTipText = gsarrLangTxt(304)
  txtSendOkFrom.ToolTipText = gsarrLangTxt(305)
  txtSendOkFromRealname.ToolTipText = gsarrLangTxt(347)
  optConnectDefault.ToolTipText = gsarrLangTxt(306)
  txtBrowserString.ToolTipText = gsarrLangTxt(307)
  optUseProxy.ToolTipText = gsarrLangTxt(308)
  chkUseProxyAuth.ToolTipText = gsarrLangTxt(322)
  optConnectDirect.ToolTipText = gsarrLangTxt(513)
  chkUseIECookies.ToolTipText = gsarrLangTxt(521)
  chkUseCurl.ToolTipText = gsarrLangTxt(533)
  optUseHTML.ToolTipText = gsarrLangTxt(309)
  optUseTime.ToolTipText = gsarrLangTxt(310)
  optUseSntp.ToolTipText = gsarrLangTxt(312)
  txtNtpServer.ToolTipText = gsarrLangTxt(311)
  chkDay.ToolTipText = gsarrLangTxt(315)
  chkPreGeb.ToolTipText = gsarrLangTxt(316)
  chkStart.ToolTipText = gsarrLangTxt(341)
  chkRepeatEvery.ToolTipText = gsarrLangTxt(502)
  txtTimeSyncIntervall.ToolTipText = gsarrLangTxt(503)
  chkKeinHinweisNachZeitsync.ToolTipText = gsarrLangTxt(342)
  btnTestNtp.ToolTipText = gsarrLangTxt(317)
  chkAutoLogin.ToolTipText = gsarrLangTxt(320)
  btnVerwerfen.ToolTipText = gsarrLangTxt(321)
  chkUseAuth.ToolTipText = gsarrLangTxt(325)
  chkPopUseSSL.ToolTipText = gsarrLangTxt(522)
  chkSmtpUseSSL.ToolTipText = gsarrLangTxt(522)
  chkPopEncryptedOnly.ToolTipText = gsarrLangTxt(523)
  chkSendTestMail.ToolTipText = gsarrLangTxt(338)
  txtSmtpServer.ToolTipText = gsarrLangTxt(326)
  btnTestPop.ToolTipText = gsarrLangTxt(327)
  txtPopZykl.ToolTipText = gsarrLangTxt(328)
  txtPopUser.ToolTipText = gsarrLangTxt(329)
  txtPopPass.ToolTipText = gsarrLangTxt(330)
  txtPopTimeOut.ToolTipText = gsarrLangTxt(331)
  btnSpeichern.ToolTipText = gsarrLangTxt(332)
  btnUebernehmen.ToolTipText = gsarrLangTxt(333)
  btnAbbruch.ToolTipText = gsarrLangTxt(334)
  btnTestSmtp.ToolTipText = gsarrLangTxt(335)
  chkAutoAktualisieren.ToolTipText = gsarrLangTxt(336)
  'sh
  chkAutoAktNext.ToolTipText = gsarrLangTxt(704)
  chkAktualXvor.ToolTipText = gsarrLangTxt(705)
  txtAktXminvor.ToolTipText = gsarrLangTxt(706)
  'f_AktXminvorCycle.ToolTipText = gsarrLangTxt(707)
  chkAutoWarnNoBid.ToolTipText = gsarrLangTxt(514)
  chkEditShippingOnClick.ToolTipText = gsarrLangTxt(532)
  txtOdbcDB.ToolTipText = gsarrLangTxt(515)
  txtOdbcPass.ToolTipText = gsarrLangTxt(516)
  txtOdbcProvider.ToolTipText = gsarrLangTxt(517)
  txtOdbcUser.ToolTipText = gsarrLangTxt(518)
  txtOdbcZyklus.ToolTipText = gsarrLangTxt(519)
  
  chkPostGebAktualisieren2.ToolTipText = gsarrLangTxt(337)
  chkWarnenBeimBeenden.ToolTipText = gsarrLangTxt(343)
  chkBeendenNachAuktion.ToolTipText = gsarrLangTxt(344)
  chkShowTitleDateTime.ToolTipText = gsarrLangTxt(500)
  chkShowTitleTimeLeft.ToolTipText = gsarrLangTxt(501)
  btnBrowseSound.ToolTipText = gsarrLangTxt(504)
  btnTestPlaySound.ToolTipText = gsarrLangTxt(505)
  btnBrowseSoundSuccess.ToolTipText = gsarrLangTxt(504)
  btnTestPlaySoundSuccess.ToolTipText = gsarrLangTxt(505)
  btnBrowseSoundFail.ToolTipText = gsarrLangTxt(504)
  btnTestPlaySoundFail.ToolTipText = gsarrLangTxt(505)
  chkCleanStatus.ToolTipText = gsarrLangTxt(506)
  txtCleanStatusTime.ToolTipText = gsarrLangTxt(507)
  chkUseNewWin.ToolTipText = gsarrLangTxt(508)
  chkUseInline.ToolTipText = gsarrLangTxt(524)
  cboToolbarSize.ToolTipText = gsarrLangTxt(525)
  'optUseInt.ToolTipText = gsarrLangTxt(526) 'MD-Marker 20090325 , Ctrl wurde entfernt
  optUseExt.ToolTipText = gsarrLangTxt(527)
  optStartNormal.ToolTipText = gsarrLangTxt(528)
  optStartMin.ToolTipText = gsarrLangTxt(529)
  optStartMax.ToolTipText = gsarrLangTxt(530)
  txtInlineBrowserDelay.ToolTipText = gsarrLangTxt(531)
  cboInlineBrowserModifierKey.ToolTipText = gsarrLangTxt(541)
  chkShowShippingCosts.ToolTipText = gsarrLangTxt(535)
  chkShowWeekday.ToolTipText = gsarrLangTxt(540)
  chkShowFocusRect.ToolTipText = gsarrLangTxt(539)
  
  lblReloadDescription.Caption = gsarrLangTxt(414)
  lblReLogin.Caption = gsarrLangTxt(742)

End Sub

Private Sub BrowseSounds(ByRef oZielTextBox As TextBox)
    
    Static sPath As String
    Dim sFile As String
    Dim sSavePath As String
    
    If sPath = "" Then sPath = CurDir
    sSavePath = CurDir
    sFile = Dialog_Open$(Me, gsarrLangTxt(453), sPath, "", gsarrLangTxt(454) & " (*.wav)|*.wav", "")
    ChDrive sSavePath
    ChDir sSavePath
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    If Len(sFile) Then oZielTextBox.Text = sPath & sFile
    
End Sub

Private Sub ReadIconSetNames()
    
    On Error Resume Next
    Dim sTmp As String
    
    ReDim gsarrIconSet(0 To 0) As String
    
    sTmp = Dir(App.Path & "\Icons\*", vbDirectory)
    Do While Len(sTmp) > 0
        If sTmp <> "." And sTmp <> ".." Then
            If (GetAttr(App.Path & "\Icons\" & sTmp) And vbDirectory) = vbDirectory Then
                'DebugPrint "Dateiname: " & sTmp
                ReDim Preserve gsarrIconSet(0 To UBound(gsarrIconSet()) + 1) As String
                gsarrIconSet(UBound(gsarrIconSet)) = sTmp
                'DebugPrint "IconSetArr-" & UBound(gsarrIconSet) & ": " & sTmp
            End If
        End If
        sTmp = Dir()
    Loop
    
End Sub

