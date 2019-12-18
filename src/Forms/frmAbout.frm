VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " About"
   ClientHeight    =   3705
   ClientLeft      =   6525
   ClientTop       =   5715
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmAbout"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   3705
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSep 
      Height          =   30
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Frame fraSep 
      Height          =   30
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   1125
      Width           =   3975
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&O K"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Timer tmrAbout 
      Enabled         =   0   'False
      Left            =   120
      Top             =   3240
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "Langfile - Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   1920
      TabIndex        =   10
      Top             =   2595
      UseMnemonic     =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyfile - Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   1920
      TabIndex        =   9
      Top             =   2310
      UseMnemonic     =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "Programm - Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   1920
      TabIndex        =   8
      Top             =   2025
      UseMnemonic     =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "Homepage : http://www.bid-o-matic.org"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1920
      MouseIcon       =   "frmAbout.frx":38E3
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   7
      Top             =   1395
      UseMnemonic     =   0   'False
      Width           =   4965
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   4
      Left            =   2640
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   6
      Top             =   1230
      UseMnemonic     =   0   'False
      Width           =   3045
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1920
      TabIndex        =   5
      Top             =   1230
      UseMnemonic     =   0   'False
      Width           =   675
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "OpenSource - Version, feel free to modify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   4005
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Der Bietautomat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   4005
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Biet-O-Matic (JARVIS-7)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4005
   End
End
Attribute VB_Name = "frmAbout"
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
'MD-Fertig 20090324
Private Const miMAXTIMERCOUNT As Integer = 1
'...
Private mbSplashMode As Boolean
Private mbSpendeActiv As Boolean
Private miTimerCount As Integer
Private Sub btnOk_Click()
  Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
      mbSplashMode = False
      mbSpendeActiv = False
      Unload Me
  End If
End Sub

Private Sub Form_Load()
Dim lPosLeft As Long, lPosTop As Long, lPosWidth As Long, lPosHeight As Long

lPosLeft = glPosLeft
lPosTop = glPosTop
lPosWidth = glPosWidth
lPosHeight = glPosHeight

If FormLoaded("frmHaupt") Then
    If frmHaupt.WindowState = vbMinimized Then
        lPosLeft = 0
        lPosTop = 0
        lPosWidth = Screen.Width
        lPosHeight = Screen.Height
    Else
        If frmHaupt.Visible Then
            lPosLeft = frmHaupt.Left
            lPosTop = frmHaupt.Top
            lPosWidth = frmHaupt.ScaleWidth
            lPosHeight = frmHaupt.ScaleHeight
        End If
    End If
End If
'...
With Me
    Set .Icon = MyLoadResPicture(105, 16)
    Call SendMessage(.hWnd, WM_SETICON, 0, .Icon)
    Call .Move((lPosLeft + (lPosWidth \ 2) - (.Width \ 2)), (lPosTop + (lPosHeight \ 2) - (.Height \ 2)), .Width, .Height)
End With
'...
lblAbout(0).Caption = gsarrLangTxt(215)
lblAbout(1).Caption = gsarrLangTxt(211)
lblAbout(2).Caption = gsarrLangTxt(214)
lblAbout(3).Visible = mbSpendeActiv  'News
With lblAbout(4) 'Spende
    .Visible = lblAbout(3).Visible 'News
    If .Visible Then
        Set .MouseIcon = lblAbout(5).MouseIcon
    End If
End With
lblAbout(5).Caption = gsarrLangTxt(213) & " : " & gsBOMUrlHP
lblAbout(7).Caption = "BOM - Version :   " & "V " & GetBOMVersion()
lblAbout(8).Caption = gsarrLangTxt(217) & " :" & "   V " & GetKeywordsFileVersion()
lblAbout(9).Caption = gsarrLangTxt(218) & " :" & " V " & GetLanguageFileVersion()
'...
btnOk.Caption = gsarrLangTxt(219)
'...
With tmrAbout
    .Enabled = mbSplashMode Or mbSpendeActiv
    If .Enabled Then .Interval = 1000
End With
'...
If mbSplashMode Then
    btnOk.Visible = False
    With Me
        .Caption = ""
        .BorderStyle = vbBSNone
        .Height = 3000
    End With
End If

gbAboutIsUp = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'...
If mbSplashMode Then
    Cancel = 1
Else
    With tmrAbout
        .Interval = 0
        .Enabled = False
    End With
    '...
    gbAboutIsUp = False
    Unload Me: Set frmAbout = Nothing
End If
'...
End Sub
Public Sub SetShowSplashOnly()
'Werden hier Controls oder die Form angesprochen, _
wird -> Sofort <- die Prozedur Form_Load von hier _
aufgerufen und das ist nicht gewünscht.
mbSplashMode = True
End Sub

Public Sub SetSpendeActiv()
'Werden hier Controls oder die Form angesprochen, _
wird -> Sofort <- die Prozedur Form_Load von hier _
aufgerufen und das ist nicht gewünscht.
mbSpendeActiv = True
End Sub

Private Sub lblAbout_Click(Index As Integer)

If Index = 4 Then 'Spende
    'Call ExecuteDoc(Me.hWnd, gsBOMUrlHP & "/hp/h10.html")
ElseIf Index = 5 Then 'Homepage
    Call ExecuteDoc(Me.hWnd, gsBOMUrlHP)
End If

End Sub

Private Sub tmrAbout_Timer()
'...
If mbSpendeActiv Then
    With lblAbout(4) 'Spendenaktion
        If .ForeColor = vbRed Then
            .ForeColor = vbBlack
        Else
            .ForeColor = vbRed
        End If
    End With
End If
'...
If mbSplashMode Then
    miTimerCount = miTimerCount + 1
    If miMAXTIMERCOUNT >= miTimerCount Then
        btnOk.Caption = CStr(miMAXTIMERCOUNT - miTimerCount)
    Else
        mbSplashMode = False
        mbSpendeActiv = False
        Unload Me
    End If
End If
'...
End Sub


