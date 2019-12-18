VERSION 5.00
Begin VB.Form frmKommentar 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Kommentar eingeben"
   ClientHeight    =   2745
   ClientLeft      =   7305
   ClientTop       =   6660
   ClientWidth     =   5190
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmKommentar.frx":0000
   LinkTopic       =   "frmKommentar"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   -10000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1000
   End
   Begin VB.TextBox txtNotiz 
      CausesValidation=   0   'False
      Height          =   1695
      Left            =   630
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      ToolTipText     =   " max 200 Zeichen"
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton btnAction 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgNotiz 
      Height          =   480
      Left            =   120
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lblNotiz 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Kommentar zu dem Artikel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   3015
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmKommentar"
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
Private miArtikelID As Integer

Private Sub btnAction_Click()
    
    With gtarrArtikelArray(miArtikelID)
    'Zeilenumbrüche werden nun mit abgespeichert
        .Kommentar = txtNotiz.Text
        .LastChangedId = GetChangeID()
    End With
    
    Unload Me
    
End Sub

Private Sub btnCancel_Click()

  Unload Me

End Sub

Private Sub Form_Load()
    
    With Me
        Set .Icon = MyLoadResPicture(105, 16)
        Call SendMessage(.hWnd, WM_SETICON, 0, .Icon)
        Set imgNotiz.Picture = .Icon
        .Caption = gsarrLangTxt(221)
    End With
    
    txtNotiz.Text = gtarrArtikelArray(miArtikelID).Kommentar
    lblNotiz.Caption = gsarrLangTxt(221) & gtarrArtikelArray(miArtikelID).Artikel & gsarrLangTxt(222)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set imgNotiz.Picture = Nothing
    
    Unload Me: Set frmKommentar = Nothing
    
End Sub

Public Property Let SetArtikelID(ByVal iArtikelID As Integer)
miArtikelID = iArtikelID
End Property

Private Sub txtNotiz_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyA And Shift = 2 Then
    With txtNotiz
      .SelStart = 0
      .SelLength = Len(txtNotiz.Text)
    End With
  End If

End Sub
