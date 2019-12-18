VERSION 5.00
Begin VB.Form frmSecurityToken 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Sicherheitsschluessel"
   ClientHeight    =   1950
   ClientLeft      =   7110
   ClientTop       =   6060
   ClientWidth     =   3870
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmSecurityToken.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer tmrToken 
      Enabled         =   0   'False
      Left            =   120
      Top             =   1320
   End
   Begin VB.TextBox txtToken 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton btnToken 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1388
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblToken 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte einen Sicherheitsschluessel für den User xyz eingeben:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgToken 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   760
      Width           =   480
   End
End
Attribute VB_Name = "frmSecurityToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msFrmRetValue As String
Private msSecurityToken As String
Private miTimeLeft As Integer
Private miMousePointer As Integer

Private Sub btnToken_Click()

    tmrToken.Enabled = False
    msFrmRetValue = txtToken.Text
    Me.Hide
    
End Sub

Private Sub Form_Load()
    
    Set imgToken.Picture = Me.Icon
    
    btnToken.Caption = gsarrLangTxt(352) & " [" & CStr(miTimeLeft) & "]"
    
    With tmrToken
        .Enabled = True
        .Interval = 1000
    End With
    
    lblToken.Caption = Replace(gsarrLangTxt(743), "%USER%", msSecurityToken)
    miMousePointer = Screen.MousePointer
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    With tmrToken
        .Interval = 0
        .Enabled = False
    End With
    
    If UnloadMode = vbFormControlMenu Then
        msFrmRetValue = txtToken.Text
        Me.Hide
    Else
        txtToken.Text = ""
        Set imgToken.Picture = Nothing
        Unload Me: Set frmSecurityToken = Nothing
    End If
    Screen.MousePointer = miMousePointer
    
End Sub

Private Sub tmrToken_Timer()

    miTimeLeft = miTimeLeft - 1
    btnToken.Caption = gsarrLangTxt(352) & " [" & CStr(miTimeLeft) & "]"
    If miTimeLeft <= 0 Then
        tmrToken.Enabled = False
        txtToken.Text = ""
        msFrmRetValue = txtToken.Text
        Me.Hide
    End If
    
End Sub

Public Property Get SecurityToken() As String
SecurityToken = msFrmRetValue
End Property

Public Property Let SecurityToken(ByVal sTextIn As String)
    
    miTimeLeft = 60
    msSecurityToken = sTextIn
    
End Property
