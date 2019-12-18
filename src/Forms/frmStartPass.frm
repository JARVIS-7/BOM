VERSION 5.00
Begin VB.Form frmStartPass 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " Password"
   ClientHeight    =   1830
   ClientLeft      =   7515
   ClientTop       =   6900
   ClientWidth     =   4185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmStartPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmStartPass"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btnPass 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   1545
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtPass 
      CausesValidation=   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Image imgPass 
      Height          =   480
      Left            =   120
      Top             =   645
      Width           =   480
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte das Passwort für den User xyz eingeben"
      Height          =   540
      Left            =   720
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "frmStartPass"
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
Private Sub btnPass_Click()
    
    Call CheckPasswort(txtPass.Text)
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    'MD-Marker 20090323 , Workaround bis frmStartPass von _
    Sub_Main geladen wird.
    
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0: gbExplicitEnd = True
        End 'MD-Marker
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0: Call btnPass_Click
    End If
    
End Sub

Private Sub Form_Load()
    
    Set imgPass.Picture = Me.Icon
    
    Me.Caption = gsarrLangTxt(142)
    lblPass.Caption = gsarrLangTxt(350) & " " & vbCrLf & gsUser & " " & gsarrLangTxt(351)
    btnPass.Caption = gsarrLangTxt(352)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set imgPass.Picture = Nothing
    Unload Me: Set frmStartPass = Nothing
    
End Sub

Private Sub CheckPasswort(sPassWort As String)
    
    sPassWort = Trim$(sPassWort)
    
    gbExplicitEnd = CBool((Len(sPassWort) <> Len(gsPass)) Or (sPassWort <> gsPass))
    
    If gbExplicitEnd Then
        Call MsgBox(gsarrLangTxt(353), vbOKOnly Or vbCritical, gsarrLangTxt(142))
        End 'MD-Marker 20090323
    Else
        Unload Me
    End If
        
End Sub
