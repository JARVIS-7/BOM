VERSION 5.00
Begin VB.Form frmChooseUser 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Account wählen..."
   ClientHeight    =   1470
   ClientLeft      =   7020
   ClientTop       =   6240
   ClientWidth     =   4110
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmChooseUser.frx":0000
   LinkTopic       =   "frmChooseUser"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnAction 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnAction 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboChooseUser 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Image imgUser 
      Height          =   480
      Left            =   120
      Top             =   280
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmChooseUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAction_Click(Index As Integer)
    
    Select Case Index
        Case 0: Unload Me
        Case Else
            
            With gtarrArtikelArray(giArtChoose)
                .UserAccount = gtarrUserArray(cboChooseUser.ListIndex + 1).UaUser
                .LastChangedId = GetChangeID()
            End With
            
            Unload Me
            
    End Select
    
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    
    With Me
        Set imgUser.Picture = Me.Icon
        .Caption = gsarrLangTxt(728)
    End With
    
    Label1.Caption = gsarrLangTxt(360)
    btnAction(0).Caption = gsarrLangTxt(719) 'Abbrechen
    btnAction(1).Caption = gsarrLangTxt(219) 'Ok
    '...
    With cboChooseUser
        
        For i = LBound(gtarrUserArray()) To UBound(gtarrUserArray())
            If Len(gtarrUserArray(i).UaUser) Then 'MD-Marker , workaround weil Userarray-Index mit '0' anfängt aber erst bei 'Index=1' genutzt wird
                Call .AddItem(gtarrUserArray(i).UaUser)
            End If
        Next 'i
        
        If Len(gtarrArtikelArray(giArtChoose).UserAccount) Then
            .ListIndex = UsrAccToIndex(gtarrArtikelArray(giArtChoose).UserAccount) - 1
        Else
            .ListIndex = giDefaultUser - 1
        End If
        
    End With 'cboChooseUser
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set imgUser.Picture = Nothing
    
    Unload Me: Set frmChooseUser = Nothing
    
End Sub

