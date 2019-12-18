VERSION 5.00
Begin VB.Form frmDummy 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   6540
   ClientTop       =   4425
   ClientWidth     =   4680
   Icon            =   "frmDummy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   4
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   3
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   2
      Left            =   1080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   1
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   720
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim i As Integer
    
    For i = Picture1.LBound To Picture1.UBound
        Set Picture1.Item(i).Picture = LoadPicture("")
    Next 'i
    
    Unload Me: Set frmDummy = Nothing
    
End Sub


Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call modSubclass.UnSubclass(Me.hWnd)
    Unload Me
    End 'MD-Marker
    
End Sub

Public Sub TimerJetzt()
    Timer1_Timer
End Sub
