VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug"
   ClientHeight    =   3615
   ClientLeft      =   6705
   ClientTop       =   5100
   ClientWidth     =   6375
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "frmDebug"
   LockControls    =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   6375
   Visible         =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
    
    With Me
        glDebugWindowLeft = .Left
        glDebugWindowTop = .Top
        glDebugWindowWidth = .Width
        glDebugWindowHeight = .Height
    End With
    
End Sub

Private Sub Form_Load()

    'MD-Marker 20090323 , bei dem ersten Aufruf wurde die _
    Settings.ini noch nicht geladen, daher sind alle Variablen=0
    
    If glDebugWindowWidth > 0 Then
        Call Me.Move(glDebugWindowLeft, glDebugWindowTop, glDebugWindowWidth, glDebugWindowHeight)
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = vbFormControlMenu Then
        
        Cancel = 1
        Me.Hide
        
    Else
        
        With Me
            glDebugWindowLeft = .Left
            glDebugWindowTop = .Top
            glDebugWindowWidth = .Width
            glDebugWindowHeight = .Height
        End With
        
        Unload Me: Set frmDebug = Nothing
        
    End If
    
End Sub

Private Sub Form_Resize()
    
    With Me
        If .WindowState <> vbMinimized Then
            If .ScaleWidth > 0 And Me.ScaleWidth > (2 * List1.Left) Then
                If .ScaleHeight > 0 And Me.ScaleHeight > (2 * List1.Top) Then
                    With List1
                        Call .Move(.Left, .Top, Me.ScaleWidth - (2 * .Left), Me.ScaleHeight - (2 * .Top))
                    End With
                End If
            End If
        End If
    End With

End Sub

Public Sub SetSize()
        
    Form_Load
    
End Sub

