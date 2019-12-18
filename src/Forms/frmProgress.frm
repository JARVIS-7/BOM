VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   390
   ClientLeft      =   6405
   ClientTop       =   6780
   ClientWidth     =   4455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "frmProgress"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Sanduhr
   ScaleHeight     =   390
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrStartTime As Single
Private mrLastRefresh As Single

Public Sub InitProgress(ByVal lMin As Long, ByVal lMax As Long)
    
    Dim lTmp As Long
    
    If lMin > lMax Then
        lTmp = lMin
        lMin = lMax
        lMax = lTmp
    End If
    
'19.04.09 - schurik
' Fehler beim Beenden des Programmes abfangen, wenn beide Werte gleich 1
    If lMin = lMax Then
        If lMax >= 32767 Then
        'lMax ist ausgemaxxt. lMin um 1 kleiner setzen
            lMin = lMax - 1
        Else
        ' lMax um 1 erhöhen
            lMax = lMax + 1
        End If
    End If
    
    With ProgressBar
        .Min = CInt(lMin)
        .Max = CInt(lMax)
    End With
    
    With frmHaupt
        Call Me.Move((.Width - Me.Width) \ 2 + .Left, (.Height - Me.Height) \ 2 + .Top, Me.Width, Me.Height)
    End With
    mrStartTime = Timer
    
End Sub

Public Sub Step(Optional ByVal iStepWidth As Integer = 1)
    
    Call Progress(ProgressBar.Value + iStepWidth)
    
End Sub

Private Sub Progress(ByVal lValue As Long)
        
    On Error Resume Next
    Dim rTime As Single
    
    With ProgressBar
        If lValue = 10 Then
            rTime = Timer
            If (((rTime - mrStartTime) * (.Max - .Min)) \ 10) > 1 Then
                If frmHaupt.WindowState <> vbMinimized Then 'Me.Show
                    If frmHaupt.Visible Then Me.Show vbModeless, frmHaupt
                End If
            End If
        End If
        
        .Value = CInt(lValue)
        rTime = Timer
        If (mrLastRefresh + 0.5) < rTime Then
            mrLastRefresh = Timer
            .Refresh
        End If
    End With 'ProgressBar
    
End Sub

Public Sub TerminateProgress()
    
    Me.Hide
    Unload Me
    DoEvents
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Unload Me: Set frmProgress = Nothing
    
End Sub

