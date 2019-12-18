VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSLWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlProcessID As Long

Public Function StartSSLWrapper(ByVal sCmd As String, ByVal bHide As Boolean, ByVal sServer As String, ByVal lPort As Long, ByVal lStartupDelay As Long) As Boolean
    
    sCmd = Replace(sCmd, "%SERVER%", sServer)
    sCmd = Replace(sCmd, "%PORT%", lPort)
    
    If bHide Then
        mlProcessID = ShellStart(sCmd, vbHide)
    Else
        mlProcessID = ShellStart(sCmd, vbNormalNoFocus)
    End If
    
    StartSSLWrapper = CBool(mlProcessID <> 0)
    Call Sleep(lStartupDelay)
    
End Function

Private Sub Class_Terminate()
    
    Call ShellStop(mlProcessID)
    
End Sub
