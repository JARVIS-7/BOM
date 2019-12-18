Attribute VB_Name = "modShellEx"
Option Explicit

Private Const PROCESS_TERMINATE As Long = &H1
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const STATUS_PENDING As Long = &H103&

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Function ShellWithWait(sPathName As String, eWindowStyle As VbAppWinStyle) As Boolean
    
    'MD-Marker , Function wird nicht aufgerufen
    
    'Dim hProcess As Long
    'Dim ProcessId As Long
    'Dim ExitCode As Long
    
    'ProcessId = Shell(PathName, WindowStyle)
    'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId)
    
    'Do
    '    Call GetExitCodeProcess(hProcess, ExitCode)
    '    DoEvents
    '    Sleep 10
    'Loop While ExitCode = STATUS_PENDING
    
    'Call CloseHandle(hProcess)
    
End Function

Public Function ShellTest(sPathName As String, eWindowStyle As VbAppWinStyle) As Boolean
    
    On Error GoTo ERR_EXIT
    Dim lProcessID As Long
    
    
    
    lProcessID = ShellStart(sPathName, eWindowStyle)
    If lProcessID <> 0 Then
        Call ShellStop(lProcessID)
        ShellTest = True
    End If
    
ERR_EXIT:
    
End Function

Public Function ShellStart(sPathName As String, eWindowStyle As VbAppWinStyle) As Long

  On Error GoTo ERR_EXIT
  
  Debug.Print sPathName
  
  ShellStart = Shell(sPathName, eWindowStyle)
  
  Exit Function
  
ERR_EXIT:
    Debug.Print Err.Description
End Function

Public Sub ShellStop(lProcessID As Long)
    
    Dim lProcess As Long
    
    If lProcessID <> 0 Then
        lProcess = OpenProcess(PROCESS_TERMINATE, 0&, lProcessID)
        If lProcess <> 0 Then
            Call TerminateProcess(lProcess, 0&)
            Call CloseHandle(lProcess)
        End If
    End If
    
End Sub

Public Function ShellStillRunning(lProcessID As Long) As Boolean
    
    Dim lProcess As Long
    Dim lExitCode As Long
    
    lProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0&, lProcessID)
    If lProcess <> 0 Then
        Call GetExitCodeProcess(lProcess, lExitCode)
        ShellStillRunning = CBool(lExitCode = STATUS_PENDING)
        Call CloseHandle(lProcess)
    End If

End Function


