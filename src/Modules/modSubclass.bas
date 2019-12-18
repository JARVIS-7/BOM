Attribute VB_Name = "modSubclass"
Option Explicit

Public Sub Subclass(lHwnd As Long)
    
    If GetProp(lHwnd, "oldWndProc") = 0 Then
        'Save the old window proc
        Call SetProp(lHwnd, "oldWndProc", GetWindowLong(lHwnd, GWL_WNDPROC))
        'Subclass
        Call SetWindowLong(lHwnd, GWL_WNDPROC, AddressOf NewWndProc)
    End If
    
End Sub
 
Public Sub UnSubclass(lHwnd As Long)

    Dim lOldProc As Long
    
    lOldProc = GetProp(lHwnd, "oldWndProc")
    If lOldProc <> 0 Then
        'Unsubclass
        Call SetWindowLong(lHwnd, GWL_WNDPROC, lOldProc)
        'Clean up properties
        Call RemoveProp(lHwnd, "oldProc")
    End If
    
End Sub

Public Function NewWndProc(ByVal lHwnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim tCDS As COPYDATASTRUCT
    Dim b() As Byte
    Dim lOldProc As Long
    Dim sCommand As String
    
    Select Case lMsg
        Case WM_COPYDATA
            'Copy for processing:
            Call CopyMemory(tCDS, ByVal lParam, Len(tCDS))
            If (tCDS.cbData > 0) Then
                ReDim b(0 To tCDS.cbData - 1) As Byte
                Call CopyMemory(b(0), ByVal tCDS.lpData, tCDS.cbData)
                sCommand = StrConv(b, vbUnicode)
                
                'We've got the info, now do it:
                'Debug.Print sCommand
                NewWndProc = ParseCommand(sCommand)
                Exit Function
            'Else
                'no data.  This is only sent by the main
                'module if it detects this window is hidden.
                'since this can't occur in this project,
                'this won't occur.  However, in a project
                'where your main window can be hidden, you
                'would make your window visible and activate
                'it here.
            End If
        Case WM_POWERBROADCAST
            If wParam = PBT_APMSUSPEND Then
                gdatFallAsleepDate = Now
                gbWarSchonWach = False
                gbSuspendNachAuktionAktiv = False
                giSuspendState = 1
            End If
            
            If wParam = PBT_APMRESUMEAUTOMATIC Or wParam = PBT_APMRESUMESUSPEND Then
                gbSuspendNachAuktionAktiv = False
                giSuspendState = 2
            End If
        Case Else
    End Select
    
    '''''''''''''''''''''''''''''''''''
    
    lOldProc = GetProp(lHwnd, "oldWndProc")
    If lOldProc <> 0 Then NewWndProc = CallWindowProc(lOldProc, lHwnd, lMsg, wParam, lParam)

End Function


