Attribute VB_Name = "modflashWex"
'Const für Window/Taskbar blinken
'Stop flashing
Private Const FLASHW_STOP As Long = 0
'Flash the window caption
Private Const FLASHW_CAPTION As Long = &H1
'Flash the taskbar button
Private Const FLASHW_TRAY As Long = &H2
'Flash both the window caption and taskbar button
Private Const FLASHW_ALL As Long = (FLASHW_CAPTION Or FLASHW_TRAY)
'Flash continuously, until the FLASHW_STOP flag is set
Private Const FLASHW_TIMER As Long = &H4
'Flash continuously until the window comes to the foreground
Private Const FLASHW_TIMERNOFG As Long = &HC

Private FLASHW_FLAGS As Long

Private Type FLASHWINFO
   cbSize As Long
   hWnd As Long
   dwFlags As Long
   uCount As Long
   dwTimeout As Long
End Type

Private Declare Function FlashWindowEx Lib "user32" _
                        (pflashwininfo As FLASHWINFO) As Long
                        
Private Declare Function GetForegroundWindow Lib "user32" () As Long

                        
Private Sub FlashBeginUntilActive(ByVal lHwnd As Long, _
                                  ByVal lFrequency As Long, _
                                  ByVal lCount As Long)

   Dim fwi As FLASHWINFO

   If (lFrequency >= 0) Then
         With fwi
            .cbSize = Len(fwi)
            .hWnd = lHwnd
            .dwFlags = FLASHW_TIMERNOFG Or FLASHW_FLAGS
            .dwTimeout = lFrequency
            .uCount = lCount
         End With
         Call FlashWindowEx(fwi)
   End If
End Sub

Private Sub FlashBeginCount(ByVal lHwnd As Long, _
                            ByVal lFrequency As Long, _
                            ByVal lCount As Long)

   Dim fwi As FLASHWINFO
   
   If (lCount > 0) Then
      With fwi
         .cbSize = Len(fwi)
         .hWnd = lHwnd
         .dwFlags = FLASHW_FLAGS
         .dwTimeout = lFrequency
         .uCount = lCount
      End With
      Call FlashWindowEx(fwi)
   End If
End Sub
                        
Public Sub FlashIt(ByVal frm As Form, _
                     Optional ByVal iFlashCount As Integer = 15, _
                     Optional ByVal iFrequenz As Integer = 0)

gbWarningflag = True
FLASHW_FLAGS = 0&
If GetForegroundWindow <> frm.hWnd Then
    FLASHW_FLAGS = FLASHW_FLAGS Or FLASHW_TRAY
    Call FlashBeginUntilActive(frm.hWnd, CLng(iFrequenz), -1)
Else
    FLASHW_FLAGS = FLASHW_CAPTION
    Call FlashBeginCount(frm.hWnd, CLng(iFrequenz), CLng(iFlashCount))
End If

End Sub
