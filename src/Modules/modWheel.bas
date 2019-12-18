Attribute VB_Name = "modWheel"
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
' $author: internet$
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
'
'MausWheel - Zugriff.. gefunden im Netz
'
Public Const GWL_WNDPROC As Long = -4&

Private Const WM_MOUSEWHEEL As Long = &H20A
'Private Const WM_MOUSELAST = &H20A
'Private Const WHEEL_DELTA = 120 '/* Value for rolling one detent */
 
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
'Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
 Public Function HIWORD(lLongIn As Long) As Integer
    '
    ' Mask off low word then do integer divide to
    ' shift right by 16.
    '
    HIWORD = (lLongIn And &HFFFF0000) \ &H10000
 End Function
 Public Function LOWORD(lLongIn As Long) As Integer
    '
    ' Low word retrieved by masking off high word.
    ' If low word is too large, twiddle sign bit.
    '
    If (lLongIn And &HFFFF&) > &H7FFF Then
       LOWORD = (lLongIn And &HFFFF&) - &H10000
    Else
       LOWORD = lLongIn And &HFFFF&
    End If
 End Function
 Private Function MWheelProc(ByVal lHwnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim lOldProc As Long
    Dim lCtlWnd As Long
    Dim lCtlPtr As Long
    Dim oIntObj As Object 'Intermediate object in between
    Dim oMWObject As ctlMWheel 'pointer and mousewheel control
    
    lCtlWnd = GetProp(lHwnd, "WheelWnd")
    lCtlPtr = GetProp(lCtlWnd, "WheelPtr")
    lOldProc = GetProp(lCtlWnd, "OldWheelProc")
    If lMsg = WM_MOUSEWHEEL Then
        Call CopyMemory(oIntObj, lCtlPtr, 4&)
        Set oMWObject = oIntObj
        Call oMWObject.WndProc(lHwnd, lMsg, wParam, lParam)
        Set oMWObject = Nothing
        Call CopyMemory(oIntObj, 0&, 4&)
    Else
        MWheelProc = CallWindowProc(lOldProc, lHwnd, lMsg, wParam, lParam)
    End If
    
 End Function
 Public Sub SubclassMW(oMWCtl As ctlMWheel, lParentWnd As Long)
    
    If GetProp(oMWCtl.hWnd, "OldWheelProc") = 0 Then
        'Save the old window proc of the control's parent
        Call SetProp(oMWCtl.hWnd, "OldWheelProc", GetWindowLong(lParentWnd, GWL_WNDPROC))
        'Object pointer to the control
        Call SetProp(oMWCtl.hWnd, "WheelPtr", ObjPtr(oMWCtl))
        'Save control's hWnd in its parent data
        Call SetProp(lParentWnd, "WheelWnd", oMWCtl.hWnd)
        'Subclass the control's parent
        Call SetWindowLong(lParentWnd, GWL_WNDPROC, AddressOf MWheelProc)
    End If
    
 End Sub
 Public Sub UnSubclassMW(oMWCtl As ctlMWheel, lParentWnd As Long)
    
    Dim lOldProc As Long
    
    lOldProc = GetProp(oMWCtl.hWnd, "OldWheelProc")
    If lOldProc <> 0 Then
        'Unsubclass control's parent
        Call SetWindowLong(lParentWnd, GWL_WNDPROC, lOldProc)
        'Clean up properties
        Call RemoveProp(lParentWnd, "WheelWnd")
        Call RemoveProp(oMWCtl.hWnd, "WheelPtr")
        Call RemoveProp(oMWCtl.hWnd, "OldWheelProc")
    End If
    
 End Sub
