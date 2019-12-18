VERSION 5.00
Begin VB.UserControl ctlMWheel 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ctlMWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
' found in den Tiefen des Zwischennetzes ..
' thx to ?
'
 Private mlCapHwnd As Long
 Private mbSubclassed As Boolean
 
 Event WheelScroll(Shift As Integer, zDelta As Integer, _
    X As Single, Y As Single)
Private Sub UserControl_Resize()
    
    Call UserControl.Size(32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY)
    
End Sub
 Public Sub DisableWheel()
     
    If mbSubclassed Then
        If mlCapHwnd <> 0 Then
            Call modWheel.UnSubclassMW(Me, mlCapHwnd)
            mbSubclassed = False
        End If
    End If
    
 End Sub
 Public Sub EnableWheel()
 
    If mlCapHwnd <> 0 Then
        mbSubclassed = True
        Call modWheel.SubclassMW(Me, mlCapHwnd)
    End If
    
 End Sub
 Friend Property Get hWnd() As Long
    hWnd = UserControl.hWnd
 End Property
 Public Property Get hWndCapture() As Long
    hWndCapture = mlCapHwnd
 End Property
 Public Property Let hWndCapture(ByVal lValue As Long)
      mlCapHwnd = lValue
 End Property
 Friend Sub WndProc(ByVal lHwnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
      
      
      Dim iShift As Integer
      Dim iDelta As Integer
      Dim rPosX As Single, rPosY As Single
      
      iShift = LOWORD(wParam)
      iDelta = HIWORD(wParam)
      rPosX = CSng(LOWORD(lParam))
      rPosY = CSng(HIWORD(lParam))
      
      RaiseEvent WheelScroll(iShift, iDelta / 120, rPosX, rPosY)
      
 End Sub


