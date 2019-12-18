VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopupMenu"
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
' $author: susi/hjs$
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
'*** Baut dynamisch ein Popup-Menü *********************
'*******************************************************
Option Explicit
'
Private mlHwnd As Long

Private Type POINT
    X As Long
    Y As Long
End Type
'
Private Const MF_ENABLED As Long = &H0&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_STRING As Long = &H0&
Private Const MF_GRAYED As Long = &H1
Private Const MF_CHECKED As Long = &H8&
Private Const TPM_RIGHTBUTTON As Long = &H2&
Private Const TPM_LEFTALIGN As Long = &H0&
Private Const TPM_NONOTIFY As Long = &H80&
Private Const TPM_RETURNCMD As Long = &H100&
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal sCaption As String) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, nIgnored As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd&) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
'
Public Function Popup(lDefaultMenuItem As Long, ParamArray param()) As Long
    
    Dim iMenu As Long
    Dim hMenu As Long
    Dim nMenus As Long
    Dim p As POINT
    Dim lMfAddon As Long
    Dim sMenuCaption As String
    
    'get the current cursor pos in screen coordinates
    Call GetCursorPos(p)
    
    'create an empty popup menu
    hMenu = CreatePopupMenu()
    
    'determine # of strings in paramarray, element 0 contains number of default-menu
    nMenus = UBound(param()) + 1
    
    'put each string in the menu
    For iMenu = 1 To nMenus
        'the AppendMenu function has been superseeded by the InsertMenuItem
        'function, but it is a bit easier to use.
        If Trim$(CStr(param(iMenu - 1))) = "-" Then
            'if the parameter is a single dash, a separator is drawn
            Call AppendMenu(hMenu, MF_SEPARATOR, iMenu, "")
        Else
            sMenuCaption = CStr(param(iMenu - 1))
            lMfAddon = 0
            If sMenuCaption Like "CHECKED_*" Then
                lMfAddon = lMfAddon Or MF_CHECKED
                sMenuCaption = Mid(sMenuCaption, 9)
            End If
            
            If sMenuCaption Like "GRAYED_*" Then
                lMfAddon = lMfAddon Or MF_GRAYED
                sMenuCaption = Mid(sMenuCaption, 8)
            End If
            Call AppendMenu(hMenu, MF_STRING Or MF_ENABLED Or lMfAddon, iMenu, sMenuCaption)
        End If
    Next iMenu
    
    'mark the default-menu (if given)
    If lDefaultMenuItem > 0 Then
        Call SetMenuDefaultItem(hMenu, lDefaultMenuItem - 1, 1)
    End If
    
    'show the menu at the current cursor location;
    'the flags make the menu aligned to the right (!); enable the right button to select
    'an item; prohibit the menu from sending messages and make it return the index of
    'the selected item.
    'the TrackPopupMenu function returns when the user selected a menu item or cancelled
    'the window handle used here may be any window handle from your application
    'the return value is the (1-based) index of the menu item or 0 in case of cancelling
    If mlHwnd <> -1 Then
        'Call SetForegroundWindow(mlHwnd)
        iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON Or TPM_LEFTALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD, p.X, p.Y, 0, mlHwnd, 0)
    Else
        iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON Or TPM_LEFTALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD, p.X, p.Y, 0, GetForegroundWindow(), 0)
    End If
    
    'release and destroy the menu (for sanity)
    Call DestroyMenu(hMenu)
    
    'return the selected menu item's index
    Popup = iMenu
    
End Function

Private Sub Class_Initialize()
    mlHwnd = -1
End Sub

' hwnd, if -1(default) then GetForegroundWindow() is called for a hwnd
' otherwise, the provided hwnd is used

Public Property Let hWnd(lNewHwnd As Long)
mlHwnd = lNewHwnd
End Property

Public Property Get hWnd() As Long
    hWnd = mlHwnd
End Property

