Attribute VB_Name = "modCC6Effects"
'//////////////////////////////////////////////////////////////////////////////
'//   Biet-O-Matic (Bid-O-Matic)                                             //
'//                                                                          //
'//   This program is free software; you can redistribute it and/or modify   //
'//   it under the terms of the GNU General Public License as published by   //
'//   the Free Software Foundation; either version 2 of the License, or      //
'//   (at your option) any later version.                                    //
'//                                                                          //
'//   This program is distributed in the hope that it will be useful,        //
'//   but WITHOUT ANY WARRANTY; without even the implied warranty of         //
'//   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the          //
'//   GNU General Public License for more details.                           //
'//                                                                          //
'//   You should have received a copy of the GNU General Public License      //
'//   along with this program; if not, write to the Free Software            //
'//   Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.              //
'//                                                                          //
'//   Main language: german                                                  //
'//   Compiled under VB6 SP5 german                                          //
'//   Contact: visit http://bom.sourceforge.net                              //
'//////////////////////////////////////////////////////////////////////////////
Option Explicit

'__________________________________________________________________
'
'--- Dieses Modul dient dazu, bestimmte Funktionalitäten, die den
'--- Common Controls 5 (SP2) fehlen, per Api nachzurüsten.
'--- zB. CC 5 - Toolbar Effekte wie in einer Coolbar generieren
'--- (sich erhebende Buttons beim Überfahren mit der Maus und
'--- senkrechte Trennstriche als Seperatoren),
'--- CC5 ListViews mit Checkboxes usw.
'--- Dieser Code ist für Common Control 6 Controls überflüssig,
'--- da diese dergleichen Eigenschaften bereits besitzen.
'___________________________________________________________________

Public Const WM_USER = &H400

'Toolbar Constants
Public Const TTM_ACTIVATE = (WM_USER + 1)
Public Const TBM_GETTOOLTIPS = (WM_USER + 30)
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800

'ListView Constants:
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETITEMSTATE As Long = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)

Public Const LVS_EX_GRIDLINES As Long = &H1
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVS_EX_FULLROWSELECT As Long = &H20 'applies to report mode only

Public Const LVIF_STATE As Long = &H8
Public Const LVIS_STATEIMAGEMASK As Long = &HF000
 
Public Const MAX_PATH As Long = 260

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, _
        ByVal hwndChildAfter As Long, ByVal lpszClass As String, _
        ByVal lpszWindow As String) As Long

Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

'Windows-Anzeige nicht erneuern (ausschalten, einschalten)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long)

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'die VB5-Toolbar per API zu einer VB6-Toolbar machen:
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'das Modul schaltet abwechselnd zwischen den Anzeige-Stadien (flach oder erhoben)
'hin und her
Public Sub SwitchToolbarStyle(Toolbar As ComctlLib.Toolbar)

   Dim style As Long
   Dim hToolbar As Long
   
  'get the handle of the toolbar
   hToolbar = FindWindowEx(Toolbar.hWnd, 0&, "ToolbarWindow32", vbNullString)
   
  'retrieve the toolbar styles
   style = SendMessage(hToolbar, TB_GETSTYLE, 0&, ByVal 0&)
   
  'Set the new style flag
   If style And TBSTYLE_FLAT Then
         style = style Xor TBSTYLE_FLAT
   Else: style = style Or TBSTYLE_FLAT
   End If
   
  'apply the new style to the toolbar
   Call SendMessage(hToolbar, TB_SETSTYLE, 0, ByVal style)
   Toolbar.Refresh
   
End Sub

'schaltet einfach nur den flachen Style ein
Public Sub ToolbarFlat(Toolbar As ComctlLib.Toolbar)
   Dim style As Long
   Dim hToolbar As Long
   
  'get the handle of the toolbar
   hToolbar = FindWindowEx(Toolbar.hWnd, 0&, "ToolbarWindow32", vbNullString)
   
  'retrieve the toolbar styles
   style = SendMessage(hToolbar, TB_GETSTYLE, 0&, ByVal 0&)
   
  'Set the flat style flag
   style = style Or TBSTYLE_FLAT
   
  'apply the new style to the toolbar
   Call SendMessage(hToolbar, TB_SETSTYLE, 0, ByVal style)
   Toolbar.Refresh
End Sub

'-----------------------------------------
'Anwendungsbeispiele:
'
'- einfach:
'   ToolbarFlat Toolbar1

'- hin- und herschalten:
'   zB. einem Command Button im Clickereignis folgenden Code hinzufügen:
        
'      SwitchToolbarStyle Toolbar1
'
'------------------------------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Dem CC5 Slider Tooltips verpassen:
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function Slider_ActivateToolTips(hwndSlider As Long, _
                                        bEnabled As Boolean) As Long
   
   Dim hToolTips As Long
   
  'das Handle zum Tooltip Control, welches
  'zum Slider gehört, ermitteln
   hToolTips = SendMessage(hwndSlider, _
                           TBM_GETTOOLTIPS, _
                           ByVal 0&, _
                           ByVal 0&)
   
   If hToolTips <> 0 Then
      'Tooltip Control de-/aktivieren
       Slider_ActivateToolTips = SendMessage(hToolTips, _
                                             TTM_ACTIVATE, _
                                             ByVal Abs(bEnabled), _
                                             ByVal 0&)
   End If
End Function

'--------------------------
'Anwendungsbeispiel:
'Slider_ActivateToolTips Slider1.hWnd, True



'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' CC5 ListView
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'CC5 Listview Checkboxes anzeigen/entfernen
Public Sub ShowCheckBoxes(LView As ListView, CheckboxesVisible As Boolean)
     Call SendMessage(LView.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, _
                      LVS_EX_CHECKBOXES, ByVal CheckboxesVisible)
End Sub

'Eine ListView-Checkbox per Code setzen (an/abschalten)
Public Sub SetCheck(LVhwnd As Long, ByVal LItemIndex As Long, bState As Boolean)

    Dim LV As LV_ITEM

    With LV
      .mask = LVIF_STATE
      .state = IIf(bState, &H2000, &H1000)
      .stateMask = LVIS_STATEIMAGEMASK
    End With

    Call SendMessage(LVhwnd, LVM_SETITEMSTATE, LItemIndex, LV)

End Sub

'Ermitteln, ob eine Checkbox angehakt ist oder nicht
Public Function IsChecked(LView As ListView, ByVal LVItemIndex As Long) As Boolean
Dim l As Long

l = SendMessage(LView.hWnd, LVM_GETITEMSTATE, LVItemIndex, ByVal LVIS_STATEIMAGEMASK)

     'when an item is checked, the LVM_GETITEMSTATE call
     'returns 8192 (&H2000&).
      If (l And &H2000&) Then
           IsChecked = True
      Else
           IsChecked = False
      End If
End Function
