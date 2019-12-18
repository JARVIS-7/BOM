Attribute VB_Name = "modTVMouseOver"
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

'Modul, um einem Treeview - Objekt folgende Eigenschaften
'verpassen zu können:
'- MouseOver-Effekt (Mauszeiger "Hand" u unterstrichener Text)
'- Hintergrundfarbe des Treeview-Objektes einstellbar
'- Textfarbe änderbar
'- Kontrollkästchen anzeigen
'------------------------------------------------------------------

'API Call for Sending the messages
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Constants for changing the treeview
Public Const GWL_STYLE = -16&
Public Const TVM_SETBKCOLOR = 4381&
Public Const TVM_GETBKCOLOR = 4383&
Public Const TVS_HASLINES = 2&
Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVS_CHECKBOXES = &H100
Public Const TVS_TRACKSELECT = &H200

Public Sub SetTreeViewAttrib(c As TreeView, ByVal Attrib As Long)
    Const GWL_STYLE As Long = -16
    Dim rStyle As Long
    rStyle = GetWindowLong(c.hWnd, GWL_STYLE)
    rStyle = rStyle Or Attrib
    Call SetWindowLong(c.hWnd, GWL_STYLE, rStyle)
End Sub


'**********************************************************************
'Anwendung:
'**********************************************************************
    
    'Hintergrund des Textviews ändern:
    'Send a message to the treeview telling it to
    'change the background It uses an RGB colour setting
    
' Call SendMessage(TreeView1.hWnd, TVM_SETBKCOLOR, 0, ByVal RGB(255, 204, 0))


    'Textfarbe ändern
    'Same as above except this one tells it to change
    'the text colour
    
' Call SendMessage(TreeView1.hWnd, TVM_SETTEXTCOLOR, 0, ByVal RGB(0, 127, 0))
    
    
    'Mouseover-Effekt hinzufügen
    'Add the track selection
    
' Call SetTreeViewAttrib(TreeView1, TVS_TRACKSELECT)


    'Checkboxes hinzufügen

'  Call SetTreeViewAttrib(TreeView1, TVS_CHECKBOXES)

