Attribute VB_Name = "modXPStyles"
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

'Um XP Styles in den eigenen Anwendungen benutzen zu können sind eine Manifest-Datei
'und ein Aufruf im Programm nötig.
'Folgende Manifest-Datei kann als Vorlage gelten:
'Zu beachten ist, dasz die Grösze der Manifest-Datei in Bytes durch 4 teilbar sein musz,
'zur Not einfach mit Leerzeichen auffüllen.

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
'<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
'    <assemblyIdentity
'        version="1.0.0.0"
'        processorArchitecture="X86"
'        name="CompanyName.ProductName.YourAppName"
'        type="win32" />
'    <description>Your application description here</description>
'    <dependency>
'        <dependentAssembly>
'            <assemblyIdentity
'                type="win32"
'                name="Microsoft.Windows.Common-Controls"
'                version="6.0.0.0"
'                processorArchitecture="X86"
'                publicKeyToken="6595b64144ccf1df"
'                language="*" />
'        </dependentAssembly>
'    </dependency>
'</assembly>


'---------------------------------------------------------------------------
' Sicherstellen, dasz unter XP die Common Controls Version 6 benutzt werden:
' IntitCommonControlsVB z.B. in der Intitialisierungsmethode von Form1 aufrufen:
'
' Private Sub Form_Initialize()
'    InitCommonControlsVB
' End Sub

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

'---------------------------------------------------------------

'XP-Styles für ein Control einschalten:
Declare Function ActivateWindowTheme Lib "uxtheme" _
    Alias "SetWindowTheme" ( _
    ByVal hWnd As Long, _
    Optional ByVal pszSubAppName As Long = 0, _
    Optional ByVal pszSubIdList As Long = 0) As Long

'XP-Style für ein Control ausschalten:
Declare Function DeactivateWindowTheme Lib "uxtheme" _
         Alias "SetWindowTheme" ( _
     ByVal hWnd As Long, _
     Optional ByRef pszSubAppName As String = " ", _
     Optional ByRef pszSubIdList As String = " ") As Long
''Beispiel XP-Style ein/ausschalten:
'' Deaktivieren:
'DeactivateWindowTheme Command1.hWnd
'' Aktivieren:
'ActivateWindowTheme Command1.hWnd

'-----------------------------------------------------------------
     
'Feststellen, ob die Anwendung, von der die API aufgerufen wird,
'Visual Styles verwendet. Das kann natürlich nur sein, wenn eine
'Manifest-Datei vorhanden ist und global ein Theme mit XP-Styles
'aktiviert ist.
Declare Function IsAppThemed Lib "UxTheme.dll" () As Boolean

''Anwendung:
'If IsAppThemed Then
'    MsgBox "Anwendung mit Styles!", vbInformation
'Else
'    MsgBox "Anwendung ohne Styles!", vbInformation
'End If

'------------------------------------------------------------------

'systemweit (windowsweit) XP-Styles an oder abschalten
Declare Function EnableTheming Lib "UxTheme.dll" (ByVal b As Boolean) As Long
''Anwendung:
'EnableTheming False

'--------------------------------------------------------------------

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

''XP-Style für eine ganze Form an/ausschalten
'Sub ShowXPStyles(Form As Form, Optional Modal, Optional OwnerForm As Form)
'On Error Resume Next
'Dim contr As Control
'
'For Each contr In Form.Controls
'    If XPStyle Then
'        ActivateWindowTheme contr.hWnd
'    Else
'        DeactivateWindowTheme contr.hWnd
'    End If
'Next contr
'
''Neuzeichnen
''LockWindowUpdate Form.hWnd
' Form.Hide
' Form.Show Modal, OwnerForm
''LockWindowUpdate 0
'End Sub

'Prüfen, ob eine Form Modal geladen ist oder nicht
Public Function IsFormModal(Form As Form) As Boolean
  Dim nTestForm As Form
  
  If Forms.count = 1 Then
    Exit Function
  End If
  For Each nTestForm In Forms
    If Not (nTestForm Is Form) Then
      If nTestForm.Enabled Then
        Exit Function
      End If
    End If
  Next
  IsFormModal = True
End Function


