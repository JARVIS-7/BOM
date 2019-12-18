VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmChooseArtList 
   Caption         =   " Wählen Sie die Artikel die Sie hinzufügen wollen..."
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   ControlBox      =   0   'False
   Icon            =   "frmChooseArtList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10455
   StartUpPosition =   1  'Fenstermitte
   Begin ComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox f_BidGroupValue 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Bietgruppe: (NR;Anzahl)"
      Top             =   210
      Width           =   1215
   End
   Begin VB.TextBox f_GrPrice 
      Enabled         =   0   'False
      Height          =   295
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "Gebotspreis für die Bietgruppe"
      Top             =   570
      Width           =   1215
   End
   Begin VB.CheckBox f_GroupCheck 
      Caption         =   " Auswahl als Bietgruppe hinzufügen:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton CB_unmarkall 
      Caption         =   "Alle Markierungen entfernen"
      Height          =   255
      Left            =   5240
      TabIndex        =   4
      Top             =   5640
      Width           =   5100
   End
   Begin VB.CommandButton CB_markall 
      Caption         =   "Alle Artikel markieren"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   5100
   End
   Begin VB.CommandButton CB_addmarked 
      Caption         =   "Artikel hinzufügen"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CB_cancel 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Gebotspreis der Bietgruppe:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   415
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmChooseArtList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim AnzChecked As Integer

Private Sub f_BidGroupValue_Validate(Cancel As Boolean)
If Len(Trim(f_BidGroupValue)) > 0 Then
  If InStr(1, Trim(f_BidGroupValue), ";", vbTextCompare) < 1 Then Cancel = True
End If
End Sub

Private Sub f_GroupCheck_Click()

f_BidGroupValue.Enabled = f_GroupCheck > 0
f_GrPrice.Enabled = f_GroupCheck > 0
Label1.Enabled = f_GroupCheck > 0

End Sub

Private Sub f_GrPrice_Validate(Cancel As Boolean)
If Len(Trim(f_GrPrice)) > 0 Then
  If Not IsNumeric(Trim(f_GrPrice)) Then Cancel = True
End If
End Sub

Private Sub Form_Load()
Dim X As Integer

On Error Resume Next

'CC5 ListView mit Checkboxes? - naaaaaaaa gut ;-)
ShowCheckBoxes ListView1, True

Me.CB_markall = LangTxt(755)
Me.CB_unmarkall = LangTxt(756)
Me.CB_addmarked.Caption = LangTxt(757)
Me.Caption = LangTxt(758)
Me.CB_cancel.Caption = LangTxt(719)

With ListView1
  
  .ColumnHeaders.Add , , LangTxt(750), (.Width / 100) * 14
  .ColumnHeaders.Add , , LangTxt(751), (.Width / 100) * 38
  .ColumnHeaders.Add , , LangTxt(752), (.Width / 100) * 20
  .ColumnHeaders.Add , , LangTxt(753), (.Width / 100) * 10
  .ColumnHeaders.Add , , LangTxt(754), (.Width / 100) * 14
  
  .ColumnHeaders.Item(4).Alignment = lvwColumnCenter
  .ColumnHeaders.Item(5).Alignment = lvwColumnRight
  .View = lvwReport
        
  For X = 1 To UBound(listing_ArtNr)
    .ListItems.Add , "Artikel" & X, listing_ArtNr(X)
    .ListItems(X).SubItems(1) = Listarray(X).LD_Titel
    .ListItems(X).SubItems(2) = Listarray(X).LD_Preis
    .ListItems(X).SubItems(3) = Listarray(X).LD_Gebote
    .ListItems(X).SubItems(4) = Listarray(X).LD_Zeit
  Next X
  
End With

End Sub

Private Sub Form_Resize()
Dim listColwidth As Integer
Dim listcolheight As Integer

If Me.Height < 3000 Or Me.Width < 7000 Then
  If Me.Height < 3000 Then Me.Height = 3000
  If Me.Width < 7000 Then Me.Width = 7000
Else
  With ListView1
    listColwidth = (Me.Width - (.Left + 100)) / 100
    listcolheight = Me.Height - (.Top + (3 * Me.CB_markall.Height))
    .Width = listColwidth * 99
    .Height = listcolheight
    .ColumnHeaders.Item(1).Width = listColwidth * 13
    .ColumnHeaders.Item(2).Width = listColwidth * 38
    .ColumnHeaders.Item(3).Width = listColwidth * 20
    .ColumnHeaders.Item(4).Width = listColwidth * 10
    .ColumnHeaders.Item(5).Width = listColwidth * 13.5
  
    Me.CB_markall.Width = .Width / 2
    Me.CB_unmarkall.Width = .Width / 2
    Me.CB_unmarkall.Left = .Left + Me.CB_markall.Width
    Me.CB_markall.Top = .Top + .Height + 10
    Me.CB_unmarkall.Top = .Top + .Height + 10
    Me.CB_addmarked.Left = .Left + .Width - Me.CB_addmarked.Width
    Me.CB_cancel.Left = .Left + .Width - Me.CB_cancel.Width
  End With
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'Nach Spalten sortieren
Dim NewSub As Long
Dim I As Integer, X As Byte, Y As Byte
Dim Li As ListItem
Dim pos As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim pos3 As Integer
Dim temp_sub As String
Dim maxlen As Integer
Dim sec_pos3 As String
Dim min_pos As String
Dim hour_pos1 As String
Dim day_pos2 As String
Dim temp_DT
Dim tz_pos As Integer, tz_pos1 As Integer
Dim NowVal1 As Double
Dim Diff1 As Double
Dim temp_datetime As Double

On Error Resume Next
   
    With ListView1
        'ListView ruhig halten, Sichtbarkeit bleibt trotzdem erhalten
        .Visible = False
        
        'zu sortierende Spalte bestimmen
        .SortKey = ColumnHeader.Index - 1
        
        'Dummy-Spalte einfügen mit Breite 0
        .ColumnHeaders.Add , , "Dummy", 0
        
        'Nummer der Dummy-Spalte
        NewSub = .ColumnHeaders.count - 1
        
        'Index Spalte
        Select Case ColumnHeader.Index
          Case 3
            'Sortiere nach Preis
            For I = 1 To .ListItems.count
              If Len(.ListItems(I).SubItems(2)) > maxlen Then
                maxlen = Len(.ListItems(I).SubItems(2))
              End If
            Next I
            
            'preis rausziehn und auf maxlen füllen
            For I = .ListItems.count To 1 Step -1
                Set Li = .ListItems(I)
                temp_sub = Li.SubItems(2) & " "
                pos = InStr(1, temp_sub, " ", vbTextCompare)
                pos1 = InStr(pos + 1, temp_sub, " ", vbTextCompare)
                temp_sub = Mid(temp_sub, pos, pos1 - pos)
                temp_sub = Replace(temp_sub, ",", "", 1, -1, vbTextCompare)
                temp_sub = Replace(temp_sub, ".", "", 1, -1, vbTextCompare)
                If Len(temp_sub) < maxlen Then temp_sub = Space(maxlen - Len(temp_sub)) & temp_sub
                Li.SubItems(NewSub) = temp_sub
            Next I
            
            'zu sortierende Spalte umbiegen
            .SortKey = NewSub
          
          Case 4
            'Sortiere nach Gebote
            For I = 1 To .ListItems.count
              temp_sub = .ListItems(I).SubItems(3)
                temp_sub = Replace(temp_sub, "Sofortkauf", "0", 1, -1, vbTextCompare)
                temp_sub = Replace(temp_sub, "-", "0", 1, -1, vbTextCompare)
              If Len(temp_sub) > maxlen Then
                maxlen = Len(temp_sub)
              End If
            Next I
            
            For I = .ListItems.count To 1 Step -1
                Set Li = .ListItems(I)
                temp_sub = Li.SubItems(3)
                temp_sub = Replace(temp_sub, "Sofortkauf", "0", 1, -1, vbTextCompare)
                temp_sub = Replace(temp_sub, "-", "0", 1, -1, vbTextCompare)
                If Len(temp_sub) < maxlen Then
                  temp_sub = Space(maxlen - Len(temp_sub)) & temp_sub
                End If
                Li.SubItems(NewSub) = temp_sub
            Next I
            
            'zu sortierende Spalte umbiegen
            .SortKey = NewSub
          
          Case 5
            'Sortiere nach Verbleibende Zeit
            For I = .ListItems.count To 1 Step -1
                Set Li = .ListItems(I)
                temp_sub = Li.SubItems(4)
                
                temp_sub = Replace(temp_sub, Chr(160), " ", 1, -1, vbTextCompare)
                                 
                'vorher abfrage ob timezone vorhanden->date/time
                pos = 0
                tz_pos = InStr(1, temp_sub, ansTime_1, vbTextCompare)
                If tz_pos > 0 Then pos = tz_pos
                tz_pos1 = InStr(1, temp_sub, ansTime_2, vbTextCompare)
                If tz_pos1 > 0 Then pos = tz_pos1
                
                If pos > 0 Then
                  temp_sub = Trim(Mid(temp_sub, 1, pos - 1))
                  temp_datetime = CDbl(CDate(temp_sub))
                  NowVal1 = Now()
                  Diff1 = temp_datetime - NowVal1
                  temp_sub = "0" & Int(Diff1) & _
                      Replace(CStr(TimeValue(CDate(Diff1))), ":", "", , , vbTextCompare)
                
                Else
                  
                  If InStr(1, temp_sub, ",", vbTextCompare) < 1 Then
                    temp_sub = Replace(temp_sub, " ", ",", 1, -1, vbTextCompare)
                  End If
                  
                  temp_DT = Split(temp_sub, ",", , vbTextCompare)
                 
                 'time/date-parts
                  For X = 0 To UBound(temp_DT)
                    
                    'sort-part-delims
                    For Y = 1 To 4
                      pos3 = InStr(1, LCase(temp_DT(X)), ansMultiSortDelimX(Y))
                      If pos3 > 0 Then
                        Select Case Y
                          Case 1
                            sec_pos3 = temp_DT(X)
                            sec_pos3 = Trim(Mid(temp_DT(X), 1, pos3 - 1))
                            'If Len(sec_pos3) < 2 Then sec_pos3 = "0" & sec_pos3
                            Exit For
                          Case 2
                            min_pos = temp_DT(X)
                            min_pos = Trim(Mid(temp_DT(X), 1, pos3 - 1))
                            'If Len(min_pos) < 2 Then min_pos = "0" & min_pos
                            Exit For
                          Case 3
                            hour_pos1 = temp_DT(X)
                            hour_pos1 = Trim(Mid(temp_DT(X), 1, pos3 - 1))
                            'If Len(hour_pos1) < 2 Then hour_pos1 = "0" & hour_pos1
                            Exit For
                          Case 4
                            day_pos2 = temp_DT(X)
                            day_pos2 = Trim(Mid(temp_DT(X), 1, pos3 - 1))
                            'If Len(day_pos2) < 2 Then day_pos2 = "0" & day_pos2
                            Exit For
                        End Select
                      End If
                      
                    Next Y
                    
                  Next X
                  
                  'Auffüllen
                  sec_pos3 = String(2 - Len(sec_pos3), "0") & sec_pos3
                  min_pos = String(2 - Len(min_pos), "0") & min_pos
                  hour_pos1 = String(2 - Len(hour_pos1), "0") & hour_pos1
                  day_pos2 = String(2 - Len(day_pos2), "0") & day_pos2

                  temp_sub = day_pos2 & hour_pos1 & min_pos & sec_pos3
                  sec_pos3 = "": min_pos = "": hour_pos1 = "": day_pos2 = ""

                End If 'pos
                 
                'Debug.Print "temp_sub: " & temp_sub
                'Debug.Print "------------------" & vbCrLf
                 
                Li.SubItems(NewSub) = temp_sub
            Next I
            
            'zu sortierende Spalte umbiegen
            .SortKey = NewSub
        End Select
        
        'SortOrder bestimmen Asc oder Desc
        If .SortOrder = lvwAscending Then
          .SortOrder = lvwDescending
        Else
          .SortOrder = lvwAscending
        End If
            
        'Sortieren
        .Sorted = True
      
        'Dummy-Spalte entfernen
        .ColumnHeaders.Remove .ColumnHeaders.count
        
        'Auf 1. Zeile scrollen
        .ListItems(1).Selected = True
        .ListItems(1).EnsureVisible
            
        'sichtbar machen
        .Visible = True
    End With
End Sub

Private Sub CB_addmarked_Click()
Dim X As Integer
Dim pos As Integer

ReDim listing_ArtNr(0)
With ListView1
  For X = 1 To .ListItems.count
    'If .ListItems.Item(x).Checked = True Then '.checked gibts nicht im CC5 Listview
    If IsChecked(ListView1, X - 1) = True Then
      If UBound(listing_ArtNr) < X Then ReDim Preserve listing_ArtNr(X)
      listing_ArtNr(X) = .ListItems.Item(X).text
    End If
  Next X
End With

If Me.f_GroupCheck > 0 Then
  If Len(f_BidGroupValue) > 0 Then
    pos = InStr(1, f_BidGroupValue, ";", vbTextCompare)
    If pos > 0 Then
      If Trim(CInt(Mid(f_BidGroupValue, pos + 1))) > AnzChecked Then
        ListBidGroup = Trim(Mid(f_BidGroupValue, 1, pos)) & CStr(AnzChecked)
      Else
        ListBidGroup = Trim(f_BidGroupValue)
      End If
    Else
      ListBidGroup = ""
    End If
  Else
    ListBidGroup = ""
  End If

  Debug.Print "ListBidGroup: " & ListBidGroup
  
  If Len(Trim(f_GrPrice)) > 0 Then
    ListBidPrice = Trim(f_GrPrice)
  Else
    ListBidPrice = ""
  End If
Else
  ListBidGroup = ""
  ListBidPrice = ""
End If
Unload frmChooseArtList
End Sub

Private Sub CB_cancel_Click()
ReDim listing_ArtNr(0)
Unload frmChooseArtList
End Sub

Private Sub CB_markall_Click()
'Dim x As Integer
'
'With ListView1
'  For x = 1 To .ListItems.count
'    .ListItems.Item(x).Checked = True
'  Next x
'End With
'AnzChecked = x

SetCheckAllItems True
AnzChecked = ListView1.ListItems.count

Me.CB_addmarked.Enabled = True
f_GroupCheck.Enabled = True

End Sub

Private Sub CB_unmarkall_Click()
'Dim x As Integer
'
'With ListView1
'  For x = 1 To .ListItems.count
'    .ListItems.Item(x).Selected = False
'  Next x
'End With

SetCheckAllItems False
AnzChecked = 0

Me.CB_addmarked.Enabled = False
f_GroupCheck = 0
f_GroupCheck_Click
f_GroupCheck.Enabled = False

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
Me.Caption = " " & Item.SubItems(1)
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Long

If KeyCode = vbKeySpace Then
  AnzChecked = 0
  With ListView1
    For I = 1 To .ListItems.count
      'If .ListItems.Item(I).Checked = True Then AnzChecked = AnzChecked + 1
      If IsChecked(ListView1, I - 1) = True Then AnzChecked = AnzChecked + 1
    Next I
  End With
End If

If AnzChecked > 0 Then
  Me.CB_addmarked.Enabled = True
  f_GroupCheck.Enabled = True
Else
  Me.CB_addmarked.Enabled = False
  f_GroupCheck = 0
  f_GroupCheck_Click
  f_GroupCheck.Enabled = False
End If
Debug.Print "Anzchecked: " & AnzChecked

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Long

If Button = 1 Then
AnzChecked = 0
  With ListView1
    For I = 1 To .ListItems.count
      'If .ListItems.Item(I).Checked = True Then AnzChecked = AnzChecked + 1
      If IsChecked(ListView1, I - 1) = True Then AnzChecked = AnzChecked + 1
    Next I
  End With
End If

If AnzChecked > 0 Then
  Me.CB_addmarked.Enabled = True
  f_GroupCheck.Enabled = True
Else
  Me.CB_addmarked.Enabled = False
  f_GroupCheck = 0
  f_GroupCheck_Click
  f_GroupCheck.Enabled = False
End If
Debug.Print "Anzchecked: " & AnzChecked
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Hilfsfunktionen

'allen ListView-Einträgen ein Häkchen verpassen (True)
'oder keinem (False)
Private Sub SetCheckAllItems(bState As Boolean)

   Dim LV As LV_ITEM
   Dim lvCount As Long
   Dim lvIndex As Long
   Dim lvState As Long
   
  'because IIf is less efficient than a
  'traditional If..Then..Else statement, just call
  'once to save the state mask to a local variable
   lvState = IIf(bState, &H2000, &H1000)
   
  'listview has 0 to count -1 items
   lvCount = ListView1.ListItems.count - 1
   
   Do
         
      With LV
         .mask = LVIF_STATE
         .state = lvState
         .stateMask = LVIS_STATEIMAGEMASK
      End With
      
      Call SendMessage(ListView1.hWnd, LVM_SETITEMSTATE, lvIndex, LV)
      lvIndex = lvIndex + 1
   
   Loop Until lvIndex > lvCount
  
End Sub

