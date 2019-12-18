Attribute VB_Name = "modSearchList"
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

Public Function Get_Artikel_Liste(ByVal str As String) As Integer
Dim pos As Long, pos1 As Long, pos2 As Long
Dim pos3 As Long, pos4 As Long
Dim start_tr As Long
Dim end_tr As Long
Dim item_start_pos  As Long
Dim temp_ergebnis As String
Dim Ergebnis As String
Dim temp_erg As String
Dim temp_row As String
Dim rest_temprow As String
Dim SaveFile As String
Dim SaveFile_sjs As String
Dim l_Artikel As String
Dim I As Integer
Dim X As Integer
Dim art_nr1 As String
Dim art_nr2 As String
Dim IsListurl As Boolean
Dim temp_tr As Variant
Dim DelimX As String

On Error GoTo errhdl

ReDim Listarray(0)
ReDim listing_ArtNr(0)
'Ergebnis_SuchListe = ""

I = 1
Do
  If InStr(1, str, ansMultiAddX(I)) > 0 Then
    IsListurl = True
  End If
  I = I + 1
Loop While I < UBound(ansMultiAddX) + 1

If IsListurl Then
  
  SaveFile = TempPfad & "\listing-1.html"
  SaveFile_sjs = TempPfad & "\listing-1_komplett_ohne_script.html"
  
  'Delete-Flags
  I = 1
  Do While I < UBound(ansMultiFlagsDelX) + 1
    If InStr(1, str, ansMultiFlagsDelX(I), vbTextCompare) Then
    'str = Replace(str, "&rdir=0", "", , , vbTextCompare)
    str = Replace(str, ansMultiFlagsDelX(I), "", , , vbTextCompare)
    End If
    I = I + 1
  Loop
  'Debug.Print "str nach del: " & str
  
  'Add-Flags
  I = 1
  Do While I < UBound(ansMultiFlagsAddX) + 1
    str = str & ansMultiFlagsAddX(I)
    I = I + 1
  Loop
  'Debug.Print "str nach add: " & str
  
  'Read ArtikelListe + scripts entfernen + umwandeln
  Ergebnis = StripJavaScript(ShortPost(str))
  temp_erg = HtmlZeichenconvert(Ergebnis)
  
  'erstes vorkommendes item= + suche <tr vor item= -> anfang erster Artikel bis ende
  pos = InStr(1, temp_erg, ansMultiListItem, vbTextCompare)
  pos1 = InStrRev(LCase(temp_erg), "<tr", pos, vbTextCompare)
  temp_erg = Mid(temp_erg, pos1)
  
  SaveToFile temp_erg, SaveFile_sjs
  
  I = 0
  'DelimX = ""
  
  Do While Len(temp_erg) > 0
    I = I + 1
 
    'Gültige Tablerows finden
    pos = InStr(1, temp_erg, ansMultiListItem, vbTextCompare)
    item_start_pos = pos
    
    'ArtNr-Delimiter ermitteln valid: string or chr(x)
    If Len(DelimX) = 0 Then
      Do While X < UBound(ansMultiDelimitX) + 1
        If IsNumeric(ansMultiDelimitX(X)) Then
          DelimX = Chr(CInt(ansMultiDelimitX(X)))
        Else
          DelimX = ansMultiDelimitX(X)
        End If
        'pos1 = InStr(pos + 5, temp_erg, DelimX, vbTextCompare)
        pos1 = InStr(pos + Len(ansMultiListItem), temp_erg, DelimX, vbTextCompare)
        'art_nr1 = Trim(Mid(temp_erg, pos + 5, pos1 - pos - 5))
        art_nr1 = Trim(Mid(temp_erg, pos + Len(ansMultiListItem), pos1 - pos - Len(ansMultiListItem)))
        If IsNumeric(art_nr1) Then Exit Do
        X = X + 1
      Loop
    Else
      'pos1 = InStr(pos + 5, temp_erg, DelimX, vbTextCompare)
      pos1 = InStr(pos + Len(ansMultiListItem), temp_erg, DelimX, vbTextCompare)
      'art_nr1 = Trim(Mid(temp_erg, pos + 5, pos1 - pos - 5))
      art_nr1 = Trim(Mid(temp_erg, pos + Len(ansMultiListItem), pos1 - pos - Len(ansMultiListItem)))
    End If
    
    If Not IsNumeric(art_nr1) Then
      Get_Artikel_Liste = 0
      Exit Function
    End If
        
    'nur erste Artikelnummer ins array ->folgende ungleiche(artnr2) nur für split-tr
    If UBound(listing_ArtNr) < I Then ReDim Preserve listing_ArtNr(I)
    listing_ArtNr(I) = Val(Trim(art_nr1))
    
    
    Do While pos > 0
      'x.item=
      pos = InStr(pos1 + 1, temp_erg, ansMultiListItem, vbTextCompare)
      'x. delimiter
      pos1 = InStr(pos + Len(ansMultiListItem), temp_erg, DelimX, vbTextCompare)
      'x. Artikelnummer
      'art_nr2 = Trim(Mid(temp_erg, pos + 5, pos1 - pos - 5))
      art_nr2 = Trim(Mid(temp_erg, pos + Len(ansMultiListItem), pos1 - pos - Len(ansMultiListItem)))
      'falscher delimiter, je nach Liste verschieden ->guckst du weidäär
          
      If Val(art_nr1) <> Val(art_nr2) Then
        Exit Do
      End If
    Loop 'pos > 0
  
    'wenn nächste Artikelnummer vorhanden pos > 0
    If pos > 0 Then
      start_tr = InStrRev(temp_erg, "<tr", item_start_pos, vbTextCompare)
      end_tr = InStrRev(temp_erg, "</tr>", pos, vbTextCompare)
      temp_row = Mid(temp_erg, start_tr, (end_tr + 5) - start_tr)
      temp_erg = Mid(temp_erg, end_tr + 5) 'rest for next valid tr
    Else
      'keine ArtNR -> tags eliminieren->erstes </tr> nach start für abschluss suchen
      start_tr = InStrRev(temp_erg, "<tr", item_start_pos, vbTextCompare)
      temp_erg = Del_InRowTables(Mid(temp_erg, start_tr))
      start_tr = InStrRev(temp_erg, "<tr", item_start_pos, vbTextCompare)
      pos4 = InStr(start_tr, temp_erg, "</tr>", vbTextCompare)
      temp_row = Mid(temp_erg, start_tr, (pos4 + 5) - start_tr)
      temp_erg = "" 'ende HTML
    End If
    
    'InRowTables vorhanden?
    If InStr(1, LCase(temp_row), "table", vbTextCompare) > 0 Then
      rest_temprow = Del_InRowTables(temp_row)
    Else
      rest_temprow = temp_row
    End If
    
    'In TD's splitten
    temp_tr = Split(rest_temprow, "</td>", -1, vbTextCompare)
    
    'Infos rausziehn
    Call split_to_listdata1(temp_tr, I)
    
    SaveToFile rest_temprow, TempPfad & "\rest_temprow-" & I & ".html" 'gültige tr's

    temp_ergebnis = temp_ergebnis & rest_temprow
  
  Loop 'Len(temp_erg) > 0
  
  SaveToFile temp_ergebnis, SaveFile
  'Ergebnis_SuchListe = temp_ergebnis
  
  Get_Artikel_Liste = UBound(listing_ArtNr)
Else 'IsListUrl
  Get_Artikel_Liste = 0
End If 'IsListUrl

errhdl:
If error_handling(Err.Number, Err.Description, "Get_Artikel_Liste") Then
  Exit Function
Else
  Err.Clear
  Resume Next
End If

End Function

Private Sub split_to_listdata1(ByVal temp_tr, I As Integer)
Dim temp As String
Dim X As Integer
Dim pos As Long
Dim pos1 As Long

On Error Resume Next

If UBound(Listarray) < I Then ReDim Preserve Listarray(I)

'TR auf Struktur prüfen
If UBound(temp_tr) > 4 Then
  
  'Tags eliminieren->Rest to Array
  For X = 1 To UBound(temp_tr) - 1
    'dbg "x", x
    pos = InStr(1, temp_tr(X), "<", vbTextCompare)
    'dbg "pos", pos
    Do While pos > 0
      pos1 = InStr(pos, temp_tr(X), ">", vbTextCompare)
      'dbg "pos1", pos1
      temp_tr(X) = Mid(temp_tr(X), pos1 + 1)
      pos = InStr(1, temp_tr(X), "<", vbTextCompare)
      'dbg "pos", pos
      If pos >= 1 Then
        temp = temp & " " & Mid(temp_tr(X), 1, pos - 1)
      Else
        If Len(temp_tr(X)) > 0 Then temp = temp & " " & temp_tr(X)
      End If
      'temp = Replace(temp, Chr(8), " ", , , vbTextCompare)
      'temp = Replace(temp, Chr(9), " ", , , vbTextCompare)
      'temp = Replace(temp, Chr(10), " ", , , vbTextCompare)
      'temp = Replace(temp, Chr(13), " ", , , vbTextCompare)
      temp = DelSZ(temp)
    Loop
  
    'Einsortieren
    Select Case X
      Case 1      'Titel
        Listarray(I).LD_Titel = temp
      
      Case 2      'Preis
        pos = InStr(1, temp, " ", vbTextCompare) + 1
        pos1 = InStr(pos, temp, " ", vbTextCompare)
        If pos1 <= 0 Then
          Listarray(I).LD_Preis = temp
        Else
          Listarray(I).LD_Preis = Mid(temp, 1, pos1) & "/ " & Mid(temp, pos1 + 1)
        End If
        
      Case 3      'Gebote
        If Len(Trim(temp)) > 0 Then
          Listarray(I).LD_Gebote = temp
        Else
          Listarray(I).LD_Gebote = "Sofortkauf"
        End If
        
      Case 4      'Verbleibende Zeit
        Listarray(I).LD_Zeit = temp
      
      'Case Else
        'Debug.Print "case else in split_to_listdata: x=" & X
        'Debug.Print "temp: " & temp
    End Select
    temp = ""
  Next X
    
End If

End Sub

Private Function Del_InRowTables(ByVal str As String) As String
Dim pos As Long, pos1 As Long, pos2 As Long, pos3 As Long, pos4 As Long
Dim temp_table As String
Dim temp_row As String

On Error Resume Next

temp_row = str

pos = InStr(1, temp_row, "</table>", vbTextCompare)

Do While pos > 0
  
  pos1 = InStrRev(LCase(temp_row), "<table", pos - 1, vbTextCompare)
  
  If pos1 <= 0 Then
    temp_table = Mid(temp_row, 1, pos - 1) 'table-/table inkl. tables
    Exit Do
  End If
  
  temp_table = Mid(temp_row, pos1, (pos - pos1) + 8) 'table-/table inkl. tables
  pos3 = InStr(1, temp_table, "<", vbTextCompare) '1stes < innerhalb table
  
  Do While pos3 > 0 'widerholen solange < vorhanden
    pos4 = InStr(pos3 + 1, temp_table, ">", vbTextCompare) 'nächstes > nach <
    temp_table = Mid(temp_table, pos4 + 1)            'abschneiden bis >
    pos3 = InStr(1, temp_table, "<", vbTextCompare)  'nächstes < suchen
    If pos3 >= 1 Then
      temp_table = temp_table & " " & Mid(temp_table, 1, pos3 - 1)
      temp_table = Trim(temp_table)
    End If
  Loop
  
  temp_row = Mid(temp_row, 1, pos1 - 1) & temp_table & Mid(temp_row, pos + 8)
  pos = InStr(1, temp_row, "</table>", vbTextCompare)

Loop

Del_InRowTables = temp_row

End Function

Public Function error_handling(ByVal err_nr As Integer, _
                                ByVal er_desc As String, _
                                Optional Name As String, _
                                Optional err_source) As Boolean

'If Err.Number <> 0 And Err.Number <> 20 Then
Debug.Print "Fehler in: " & Name & vbCrLf & err_nr & " " & er_desc & vbCrLf & "Source: " & err_source
  'Err.Clear
  error_handling = False
'endif
dbg "error_handling", error_handling
End Function

Private Sub dbg(ByVal Name As String, Optional ByVal wert As Variant, Optional ByVal nl As Boolean = False)
Debug.Print Name & ": " & wert
End Sub

