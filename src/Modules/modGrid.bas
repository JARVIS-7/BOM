Attribute VB_Name = "modGrid"
Option Explicit

'Column-Consts->noch ändern in Usersettings + Width
Public Const Scolindex = 1      'Array-Index
Public Const gChoose = 2        'Auswahl
Public Const gArtnr = 3         'Artikelnummer
Public Const gEzeit = 4         'Ende-Zeit
Public Const gTitel = 5         'Beschreibung
Public Const gComment = 6       'Kommentar
Public Const gPrice = 7         'Aktueller-Preis
Public Const gAnzbids = 8       'Anzahl Gebote
Public Const gSeller = 9        'Verkäufer
Public Const gRating = 10       'Bewertung Anzahl / Prozent
Public Const gShipping = 11     'Versandkosten
Public Const gRzeit = 12        'Rest-Zeit
Public Const gBidprice = 13     'Gebot
Public Const gCurrency = 14     'Währung
Public Const gGroup = 15        'Gruppe+Anzahl
Public Const gStat = 16         'Status

'Spaltenbreiten + Anzeige y/n ohne Indexspalte u. Statusspalte(eingepasst)
Public Const DefaultGridwidths As String = "20,80,80,200,65,80,40,60,65,55,80,50,30,50,0"
Public Const DefaultGridshows As String = "0,1,1,1,0,1,0,0,0,0,1,1,1,1,1"
Public Usr_Widths, Usr_Shows
'

Public Sub InitGrid(F As Form)
'Col-Konstanten

With frmHaupt.grid
  F.ImageList2.MaskColor = .BackColor
  .Top = F.Toolbar1.Top + F.Toolbar1.Height + 15
  .Height = F.ScaleHeight - (F.Toolbar1.Top + F.Toolbar1.Height + F.StatusBar1.Height)
  .Left = F.Toolbar1.Left
  .Width = F.Toolbar1.Width
  
  'Grid-Eigenschaften
  '.HotTrack = True
  'Alpha -Blending
  '.SelectionAlphaBlend = True
'  .SelectionOutline = False
'  .DrawFocusRectangle = False
  .HighlightForeColor = vbBlack
  .HighlightBackColor = vbWhite
  .BorderStyle = ecgBorderStyle3d
  .BackColor = vbWhite
  .GridLines = True
  .GridLineMode = ecgGridFillControl 'ecgGridStandard
  .GridLineColor = vb3DShadow
  .AlternateRowBackColor = RGB(252, 252, 230)
  .ImageList = frmHaupt.ImageList2
  Set .BackgroundPicture = frmHaupt.Picture2.Picture
  
  .DefaultRowHeight = 35 'Setformsize noch anpassen
  .Editable = True
  .SingleClickEdit = False
  .StretchLastColumnToFit = True
  .SplitRow = 0
  .AllowGrouping = False
  .HighlightSelectedIcons = False
  .Draw
  .Redraw = False
  
  'Spalten hinzufügen
  .AddColumn "Scolindex", , , , 15, 1, True, , , , , CCLSortNumeric
  .AddColumn "gChoose", " ", ecgHdrTextALignLeft, , Usr_Widths(0), Usr_Shows(0), , , , , , CCLSortIcon
  .AddColumn "gArtnr", "Artikelnummer", ecgHdrTextALignRight, , Usr_Widths(1), Usr_Shows(1), False, , False, , , CCLSortStringNoCase
  .AddColumn "gEzeit", "Endet", ecgHdrTextALignLeft, , Usr_Widths(2), Usr_Shows(2), False, , False, , , CCLSortDate
  .AddColumn "gTitel", "Beschreibung", ecgHdrTextALignLeft, , Usr_Widths(3), Usr_Shows(3), False, , False, , , CCLSortStringNoCase
  .AddColumn "gComment", "Kommentar", ecgHdrTextALignLeft, , Usr_Widths(4), Usr_Shows(4), False, , False, , , CCLSortString
  .AddColumn "gPrice", "Akt.Preis", ecgHdrTextALignCentre, , Usr_Widths(5), Usr_Shows(5), False, , False, , , CCLSortNumeric
  .AddColumn "gAnzbids", "Anz. Gebote", ecgHdrTextALignRight, , Usr_Widths(6), Usr_Shows(6), False, , False, , , CCLSortNumeric
  .AddColumn "gSeller", "Verkäufer", ecgHdrTextALignLeft, , Usr_Widths(7), Usr_Shows(7), False, , False, , , CCLSortString
  .AddColumn "gRating", "Bewertung", ecgHdrTextALignLeft, , Usr_Widths(8), Usr_Shows(8), False, , False, , , CCLSortString
  .AddColumn "gShipping", "Versand", ecgHdrTextALignRight, , Usr_Widths(9), Usr_Shows(9), False, , False, , , CCLSortNumeric
  .AddColumn "gRzeit", "Restzeit", ecgHdrTextALignLeft, , Usr_Widths(10), Usr_Shows(10), False, , False, , , CCLSortDate
  .AddColumn "gBidprice", "Gebot", ecgHdrTextALignRight, , Usr_Widths(11), Usr_Shows(11), False, , False, , , CCLSortString
  .AddColumn "gCurrency", "WE", ecgHdrTextALignLeft, , Usr_Widths(12), Usr_Shows(12), False, , False, , , CCLSortString
  .AddColumn "gGroup", "Gruppe", ecgHdrTextALignLeft, , Usr_Widths(13), Usr_Shows(13), False, , False, , , CCLSortString
  .AddColumn "gStat", "Status", ecgHdrTextALignLeft, , , Usr_Shows(14), False, , False, , , CCLSortString

  'zusätzliche Spalten am Anfang ausgeblendet
  'Array-Index Spalte(Scolindex)->Abgriff Index nach sortierung

  frmHaupt.SetIndexCol
  .Redraw = True
  
End With


End Sub
