Attribute VB_Name = "modMZeit"
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
' $author: hjs$
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
'*****************************************************************************
'*** Wandlungsroutinen von / nach Normalzeit. Bezugszeit 01..Jan 2000, 00:00:00
'*****************************************************************************
Option Explicit

'Bezugszeit 1. Jan. 00
Private Normzeit As Date '= "01.01.00 00:00:00"

Public Function To_Normzeit(zeit As Date) As Long
'*****************************************************************************
'*** Wandlungsroutine Date nach Normalzeit. Bezugszeit 01.Jan 2000, 00:00:00
'*****************************************************************************

Normzeit = DateSerial(2000, 1, 1)
To_Normzeit = DateDiff("s", Normzeit, zeit)

End Function

Public Function From_Normzeit(zeit As Long) As Date
'*****************************************************************************
'*** Wandlungsroutine Normalzeit nach Date. Bezugszeit 01.Jan 2000, 00:00:00
'*****************************************************************************

Normzeit = DateSerial(2000, 1, 1)
From_Normzeit = DateAdd("s", zeit, Normzeit)

End Function

