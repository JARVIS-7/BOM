Attribute VB_Name = "modMain"
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

'
Option Explicit


Private Sub Main()
'...
'...

    glThreadID = App.ThreadID
    
    If (WeAreAlone() Or Command() = "" Or IsJobCommand()) Then
        Call SetAppDataPath
        Load frmHaupt
        Call TagWindow(frmDummy.hWnd)
        
    Else
        
        Call ProcessCmdline
    
    End If
    
End Sub

