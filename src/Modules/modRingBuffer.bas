Attribute VB_Name = "modRingBuffer"
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
Option Explicit
'
' Ringbuff- Zugriffe, für Mails und Artikel, lg 14.05.03
'
'Ringbuffer- Typ 0815
'auf 1000 aufgebohrt lg
Private Type udtRingBuffer
    ReadCount As Integer
    WriteCount As Integer
    BuffSize As Integer
    Data(1000) As String
End Type

'Ringbuffer für Auktionsendemails
Private mtMailBuffer As udtRingBuffer
'Ringbuffer für Artikellesen, lg 14.05.03
Private mtArtikelBuffer As udtRingBuffer

Public Sub InitRingBuffs()
    
    mtMailBuffer.BuffSize = UBound(mtMailBuffer.Data())
    mtArtikelBuffer.BuffSize = UBound(mtArtikelBuffer.Data())
    
End Sub

Public Sub InsertMailBuff(sTxt As String)
        
    On Error Resume Next
    
    With mtMailBuffer
        .ReadCount = .ReadCount + 1
        If .ReadCount > UBound(.Data()) Then
            .ReadCount = 1
        End If
        .Data(.ReadCount) = sTxt
    End With
    frmHaupt.MailBuffTimer.Enabled = True
    
End Sub

Public Function ReadMailBuff(sTxt As String) As Boolean
    
    On Error Resume Next
    
    ReadMailBuff = False
    frmHaupt.MailBuffTimer.Enabled = False 'erstmal ausmachen, lg 14.05.03
    
    With mtMailBuffer
        If .WriteCount <> .ReadCount Then
            .WriteCount = .WriteCount + 1
            If .WriteCount > UBound(.Data()) Then
                .WriteCount = 1
            End If
            sTxt = .Data(.WriteCount)
            ReadMailBuff = True
        End If
    End With
    'in .Tag hinterlegen, wie es nach der evtl. Aktion weitergeht, lg 14.05.03
    frmHaupt.MailBuffTimer.Tag = CStr(ReadMailBuff)
    
End Function

Public Sub InsertArtikelBuff(sTxt As String)
        
    On Error Resume Next
    
    With mtArtikelBuffer
        .ReadCount = .ReadCount + 1
        If .ReadCount > UBound(.Data()) Then
            .ReadCount = 1
        End If
        .Data(.ReadCount) = sTxt
    End With
    
    frmHaupt.ArtikelBuffTimer.Enabled = True
    Call DebugPrint("InsertArtikelBuff: " & sTxt, 3)
    
End Sub

Public Function ReadArtikelBuff(sTxt As String) As Boolean
    
    On Error Resume Next
    
    ReadArtikelBuff = False
    frmHaupt.ArtikelBuffTimer.Enabled = False 'erstmal ausmachen
    
    With mtArtikelBuffer
        If .WriteCount <> .ReadCount Then
            .WriteCount = .WriteCount + 1
            If .WriteCount > UBound(.Data()) Then
                .WriteCount = 1
            End If
            sTxt = .Data(.WriteCount)
            ReadArtikelBuff = True
        End If
    End With
    'in .Tag hinterlegen, wie es nach der evtl. Aktion weitergeht, lg 14.05.03
    frmHaupt.ArtikelBuffTimer.Tag = CStr(ReadArtikelBuff)
    
End Function

