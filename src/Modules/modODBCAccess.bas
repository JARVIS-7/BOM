Attribute VB_Name = "modODBCAccess"
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
' Zugriff auf ODBC- Database(s)
'
Dim Cn As Object 'ADODB.Connection

Public Sub ODBC_Connect()
Dim ConnectString As String
On Error GoTo errhdl

Set Cn = New ADODB.Connection

'neuer Zugriff ist flexibler ..
'
DoEvents
ConnectString = "PROVIDER=" & gsOdbcProvider & ";"

If Trim(gsOdbcDb) <> "" Then
   ConnectString = ConnectString & "DATA SOURCE=" & gsOdbcDb & ";"
End If

If gsOdbcUser <> "" Then
    ConnectString = ConnectString & "USER ID=" & gsOdbcUser & ";PASSWORD=" & gsOdbcPass & ";"
End If

Cn.Open ConnectString
DoEvents

If Cn.State <> 0 Then
    Exit Sub
End If

errhdl:

gsOdbcStopRead = True

If Not gbAutoMode Then
    MsgBox "Keine Verbindung zu DB möglich" & vbCrLf _
      & "Provider=" & gsOdbcProvider & vbCrLf _
      & "DataSource=" & gsOdbcDb & vbCrLf _
      & "Fehler: " & Err.Description
End If

Set Cn = Nothing

End Sub

Public Sub ODBC_ResetConnection()

On Error Resume Next

Cn.Close
Set Cn = Nothing

End Sub

Public Function ODBC_Check() As Boolean

ODBC_Check = False
If Not Cn Is Nothing Then
    ODBC_Check = Cn.State <> 0
End If

End Function
Public Sub ODBC_ReadNew()
'
' lesen aus der DB
'
Dim sSQL     As String
Dim rc      As New ADODB.Recordset
Dim i       As Integer
Dim bFound As Boolean
Dim sWert As String
Dim sArtikelNr As String

On Error GoTo errhdl

If Cn Is Nothing Then
    ODBC_Connect
End If

If Cn.State = 0 Then Exit Sub

sSQL = "SELECT * FROM BOMData"

With rc
    .ActiveConnection = Cn
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = sSQL
    .Open
End With

If rc.RecordCount > 0 Then
        rc.MoveFirst
        Do Until rc.EOF Or gsOdbcStopRead
            DoEvents
            'Satz auswerten und eintragen
            'Prüfen, ob neu:
            For i = 1 To giAktAnzArtikel
                bFound = False
                If gtarrArtikelArray(i).Artikel = rc.Fields("Artikel") Then
                    bFound = True
                    Exit For
                End If
            Next i
            
            If Not bFound Then
                'neuer Satz
                frmHaupt.AddArtikel rc.Fields("Artikel")
            End If
            
            'Daten ggf. updaten
            'Gebot, Kommentar und Gruppe übernehmen
            For i = 1 To giAktAnzArtikel
                sArtikelNr = gtarrArtikelArray(i).Artikel
                If gtarrArtikelArray(i).Artikel = rc.Fields("Artikel") Then
                
                    sWert = "Gebot"
                    If Not IsNull(rc.Fields("Gebot")) Then
                        If gtarrArtikelArray(i).Gebot <> rc.Fields("Gebot") Then
                            gtarrArtikelArray(i).Gebot = rc.Fields("Gebot")
                        End If
                    Else
                        gtarrArtikelArray(i).Gebot = 0
                    End If
                    sWert = "Status"
                    If Not IsNull(rc.Fields("Status")) Then
                        gtarrArtikelArray(i).Status = rc.Fields("Status")
                    End If
                    sWert = "Gruppe"
                    If Not IsNull(rc.Fields("Gruppe")) Then
                        gtarrArtikelArray(i).Gruppe = rc.Fields("Gruppe")
                    Else
                        gtarrArtikelArray(i).Gruppe = ""
                    End If
                    sWert = "Kommentar"
                    If Not IsNull(rc.Fields("Kommentar")) Then
                        gtarrArtikelArray(i).Kommentar = rc.Fields("Kommentar")
                    Else
                        gtarrArtikelArray(i).Kommentar = ""
                    End If
                    sWert = "EbayUser"
                    If Not IsNull(rc.Fields("EbayUser")) Then
                        gtarrArtikelArray(i).eBayUser = Trim(rc.Fields("EbayUser"))
                        gtarrArtikelArray(i).UserAccount = Trim(rc.Fields("EbayUser"))
                    Else
                        gtarrArtikelArray(i).eBayUser = ""
                    End If
                    sWert = "EbayPass"
                    If Not IsNull(rc.Fields("EbayPass")) Then
                        gtarrArtikelArray(i).eBayPass = Trim(rc.Fields("EbayPass"))
                    Else
                        gtarrArtikelArray(i).eBayPass = ""
                    End If
                    sWert = "UseToken"
                    On Error Resume Next
                    If Not IsNull(rc.Fields("UseToken")) Then
                        gtarrArtikelArray(i).UseToken = Trim(rc.Fields("UseToken"))
                    Else
                        gtarrArtikelArray(i).UseToken = ""
                    End If
                    On Error GoTo errhdl

                    Exit For
                End If
            Next i

            'nächsten Satz
            rc.MoveNext
        Loop
End If
rc.Close
Exit Sub

errhdl:

Dim ErrorString As String
ErrorString = Err.Number & ": " & Err.Description & vbCrLf & _
"aufgetreten beim Lesen von Artikel " & sArtikelNr & ", Wert: " & sWert
DebugPrint "ReadNew: ODBC-Err: " & ErrorString

Err.Clear
On Error Resume Next
If rc.State = 1 Then
    rc.CancelUpdate
    rc.Close
End If

End Sub

Public Sub ODBC_UpdateArtikel()
'
' Update der DB
'
Dim i As Integer
Dim sSQL     As String
Dim rc      As New ADODB.Recordset
Dim sWert As String
Dim sArtikelNr As String

On Error GoTo errhdl

If Cn.State = 0 Then
    ODBC_Connect
End If

If Cn.State = 0 Then Exit Sub

With rc
    .ActiveConnection = Cn
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
End With

'und Update/ Eintrag der neuen Artikel
For i = 1 To giAktAnzArtikel
    DoEvents
    sSQL = "select * from BOMData where Artikel='" & gtarrArtikelArray(i).Artikel & "'"

    rc.Source = sSQL
    rc.Open
    
    sArtikelNr = gtarrArtikelArray(i).Artikel
    If rc.RecordCount > 0 Then

        sWert = "Titel"
        If IsNull(rc.Fields("Titel").Value) Then
            rc.Fields("Titel").Value = gtarrArtikelArray(i).Titel
        End If
        sWert = "Endezeit"
        If IsNull(rc.Fields("Endezeit").Value) Then
            rc.Fields("Endezeit").Value = gtarrArtikelArray(i).EndeZeit
        End If
        sWert = "Aktpreis"
        rc.Fields("Aktpreis").Value = gtarrArtikelArray(i).AktPreis
        sWert = "WE"
        rc.Fields("WE").Value = gtarrArtikelArray(i).WE
        'Rc.Fields("Gebot").Value = gtarrArtikelArray(i).Gebot
        'Rc.Fields("Gruppe").Value = gtarrArtikelArray(i).Gruppe
        sWert = "Status"
        rc.Fields("Status").Value = gtarrArtikelArray(i).Status
        sWert = "AnzGebote"
        rc.Fields("AnzGebote").Value = gtarrArtikelArray(i).AnzGebote
        sWert = "Bieter"
        rc.Fields("Bieter").Value = gtarrArtikelArray(i).Bieter
        sWert = "Verkaeufer"
        rc.Fields("Verkaeufer").Value = gtarrArtikelArray(i).Verkaeufer
On Error Resume Next 'evtl. ist das Feld Versand noch nicht vorhanden, dann übergehen, lg 06.09.03
        rc.Fields("Versand").Value = gtarrArtikelArray(i).Versand
On Error GoTo errhdl
        sWert = "Kommentar"
        rc.Fields("Kommentar").Value = gtarrArtikelArray(i).Kommentar
        sWert = "kompletten Datensatz aktualisieren"
        rc.Update
    End If    'Satz vorhanden
    rc.Close
    If gsOdbcStopRead Then Exit For
Next i

Exit Sub

errhdl:

Dim ErrorString As String
ErrorString = Err.Number & ": " & Err.Description & vbCrLf & _
"aufgetreten beim Schreiben von Artikel " & sArtikelNr & ", Wert: " & sWert
DebugPrint "Update: ODBC-Err: " & ErrorString

Err.Clear
On Error Resume Next
If rc.State = 1 Then
    rc.CancelUpdate
    rc.Close
End If
End Sub

Public Sub ODBC_RemoveArtikel()

Dim i As Integer
Dim sSQL     As String
Dim rc      As New ADODB.Recordset

On Error GoTo errhdl

If Cn.State = 0 Then
    ODBC_Connect
End If

With rc
    .ActiveConnection = Cn
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
End With

'wir räumen die alten Artikel ab:
i = 1
While i <= giAktAnzArtikel And Not gsOdbcStopRead
    DoEvents
    sSQL = "select * from BOMData where Artikel='" & gtarrArtikelArray(i).Artikel & "'"

    rc.Source = sSQL
    rc.Open
    
     
    If rc.RecordCount = 0 Then
        frmHaupt.RemoveArtikel i
        i = 1 'pervers ;-)
    Else
        i = i + 1
    End If
    
    rc.Close
Wend
Exit Sub

errhdl:

Dim ErrorString As String
ErrorString = "Error Nummer " & Err.Number & ": " & Err.Description
DebugPrint "Remove: ODBC-Err: " & ErrorString

Err.Clear
On Error Resume Next
rc.Cancel
If rc.State = 1 Then
    rc.Close
End If
End Sub

