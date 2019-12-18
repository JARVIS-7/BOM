VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ctlSNTP 
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   ScaleHeight     =   1380
   ScaleWidth      =   2205
   Begin VB.Timer TimerTimeout 
      Left            =   720
      Top             =   840
   End
   Begin MSWinsockLib.Winsock wscWinsock 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "SNTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "ctlSNTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'*
'* A C# SNTP Client
'*
'* Copyright (C)2001-2003 Valer BOCAN <vbocan@dataman.ro>
'* All Rights Reserved
'*
'* VB.NET port by Ray Frankulin <random0000@cox.net>
'*
'* You may download the latest version from http://www.dataman.ro/sntp
'* If you find this class useful and would like to support my existence, please have a
'* look at my Amazon wish list at
'* http://www.amazon.com/exec/obidos/wishlist/ref=pd_wt_3/103-6370142-9973408
'* or make a donation to my Delta Forth .NET project, at
'* http://shareit1.element5.com/product.html?productid=159082&languageid=1&stylefrom=159082&backlink=http%3A%2F%2Fwww.dataman.ro&currencies=USD
'*
'* Last modified: October 3, 2005
'*
'* Permission is hereby granted, free of charge, to any person obtaining a
'* copy of this software and associated documentation files (the
'* "Software"), to deal in the Software without restriction, including
'* without limitation the rights to use, copy, modify, merge, publish,
'* distribute, and/or sell copies of the Software, and to permit persons
'* to whom the Software is furnished to do so, provided that the above
'* copyright notice(s) and this permission notice appear in all copies of
'* the Software and that both the above copyright notice(s) and this
'* permission notice appear in supporting documentation.
'*
'* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
'* OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
'* MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT
'* OF THIRD PARTY RIGHTS. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
'* HOLDERS INCLUDED IN THIS NOTICE BE LIABLE FOR ANY CLAIM, OR ANY SPECIAL
'* INDIRECT OR CONSEQUENTIAL DAMAGES, OR ANY DAMAGES WHATSOEVER RESULTING
'* FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT,
'* NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION
'* WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
'*
'* Disclaimer
'* ----------
'* Although reasonable care has been taken to ensure the correctness of this
'* implementation, this code should never be used in any application without
'* proper verification and testing. I disclaim all liability and responsibility
'* to any person or entity with respect to any loss or damage caused, or alleged
'* to be caused, directly or indirectly, by the use of this SNTPClient class.
'*
'* Comments, bugs and suggestions are welcome.
'*
'* Update history:
'*
'* October 3, 2005
'* - Translated into VB6 by Lasse
'*
'* September 20, 2003
'* - Renamed the class from NTPClient to SNTPClient.
'* - Fixed the RoundTripDelay and LocalClockOffset properties.
'*   Thanks go to DNH <dnharris@csrlink.net>.
'* - Fixed the PollInterval property.
'*   Thanks go to Jim Hollenhorst <hollenho@attbi.com>.
'* - Changed the ReceptionTimestamp variable to mdatDestinationTimeStamp to follow the standard
'*   more closely.
'* - Precision property is now shown is seconds rather than milliseconds in the
'*   ToString method.
'*
'* May 28, 2002
'* - Fixed a bug in the Precision property and the SetTime function.
'*   Thanks go to Jim Hollenhorst <hollenho@attbi.com>.
'*
'* March 14, 2001
'* - First public release.
'*/

'Leap indicator field values
Private Enum LeapIndicator_
    NoWarning       '0 - No warning
    LastMinute61    '1 - Last minute has 61 seconds
    LastMinute59    '2 - Last minute has 59 seconds
    Alarm           '3 - Alarm condition (clock not synchronized)
End Enum

'Mode field values
Enum Mode_
    SymmetricActive     '1 - Symmetric active
    SymmetricPassive    '2 - Symmetric pasive
    Client              '3 - Client
    Server              '4 - Server
    Broadcast           '5 - Broadcast
    unknown             '0, 6, 7 - Reserved
End Enum

'Stratum field values
Enum Stratum_
    Unspecified         '0 - unspecified or unavailable
    PrimaryReference    '1 - primary reference (e.g. radio-clock)
    SecondaryReference  '2-15 - secondary reference (via NTP or SNTP)
    Reserved            '16-255 - reserved
End Enum

'/// <summary>
'/// SNTPClient is a VB.NET# class designed to connect to time servers on the Internet and
'/// fetch the current date and time. Optionally, it may update the time of the local system.
'/// The implementation of the protocol is based on the RFC 2030.
'///
'/// Public class members:
'///
'/// LeapIndicator - Warns of an impending leap second to be inserted/deleted in the last
'/// minute of the current day. (See the _LeapIndicator enum)
'///
'/// VersionNumber - Version number of the protocol (3 or 4).
'///
'/// Mode - Returns mode. (See the _Mode enum)
'///
'/// Stratum - Stratum of the clock. (See the _Stratum enum)
'///
'/// PollInterval - Maximum interval between successive messages
'///
'/// Precision - Precision of the clock
'///
'/// RootDelay - Round trip time to the primary reference source.
'///
'/// RootDispersion - Nominal error relative to the primary reference source.
'///
'/// ReferenceID - Reference identifier (either a 4 character string or an IP address).
'///
'/// ReferenceTimestamp - The time at which the clock was last set or corrected.
'///
'/// OriginateTimestamp - The time at which the request departed the client for the server.
'///
'/// ReceiveTimestamp - The time at which the request arrived at the server.
'///
'/// Transmit Timestamp - The time at which the reply departed the server for client.
'///
'/// RoundTripDelay - The time between the departure of request and arrival of reply.
'///
'/// LocalClockOffset - The offset of the local clock relative to the primary reference
'/// source.
'///
'/// Initialize - Sets up data structure and prepares for connection.
'///
'/// Connect - Connects to the time server and populates the data structure.
'///  It can also update the system time.
'///
'/// IsResponseValid - Returns true if received data is valid and if comes from
'/// a NTP-compliant time server.
'///
'/// ToString - Returns a string representation of the object.
'///
'/// -----------------------------------------------------------------------------
'/// Structure of the standard NTP header (as described in RFC 2030)
'///                       1                   2                   3
'///   0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |LI | VN  |Mode |    Stratum    |     Poll      |   Precision   |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                          Root Delay                           |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                       Root Dispersion                         |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                     Reference Identifier                      |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                                                               |
'///  |                   Reference Timestamp (64)                    |
'///  |                                                               |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                                                               |
'///  |                   Originate Timestamp (64)                    |
'///  |                                                               |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                                                               |
'///  |                    Receive Timestamp (64)                     |
'///  |                                                               |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                                                               |
'///  |                    Transmit Timestamp (64)                    |
'///  |                                                               |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                 Key Identifier (optional) (32)                |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///  |                                                               |
'///  |                                                               |
'///  |                 Message Digest (optional) (128)               |
'///  |                                                               |
'///  |                                                               |
'///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///
'/// -----------------------------------------------------------------------------
'///
'/// SNTP Timestamp Format (as described in RFC 2030)
'///                         1                   2                   3
'///     0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1
'/// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'/// |                           Seconds                             |
'/// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'/// |                  Seconds Fraction (0-padded)                  |
'/// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'///
'/// </summary>


'// NTP Data Structure Length
Private Const SNTPDataLength As Byte = 47

'// NTP Data Structure (as described in RFC 2030)
Private SNTPData(SNTPDataLength) As Byte

'// Offset constants for timestamps in the data structure
Private Const offReferenceID As Byte = 12
Private Const offReferenceTimestamp As Byte = 16
Private Const offOriginateTimestamp As Byte = 24
Private Const offReceiveTimestamp As Byte = 32
Private Const offTransmitTimestamp As Byte = 40

'// Destination Timestamp
Private mdatDestinationTimeStamp As Date

'ACHTUNG SYSTEMTIME IST UTC
Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Const mlMS2VBDate As Long = 86400000

Private mlTimeOut As Long
Private msTimeOut As String
Private msTimeServer As String

Private Property Get TimeOut() As Long
    
    If mlTimeOut < 1 Then mlTimeOut = 1000
    TimeOut = mlTimeOut
    
End Property

Public Property Let TimeOut(lTimeOut As Long)
    mlTimeOut = lTimeOut
End Property

Public Property Get TimeServer() As String

    TimeServer = msTimeServer

End Property

Public Property Let TimeServer(sTimeServer As String)

    msTimeServer = sTimeServer

End Property

Public Property Get LastError() As String

    LastError = msTimeOut

End Property

Public Property Get LastLapse() As Long

    LastLapse = LocalClockOffset()

End Property

'Leap Indicator
Private Property Get LeapIndicator() As LeapIndicator_
    'Isolate the two most significant bits
    Dim bytVal As Byte
    
    bytVal = CByte(SNTPData(0) / 64)
    Select Case bytVal
        Case 0: LeapIndicator = LeapIndicator_.NoWarning
        Case 1: LeapIndicator = LeapIndicator_.LastMinute61
        Case 2: LeapIndicator = LeapIndicator_.LastMinute59
        Case 3: LeapIndicator = LeapIndicator_.Alarm
        Case Else: LeapIndicator = LeapIndicator_.Alarm
    End Select
    
End Property

' Version Number
Private Property Get VersionNumber() As Byte
    'Isolate bits 3 - 5
    VersionNumber = (SNTPData(0) And &H38) / 8
End Property

Private Property Get Mode() As Mode_
    
    'Isolate bits 0 - 3
    Dim bytVal As Byte
    
    bytVal = CByte(SNTPData(0) And &H7)
    Select Case bytVal
        Case 0, 6, 7
            Mode = Mode_.unknown
        Case 1
            Mode = Mode_.SymmetricActive
        Case 2
            Mode = Mode_.SymmetricPassive
        Case 3
            Mode = Mode_.Client
        Case 4
            Mode = Mode_.Server
        Case 5
            Mode = Mode_.Broadcast
        Case Else
    End Select
    
End Property

'Stratum
Private Property Get Stratum() As Stratum_

    Dim bytVal As Byte
    
    bytVal = CByte(SNTPData(1))
    
    If (bytVal = 0) Then
        Stratum = Stratum_.Unspecified
    ElseIf (bytVal = 1) Then
        Stratum = Stratum_.PrimaryReference
    ElseIf (bytVal <= 15) Then
        Stratum = Stratum_.SecondaryReference
    Else
        Stratum = Stratum_.Reserved
    End If
    
End Property

'Poll Interval
Private Property Get PollInterval() As Long
    '// Thanks to Jim Hollenhorst <hollenho@attbi.com>
    PollInterval = 2 ^ SNTPData(2)
End Property

'Precision (in milliseconds)
Private Property Get Precision() As Double
    Precision = 2 ^ SNTPData(3)
End Property

'Root Delay (in milliseconds)
Private Property Get RootDelay() As Double

    Dim curTmp As Currency
    
    curTmp = 256 * (256 * (256 * SNTPData(4) + SNTPData(5)) + SNTPData(6)) + SNTPData(7)
    RootDelay = 1000 * ((curTmp) / &H10000)
    
End Property

'Root Dispersion (in milliseconds)
Private Property Get RootDispersion() As Double

    Dim curTmp As Currency
    
    curTmp = 256 * (256 * (256 * SNTPData(8) + SNTPData(9)) + SNTPData(10)) + SNTPData(11)
    RootDispersion = 1000 * ((curTmp) / &H10000)
    
End Property

'Reference Identifier
Private Property Get ReferenceID() As String
    
    Dim sTmp As String, fTime As Double, fOffset As Double
    
    Select Case Stratum
        Case Stratum_.PrimaryReference Or Stratum_.Unspecified
            If SNTPData(offReferenceID + 0) <> 0 Then sTmp = sTmp & Chr(SNTPData(offReferenceID + 0))
            If SNTPData(offReferenceID + 1) <> 0 Then sTmp = sTmp & Chr(SNTPData(offReferenceID + 1))
            If SNTPData(offReferenceID + 2) <> 0 Then sTmp = sTmp & Chr(SNTPData(offReferenceID + 2))
            If SNTPData(offReferenceID + 3) <> 0 Then sTmp = sTmp & Chr(SNTPData(offReferenceID + 3))
        Case Stratum_.SecondaryReference
            Select Case VersionNumber
                Case 3 '// Version 3, Reference ID is an IPv4 address
                    sTmp = SNTPData(offReferenceID + 0) & "." & SNTPData(offReferenceID + 1) & "." & SNTPData(offReferenceID + 2) & "." & SNTPData(offReferenceID + 3)
                Case 4 '// Version 4, Reference ID is the timestamp of last update
                    fTime = ComputeDate(GetMilliSeconds(offReferenceID))
                    '// Take care of the time zone
                    fOffset = GetUTCOffset()
                    sTmp = DateAdd("n", 60 * fOffset, fTime)
                Case Else
                    sTmp = "N/A"
            End Select
    End Select
    ReferenceID = sTmp
    
End Property

'// Reference Timestamp
Private Property Get ReferenceTimestamp() As Date

    Dim datTime As Date
    Dim fOffSpan As Double
    
    datTime = ComputeDate(GetMilliSeconds(offReferenceTimestamp))
    
    '// Take care of the time zone
    fOffSpan = GetUTCOffset()
    ReferenceTimestamp = datTime + fOffSpan / 24
    
End Property

'// Originate Timestamp
Private Property Get OriginateTimestamp() As Date
    OriginateTimestamp = ComputeDate(GetMilliSeconds(offOriginateTimestamp))
End Property

'// Receive Timestamp
Private Property Get ReceiveTimestamp() As Date

    Dim datTime As Date
    Dim fOffSpan As Double
    
    datTime = ComputeDate(GetMilliSeconds(offReceiveTimestamp))
    
    'Take care of the time zone
    fOffSpan = GetUTCOffset()
    ReceiveTimestamp = datTime + fOffSpan / 24
    
End Property

'// Transmit Timestamp
Private Property Get TransmitTimestamp() As Date
        
    Dim datTime As Date
    Dim fOffSpan As Double
    
    datTime = ComputeDate(GetMilliSeconds(offTransmitTimestamp))
    'Take care of the time zone
    fOffSpan = GetUTCOffset()
    TransmitTimestamp = datTime + fOffSpan / 24
    
End Property

Private Property Let TransmitTimestamp(ByVal datValue As Date)
    Call SetDate(offTransmitTimestamp, datValue)
End Property

'// Round trip delay (in milliseconds)
Private Property Get RoundTripDelay() As Currency

    '// Thanks to DNH <dnharris@csrlink.net>
    Dim fSpan As Double
    
    fSpan = mdatDestinationTimeStamp - OriginateTimestamp - (ReceiveTimestamp - TransmitTimestamp)
    RoundTripDelay = fSpan * mlMS2VBDate
    
End Property

'// Local clock offset (in milliseconds)
Private Property Get LocalClockOffset() As Currency
    
    '// Thanks to DNH <dnharris@csrlink.net>
    Dim fSpan As Double
    
    fSpan = ReceiveTimestamp - OriginateTimestamp + (TransmitTimestamp - mdatDestinationTimeStamp)
    LocalClockOffset = fSpan * (mlMS2VBDate / 2)
    
End Property

'// Compute date, given the number of milliseconds since January 1, 1900
Private Function ComputeDate(ByVal curMilliseconds As Currency) As Date

    Dim fSpan As Double
    Dim datTime As Date
    
    fSpan = CDbl(curMilliseconds / mlMS2VBDate)
    
    datTime = #1/1/1900#
    datTime = datTime + fSpan
    ComputeDate = datTime
    
End Function

'// Compute the number of milliseconds, given the Offset of a 8-byte array
Private Function GetMilliSeconds(ByVal bytOffset As Byte) As Currency
        
    Dim intPart As Currency, fractPart As Currency
    Dim i As Long
    
    For i = 0 To 3
        intPart = 256 * intPart + SNTPData(bytOffset + i)
    Next
    
    For i = 4 To 7
        fractPart = 256 * fractPart + SNTPData(bytOffset + i)
    Next
    
    GetMilliSeconds = Int(intPart * 1000 + (fractPart / 4294967.296)) '* 1000 / &H80000000 / &H80000000)
    
End Function

'// Compute the 8-byte array, given the date
Private Sub SetDate(ByVal bytOffset As Byte, ByVal datValue As Date)

    Dim i As Integer, intPart As Currency, fractPart As Currency
    Dim datStartOfCentury As Date
    Dim curMilliseconds As Currency
    Dim curTmp As Currency
    
    datStartOfCentury = #1/1/1900#
    
    curMilliseconds = Int((datValue - datStartOfCentury) * mlMS2VBDate)
    intPart = Int(curMilliseconds / 1000)
    fractPart = Int(CurMod(curMilliseconds, 1000) * 4294967.296) '* &H80000000 * &H80000000) / 1000
    
    curTmp = intPart
    
    For i = 3 To 0 Step -1
        SNTPData(bytOffset + i) = Int(CurMod(curTmp, 256))
        curTmp = Int(curTmp / 256)
    Next
    
    curTmp = fractPart
    
    For i = 7 To 4 Step -1
        SNTPData(bytOffset + i) = Int(CurMod(curTmp, 256))
        curTmp = Int(curTmp / 256)
    Next
    
End Sub

Private Function GetTimeUTC() As Date

    Dim ST As SYSTEMTIME
    
    GetSystemTime ST
    With ST
      GetTimeUTC = myDateSerial(.wYear, .wMonth, .wDay) + myTimeSerial(.wHour, .wMinute, .wSecond) + .wMilliseconds / mlMS2VBDate
    End With
    
End Function

Private Function GetTimeLocal() As Date

    Dim ST As SYSTEMTIME
    
    GetLocalTime ST
    With ST
      GetTimeLocal = myDateSerial(.wYear, .wMonth, .wDay) + myTimeSerial(.wHour, .wMinute, .wSecond) + .wMilliseconds / mlMS2VBDate
    End With
    
End Function

'// Initialize the NTPClient data
Private Sub Initialize()
    
    Dim i As Long
    
    'Set version number to 4 and Mode to 3 (client)
    SNTPData(0) = &H1B
    
    'Initialize all other fields with 0
    For i = 1 To 47
        SNTPData(i) = 0
    Next
    'Initialize the transmit timestamp
    TransmitTimestamp = GetTimeLocal()
    
End Sub

'// Connect to the time server and update system time
Public Function SyncTime() As Boolean
        
    msTimeOut = ""
    If Len(TimeServer) = 0 Then
        msTimeOut = "No server set"
    Else
    
        With wscWinsock
            .Close
            .LocalPort = 0 ' wird automatisch vergeben
            .RemotePort = 123
            .RemoteHost = TimeServer
            .protocol = sckUDPProtocol
        End With
        
        Call Initialize
        
        On Error Resume Next
        Call wscWinsock.SendData(SNTPData())
        If Err.Number <> 0 Then msTimeOut = Err.Description
        On Error GoTo 0
        
        If msTimeOut = "" Then
        
            TimerTimeout.Interval = TimeOut
            TimerTimeout.Enabled = True
            
            Do While TimerTimeout.Enabled
                DoEvents
            Loop
            
            If IsResponseValid() Then
                If SetTime() Then
                    SyncTime = True
                Else
                    msTimeOut = "No permission to set time"
                End If
            Else
                If msTimeOut = "" Then
                    msTimeOut = "Invalid server response"
                End If
            End If
        
        End If
    End If
End Function

'// Check if the response from server is valid
Private Function IsResponseValid() As Boolean

    If (Mode <> Mode_.Server) Then
        IsResponseValid = False
    Else
        IsResponseValid = True
    End If
    
End Function

'// Set system time according to transmit timestamp
Private Function SetTime() As Boolean

    Dim ST As SYSTEMTIME
    Dim datTrTs As Date
    Dim lMs As Long
    Dim fOffSpan As Double
    
    'Zeit holen & Korrektur addieren
    datTrTs = GetTimeUTC()
    datTrTs = datTrTs + LocalClockOffset() / mlMS2VBDate
    
    'Millisekunden rausrechnen, sonst wird falsch gerundet
    lMs = Int((datTrTs - Int(datTrTs)) * mlMS2VBDate) Mod 1000
    datTrTs = datTrTs - lMs / mlMS2VBDate
    
    'call DebugPrint( "LocalClockOffset: " & LocalClockOffset())
    
    With ST
      .wYear = Year(datTrTs)
      .wMonth = Month(datTrTs)
      .wDay = Day(datTrTs)
      .wHour = Hour(datTrTs)
      .wMinute = Minute(datTrTs)
      .wSecond = Second(datTrTs)
      .wMilliseconds = lMs
    End With
    
    If SetSystemTime(ST) Then SetTime = True
    
End Function

Private Function GetUTCOffset() As Double

  GetUTCOffset = DateDiff("n", GetTimeUTC(), GetTimeLocal()) / 60
  
End Function

Private Sub TimerTimeout_Timer()

  msTimeOut = "Timeout"
  TimerTimeout.Enabled = False

End Sub

Private Sub wscWinsock_DataArrival(ByVal bytesTotal As Long)

    Dim i As Integer
    Dim sData As String
    
    On Error Resume Next
    wscWinsock.GetData sData
    If Err.Number <> 0 Then msTimeOut = Err.Description
    On Error GoTo 0
    
    mdatDestinationTimeStamp = GetTimeLocal()
    
    If Len(sData) = UBound(SNTPData()) - LBound(SNTPData()) + 1 Then
      For i = 1 To Len(sData)
        SNTPData(i - 1 + LBound(SNTPData())) = Asc(Mid(sData, i, 1))
      Next
    End If
    
    TimerTimeout.Enabled = False
    
'    If IsResponseValid() Then
'        DebugPrint "Delta: " & LocalClockOffset()
'        SetTime
'    Else
'        DebugPrint "Invalid response from " & wscWinsock.RemoteHost
'    End If
                
End Sub

Private Function CurMod(ByVal curValue As Currency, ByVal lDiv As Long) As Currency

    CurMod = curValue - Int(curValue / lDiv) * lDiv

End Function

Private Function GetPrecisionDate(ByVal datValue As Date) As String

    GetPrecisionDate = Format$(datValue, "Short Date") & " " & Format(datValue, "Long Time") & " " & Int((datValue - Int(datValue)) * mlMS2VBDate) Mod 1000

End Function

' wir brauchen diese 'nachgebauten' Funktionen, weil die Originale unter WINE (Linux) buggy sind!
Private Function myDateSerial(ByVal iYear As Integer, ByVal iMonth As Integer, ByVal iDay As Integer) As Date

  myDateSerial = #1/1/2000#
  myDateSerial = DateAdd("yyyy", iYear - 2000, myDateSerial)
  myDateSerial = DateAdd("m", iMonth - 1, myDateSerial)
  myDateSerial = DateAdd("d", iDay - 1, myDateSerial)

End Function

Private Function myTimeSerial(ByVal iHour As Integer, ByVal iMinute As Integer, ByVal iSecond As Integer) As Date

  myTimeSerial = DateAdd("h", iHour, myTimeSerial)
  myTimeSerial = DateAdd("n", iMinute, myTimeSerial)
  myTimeSerial = DateAdd("s", iSecond, myTimeSerial)

End Function

