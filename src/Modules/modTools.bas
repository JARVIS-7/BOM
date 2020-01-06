Attribute VB_Name = "modTools"
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
' gesammelte Werke allgemeiner Natur ..
' teils aus dem Net
' muss mal sortiert und aufgeteilt werden :-)
'


' Sound abspielen
Private Declare Function sndPlaySoundA Lib "winmm.dll" ( _
    ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'wie viele Millisekunden läuft Windows
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'Shell- Calls
Public Declare Function ShellExecute Lib "shell32.dll" _
                    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal _
                    lpOperation As String, ByVal lpFile As String, ByVal _
                    lpParameters As String, ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long
                    
'Private Declare Function GetWindowsDirectory Lib "kernel32" _
                    Alias "GetWindowsDirectoryA" (ByVal lpBuffer As _
                    String, ByVal nSize As Long) As Long

'Const SW_RESTORE As Long = &H9&

Private Declare Function GetEnvironmentVariable Lib _
                               "kernel32.dll" Alias "GetEnvironmentVariableA" (ByVal _
                               lpName As String, ByVal lpBuffer As String, ByVal _
                               nSize As Long) As Long

'Private Declare Function SetEnvironmentVariable Lib _
                               "kernel32.dll" Alias "SetEnvironmentVariableA" (ByVal _
                               lpName As String, ByVal lpValue As String) As Long
                               
'Versionsabfrage
Private Declare Function GetVersionEx Lib "kernel32" Alias _
                   "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) _
                   As Long
                   
'Tempverzeichnis
Private Declare Function GetTempPath Lib "kernel32" Alias _
    "GetTempPathA" (ByVal nBufferLength As Long, ByVal _
    lpBuffer As String) As Long
                   
'ShutDown
Private Type LUID
  UsedPart As Long
  IgnoredForNowHigh32BitPart As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  TheLuid As LUID
  Attributes As Long
End Type

' Konstante für ExitWindowsEx:

' Prozesse des Benutzers werden beendet, dann der User
' ausgeloggt
'Private Const EWX_LOGOFF = 0

' Prozesse des Benutzers werden beendet, dann das System
' heruntergefahren.
'Private Const EWX_SHUTDOWN = 1

' Prozesse des Benutzers werden beendet, dann das System neu
' gestartet.
'Private Const EWX_REBOOT = 2

' Prozesse des Benutzers werden ohne Rückfrage beendet
'Private Const EWX_FORCE = 4

' Schaltet (bei Hardware-Unterstützung dieses Features) nach
' dem Herunterfahren die Stromversorgung auf StandBy
'Private Const EWX_POWEROFF = 8

' Windows 2000: Prozessen des Benutzers wird eine Aufforderung
' zur Beendigung gesendet. Wenn sie nicht reagieren, werden sie
' gezwungen beendet.
'Private Const EWX_FORCEIFHUNG = 16

'als ENums:
Public Enum ShutDownActionsEnum
  [saShutdown] = 1&
  [saReboot] = 2&
  [saLogOff] = 4&
  [saPowerOff] = 8&
  [saForceIfHung] = 16&
End Enum

'Public Enum ForceModes
'  NoForce = 0
'  Force = 1
'  ForceIfHung_Win2K = 2
'End Enum

                                    
Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
        dwOptions As Long, ByVal dwReserved As Long) As Long
        
Private Declare Function GetCurrentProcess Lib "kernel32" () _
        As Long
        
Private Declare Function OpenProcessToken Lib "advapi32" ( _
        ByVal ProcessHandle As Long, ByVal DesiredAccess As _
        Long, TokenHandle As Long) As Long
        
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
        Alias "LookupPrivilegeValueA" (ByVal lpSystemName As _
        String, ByVal lpName As String, lpLuid As LUID) As Long
        
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
        (ByVal TokenHandle As Long, ByVal DisableAllPrivileges _
        As Long, NewState As TOKEN_PRIVILEGES, ByVal _
        BufferLength As Long, PreviousState As _
        TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal _
        hObject As Long) As Long
        

                   
' Fenster on Top halten
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
                     
Private Declare Function SetWindowPos Lib "user32" (ByVal _
                            hWnd As Long, ByVal hWndInsertAfter As Long, ByVal _
                            X As Long, ByVal Y As Long, ByVal cx As Long, ByVal _
                            cy As Long, ByVal wFlags As Long) As Long
                                               
Private Type OSVERSIONINFO
             dwOSVersionInfoSize As Long
             dwMajorVersion As Long
             dwMinorVersion As Long
             dwBuildNumber As Long
             dwPlatformId As Long
             szCSDVersion As String * 128
End Type

' 1.8.0 Zugriffe auf INI- Dateien
Private Declare Function WritePrivateProfileString Lib _
        "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal _
        lpKeyName As Any, ByVal lpString As Any, ByVal _
        lpFileName As String) As Long
        
Private Declare Function GetPrivateProfileString Lib _
        "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal _
        lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize _
        As Long, ByVal lpFileName As String) As Long

'Private Declare Function WritePrivateProfileSection Lib _
        "kernel32" Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpString As _
        String, ByVal lpFileName As String) As Long
        
'Private Declare Function GetPrivateProfileSection Lib _
        "kernel32" Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpReturnedString _
        As String, ByVal nSize As Long, ByVal lpFileName _
        As String) As Long

Private Declare Function FindWindow Lib "user32.dll" Alias _
                        "FindWindowA" (ByVal lpClassName As String, _
                                       ByVal lpWindowName As String) As Long
                               
'Datentyp Rechteck (Koordinaten)
Private Type udtRectX
    rX1 As Long 'links
    rY1 As Long  'oben
    rX2 As Long  'rechts
    rY2 As Long   'unten
End Type

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As udtRectX) As Long


'Public Const SM_XVIRTUALSCREEN = 76    'virtual desktop left
'Public Const SM_YVIRTUALSCREEN = 77    'virtual top
Private Const SM_CXVIRTUALSCREEN As Long = 78  'virtual width
Private Const SM_CYVIRTUALSCREEN As Long = 79  'virtual height
 
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Constants that will be used in the API functions
'Public Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE As Long = -11&

' Declare the needed API functions
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
'Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
'Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, ByRef lpNumberOfCharsWritten As Long, lpReserved As Long) As Long

Public Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
                         ByVal CodePage As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpWideCharStr As Long, _
                         ByVal cchWideChar As Long, _
                         ByVal lpMultiByteStr As Long, _
                         ByVal cbMultiByte As Long, _
                         ByVal lpDefaultChar As Long, _
                         ByVal lpUsedDefaultChar As Long) As Long
                         
Public Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
                         ByVal CodePage As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpMultiByteStr As Long, _
                         ByVal cbMultiByte As Long, _
                         ByVal lpWideCharStr As Long, _
                         ByVal cchWideChar As Long) As Long
                         
Public Const CP_UTF8 As Long = 65001

Public Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type


'Windows API function declarations

'Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
    lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Const STD_INPUT_HANDLE = -10&


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

' Special Folders
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Type SHITEMID
  cb As Long
  abID As Byte
End Type
Public Type ITEMIDLIST
  mkid As SHITEMID
End Type
Public Enum spfSpecialFolderConstants
  spfAppdata = &H1A
  spfLOCAL_APPDATA = &H1C
End Enum

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Const WM_SETICON As Long = &H80

Public Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
    (ByVal hInst As Long, ByVal lpsz As String, _
    ByVal iType As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal fOptions As Long) As Long
' iType options:
'Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON As Long = 1&
'Private Const IMAGE_CURSOR = 2
' fOptions flags:
Private Const LR_LOADMAP3DCOLORS As Long = &H1000
Private Const LR_LOADFROMFILE As Long = &H10
'Private Const LR_LOADTRANSPARENT = &H20

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Public Sub SetIcon(lHwnd As Long, vntHIcon As Variant)
    
    On Error Resume Next
    
    Set frmDummy.Icon = vntHIcon
    Call SendMessage(lHwnd, WM_SETICON, 0, vntHIcon)
    
End Sub

Public Sub SetFont(oFrm As Form)
    
    On Error Resume Next
    Dim i As Integer
    
    
    For i = 0 To oFrm.Controls.Count - 1
        With oFrm.Controls(i)
            .Font.Size = giDefaultFontSize
            '.FontSize = giDefaultFontSize
            '.HeadFont.Size = giDefaultFontSize
            .Font.Name = gsGlobFontName
            '.FontName = gsGlobFontName
        End With
    Next i
    
    With oFrm
        .Font.Size = giDefaultFontSize
        '.FontSize = giDefaultFontSize
        '.HeadFont.Size = giDefaultFontSize
        .Font.Name = gsGlobFontName
        '.FontName = gsGlobFontName
    End With
    oFrm.PanelRepaint

End Sub



Public Function ExecuteDoc(lHwnd As Long, sDocFilename As String, Optional sParam As String, Optional bQuiet As Boolean = False) As Boolean
    
    Dim sDefBrowser As String
    Dim lPosStart As Long
    Dim lPosEnd As Long
    
    'mal sehen ob es ohne was geht ..
    On Error GoTo errhdl
    
    If gbBrowseInNewWindow And (sDocFilename Like "http://*" Or sDocFilename Like "*.htm" Or sDocFilename Like "*.html") Then
        
        If modRegistry.GetValue(HKEY_CLASSES_ROOT, "http\shell\open\command", "", sDefBrowser) Then
            lPosStart = InStr(1, sDefBrowser, ":\") - 1
            lPosEnd = InStr(lPosStart + 1, LCase(sDefBrowser), ".exe") + 4
            If lPosStart > 0 And lPosEnd > lPosStart Then
                sDefBrowser = Mid(sDefBrowser, lPosStart, lPosEnd - lPosStart)
            Else
                lPosStart = InStrRev(sDefBrowser, "\")
                If lPosStart > 0 Then
                    lPosEnd = InStr(lPosStart, sDefBrowser, " ")
                    If lPosEnd > 0 Then
                        sDefBrowser = Left(sDefBrowser, lPosEnd - 1)
                    End If
                End If
            End If
        End If
        ExecuteDoc = 32 < ShellExecute(lHwnd, "open", sDefBrowser, """" & sDocFilename & """", App.Path, 1)
    End If
    
    If Not ExecuteDoc Then ExecuteDoc = 32 < ShellExecute(lHwnd, "open", sDocFilename, sParam, App.Path, 1)
    
    If Not ExecuteDoc And Not bQuiet Then MsgBox Replace(gsarrLangTxt(29), "%FILE%", sDocFilename, vbTextCompare), vbExclamation
    
Done:
On Error GoTo 0
Exit Function
errhdl:
Resume Done
End Function

Public Sub QuickSortDate(tarrFeld() As udtArtikelZeile, ByVal LB As Long, ByVal UB As Long, bAscending As Boolean)
    
    Dim p1 As Long, p2 As Long
    Dim datRef As Date
    Dim tTmp As udtArtikelZeile
    
    p1 = LB
    p2 = UB
    datRef = tarrFeld((p1 + p2) / 2).EndeZeit
                 
    Do
        If bAscending Then
            Do While (tarrFeld(p1).EndeZeit < datRef): p1 = p1 + 1: Loop
            Do While (tarrFeld(p2).EndeZeit > datRef): p2 = p2 - 1: Loop
        Else
            Do While (tarrFeld(p1).EndeZeit > datRef): p1 = p1 + 1: Loop
            Do While (tarrFeld(p2).EndeZeit < datRef): p2 = p2 - 1: Loop
        End If
        
        If p1 <= p2 Then
            tTmp = tarrFeld(p1)
            tarrFeld(p1) = tarrFeld(p2)
            tarrFeld(p2) = tTmp
                        
            p1 = p1 + 1
            p2 = p2 - 1
        End If
    Loop Until (p1 > p2)
    
    If LB < p2 Then Call QuickSortDate(tarrFeld(), LB, p2, bAscending)
    If p1 < UB Then Call QuickSortDate(tarrFeld(), p1, UB, bAscending)
    
End Sub


Private Function GetEnv(sEnvName As String) As String
    
    'MD-Marker , Function wird nicht aufgerufen
    '
    'Dim Buffer As String, l As Long
    '
    'l = 256
    'Buffer = String$(l, Chr$(0))
    '
    'l = GetEnvironmentVariable(EnvName, Buffer, l)
    '
    'If l <> 0 Then
        'GetEnv = Left(Buffer, l)
    'Else
        'GetEnv = ""
    'End If
    '
End Function

Public Function Str2Date(ByVal sDateString As String, ByVal sDateFormat As String) As Date
    
    On Error Resume Next
    Dim lDayVal As Long
    Dim lMonthVal As Long
    Dim lYearVal As Long
    Dim lHourVal As Long
    Dim lMinuteVal As Long
    Dim lSecondVal As Long
    Dim sTmp As String
    
    
    Do While sDateFormat > "" And sDateString > ""
    
        'DebugPrint "Str2Date,DateString: " & sDateString
        'DebugPrint "Str2Date,DateFormat: " & sDateFormat
        
        Do While sDateString > "" And Not IsNumeric(Right(sDateString, 1))
            sDateString = Left(sDateString, Len(sDateString) - 1)
        Loop
        
        sTmp = ""
        Do While sDateString > "" And IsNumeric(Right(sDateString, 1))
            sTmp = Right(sDateString, 1) & sTmp
            sDateString = Left(sDateString, Len(sDateString) - 1)
        Loop
        
        'DebugPrint "Str2Date,tmp: " & sTmp
        
        Select Case Right(sDateFormat, 1)
            Case "D": lDayVal = Val(sTmp)
            Case "M": lMonthVal = Val(sTmp)
            Case "Y": lYearVal = Val(sTmp)
            Case "h": lHourVal = Val(sTmp)
            Case "m": lMinuteVal = Val(sTmp)
            Case "s": lSecondVal = Val(sTmp)
        End Select
        
        sDateFormat = Left(sDateFormat, Len(sDateFormat) - 1)
        
    Loop
    
    If lYearVal < 100 Then lYearVal = lYearVal + 2000
    
    Str2Date = myDateSerial(lYearVal, lMonthVal, lDayVal) + myTimeSerial(lHourVal, lMinuteVal, lSecondVal)
    'DebugPrint "Str2Date,Str2Date: " & Date2Str(Str2Date)
    
End Function

Public Function Date2Str(ByVal datDateValue As Date, Optional ByVal bShowWeekday As Boolean = False, Optional ByVal sSpecialDateFormat As String = "") As String
    
    'Formatiert in lokalem Format über "Short Date" und "Long Time"
    
    On Error Resume Next
    
    If sSpecialDateFormat > "" Then
        Date2Str = Format$(datDateValue, sSpecialDateFormat)
    ElseIf bShowWeekday Then
        Date2Str = Format$(datDateValue, "ddd") & ", " & Format$(datDateValue, "Short Date") & " " & Format$(datDateValue, "Long Time")
    Else
        Date2Str = Format$(datDateValue, "Short Date") & " " & Format$(datDateValue, "Long Time")
    End If
    
End Function

Public Function GetWinVersion() As String
    
    On Error GoTo errhdl
    
    Dim sOSString As String
    Dim OSVersion As OSVERSIONINFO
    Dim lBuildNr As Long
               
    sOSString = "Unbekanntes Betriebssystem"
    
    OSVersion.dwOSVersionInfoSize = Len(OSVersion)
    Call GetVersionEx(OSVersion)
    
    With OSVersion
        If (.dwBuildNumber And &HFFFF&) > &H7FFF Then
            lBuildNr = (.dwBuildNumber And &HFFFF&) - &H10000
        Else
            lBuildNr = .dwBuildNumber And &HFFFF&
        End If
               
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
        
            If .dwMajorVersion <= 4 Then
                sOSString = "Windows NT " & .dwMajorVersion
            ElseIf .dwMajorVersion = 5 Then
                If .dwMinorVersion = 1 Then
                    sOSString = "Windows XP"
                Else
                    sOSString = "Windows 2000"
                End If
            End If
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            If (.dwMajorVersion > 4) Or (.dwMajorVersion = 4 And .dwMinorVersion = 10) Then
                If lBuildNr = 1998 Then
                    sOSString = "Windows 98"
                Else
                    sOSString = "Windows 98 SE"
                End If
            ElseIf (.dwMajorVersion = 4 And .dwMinorVersion = 0) Then
                sOSString = "Windows 95"
            ElseIf (.dwMajorVersion = 4 And .dwMinorVersion = 90) Then
                sOSString = "Windows ME"
            End If
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32s Then
            sOSString = "Windows 32s"
        End If
        
        gsWinVersion = sOSString
        
End With
                 
errhdl:

GetWinVersion = sOSString

End Function

Private Function IsWinNT() As Boolean
    
    Dim tOSVERSIONINFO As OSVERSIONINFO
    
    tOSVERSIONINFO.dwOSVersionInfoSize = Len(tOSVERSIONINFO)
    
    Call GetVersionEx(tOSVERSIONINFO)
    
    IsWinNT = CBool(tOSVERSIONINFO.dwPlatformId And VER_PLATFORM_WIN32_NT)
    
End Function

Public Sub SetForeground(bSetIt As Boolean, lHandle As Long)
    
    If bSetIt Then
        Call SetWindowPos(lHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Call SetWindowPos(lHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
    
End Sub

Public Sub ShutDownWin()

    Dim ShutdownFlags As Long
    
    gbExplicitEnd = True
    
    ShutdownFlags = [saShutdown] Or [saPowerOff]
    
    If IsWinNT() Then    'Priviligien setzen
        Call SetShutdownPrivilege
    End If
    
    If gsWinVersion = "Windows 2000" Or gsWinVersion = "Windows XP" Then
        ShutdownFlags = ShutdownFlags Or [saForceIfHung]
    Else
        ShutdownFlags = ShutdownFlags Or [saLogOff]
    End If
    
    Call ExitWindowsEx(ShutdownFlags, &HFFFF)
    
End Sub

Private Sub SetShutdownPrivilege()

    Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
    Const TOKEN_QUERY As Long = &H8
    Const SE_PRIVILEGE_ENABLED As Long = &H2
    
    Dim hProcessHandle As Long
    Dim hTokenHandle As Long
    Dim PrivLUID As LUID
    Dim TokenPriv As TOKEN_PRIVILEGES
    Dim tkpDummy As TOKEN_PRIVILEGES
    Dim lDummy As Long
    
    'Ermittlung eines Prozess-Handles dieser Anwendung
    hProcessHandle = GetCurrentProcess()
    
    'Für unseren Prozess soll ein Token geändert werden.
    OpenProcessToken hProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
        TOKEN_QUERY), hTokenHandle
    
    'Die repräsentierende LUID für das "SeShutdownPrivilege" ermitteln
    Call LookupPrivilegeValue("", "SeShutdownPrivilege", PrivLUID)
    
    'Vorbereitungen auf das Ändern des Tokens
    With TokenPriv
        'Anzahl der Privilegien
        .PrivilegeCount = 1
        
        'LUID-Struktur für das Privileg
        .TheLuid = PrivLUID
        
        'Das Privileg soll gesetzt werden
        .Attributes = SE_PRIVILEGE_ENABLED
    End With
    
    'Jetzt wird das Token für diesen Prozess gesetzt, um
    'unserem Prozess das Recht für ein Herunterfahren / einen
    'Neustart zuzuteilen:
    Call AdjustTokenPrivileges(hTokenHandle, False, TokenPriv, _
        Len(tkpDummy), tkpDummy, lDummy)
    
    'Handle auf das geoeffnete Token freigeben
    Call CloseHandle(hTokenHandle)
    
End Sub

Public Sub ShowBrowser(lHwnd As Long)
    
    'MD-Marker 20090325 , frmbrowser wurde aus dem Projekt _
    entfernt , damit kann die If - Abfrage gelöscht werden. _
    Im zweiten Schritt wird die kpl. Sub gelöscht
    
    Call ExecuteDoc(lHwnd, gsGlobalUrl)
    
    'If gbUseIntBrowser Then
    '    frmBrowser.Show
    'Else
    '    Call ExecuteDoc(lHwnd, gsGlobalUrl)
    'End If
    
End Sub

Public Function RequestSema(tSema As Sema) As Boolean
'Sema geht nicht, also nur OK oder nicht

'Exit Sub

If Not tSema.is_requested Then
    tSema.is_requested = True
    RequestSema = True
    'DebugPrint tsema.sema_name & " requested"

Else
    RequestSema = False
    'DebugPrint tsema.sema_name & " request abgelehnt"
End If

End Function

Public Sub ReleaseSema(tSemaIn As Sema)

tSemaIn.is_requested = False
'DebugPrint tSemaIn.sema_name & " released"
End Sub

Public Sub InitSemas()

    gtBrowserSema.sema_Name = "BROWSER_Sema"
    gtTcpSema.sema_Name = "TCP_Sema"
    
End Sub

Public Sub SaveWindowSettings(Optional oIni As clsIni)
    
    Dim bCalledBySaveSettings As Boolean
  
  
    If oIni Is Nothing Then
        Set oIni = New clsIni
    Else
        bCalledBySaveSettings = True
    End If
    
    If Not bCalledBySaveSettings Then
        If Not oIni.ReadIni(gsAppDataPath & "\Settings.ini") Then Exit Sub
    End If
    
    With oIni
        If frmHaupt.WindowState = vbNormal Then
            Call .SetValue("Fenster", "PosTop", frmHaupt.Top)
            Call .SetValue("Fenster", "PosLeft", frmHaupt.Left)
            Call .SetValue("Fenster", "PosHeight", frmHaupt.Height)
            Call .SetValue("Fenster", "PosWidth", frmHaupt.Width)
        Else 'die letzten Werte vor dem Maximize
            Call .SetValue("Fenster", "PosTop", glPosTop)
            Call .SetValue("Fenster", "PosLeft", glPosLeft)
            Call .SetValue("Fenster", "PosHeight", glPosHeight)
            Call .SetValue("Fenster", "PosWidth", glPosWidth)
        End If
        
        Call .SetValue("Fenster", "ScrollValue", frmHaupt.VScroll1.Value)
        
        'MD-Marker 20090325 , frmBrowser aus dem Projekt entfernt
'        'Section Browser
'        Call .SetValue("Browser", "PosTop", glBrowserTop)
'        Call .SetValue("Browser", "PosLeft", glBrowserLeft)
'        Call .SetValue("Browser", "PosHeight", glBrowserHeight)
'        Call .SetValue("Browser", "PosWidth", glBrowserWidth)
        
        'Section NeuerArtikelFenster
        Call .SetValue("NeuerArtikel", "PosTop", glNeuerArtikelTop)
        Call .SetValue("NeuerArtikel", "PosLeft", glNeuerArtikelLeft)
        Call .SetValue("NeuerArtikel", "PosHeight", glNeuerArtikelHeight)
        Call .SetValue("NeuerArtikel", "PosWidth", glNeuerArtikelWidth)
        
        'Section InlineBrowser
        Call .SetValue("InlineBrowser", "PosTopPercent", gfInfoTop)
        Call .SetValue("InlineBrowser", "PosLeftPercent", gfInfoLeft)
        Call .SetValue("InlineBrowser", "PosHeightPercent", gfInfoHeight)
        Call .SetValue("InlineBrowser", "PosWidthPercent", gfInfoWidth)
    End With
    
    If Not bCalledBySaveSettings Then Call oIni.WriteIni
    
End Sub

Public Sub SaveSetting(sSection As String, sName As String, ByVal vValue As Variant)
    
    Dim oIni As clsIni
    
    Set oIni = New clsIni
    If oIni.ReadIni(gsAppDataPath & "\Settings.ini") Then
        Call CINISetValue(oIni, sSection, sName, vValue)
        oIni.WriteIni
    End If
    Set oIni = Nothing
    
End Sub

Public Function ReadSetting(sSection As String, sName As String, Optional ByVal sDefaultValue As String = "") As String
    
    On Error Resume Next
    
    Dim iRet As Integer
    Dim sFile As String
    Dim sTmp As String
    Dim oIni As clsIni
    
    sFile = gsAppDataPath & "\Settings.ini"
      
    Set oIni = New clsIni
    
    Call oIni.ReadIni(sFile)
    
    iRet = CINIGetValue(oIni, sSection, sName, sTmp)
    If iRet <= 0 Then
        ReadSetting = sDefaultValue
    Else
        ReadSetting = sTmp
    End If
    
End Function

' 1.8.0 Zugriff auf Ini- Dateien
Public Sub SaveAllSettings()
    
    On Error Resume Next
    
    Dim sTmp As String
    Dim sFile As String
    Dim i As Integer
    Dim oIni As clsIni
    
    Set oIni = New clsIni
    'es werden nur die Daten aus dem HSP gespeichert
    
    sFile = gsAppDataPath & "\Settings.ini"
    
    'Section Bieten
    'INIDeleteSection oIni, "Bieten"
    
    Call CINISetValue(oIni, "Bieten", "Useranzahl", giUserAnzahl)
    Call CINISetValue(oIni, "Bieten", "DefaultUser", giDefaultUser)
    
    If giUserAnzahl > 0 Then
        For i = 1 To giUserAnzahl
            Call CINISetValue(oIni, "Bieten", "User" & CStr(i), gtarrUserArray(i).UaUser)
            sTmp = EncodePass(gtarrUserArray(i).UaPass)
            Call CINISetValue(oIni, "Bieten", "Pass" & CStr(i), sTmp)
            Call CINISetValue(oIni, "Bieten", "UseSecurityToken" & CStr(i), gtarrUserArray(i).UaToken)
        Next i
    End If
    
    Call CINISetValue(oIni, "Bieten", "BrowserId", gsBrowserIdString)
    Call CINISetValue(oIni, "Bieten", "Vorlauf", glVorlaufGebot)
    Call CINISetValue(oIni, "Bieten", "VorlaufSnipe", gfVorlaufSnipe)
    Call CINISetValue(oIni, "Bieten", "PlaySoundOnBid", gbPlaySoundOnBid)
    Call CINISetValue(oIni, "Bieten", "SoundOnBid", gsSoundOnBid)
    Call CINISetValue(oIni, "Bieten", "SoundOnBidSuccess", gsSoundOnBidSuccess)
    Call CINISetValue(oIni, "Bieten", "SoundOnBidFail", gsSoundOnBidFail)
    Call oIni.SetValue("Bieten", "BuyItNow", gbBuyItNow)
    
    'Section Verbindung
    'INIDeleteSection oIni, "Verbindung"
    Call CINISetValue(oIni, "Verbindung", "Modem", gbUsesModem)
    Call CINISetValue(oIni, "Verbindung", "DialupRequestTimeout", giDialupRequestTimeout)
    Call CINISetValue(oIni, "Verbindung", "Vorlauf_LAN", giVorlaufLan)
    Call CINISetValue(oIni, "Verbindung", "UseProxy", gbUseProxy)
    Call CINISetValue(oIni, "Verbindung", "ProxyName", gsProxyName)
    Call CINISetValue(oIni, "Verbindung", "ProxyPort", giProxyPort)
    Call CINISetValue(oIni, "Verbindung", "UseProxyAuthentication", gbUseProxyAuthentication)
    Call CINISetValue(oIni, "Verbindung", "ProxyUser", gsProxyUser)
    sTmp = EncodePass(gsProxyPass)
    Call CINISetValue(oIni, "Verbindung", "ProxyPass", sTmp)
    Call CINISetValue(oIni, "Verbindung", "UseDirectConnect", gbUseDirectConnect)
    Call CINISetValue(oIni, "Verbindung", "UseIECookies", gbUseIECookies)
    Call CINISetValue(oIni, "Verbindung", "UseCurl", gbUseCurl)
    Call CINISetValue(oIni, "Verbindung", "HTTPTimeout", glHttpTimeOut)
    Call CINISetValue(oIni, "Verbindung", "Vorlauf_Modem", glVorlaufModem)
    Call CINISetValue(oIni, "Verbindung", "ConnectName", gsConnectName)
    Call CINISetValue(oIni, "Verbindung", "TestConnect", gbTestConnect)
    Call CINISetValue(oIni, "Verbindung", "CheckForUpdate", gbCheckForUpdate)
    Call CINISetValue(oIni, "Verbindung", "CheckForUpdateBeta", gbCheckForUpdateBeta)
    Call CINISetValue(oIni, "Verbindung", "CheckForUpdateInterval", glCheckForUpdateInterval)
    Call CINISetValue(oIni, "Verbindung", "AutoUpdateCurrencies", gbAutoUpdateCurrencies)
    Call CINISetValue(oIni, "Verbindung", "BrowserLanguage", gsBrowserLanguage)
    
    'Section Automatik
    'INIDeleteSection oIni, "Automatik"
    Call CINISetValue(oIni, "Automatik", "StartCheck", gbPassAtStart)
    Call CINISetValue(oIni, "Automatik", "AutoStart", gbAutoStart)
    Call CINISetValue(oIni, "Automatik", "AutoLogin", gbAutoLogin)
    Call CINISetValue(oIni, "Automatik", "TrayAction", gbTrayAction)
    Call CINISetValue(oIni, "Automatik", "WinShutdown", gbFileWinShutdown)
    Call CINISetValue(oIni, "Automatik", "ArtikelRefresh", gbGeboteAktualisieren)
    Call CINISetValue(oIni, "Automatik", "ArtikelRefreshCycle", giArtikelRefreshCycle)
    Call CINISetValue(oIni, "Automatik", "ArtikelRefreshPost", gbArtikelRefreshPost)
    Call CINISetValue(oIni, "Automatik", "ArtikelRefreshPost2", gbArtikelRefreshPost2)
    Call CINISetValue(oIni, "Automatik", "TimeSync", giUseTimeSync)
    Call CINISetValue(oIni, "Automatik", "TimeSyncIntervall", glTimeSyncIntervall)
    Call CINISetValue(oIni, "Automatik", "AutoSave", frmHaupt.AutoSave.Enabled)
    Call CINISetValue(oIni, "Automatik", "AutoAktualisieren", gbAutoAktualisieren)
    
    Call CINISetValue(oIni, "Automatik", "AutoAktualisierennext", gbAutoAktualisierenNext)
    Call CINISetValue(oIni, "Automatik", "AktualisierenXvor", gbAktualisierenXvor)
    Call CINISetValue(oIni, "Automatik", "AktXminvor", giAktXminvor)
    Call CINISetValue(oIni, "Automatik", "AktXminvorCycle", giAktXminvorCycle)
    Call CINISetValue(oIni, "Automatik", "ArtAktOptions", giArtAktOptions)
    Call CINISetValue(oIni, "Automatik", "ArtAktOptionsValue", giArtAktOptionsValue)
    Call CINISetValue(oIni, "Automatik", "AktualisierenOpt", giAktualisierenOpt)
    Call CINISetValue(oIni, "Automatik", "AutoWarnNoBid", gbAutoWarnNoBid)
    Call CINISetValue(oIni, "Automatik", "ConcurrentUpdates", gbConcurrentUpdates)
    Call CINISetValue(oIni, "Automatik", "UpdateAfterManualBid", gbUpdateAfterManualBid)
    Call CINISetValue(oIni, "Automatik", "QuietAfterManualBid", gbQuietAfterManualBid)
    
    Call CINISetValue(oIni, "Automatik", "KeinHinweisNachZeitsync", gbKeinHinweisNachZeitsync)
    Call CINISetValue(oIni, "Automatik", "WarnenBeimBeenden", gbWarnenBeimBeenden)
    Call CINISetValue(oIni, "Automatik", "BeendenNachAuktion", gbBeendenNachAuktion)
    
    Call CINISetValue(oIni, "Automatik", "NeuLadenBeiNichtGefunden", giReloadTimes)
    Call CINISetValue(oIni, "Automatik", "ReLogin", giReLogin)
    Call CINISetValue(oIni, "Automatik", "EditShippingOnClick", gbEditShippingOnClick)
    Call CINISetValue(oIni, "Automatik", "OpenBrowserOnClick", gbOpenBrowserOnClick)
    
    Call CINISetValue(oIni, "Automatik", "PreventSuspend", giPreventSuspend)
    Call CINISetValue(oIni, "Automatik", "WakeOnAuction", giWakeOnAuction)
    Call CINISetValue(oIni, "Automatik", "ResuspendAfterEnd", gbResuspendAfterEnd)
    Call CINISetValue(oIni, "Automatik", "ForceResuspendAfterEnd", gbForceResuspendAfterEnd)
    Call CINISetValue(oIni, "Automatik", "SleepAfterWakeup", giSleepAfterWakeup)
    Call CINISetValue(oIni, "Automatik", "Hibernate", gbHibernate)
    
    Call CINISetValue(oIni, "Automatik", "ExtCmdTimeWindow", glExtCmdTimeWindow)
    Call CINISetValue(oIni, "Automatik", "ExtCmdPreCmd", gsExtCmdPreCmd)
    Call CINISetValue(oIni, "Automatik", "ExtCmdPostCmd", gsExtCmdPostCmd)
    Call CINISetValue(oIni, "Automatik", "ExtCmdPeriodicCmd", gsExtCmdPeriodicCmd)
    Call CINISetValue(oIni, "Automatik", "ExtCmdPreTime", glExtCmdPreTime)
    Call CINISetValue(oIni, "Automatik", "ExtCmdPostTime", glExtCmdPostTime)
    Call CINISetValue(oIni, "Automatik", "ExtCmdPeriodicTime", glExtCmdPeriodicTime)
    Call CINISetValue(oIni, "Automatik", "ExtCmdWindowStyle", giExtCmdWindowStyle)
    
    Call CINISetValue(oIni, "Automatik", "SendCsvInterval", glSendCsvInterval)
    Call CINISetValue(oIni, "Automatik", "SendCsvTo", gsSendCsvTo)
    
    Call CINISetValue(oIni, "Automatik", "ReadEndedItems", gbReadEndedItems)
    Call CINISetValue(oIni, "Automatik", "BeepBeforeAuction", gbBeepBeforeAuction)
    Call CINISetValue(oIni, "Automatik", "BlockEndedItems", gbBlockEndedItems)
    Call CINISetValue(oIni, "Automatik", "BlockBuyItNowItems", gbBlockBuyItNowItems)

    
    'Section POP
    'INIDeleteSection oIni, "POP"
    Call CINISetValue(oIni, "POP", "UsePop", gbUsePop)
    Call CINISetValue(oIni, "POP", "POPZykl", giPopZyklus)
    Call CINISetValue(oIni, "POP", "POPServer", gsPopServer)
    Call CINISetValue(oIni, "POP", "POPPort", giPopPort)
    Call CINISetValue(oIni, "POP", "SMTPServer", gsSmtpServer)
    Call CINISetValue(oIni, "POP", "SMTPPort", giSmtpPort)
    Call CINISetValue(oIni, "POP", "POPUser", gsPopUser)
    sTmp = EncodePass(gsPopPass)
    Call CINISetValue(oIni, "POP", "POPPass", sTmp)
    Call CINISetValue(oIni, "POP", "POPTimeout", giPopTimeOut)
    Call CINISetValue(oIni, "POP", "Absender", gsAbsender)
    Call CINISetValue(oIni, "POP", "UseSMTPAuth", gbUseSmtpAuth)
    
    With oIni
        Call .SetValue("POP", "POPCmdSSL", gsPopCmdSSL)
        Call .SetValue("POP", "SMTPCmdSSL", gsSmtpCmdSSL)
        Call .SetValue("POP", "POPUseSSL", gbPopUseSSL)
        Call .SetValue("POP", "SMTPUseSSL", gbSmtpUseSSL)
        Call .SetValue("POP", "HideSSLWindow", gbHideSSLWindow)
        Call .SetValue("POP", "SSLStartupDelay", glSSLStartupDelay)
        Call .SetValue("POP", "POPEncryptedOnly", gbPopEncryptedOnly)
        Call .SetValue("POP", "POPSendEncryptedAcknowledgment", gbPopSendEncryptedAcknowledgment)
        Call .SetValue("POP", "POPNeedsUsername", gbPopNeedsUsername)
        Call .SetValue("POP", "POPSubjectDelimiter", gsPopSubjectDelimiter)
    End With
    
    'Section EbayServer
    'INIDeleteSection oIni, "Server"
    'INIDeleteSection oIni, "EbayServer"
    'sind in den Keyword-Settings
    
    'Section Darstellung
    'INIDeleteSection oIni, "Darstellung"
    Call CINISetValue(oIni, "Darstellung", "AnzZeilen", giMaxRowSetting)
    Call CINISetValue(oIni, "Darstellung", "StartupSize", giStartupSize)
    'Call CINISetValue(oIni, "Darstellung", "UseIntBrowser", gbUseIntBrowser) 'MD-Marker 20090325 , interner browser entfernt
    Call CINISetValue(oIni, "Darstellung", "UseWheelMouse", gbUseWheel)
    Call CINISetValue(oIni, "Darstellung", "ShowToolbar", gbShowToolbar)
    Call CINISetValue(oIni, "Darstellung", "ToolbarSize", giToolbarSize)
    Call CINISetValue(oIni, "Darstellung", "UseOperaField", gbOperaField)
    Call CINISetValue(oIni, "Darstellung", "Language", gsAktLanguage)
    Call CINISetValue(oIni, "Darstellung", "MinToTray", gbMinToTray)
    Call CINISetValue(oIni, "Darstellung", "NewItemWindowAlwaysOnTop", gbNewItemWindowAlwaysOnTop)
    Call CINISetValue(oIni, "Darstellung", "NewItemWindowOpenOnStartup", gbNewItemWindowOpenOnStartup)
    Call CINISetValue(oIni, "Darstellung", "NewItemWindowKeepsValues", gbNewItemWindowKeepsValues)
    Call CINISetValue(oIni, "Darstellung", "NewItemWindowWidgetOrdner", gsNewItemWindowWidgetOrdner)
    Call CINISetValue(oIni, "Darstellung", "IconSet", gsIconSet)
    Call CINISetValue(oIni, "Darstellung", "TrayIconDisplayTimeOnlineMode1", giTrayIconDisplayTimeOnlineMode1)
    Call CINISetValue(oIni, "Darstellung", "TrayIconDisplayTimeOnlineMode2", giTrayIconDisplayTimeOnlineMode2)
    Call CINISetValue(oIni, "Darstellung", "TrayIconDisplayTimeOfflineMode1", giTrayIconDisplayTimeOfflineMode1)
    Call CINISetValue(oIni, "Darstellung", "TrayIconDisplayTimeOfflineMode2", giTrayIconDisplayTimeOfflineMode2)
    Call CINISetValue(oIni, "Darstellung", "ShowTitleDateTime", gbShowTitleDateTime)
    Call CINISetValue(oIni, "Darstellung", "ShowTitleTimeLeft", gbShowTitleTimeLeft)
    Call CINISetValue(oIni, "Darstellung", "ShowTitleVersion", gbShowTitleVersion)
    Call CINISetValue(oIni, "Darstellung", "ShowTitleAuctionHome", gbShowTitleAuctionHome)
    Call CINISetValue(oIni, "Darstellung", "ShowTitleDefaultUser", gbShowTitleDefaultUser)
    Call CINISetValue(oIni, "Darstellung", "CleanStatus", gbCleanStatus)
    Call CINISetValue(oIni, "Darstellung", "CleanStatusTime", glCleanStatusTime)
    Call CINISetValue(oIni, "Darstellung", "BrowseInNewWindow", gbBrowseInNewWindow)
    Call CINISetValue(oIni, "Darstellung", "SortOrder", gsSortOrder)
    
    With oIni
        Call .SetValue("Darstellung", "BrowseInline", gbBrowseInline)
        Call .SetValue("Darstellung", "InlineBrowserDelay", giInlineBrowserDelay)
        Call .SetValue("Darstellung", "InlineBrowserModifierKey", giInlineBrowserModifierKey)
        Call .SetValue("Darstellung", "CommentInTitle", gbCommentInTitle)
        Call .SetValue("Darstellung", "RevisedInTitle", gbRevisedInTitle)
        Call .SetValue("Darstellung", "ShowShippingCosts", gbShowShippingCosts)
        Call .SetValue("Darstellung", "ShowFocusRect", gbShowFocusRect)
        Call .SetValue("Darstellung", "FocusRectColor", GetRgbHexFromColor(glFocusRectColor))
        Call oIni.SetValue("Darstellung", "SpecialDateFormat", gsSpecialDateFormat)
        Call .SetValue("Darstellung", "ShowWeekday", gbShowWeekday)
    End With
    
    'Section Fenster
    'INIDeleteSection oIni, "Fenster"
    Call CINISetValue(oIni, "Fenster", "FontName", gsGlobFontName)
    Call CINISetValue(oIni, "Fenster", "FontSize", giDefaultFontSize)
    Call CINISetValue(oIni, "Fenster", "FieldHeight", giDefaultHeight)

    'Section Browser
    'INIDeleteSection oIni, "Browser"
    
    'Section NeuerArtikelFenster
    'INIDeleteSection oIni, "NeuerArtikel"
    
    Call SaveWindowSettings(oIni)
    
    'Section Diverses
    'INIDeleteSection oIni, "Diverses"
    Call CINISetValue(oIni, "Diverses", "SendAuctionEnd", gbSendAuctionEnd)
    Call CINISetValue(oIni, "Diverses", "SendAuctionEndNoSuccess", gbSendAuctionEndNoSuccess)
    Call CINISetValue(oIni, "Diverses", "SendIfLow", gbSendIfLow)
    Call CINISetValue(oIni, "Diverses", "SendEndTo", gsSendEndTo)
    Call CINISetValue(oIni, "Diverses", "SendEndFrom", gsSendEndFrom)
    Call CINISetValue(oIni, "Diverses", "SendEndFromRealname", gsSendEndFromRealname)
    Call CINISetValue(oIni, "Diverses", "TestArtikelNummer", gsTestArtikel)
    Call CINISetValue(oIni, "Diverses", "Separator", gsSeparator)
    Call CINISetValue(oIni, "Diverses", "Delimiter", gsDelimiter)
    Call CINISetValue(oIni, "Diverses", "ShowSplash", gbShowSplash)
    Call CINISetValue(oIni, "Diverses", "CountDownInAutomodeOnly", gbCountDownInAutomodeOnly)
    Call CINISetValue(oIni, "Diverses", "ShowSplashOnce", 1)
    Call CINISetValue(oIni, "Diverses", "ShippingMode", giShippingMode)
    Call CINISetValue(oIni, "Diverses", "ServerStringsFile", gsServerStringsFile)
    
    With oIni
        Call .SetValue("Diverses", "ReservedPriceMarker", gsReservedPriceMarker)
        Call .SetValue("Diverses", "SendItemTo", gsSendItemTo)
        Call .SetValue("Diverses", "SendItemEncrypted", gbSendItemEncrypted)
        Call .SetValue("Diverses", "ConfirmDelete", gbConfirmDelete)
        Call .SetValue("Diverses", "LogfileMaxSize", glLogfileMaxSize)
        Call .SetValue("Diverses", "LogfileShrinkPercent", glLogfileShrinkPercent)
        Call .SetValue("Diverses", "UseIsoDate", gbUseIsoDate)
        Call .SetValue("Diverses", "UseUnixDate", gbUseUnixDate)
        Call .SetValue("Diverses", "NoEnumFonts", gbNoEnumFonts)
        Call .SetValue("Diverses", "SuppressHeader", gbSuppressHeader)
        Call .SetValue("Diverses", "IgnoreItemErrorsOnStartup", gbIgnoreItemErrorsOnStartup)
        Call .SetValue("Diverses", "LogDeletedItems", gbLogDeletedItems)
        Call .SetValue("Diverses", "BlacklistDeletedItems", gbBlacklistDeletedItems)
    End With
    
    'Section NTP
    'INIDeleteSection oIni, "NTP"
    Call CINISetValue(oIni, "NTP", "UseNTP", giUseNtp)
    Call CINISetValue(oIni, "NTP", "NTPServer", gsNtpServer)
    
    'Section ODBC
    'INIDeleteSection oIni, "ODBC"
    Call CINISetValue(oIni, "ODBC", "UseODBC", gbUsesOdbc)
    Call CINISetValue(oIni, "ODBC", "ODBC_Zyklus", giOdbcZyklus)
    Call CINISetValue(oIni, "ODBC", "ODBC_Provider", gsOdbcProvider)
    Call CINISetValue(oIni, "ODBC", "ODBC_DB", gsOdbcDb)
    Call CINISetValue(oIni, "ODBC", "ODBC_User", gsOdbcUser)
    Call CINISetValue(oIni, "ODBC", "ODBC_Pass", gsOdbcPass)
    
    'Section Debug
    'INIDeleteSection oIni, "Debug"
    Call CINISetValue(oIni, "Debug", "StripJS", gbStripJS)
    Call CINISetValue(oIni, "Debug", "LogHTML", gbLogHtml)
    Call CINISetValue(oIni, "Debug", "ShowDebugWindow", gbShowDebugWindow)
    Call CINISetValue(oIni, "Debug", "DebugLevel", giDebugLevel)
       
    Call CINISetValue(oIni, "Debug", "PosTop", glDebugWindowTop)
    Call CINISetValue(oIni, "Debug", "PosLeft", glDebugWindowLeft)
    Call CINISetValue(oIni, "Debug", "PosHeight", glDebugWindowHeight)
    Call CINISetValue(oIni, "Debug", "PosWidth", glDebugWindowWidth)
       
    Call oIni.WriteIni(sFile)
    Call SaveCurrencies
    
    Set oIni = Nothing
    
End Sub
Public Sub ReadAllSettings()

    On Error Resume Next

    Dim iTmp As Integer, iTmp1 As Integer
    Dim sTmp As String, sTmp1  As String, sTmp2 As String
    Dim i As Integer
    Dim iEmptyVal As Integer
    Dim oIni As New clsIni
    Dim oPicTmp As IPictureDisp
    Dim oTIni As New clsIni
    Dim vntWeTmp As Variant




    'es werden nur die Daten aus dem HSP gespeichert
    oIni.ReadIni gsAppDataPath & "\Settings.ini"
    
    'Section Bieten
    
    iTmp = CINIGetValue(oIni, "Bieten", "Useranzahl", sTmp)
    If iTmp <= 0 Then
        giUserAnzahl = 0
    Else
        giUserAnzahl = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Bieten", "DefaultUser", sTmp)
    If iTmp <= 0 Then
        giDefaultUser = 0
    Else
        giDefaultUser = sTmp
    End If
    
    ReDim gtarrUserArray(0 To 0) As udtUserPass
    If giUserAnzahl > 0 Then
        i = 1
        Do While i < giUserAnzahl + 1
            
            iTmp = CINIGetValue(oIni, "Bieten", "User" & i, sTmp)
            iTmp1 = CINIGetValue(oIni, "Bieten", "Pass" & i, sTmp1)
            sTmp = LCase(Trim(sTmp))
            
            If iTmp > 0 And iTmp1 > 0 Then
                ReDim Preserve gtarrUserArray(0 To i - iEmptyVal) As udtUserPass
                gtarrUserArray(i - iEmptyVal).UaUser = sTmp
                gtarrUserArray(i - iEmptyVal).UaPass = DecodePass(sTmp1)
                iTmp = CINIGetValue(oIni, "Bieten", "UseSecurityToken" & i, sTmp2)
                If iTmp <= 0 Then
                    gtarrUserArray(i - iEmptyVal).UaToken = 0
                Else
                    gtarrUserArray(i - iEmptyVal).UaToken = sTmp2
                End If
                
                If i = giDefaultUser Then
                    giDefaultUser = (i - iEmptyVal)
                    gsUser = sTmp
                    gsPass = DecodePass(sTmp1)
                    gbUseSecurityToken = gtarrUserArray(i - iEmptyVal).UaToken
                End If
            Else
                iEmptyVal = iEmptyVal + 1
            End If
            i = i + 1
        Loop
    Else   'alte ini auf neu
        iTmp = CINIGetValue(oIni, "Bieten", "User", sTmp)
        iTmp1 = CINIGetValue(oIni, "Bieten", "Pass", sTmp1)
        sTmp = LCase(Replace(sTmp, " ", ""))
        If iTmp > 0 And iTmp1 > 0 Then
            ReDim gtarrUserArray(1) As udtUserPass
            gtarrUserArray(1).UaUser = sTmp
            gsUser = sTmp
            sTmp1 = DecodePass(sTmp1)
            gtarrUserArray(1).UaPass = sTmp1
            gsPass = sTmp1
            giUserAnzahl = 1
            giDefaultUser = 1
        End If
    End If 'giUserAnzahl > 0
    
    If iEmptyVal > 0 Then
        giUserAnzahl = giUserAnzahl - iEmptyVal
        If giUserAnzahl < 1 Then
            giDefaultUser = 0
            ReDim gtarrUserArray(0 To 0) As udtUserPass
        Else
            If UBound(gtarrUserArray()) < giDefaultUser Then
                giDefaultUser = 1
            End If
            gsUser = gtarrUserArray(giDefaultUser).UaUser
            gsPass = gtarrUserArray(giDefaultUser).UaPass
            gbUseSecurityToken = gtarrUserArray(giDefaultUser).UaToken
        End If
    End If
    
    iTmp = CINIGetValue(oIni, "Bieten", "PlaySoundOnBid", sTmp)
    If iTmp <= 0 Then
        gbPlaySoundOnBid = False
    Else
        gbPlaySoundOnBid = sTmp
    End If
    iTmp = CINIGetValue(oIni, "Bieten", "SoundOnBid", gsSoundOnBid)
    iTmp = CINIGetValue(oIni, "Bieten", "SoundOnBidSuccess", gsSoundOnBidSuccess)
    iTmp = CINIGetValue(oIni, "Bieten", "SoundOnBidFail", gsSoundOnBidFail)
    
    iTmp = CINIGetValue(oIni, "Bieten", "BrowserId", gsBrowserIdString)
    If iTmp <= 0 Then
        gsBrowserIdString = "Mozilla/5.0 (Windows; U; Windows NT 5.1; en; rv:1.8.1.3) Gecko/20070309 Firefox/2.0.0.3"
    End If
    
    iTmp = CINIGetValue(oIni, "Bieten", "Vorlauf", sTmp)
    If iTmp <= 0 Then
        glVorlaufGebot = 15
    Else
    glVorlaufGebot = sTmp
    End If
    
    'Vorlaufzeit wandeln
    gfVorlaufGebotTimeVal = myTimeSerial(0, 0, 1) * glVorlaufGebot 'lg 12.05.2003
    
    iTmp = CINIGetValue(oIni, "Bieten", "VorlaufSnipe", sTmp)
    If iTmp <= 0 Then
        gfVorlaufSnipe = 3
    Else
        gfVorlaufSnipe = sTmp
    End If
    
    gbBuyItNow = CBool(oIni.GetValue("Bieten", "BuyItNow", "0"))
    
    'Section Verbindung
    iTmp = CINIGetValue(oIni, "Verbindung", "Modem", sTmp)
    If iTmp <= 0 Then
        gbUsesModem = False
    Else
        gbUsesModem = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "DialupRequestTimeout", sTmp)
    If iTmp <= 0 Then
        giDialupRequestTimeout = 10
    Else
        giDialupRequestTimeout = sTmp
    End If
    
    With oIni
        gbCheckForUpdate = CBool(.GetValue("Verbindung", "CheckForUpdate", "1"))
        gbCheckForUpdateBeta = CBool(.GetValue("Verbindung", "CheckForUpdateBeta", "0"))
        glCheckForUpdateInterval = CLng(.GetValue("Verbindung", "CheckForUpdateInterval", "4"))
    End With
    
    iTmp = CINIGetValue(oIni, "Verbindung", "AutoUpdateCurrencies", sTmp)
    If iTmp <= 0 Then
        gbAutoUpdateCurrencies = True
    Else
        gbAutoUpdateCurrencies = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "Vorlauf_LAN", sTmp)
    
    If iTmp <= 0 Then
        giVorlaufLan = 0
    Else
        giVorlaufLan = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "UseProxy", sTmp)
    If iTmp <= 0 Then
        gbUseProxy = False
    Else
        gbUseProxy = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "ProxyName", gsProxyName)
    If iTmp <= 0 Then gsProxyName = ""
    
    iTmp = CINIGetValue(oIni, "Verbindung", "ProxyPort", sTmp)
    If iTmp <= 0 Then
        giProxyPort = 80
    Else
        giProxyPort = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "UseProxyAuthentication", sTmp)
    If iTmp <= 0 Then
        gbUseProxyAuthentication = False
    Else
        gbUseProxyAuthentication = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "ProxyUser", gsProxyUser)
    If iTmp <= 0 Then gsProxyUser = ""
    
    iTmp = CINIGetValue(oIni, "Verbindung", "ProxyPass", sTmp)
    If iTmp <= 0 Then
        gsProxyPass = ""
    Else
        gsProxyPass = DecodePass(sTmp)
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "UseDirectConnect", sTmp)
    If iTmp <= 0 Then
        gbUseDirectConnect = True
    Else
        gbUseDirectConnect = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "UseIECookies", sTmp)
    If iTmp <= 0 Then
        gbUseIECookies = False
    Else
'        gbUseIECookies = sTmp ' wir wollen keine IE-Cookies mehr, das gibt nur Stress!
    End If

    gbUseCurl = oIni.GetValue("Verbindung", "UseCurl", True)
    If Not TestForCurl() Then gbUseCurl = False
    
    iTmp = CINIGetValue(oIni, "Verbindung", "HTTPTimeout", sTmp)
    If iTmp <= 0 Then
        glHttpTimeOut = 10000
    Else
        glHttpTimeOut = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "Vorlauf_Modem", sTmp)
    If iTmp <= 0 Then
        glVorlaufModem = 5
    Else
        glVorlaufModem = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "ConnectName", gsConnectName)
    If iTmp <= 0 Then gsConnectName = ""
    
    iTmp = CINIGetValue(oIni, "Verbindung", "TestConnect", sTmp)
    If iTmp <= 0 Then
        gbTestConnect = False
    Else
        gbTestConnect = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Verbindung", "BrowserLanguage", sTmp)
    If iTmp <= 0 Then
        gsBrowserLanguage = "en"
    Else
        gsBrowserLanguage = sTmp
    End If
    
    'Section Automatik
    iTmp = CINIGetValue(oIni, "Automatik", "StartCheck", sTmp)
    If iTmp <= 0 Then
        gbPassAtStart = False
    Else
        gbPassAtStart = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "AutoStart", sTmp)
    If iTmp <= 0 Then
        gbAutoStart = False
    Else
        gbAutoStart = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "AutoLogin", sTmp)
    If iTmp <= 0 Then
        gbAutoLogin = False
    Else
        gbAutoLogin = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "TrayAction", sTmp)
    If iTmp <= 0 Then
        gbTrayAction = False
    Else
        gbTrayAction = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "WinShutdown", sTmp)
    If iTmp <= 0 Then
        gbFileWinShutdown = False
    Else
        gbFileWinShutdown = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "ArtikelRefresh", sTmp)
    If iTmp <= 0 Then
        gbGeboteAktualisieren = False
    Else
        gbGeboteAktualisieren = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "ArtikelRefreshCycle", sTmp)
    If iTmp <= 0 Then
        giArtikelRefreshCycle = 0
    Else
        giArtikelRefreshCycle = sTmp
    End If

    iTmp = CINIGetValue(oIni, "Automatik", "ArtikelRefreshPost", sTmp)
    If iTmp <= 0 Then
        gbArtikelRefreshPost = True
    Else
        gbArtikelRefreshPost = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "ArtikelRefreshPost2", sTmp)
    If iTmp <= 0 Then
        gbArtikelRefreshPost2 = True
    Else
        gbArtikelRefreshPost2 = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "TimeSync", sTmp)
    If iTmp <= 0 Then
        giUseTimeSync = 14
    Else
        giUseTimeSync = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "TimeSyncIntervall", sTmp)
    If iTmp <= 0 Then
        glTimeSyncIntervall = 60
    Else
        glTimeSyncIntervall = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "AutoSave", sTmp)
    If iTmp <= 0 Then
        frmHaupt.AutoSave.Enabled = True
    Else
        frmHaupt.AutoSave.Enabled = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "AutoAktualisieren", sTmp)
    If iTmp <= 0 Then
        gbAutoAktualisieren = False
    Else
        gbAutoAktualisieren = sTmp
    End If
    
    'sh nur nächster + xminvor
    iTmp = CINIGetValue(oIni, "Automatik", "AutoAktualisierennext", sTmp)
    If iTmp <= 0 Then
        gbAutoAktualisierenNext = False
    Else
        gbAutoAktualisierenNext = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "AktualisierenXvor", sTmp)
    If iTmp <= 0 Then
        gbAktualisierenXvor = False
    Else
        gbAktualisierenXvor = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "AktXminvor", sTmp)
    If iTmp <= 0 Then
        giAktXminvor = 3
    Else
        giAktXminvor = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "AktXminvorCycle", sTmp)
    If iTmp <= 0 Then
        giAktXminvorCycle = 10
    Else
        giAktXminvorCycle = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "ArtAktOptions", sTmp)
    If iTmp <= 0 Then
        giArtAktOptions = 0
    Else
        giArtAktOptions = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "ArtAktOptionsValue", sTmp)
    If iTmp <= 0 Then
        giArtAktOptionsValue = 0
    Else
        giArtAktOptionsValue = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "AktualisierenOpt", sTmp)
    If iTmp <= 0 Then
        giAktualisierenOpt = 0
    Else
        giAktualisierenOpt = sTmp
    End If

    iTmp = CINIGetValue(oIni, "Automatik", "AutoWarnNoBid", sTmp)
    If iTmp <= 0 Then
        gbAutoWarnNoBid = False
    Else
        gbAutoWarnNoBid = sTmp
    End If
    
    With oIni
        gbConcurrentUpdates = .GetValue("Automatik", "ConcurrentUpdates", "0")
        gbUpdateAfterManualBid = .GetValue("Automatik", "UpdateAfterManualBid", "1")
        gbQuietAfterManualBid = .GetValue("Automatik", "QuietAfterManualBid", "0")
    End With
    
    iTmp = CINIGetValue(oIni, "Automatik", "KeinHinweisNachZeitsync", sTmp)
    If iTmp <= 0 Then
        gbKeinHinweisNachZeitsync = True
    Else
        gbKeinHinweisNachZeitsync = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "WarnenBeimBeenden", sTmp)
    If iTmp <= 0 Then
        gbWarnenBeimBeenden = False
    Else
        gbWarnenBeimBeenden = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "BeendenNachAuktion", sTmp)
    If iTmp <= 0 Then
        gbBeendenNachAuktion = False
    Else
        gbBeendenNachAuktion = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "NeuLadenBeiNichtGefunden", sTmp)
    If iTmp <= 0 Then
        giReloadTimes = 3
    Else
        giReloadTimes = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Automatik", "ReLogin", sTmp)
    
    If iTmp <= 0 Then
        giReLogin = 3
    Else
        giReLogin = sTmp
    End If
    
    With oIni
        gbEditShippingOnClick = .GetValue("Automatik", "EditShippingOnClick", False)
        gbOpenBrowserOnClick = .GetValue("Automatik", "OpenBrowserOnClick", True)
        
        giPreventSuspend = .GetValue("Automatik", "PreventSuspend", 7)
        giWakeOnAuction = .GetValue("Automatik", "WakeOnAuction", 4)
        gbResuspendAfterEnd = .GetValue("Automatik", "ResuspendAfterEnd", True)
        gbForceResuspendAfterEnd = .GetValue("Automatik", "ForceResuspendAfterEnd", False)
        giSleepAfterWakeup = .GetValue("Automatik", "SleepAfterWakeup", 10)
        gbHibernate = .GetValue("Automatik", "Hibernate", False)
        
        glExtCmdTimeWindow = .GetValue("Automatik", "ExtCmdTimeWindow", 30)
        gsExtCmdPreCmd = .GetValue("Automatik", "ExtCmdPreCmd", "")
        gsExtCmdPostCmd = .GetValue("Automatik", "ExtCmdPostCmd", "")
        gsExtCmdPeriodicCmd = .GetValue("Automatik", "ExtCmdPeriodicCmd", "")
        glExtCmdPreTime = .GetValue("Automatik", "ExtCmdPreTime", 60)
        glExtCmdPostTime = .GetValue("Automatik", "ExtCmdPostTime", 30)
        glExtCmdPeriodicTime = .GetValue("Automatik", "ExtCmdPeriodicTime", 3600)
        giExtCmdWindowStyle = .GetValue("Automatik", "ExtCmdWindowStyle", vbHide)
        
        glSendCsvInterval = .GetValue("Automatik", "SendCsvInterval", 0)
        gsSendCsvTo = .GetValue("Automatik", "SendCsvTo", "")
        
        gbReadEndedItems = .GetValue("Automatik", "ReadEndedItems", False)
        gbBeepBeforeAuction = .GetValue("Automatik", "BeepBeforeAuction", False)
        gbBlockEndedItems = .GetValue("Automatik", "BlockEndedItems", False)
        gbBlockBuyItNowItems = .GetValue("Automatik", "BlockBuyItNowItems", False)
    End With
    
    'Section POP
    iTmp = CINIGetValue(oIni, "POP", "UsePop", sTmp)
    If iTmp <= 0 Then
        gbUsePop = False
    Else
        gbUsePop = sTmp
    End If

    iTmp = CINIGetValue(oIni, "POP", "POPZykl", sTmp)
    If iTmp <= 0 Then
        giPopZyklus = 60
    Else
        giPopZyklus = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "POP", "POPServer", gsPopServer)
    iTmp = CINIGetValue(oIni, "POP", "SMTPServer", gsSmtpServer)
    iTmp = CINIGetValue(oIni, "POP", "POPUser", gsPopUser)
    
    iTmp = CINIGetValue(oIni, "POP", "POPPort", sTmp)
    If iTmp <= 0 Then
        giPopPort = 110
    Else
        giPopPort = Val(sTmp)
    End If
    
    iTmp = CINIGetValue(oIni, "POP", "SMTPPort", sTmp)
    If iTmp <= 0 Then
        giSmtpPort = 25
    Else
        giSmtpPort = Val(sTmp)
    End If
    
    iTmp = CINIGetValue(oIni, "POP", "POPPass", sTmp)
    If iTmp <= 0 Then
        gsPopPass = ""
    Else
        gsPopPass = DecodePass(sTmp)
    End If
    
    iTmp = CINIGetValue(oIni, "POP", "POPTimeout", sTmp)
    If iTmp <= 0 Then
        giPopTimeOut = 60
    Else
        giPopTimeOut = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "POP", "Absender", gsAbsender)
    
    iTmp = CINIGetValue(oIni, "POP", "UseSMTPAuth", sTmp)
    
    If iTmp <= 0 Then
        gbUseSmtpAuth = False
    Else
        gbUseSmtpAuth = sTmp
    End If
    
    With oIni
        gsPopCmdSSL = .GetValue("POP", "POPCmdSSL", "stunnel -c -r %SERVER%:%PORT% -d 127.0.0.1:%PORT%")
        gsSmtpCmdSSL = .GetValue("POP", "SMTPCmdSSL", "stunnel -c -r %SERVER%:%PORT% -d 127.0.0.1:%PORT%")
        gbPopUseSSL = .GetValue("POP", "POPUseSSL", False)
        gbSmtpUseSSL = .GetValue("POP", "SMTPUseSSL", False)
        gbHideSSLWindow = .GetValue("POP", "HideSSLWindow", True)
        glSSLStartupDelay = .GetValue("POP", "SSLStartupDelay", 100)
        gbPopEncryptedOnly = .GetValue("POP", "POPEncryptedOnly", False)
        gbPopSendEncryptedAcknowledgment = .GetValue("POP", "POPSendEncryptedAcknowledgment", False)
        gbPopNeedsUsername = .GetValue("POP", "POPNeedsUsername", False)
        gsPopSubjectDelimiter = .GetValue("POP", "POPSubjectDelimiter", "|")
    End With
    
    'Section EbayServer
    'raus
    
    'Section Darstellung
    iTmp = CINIGetValue(oIni, "Darstellung", "AnzZeilen", sTmp)
    If iTmp <= 0 Then
        giMaxRow = 12
        giMaxRowSetting = giMaxRow
    Else
        giMaxRow = sTmp
        giMaxRowSetting = giMaxRow
    End If
    
    iTmp = CINIGetValue(oIni, "Darstellung", "StartupSize", sTmp)
    If iTmp <= 0 Then
        giStartupSize = vbNormal
    Else
        giStartupSize = sTmp
    End If
    
    'MD-Marker 20090325 , interner Browser entfernt
    
    'iTmp = CINIGetValue(oIni, "Darstellung", "UseIntBrowser", sTmp)
    'If iTmp <= 0 Then
        'gbUseIntBrowser = True
    'Else
        'gbUseIntBrowser = sTmp
    'End If
    'gbUseIntBrowser = False

    iTmp = CINIGetValue(oIni, "Darstellung", "BrowseInNewWindow", sTmp)
    If iTmp <= 0 Then
        gbBrowseInNewWindow = False
    Else
        gbBrowseInNewWindow = sTmp
    End If
    
    With oIni
        gbBrowseInline = .GetValue("Darstellung", "BrowseInline", 1)
        giInlineBrowserDelay = .GetValue("Darstellung", "InlineBrowserDelay", "100")
        giInlineBrowserModifierKey = .GetValue("Darstellung", "InlineBrowserModifierKey", "2")
        gbCommentInTitle = .GetValue("Darstellung", "CommentInTitle", 1)
        gbRevisedInTitle = .GetValue("Darstellung", "RevisedInTitle", 1)
        gbShowShippingCosts = .GetValue("Darstellung", "ShowShippingCosts", 1)
        gbShowFocusRect = .GetValue("Darstellung", "ShowFocusRect", 1)
        glFocusRectColor = GetColorFromRgbHex(.GetValue("Darstellung", "FocusRectColor", "808080"))
        gsSpecialDateFormat = .GetValue("Darstellung", "SpecialDateFormat", "")
        gbShowWeekday = .GetValue("Darstellung", "ShowWeekday", 0)
    End With
    
    iTmp = CINIGetValue(oIni, "Darstellung", "SortOrder", sTmp)
    If iTmp <= 0 Then
        gsSortOrder = "asc"
    Else
        gsSortOrder = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Darstellung", "ShowToolbar", sTmp)
    If iTmp <= 0 Then
        gbShowToolbar = True
    Else
        gbShowToolbar = sTmp
    End If

    iTmp = CINIGetValue(oIni, "Darstellung", "ToolbarSize", sTmp)
    If iTmp <= 0 Then
        giToolbarSize = 0
    Else
        giToolbarSize = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Darstellung", "UseWheelMouse", sTmp)
    If iTmp <= 0 Then
        gbUseWheel = True
    Else
        gbUseWheel = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Darstellung", "UseOperaField", sTmp)
    If iTmp <= 0 Then
        gbOperaField = False
    Else
        gbOperaField = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Darstellung", "Language", sTmp)
    If iTmp <= 0 Then
        gsAktLanguage = "deutsch"
    Else
        If sTmp = "german" Then sTmp = "deutsch"
        gsAktLanguage = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Darstellung", "MinToTray", sTmp)
    If iTmp <= 0 Then
        gbMinToTray = True
    Else
        gbMinToTray = sTmp
    End If
    
    gbNewItemWindowAlwaysOnTop = oIni.GetValue("Darstellung", "NewItemWindowAlwaysOnTop", 0)
    gbNewItemWindowOpenOnStartup = oIni.GetValue("Darstellung", "NewItemWindowOpenOnStartup", 0)
    gbNewItemWindowKeepsValues = oIni.GetValue("Darstellung", "NewItemWindowKeepsValues", 0)
    gsNewItemWindowWidgetOrdner = oIni.GetValue("Darstellung", "NewItemWindowWidgetOrdner", "1,2,3,4,5,6")
    gsIconSet = oIni.GetValue("Darstellung", "IconSet", "")
    
    giTrayIconDisplayTimeOnlineMode1 = oIni.GetValue("Darstellung", "TrayIconDisplayTimeOnlineMode1", 10)
    giTrayIconDisplayTimeOnlineMode2 = oIni.GetValue("Darstellung", "TrayIconDisplayTimeOnlineMode2", 2)
    giTrayIconDisplayTimeOfflineMode1 = oIni.GetValue("Darstellung", "TrayIconDisplayTimeOfflineMode1", 10)
    giTrayIconDisplayTimeOfflineMode2 = oIni.GetValue("Darstellung", "TrayIconDisplayTimeOfflineMode2", 2)
    
    Set frmDummy.Picture1(0) = frmHaupt.Icon
    Set frmDummy.Picture1(1) = MyLoadResPicture(202, 16)
    frmDummy.Picture1(2).Tag = ""
    Set frmDummy.Picture1(3) = MyLoadResPicture(201, 16)
    frmDummy.Picture1(4).Tag = ""
    
    Set oPicTmp = MyLoadResPicture(203, 16)
    If Not oPicTmp Is Nothing Then Set frmDummy.Picture1(1) = oPicTmp
    Set oPicTmp = MyLoadResPicture(204, 16)
    If Not oPicTmp Is Nothing Then Set frmDummy.Picture1(2) = oPicTmp:  frmDummy.Picture1(2).Tag = "1"
    Set oPicTmp = MyLoadResPicture(205, 16)
    If Not oPicTmp Is Nothing Then Set frmDummy.Picture1(3) = oPicTmp
    Set oPicTmp = MyLoadResPicture(206, 16)
    If Not oPicTmp Is Nothing Then Set frmDummy.Picture1(4) = oPicTmp:  frmDummy.Picture1(4).Tag = "1"
    
    iTmp = CINIGetValue(oIni, "Darstellung", "ShowTitleDateTime", sTmp)
    If iTmp <= 0 Then
        gbShowTitleDateTime = True
    Else
        gbShowTitleDateTime = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Darstellung", "ShowTitleTimeLeft", sTmp)
    If iTmp <= 0 Then
        gbShowTitleTimeLeft = True
    Else
        gbShowTitleTimeLeft = sTmp
    End If
    
    gbShowTitleVersion = oIni.GetValue("Darstellung", "ShowTitleVersion", 1)
    gbShowTitleAuctionHome = oIni.GetValue("Darstellung", "ShowTitleAuctionHome", 1)
    gbShowTitleDefaultUser = oIni.GetValue("Darstellung", "ShowTitleDefaultUser", 1)
    
    iTmp = CINIGetValue(oIni, "Darstellung", "CleanStatus", sTmp)
    If iTmp <= 0 Then
        gbCleanStatus = True
    Else
        gbCleanStatus = sTmp
    End If

    iTmp = CINIGetValue(oIni, "Darstellung", "CleanStatusTime", sTmp)
    If iTmp <= 0 Then
        glCleanStatusTime = 3
    Else
        glCleanStatusTime = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Fenster", "FontName", gsGlobFontName)
    If iTmp <= 0 Then
        gsGlobFontName = "MS Sans Serif"
    End If
    
    iTmp = CINIGetValue(oIni, "Fenster", "FontSize", sTmp)
    If iTmp <= 0 Then
        giDefaultFontSize = 8
    Else
        giDefaultFontSize = Abs(Val(sTmp))
    End If
    
    iTmp = CINIGetValue(oIni, "Fenster", "FieldHeight", sTmp)
    If iTmp <= 0 Then
        giDefaultHeight = 440
    Else
        giDefaultHeight = Abs(Val(sTmp))
    End If
    
    iTmp = CINIGetValue(oIni, "Fenster", "PosTop", sTmp)
    If iTmp <= 0 Or Val(sTmp) < -400 Or Val(sTmp) > GetScreenHeight() Then
        glPosTop = (Screen.Height - 9000) / 2
    Else
        glPosTop = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Fenster", "PosLeft", sTmp)
    If iTmp <= 0 Or Val(sTmp) < -400 Or Val(sTmp) > GetScreenWidth() Then
        glPosLeft = (Screen.Width - 12000) / 2
    Else
        glPosLeft = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Fenster", "PosHeight", sTmp)
    If iTmp <= 0 Or Val(sTmp) <= 0 Then
        glPosHeight = 9000
    Else
        glPosHeight = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Fenster", "PosWidth", sTmp)
    If iTmp <= 0 Or Val(sTmp) <= 0 Then
        glPosWidth = 12000
    Else
        glPosWidth = sTmp
    End If
    
'MD-Marker 20090325 , frmBrowser aus Projekt entfernt
'    'Section Browser
'    iTmp = CINIGetValue(oIni, "Browser", "PosTop", sTmp)
'    If iTmp <= 0 Or Val(sTmp) < -400 Or Val(sTmp) > GetScreenHeight() Then
'        glBrowserTop = (Screen.Height - 9000) / 2
'    Else
'        glBrowserTop = sTmp
'    End If
'
'    iTmp = CINIGetValue(oIni, "Browser", "PosLeft", sTmp)
'    If iTmp <= 0 Or Val(sTmp) < -400 Or Val(sTmp) > GetScreenWidth() Then
'        glBrowserLeft = (Screen.Width - 12000) / 2
'    Else
'        glBrowserLeft = sTmp
'    End If
'
'    iTmp = CINIGetValue(oIni, "Browser", "PosHeight", sTmp)
'    If iTmp <= 0 Or Val(sTmp) <= 0 Then
'        glBrowserHeight = 9000
'    Else
'        glBrowserHeight = sTmp
'    End If
'
'    iTmp = CINIGetValue(oIni, "Browser", "PosWidth", sTmp)
'    If iTmp <= 0 Or Val(sTmp) <= 0 Then
'        glBrowserWidth = 12000
'    Else
'        glBrowserWidth = sTmp
'    End If
    
    'Section NeuerArtikel
    iTmp = CINIGetValue(oIni, "NeuerArtikel", "PosTop", sTmp)
    If iTmp <= 0 Or Val(sTmp) < -400 Or Val(sTmp) > GetScreenHeight() Then
        glNeuerArtikelTop = 0
    Else
        glNeuerArtikelTop = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "NeuerArtikel", "PosLeft", sTmp)
    If iTmp <= 0 Or Val(sTmp) < -400 Or Val(sTmp) > GetScreenWidth() Then
        glNeuerArtikelLeft = 0
    Else
        glNeuerArtikelLeft = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "NeuerArtikel", "PosHeight", sTmp)
    If iTmp <= 0 Or Val(sTmp) <= 0 Then
        glNeuerArtikelHeight = 2300
    Else
        glNeuerArtikelHeight = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "NeuerArtikel", "PosWidth", sTmp)
    If iTmp <= 0 Or Val(sTmp) <= 0 Then
        glNeuerArtikelWidth = 3375
    Else
        glNeuerArtikelWidth = sTmp
    End If
    
    'Section Debug
    iTmp = CINIGetValue(oIni, "Debug", "PosTop", sTmp)
    If iTmp <= 0 Or Val(sTmp) < -400 Or Val(sTmp) > GetScreenHeight() Then
        glDebugWindowTop = (Screen.Height - 9000) / 2
    Else
        glDebugWindowTop = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Debug", "PosLeft", sTmp)
    If iTmp <= 0 Or Val(sTmp) < -400 Or Val(sTmp) > GetScreenWidth() Then
        glDebugWindowLeft = (Screen.Width - 12000) / 2
    Else
        glDebugWindowLeft = sTmp
    End If

    iTmp = CINIGetValue(oIni, "Debug", "PosHeight", sTmp)
    If iTmp <= 0 Or Val(sTmp) <= 0 Then
        glDebugWindowHeight = 9000
    Else
        glDebugWindowHeight = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Debug", "PosWidth", sTmp)
    If iTmp <= 0 Or Val(sTmp) <= 0 Then
        glDebugWindowWidth = 12000
    Else
        glDebugWindowWidth = sTmp
    End If
    frmDebug.SetSize
    
    'Section InlineBrowser
    gfInfoTop = oIni.GetValue("InlineBrowser", "PosTopPercent", 0)
    gfInfoLeft = oIni.GetValue("InlineBrowser", "PosLeftPercent", 50)
    gfInfoHeight = oIni.GetValue("InlineBrowser", "PosHeightPercent", 100)
    gfInfoWidth = oIni.GetValue("InlineBrowser", "PosWidthPercent", 50)
    
    'Section Diverses
    iTmp = CINIGetValue(oIni, "Diverses", "SendAuctionEnd", sTmp)
    If iTmp <= 0 Then
        gbSendAuctionEnd = False
    Else
        gbSendAuctionEnd = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Diverses", "SendAuctionEndNoSuccess", sTmp)
    If iTmp <= 0 Then
        gbSendAuctionEndNoSuccess = gbSendAuctionEnd 'Stand vor der Aufteilung oder ganz neu
    Else
        gbSendAuctionEndNoSuccess = sTmp
    End If

    gbSendIfLow = oIni.GetValue("Diverses", "SendIfLow", 0)
    iTmp = CINIGetValue(oIni, "Diverses", "SendEndTo", gsSendEndTo)
    iTmp = CINIGetValue(oIni, "Diverses", "SendEndFrom", gsSendEndFrom)
    iTmp = CINIGetValue(oIni, "Diverses", "SendEndFromRealname", gsSendEndFromRealname)
    
    iTmp = CINIGetValue(oIni, "Diverses", "TestArtikelNummer", sTmp)
    If iTmp <= 0 Then
        gsTestArtikel = "3378262626"
    Else
        gsTestArtikel = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Diverses", "Separator", sTmp)
    If iTmp <= 0 Then
        gsSeparator = ";"
    Else
        gsSeparator = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Diverses", "Delimiter", sTmp)
    If iTmp <= 0 Then
        gsDelimiter = """"
    Else
        gsDelimiter = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Diverses", "ShowSplash", sTmp)
    If iTmp <= 0 Then
        gbShowSplash = True
    Else
        gbShowSplash = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Diverses", "CountDownInAutomodeOnly", sTmp)
    If iTmp <= 0 Then
        gbCountDownInAutomodeOnly = False
    Else
        gbCountDownInAutomodeOnly = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Diverses", "ShippingMode", sTmp)
    If iTmp <= 0 Then
        giShippingMode = 1
    Else
        giShippingMode = sTmp
    End If

    iTmp = CINIGetValue(oIni, "Diverses", "ServerStringsFile", sTmp)
    If iTmp <= 0 Then
        If Dir(App.Path & "\ServerStrings.ini") > "" Then
            oTIni.ReadIni App.Path & "\ServerStrings.ini"
            iTmp = CINIGetValue(oTIni, "Server", "AuctionHome", sTmp)
        End If
        
        If iTmp <= 0 Then sTmp = "eBay.de"
        sTmp = LCase(Trim(sTmp))
        gsServerStringsFile = "ServerStrings_" & sTmp & ".ini"
    Else
        gsServerStringsFile = Trim(sTmp)
        'new auf normale umschalten
        If LCase(gsServerStringsFile) Like "*.new.*" Then gsServerStringsFile = Replace(gsServerStringsFile, ".new", "")
    End If
        
    gsToolTipSeparator = " " & Chr(149) & " "
    
    With oIni
        gsReservedPriceMarker = .GetValue("Diverses", "ReservedPriceMarker", "%PRICE%*")
        gsSendItemTo = .GetValue("Diverses", "SendItemTo")
        gbSendItemEncrypted = .GetValue("Diverses", "SendItemEncrypted", "1")
        gbConfirmDelete = .GetValue("Diverses", "ConfirmDelete", False)
        glLogfileMaxSize = .GetValue("Diverses", "LogfileMaxSize", CLng(1) * 1024 * 1024) ' 1 MB
        glLogfileShrinkPercent = .GetValue("Diverses", "LogfileShrinkPercent", 10)
        gbLogDeletedItems = .GetValue("Diverses", "LogDeletedItems", 0)
        gbBlacklistDeletedItems = .GetValue("Diverses", "BlacklistDeletedItems", 0)
        gbUseIsoDate = .GetValue("Diverses", "UseIsoDate", 0)
        gbUseUnixDate = .GetValue("Diverses", "UseUnixDate", 0)
        gbNoEnumFonts = .GetValue("Diverses", "NoEnumFonts", 0)
        gbSuppressHeader = .GetValue("Diverses", "SuppressHeader", 0)
        gbIgnoreItemErrorsOnStartup = .GetValue("Diverses", "IgnoreItemErrorsOnStartup", 0)
    End With
    
    'Section NTP
    iTmp = CINIGetValue(oIni, "NTP", "UseNtp", sTmp)
    If iTmp <= 0 Then
        giUseNtp = 2
    Else
        If Len(Trim(sTmp)) > 1 Then 'old-value (True/False)
            If Not CBool(sTmp) Then
                giUseNtp = 0
            Else
                giUseNtp = 1
            End If
        Else
            giUseNtp = Val(sTmp)
        End If
    End If

    iTmp = CINIGetValue(oIni, "NTP", "NTPServer", sTmp)
    If iTmp <= 0 Then
        gsNtpServer = "pool.ntp.org"
    Else
        gsNtpServer = sTmp
    End If
    
    'ODBC
    iTmp = CINIGetValue(oIni, "ODBC", "UseODBC", sTmp)
    If iTmp <= 0 Then
        gbUsesOdbc = False
    Else
        gbUsesOdbc = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "ODBC", "ODBC_Zyklus", sTmp)
    
    If iTmp <= 0 Then
        giOdbcZyklus = 60
    Else
        giOdbcZyklus = sTmp
        If giOdbcZyklus <= 1 Then giOdbcZyklus = 1
    End If
    
    iTmp = CINIGetValue(oIni, "ODBC", "ODBC_Provider", sTmp)
    If iTmp <= 0 Then
        gsOdbcProvider = "Microsoft.Jet.OLEDB.4.0"
    Else
        gsOdbcProvider = Trim(sTmp)
    End If
    
    iTmp = CINIGetValue(oIni, "ODBC", "ODBC_DB", sTmp)
    If iTmp <= 0 Then
        gsOdbcDb = ""
    Else
        gsOdbcDb = Trim(sTmp)
    End If
    
    iTmp = CINIGetValue(oIni, "ODBC", "ODBC_User", sTmp)
    If iTmp <= 0 Then
        gsOdbcUser = ""
    Else
        gsOdbcUser = Trim(sTmp)
    End If
    
    iTmp = CINIGetValue(oIni, "ODBC", "ODBC_Pass", sTmp)
    If iTmp <= 0 Then
        gsOdbcPass = ""
    Else
        gsOdbcPass = Trim(sTmp)
    End If
    
    For Each vntWeTmp In gcolWeNames
        iTmp = CINIGetValue(oIni, "Currency", vntWeTmp, sTmp)
        If iTmp > 0 Then
            gcolWeValues.Remove vntWeTmp
            gcolWeValues.Add CCur(sTmp), vntWeTmp
        End If
    Next
    
    iTmp = CINIGetValue(oIni, "Debug", "StripJS", sTmp)
    If iTmp <= 0 Then
        gbStripJS = True
    Else
        gbStripJS = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Debug", "LogHTML", sTmp)
    If iTmp <= 0 Then
        gbLogHtml = False
    Else
        gbLogHtml = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Debug", "ShowDebugWindow", sTmp)
    If iTmp <= 0 Then
        gbShowDebugWindow = False
    Else
        gbShowDebugWindow = sTmp
    End If
    
    iTmp = CINIGetValue(oIni, "Debug", "DebugLevel", sTmp)
    If iTmp <= 0 Then
        giDebugLevel = 1
    Else
        giDebugLevel = sTmp
    End If
    
    Set oIni = Nothing
    Set oTIni = Nothing
    
End Sub

Public Function INISetValue(ByVal sPath As String, ByVal sSect As String, ByVal sKey As String, ByVal sValue As String) As Integer

    Dim lResult As Long
    
    'Wert schreiben
    lResult = WritePrivateProfileString(sSect, sKey, sValue, sPath)
    
    INISetValue = lResult
    
End Function

Public Function INIGetValue(ByVal sPath As String, ByVal sSect As String, ByVal sKey As String, sValue As String) As Integer

    Dim lResult As Long, sBuffer As String
    
    'Wert lesen
    sBuffer = Space$(255)
    lResult = GetPrivateProfileString(sSect, sKey, vbNullString, _
        sBuffer, Len(sBuffer), sPath)
    
    sValue = Left$(sBuffer, lResult)
    INIGetValue = lResult
    
End Function

Public Function CINISetValue(oIni As clsIni, ByVal sSect As String, ByVal sKey As String, ByVal vValue As Variant) As Integer
    
    Call oIni.SetValue(sSect, sKey, vValue)
    CINISetValue = 1
    
End Function

Public Function CINIGetValue(oIni As clsIni, ByVal sSect As String, ByVal sKey As String, ByRef sValue As String) As Integer
    
    CINIGetValue = 1
    sValue = oIni.GetValue(sSect, sKey, "<-~=not found=~->")
    If sValue = "<-~=not found=~->" Then
        sValue = ""
        CINIGetValue = 0
    End If
    
End Function

'Private Function INISetArray(ByVal Path$, ByVal Sect$, xArray() As String)
'  Dim X%, Buffer$, result&
'    'Feld in einen String mit Trennzeichen Chr$(0) umwandeln
'    For X = LBound(xArray) To UBound(xArray)
'      Buffer = Buffer & xArray(X) & Chr$(0)
'    Next X
'
'    'String schreiben
'    Buffer = Left$(Buffer, Len(Buffer) - 1)
'    result = WritePrivateProfileSection(Sect, Buffer, Path)
'End Function

'Private Sub INIGetArray(ByVal Path$, ByVal Sect$, xArray() As String)
'  Dim result&, Buffer$
'  Dim l%, p%, Z%
'    'String lesen
'    Buffer = Space(32767)
'    result = GetPrivateProfileSection(Sect, Buffer, Len(Buffer), Path)
'
'    Buffer = Left$(Buffer, result)
'
'    If Buffer <> "" Then
'      'String mit Trennzeichen Chr$(0) in ein Feld umwandeln
'      l = 1
'      ReDim xArray(0)
'      Do While l < result
'        p = InStr(l, Buffer, Chr$(0))
'        If p = 0 Then Exit Do
'
'        xArray(Z) = Mid$(Buffer, l, p - l)
'        Z = Z + 1
'        ReDim Preserve xArray(0 To Z)
'        l = p + 1
'      Loop
'    End If
'End Sub
'
'Public Function INIGetArrayEx(ByVal Path$, ByVal Sect$, Optional BufferSize As Long = 32767) As Variant
'  Dim result&, Buffer$
'  Dim l%, p%, Z%
'  Dim tmpArray As Variant
'    'String lesen
'    Buffer = Space(BufferSize)
'    result = GetPrivateProfileSection(Sect, Buffer, Len(Buffer), Path)
'
'    If result > 0 Then
'      Buffer = Left$(Buffer, result - 1)
'    Else
'      Buffer = ""
'    End If
'
'    tmpArray = Array("")
'    If Buffer <> "" Then
'      'String mit Trennzeichen Chr$(0) in ein Feld umwandeln
'      tmpArray = Split(Buffer, Chr$(0))
'    End If
'    INIGetArrayEx = tmpArray
'
'End Function

Private Sub INIDeleteKey(ByVal Path$, ByVal Sect$, ByVal Key$)

    'MD-Marker , Sub wird nicht aufgerufen
    
    'Call WritePrivateProfileString(Sect, Key, 0&, Path)
    
End Sub
 
Private Sub INIDeleteSection(ByVal sPath As String, ByVal sSect As String)
    Call WritePrivateProfileString(sSect, 0&, 0&, sPath)
End Sub

Private Function EncodePass(sTxt As String) As String

    Dim i As Integer
    Dim iChar1 As Integer
    Dim iChar2 As Integer
    Dim sBuffer As String
    Dim sPass As String
    Dim sKey As String
    
    Do While Len(sKey) < Len(sTxt): sKey = sKey & gsCKEY2: Loop
    
    For i = 1 To Len(sTxt)
        iChar1 = Asc(Mid(sTxt, i, 1))
        iChar2 = Asc(Mid(sKey, i, 1))
        sBuffer = Chr(iChar1 Xor iChar2)
        sBuffer = Hex((CInt(Asc(sBuffer)) + 17) Mod 256)
        If Len(sBuffer) = 1 Then sBuffer = "0" & sBuffer
        sPass = sPass & sBuffer
    Next i
    EncodePass = sPass
    
End Function
Public Function DecodePass(sTxt As String) As String

On Error GoTo ERROR_HANDLER

Dim i As Integer
Dim sChar1 As String
Dim sBuffer As String
Dim sPass As String
Dim sShort As String
Dim sKey As String

For i = 1 To Len(sTxt) Step 2
    sBuffer = "&h" & Mid(sTxt, i, 2)
    sShort = sShort & Chr((256 + sBuffer - 17) Mod 256)
Next i

Do While Len(sKey) < Len(sShort): sKey = sKey & gsCKEY2: Loop

For i = 1 To Len(sShort)
    sChar1 = Asc(Mid(sShort, i, 1))
    sBuffer = Asc(Mid(sKey, i, 1))
    sBuffer = sBuffer Xor sChar1
    sPass = sPass & Chr(sBuffer)
Next i

DecodePass = sPass
Exit Function
ERROR_HANDLER:
DecodePass = ""
End Function
' Artikel- CSV Schreiben

Public Function GetDynArrayFromArtikelZeile(tZeile As udtArtikelZeile) As Variant
    
    With tZeile
        GetDynArrayFromArtikelZeile = Array(.Artikel, .Titel, _
            FormatOutputDate(.EndeZeit), .AktPreis, .WE, .Gebot, .Gruppe, .Status, _
            .AnzGebote, .Bieter, .Kommentar, FormatOutputBoolean(.PostUpdateDone), _
            .Versand, .Verkaeufer, .UserAccount, .Bewertung, .Standort, .MinGebot, _
            FormatOutputBoolean(.MindestpreisNichtErreicht), _
            FormatOutputBoolean(.Ueberarbeitet), .TimeZone)
    End With
    
End Function

Private Function FormatOutputBoolean(ByVal bValue As Boolean) As String
    
    If bValue Then
        FormatOutputBoolean = "1"
    Else
        FormatOutputBoolean = "0"
    End If
    
End Function

Private Function FormatOutputDate(ByVal datDate As Date) As String
    
    If gbUseIsoDate Then
        FormatOutputDate = Date2IsoDate(datDate)
    ElseIf gbUseUnixDate Then
        FormatOutputDate = Date2UnixDate(datDate)
    Else
        FormatOutputDate = datDate
    End If
    
End Function

Private Function ConvertInputDate(ByVal sDate As String) As Date

  If gbUseIsoDate Then
    ConvertInputDate = IsoDate2Date(sDate)
  ElseIf gbUseUnixDate Then
    ConvertInputDate = UnixDate2Date(sDate)
  Else
    ConvertInputDate = sDate
  End If

End Function

Public Function BuildArtikelCSV2(Optional lFromID As Long = 0) As String

    On Error Resume Next
    
    Dim iAktRow As Integer
    Dim tmp As String
    Dim sCSVString As String
    Dim vntElement As Variant
    Dim vntDynArray As Variant
    Dim lStringSize As Long
    Dim bytTyp As Byte
    Dim tTypZeile As udtArtikelZeile
    
    
    If Not gbSuppressHeader Then
        sCSVString = sCSVString & "# Artikeldatei V 2.9.0 Separator:" & gsSeparator & " Delimiter:" & gsDelimiter & vbCrLf
        sCSVString = sCSVString & "# Wird bei jedem Speichern überschrieben!" & vbCrLf
        sCSVString = sCSVString & "# Aufbau:" & vbCrLf
        sCSVString = sCSVString & "# Artikelnr" & gsSeparator & "Titel" & gsSeparator & "Endezeit" & gsSeparator & "AktPreis" & gsSeparator & "Waehrung" & gsSeparator & "Gebot" & gsSeparator & "Gruppe" & gsSeparator & "Status" & gsSeparator & "AnzGebote" & gsSeparator & "Bieter" & gsSeparator & "Kommentar" & gsSeparator & "PostUpdateDone" & gsSeparator & "Versandkosten" & gsSeparator & "Verkaeufer" & gsSeparator & "Account" & gsSeparator & "Bewertung" & gsSeparator & "Standort" & gsSeparator & "MinGebot" & gsSeparator & "MindestPreisNichtErreicht" & gsSeparator & "Ueberarbeitet" & gsSeparator & "Zeitzone" & vbCrLf
    End If
    
    lStringSize = Len(sCSVString)
    frmProgress.InitProgress 1, giAktAnzArtikel
    
    For bytTyp = 1 To IIf(lFromID > 0, 2, 1)
        For iAktRow = 1 To IIf(bytTyp = 1, giAktAnzArtikel, UBound(gtarrRemovedArtikelArray()))
            
            If bytTyp = 1 Then
                LSet tTypZeile = gtarrArtikelArray(iAktRow)
            ElseIf bytTyp = 2 Then
                LSet tTypZeile = gtarrRemovedArtikelArray(iAktRow)
            End If
            
            If tTypZeile.LastChangedId >= lFromID Then
                vntDynArray = GetDynArrayFromArtikelZeile(tTypZeile)
                
                tmp = ""
                
                For Each vntElement In vntDynArray
                    vntElement = Replace(vntElement, gsDelimiter, gsDelimiter & gsDelimiter)
                    vntElement = Replace(vntElement, "\", "\\")
                    vntElement = Replace(vntElement, vbCrLf, "\n")
                    vntElement = Replace(vntElement, vbLf, "\n")
                    vntElement = Replace(vntElement, vbCr, "\r")
                    vntElement = Replace(vntElement, vbTab, "\t")
                    If InStr(1, vntElement, gsDelimiter) > 0 Or InStr(1, vntElement, gsSeparator) > 0 Then
                        vntElement = gsDelimiter & vntElement & gsDelimiter
                    End If
                    tmp = tmp & vntElement & gsSeparator
                Next
                
                If lStringSize + Len(tmp) + 1 > Len(sCSVString) Then
                    sCSVString = sCSVString & Space(10000)
                End If
                
                Mid(sCSVString, lStringSize + 1, Len(tmp) + 1) = Left(tmp, Len(tmp) - 1) & vbCrLf
                lStringSize = lStringSize + Len(tmp) + 1
            End If
            
            frmProgress.Step
            
        Next iAktRow
    Next bytTyp
    
    BuildArtikelCSV2 = Mid$(sCSVString, 1, lStringSize)
    
    frmProgress.TerminateProgress
    
End Function

Public Function VersionValue(sVersion As String) As Long
    
    On Error Resume Next
    Dim lMajor As Long
    Dim lMinor As Long
    Dim lRevision As Long
    Dim sBeta As String
    Dim lPos1 As Long
    Dim lPos2 As Long
    
    lPos1 = InStr(1, sVersion, ".")
    If lPos1 > 0 Then
        lMajor = Val(Mid(sVersion, 1, lPos1 - 1))
        
        lPos2 = InStr(lPos1 + 1, sVersion, ".")
        If lPos2 > 0 Then
            lMinor = Val(Mid(sVersion, lPos1 + 1, lPos2 - lPos1))
            lRevision = Val(Mid(sVersion, lPos2 + 1))
            sBeta = Mid(sVersion, lPos2 + Len(CStr(lRevision)) + 1)
        Else
            lMinor = Val(Mid(sVersion, lPos1 + 1))
        End If
    Else
        lMajor = Val(sVersion)
    End If
    
    VersionValue = ((lMajor * 1000000) + (lMinor * 1000)) + lRevision
    'If InStr(1, sBeta, "beta", vbTextCompare) Then
        'VersionValue = VersionValue - 1
    'End If

End Function

Public Sub WriteDeletedItemLog(iAktRow As Integer)
    
    On Error GoTo errhdl
    Dim sTmp As String
    Dim iFileNr As Integer
    Dim vntDynArray As Variant
    Dim vntElement As Variant
    
    vntDynArray = GetDynArrayFromArtikelZeile(gtarrArtikelArray(iAktRow))
    
    sTmp = ""
    
    For Each vntElement In vntDynArray
        vntElement = Replace(vntElement, gsDelimiter, gsDelimiter & gsDelimiter)
        vntElement = Replace(vntElement, "\", "\\")
        vntElement = Replace(vntElement, vbCrLf, "\n")
        vntElement = Replace(vntElement, vbLf, "\n")
        vntElement = Replace(vntElement, vbCr, "\r")
        vntElement = Replace(vntElement, vbTab, "\t")
        If InStr(1, vntElement, gsDelimiter) > 0 Or InStr(1, vntElement, gsSeparator) > 0 Then vntElement = gsDelimiter & vntElement & gsDelimiter
        sTmp = sTmp & vntElement & gsSeparator
    Next
    
    iFileNr = FreeFile()
    Open gsAppDataPath & "\DeletedItems.log" For Append As #iFileNr
        Print #iFileNr, Date2Str(MyNow) & "   " & sTmp
    Close #iFileNr
    
  Exit Sub
errhdl:
  MsgBox "Fehler beim Schreiben der Artikeldaten: " & vbCrLf & Err.Description
  
  Err.Clear
  On Error Resume Next
  Close #iFileNr

End Sub

Public Sub WriteBlacklistedItemLog(iAktRow As Integer)
    
    On Error GoTo errhdl
    Dim sTmp As String
    Dim iFileNr As Integer
    
    sTmp = gtarrArtikelArray(iAktRow).Artikel
    
    iFileNr = FreeFile()
    Open gsAppDataPath & "\BlacklistedItems.log" For Append As #iFileNr
        Print #iFileNr, sTmp
    Close #iFileNr
    
  Exit Sub
errhdl:
  MsgBox "Fehler beim Schreiben der Artikeldaten: " & vbCrLf & Err.Description
  
  Err.Clear
  On Error Resume Next
  Close #iFileNr

End Sub

Public Sub WriteArtikelCsv2(Optional sData As String = "")
      
    'Kopiert die Artikelstruktur in ein Array und baut daraus ein CSV-File zusammen
    On Error GoTo errhdl
    Dim iFileNr As Integer
    
    Call BackupArtikelCsv
  
  
    'Zum Löschen:
    iFileNr = FreeFile()
    Open gsAppDataPath & "\Artikel.csv" For Output As #iFileNr
    Close #iFileNr
    'und CSV schreiben
    
    If sData = "" Then
        sData = BuildArtikelCSV2()
    End If
    
    iFileNr = FreeFile()
    Open gsAppDataPath & "\Artikel.csv" For Output As #iFileNr
        Print #iFileNr, sData;
    Close #iFileNr
    
    gsLastSavedCrc = Crc32(sData)
    
  Exit Sub
  
errhdl:
  MsgBox "Fehler beim Schreiben der Artikeldaten: " & vbCrLf & Err.Description
  
  Err.Clear
  On Error Resume Next
  Close #iFileNr

End Sub

Public Sub AddCsvArtikel2(sCSVArtikel As String, Optional bConvertValues As Boolean = False)
    
    'Zerlegt ein CSV-File Zeilenweise in Arrays und kopiert diese in die Artikelstruktur
    
    On Error GoTo errhdl
    
    Dim lRet As VbMsgBoxResult
    Dim sFehlermeldung As String
    Dim vntLine As Variant
    Dim vntLines As Variant
    Dim vntDynArray As Variant
    Dim iAktRow As Integer
    Dim tZeile As udtArtikelZeile
    Dim bIsNew As Boolean
    Dim i As Long
    Dim sCSVSeparator As String
    Dim sCSVDelimiter As String
    
    
    sCSVSeparator = GetArtikelFileSeparator(Left(sCSVArtikel, 100))
    sCSVDelimiter = GetArtikelFileDelimiter(Left(sCSVArtikel, 100))
    If sCSVSeparator = vbNullString Then sCSVSeparator = gsSeparator
    If sCSVDelimiter = vbNullString Then sCSVDelimiter = gsDelimiter
    
    vntLines = Split(sCSVArtikel, vbCrLf)
    frmProgress.InitProgress 0, UBound(vntLines) - LBound(vntLines) + 1
    
    For Each vntLine In vntLines
        
        If Left(vntLine, 1) <> "#" And Len(vntLine) > 5 Then
            vntDynArray = CSV2Array(vntLine, sCSVSeparator, sCSVDelimiter)
            If UBound(vntDynArray) < 21 Then ReDim Preserve vntDynArray(1 To 21) As Variant
            
            If vntDynArray(4) = "" Then vntDynArray(4) = 0
            If vntDynArray(6) = "" Then vntDynArray(6) = 0
            If vntDynArray(18) = "" Then vntDynArray(18) = 0
            If vntDynArray(19) = "" Then vntDynArray(19) = 0
            If vntDynArray(20) = "" Then vntDynArray(20) = 0
            If vntDynArray(21) = "" Then vntDynArray(21) = GetUTCOffset()
            
            With tZeile
                .Artikel = ""
                .Titel = "-"
                .EndeZeit = myDateSerial(3000, 1, 1)
                .AktPreis = 0
                .WE = ""
                .Gebot = 0
                .Gruppe = ""
                .Status = 0
                .AnzGebote = 0
                .Bieter = ""
                .Kommentar = ""
                .PostUpdateDone = False
                .Versand = ""
                .Verkaeufer = ""
                .UserAccount = ""
                .Bewertung = ""
                .Standort = ""
                .MinGebot = 0
                .MindestpreisNichtErreicht = False
                .Ueberarbeitet = False
                .TimeZone = 0
                
                .Artikel = vntDynArray(1)
                .Titel = vntDynArray(2)
                .EndeZeit = ConvertInputDate(vntDynArray(3))
                .AktPreis = IIf(bConvertValues, EbayString2Float(vntDynArray(4)), vntDynArray(4))
                .WE = vntDynArray(5)
                .Gebot = vntDynArray(6)
                .Gruppe = vntDynArray(7)
                .Status = vntDynArray(8)
                .AnzGebote = vntDynArray(9)
                .Bieter = vntDynArray(10)
                .Kommentar = vntDynArray(11)
                .PostUpdateDone = vntDynArray(12)
                .Versand = vntDynArray(13)
                .Verkaeufer = vntDynArray(14)
                If Len(vntDynArray(15)) = 0 Then
                    .UserAccount = vntDynArray(15)
                Else
                    .UserAccount = UsrAccTest(CStr(vntDynArray(15)))
                End If
                .Bewertung = vntDynArray(16)
                .Standort = vntDynArray(17)
                .MinGebot = vntDynArray(18)
                .MindestpreisNichtErreicht = vntDynArray(19)
                .Ueberarbeitet = vntDynArray(20)
                .TimeZone = vntDynArray(21)
                         
                If GetNumericPart(.Artikel) <> "" Then
                    bIsNew = True
                    'wir wollen den Artikel nicht doppelt haben:
                    For i = 1 To giAktAnzArtikel
                        If gtarrArtikelArray(i).Artikel = .Artikel Then
                            bIsNew = False
                        End If
                    Next i
                              
                    If bIsNew Then
                        iAktRow = giAktAnzArtikel + 1
                    
                        If iAktRow > UBound(gtarrArtikelArray) Then
                            ReDim Preserve gtarrArtikelArray(iAktRow + 100) As udtArtikelZeile
                        End If
                        LSet gtarrArtikelArray(iAktRow) = tZeile
                        giAktAnzArtikel = iAktRow
                    End If
                End If
            End With 'tZeile
        End If 'kein #
        frmProgress.Step
    Next
    
    If iAktRow > 0 Then ReDim Preserve gtarrArtikelArray(iAktRow) As udtArtikelZeile
    frmProgress.TerminateProgress
    
  Exit Sub

errhdl:
  sFehlermeldung = Err.Description
  Err.Clear
  If gbIgnoreItemErrorsOnStartup Then
    lRet = vbYes
  Else
    If FormLoaded("frmAbout") Then frmAbout.Hide
    If tZeile.Artikel <> "" Then
      MsgBox gsarrLangTxt(440) & " " & tZeile.Artikel & ": " & vbCrLf & sFehlermeldung, vbCritical  'TODO
    Else
      MsgBox gsarrLangTxt(441) & ": " & vbCrLf & sFehlermeldung
    End If
    lRet = MsgBox(gsarrLangTxt(442), vbQuestion + vbYesNoCancel)
  End If
  If lRet = vbCancel Then Exit Sub
  If lRet = vbYes Then On Error Resume Next
  Resume Next
  
End Sub

Private Function CSV2Array(ByVal sCSVData As String, sCSVSeparator As String, sCSVDelimiter As String) As Variant
    
    'Zerlegt eine CSV-Zeile in ein Array
    
    Dim sTmp As String
    Dim lArrayCounter As Long
    Dim sarrDynArray() As String
    
    ReDim sarrDynArray(1 To 1) As String
    
    Do While sCSVData <> vbNullString
        If Left(sCSVData, 1) = sCSVDelimiter Then 'Wert beginnt mit Delimiter, nach End-Delimiter suchen:
            sTmp = vbNullString
            Do
                sCSVData = Mid(sCSVData, 2)
                sTmp = sTmp & GetBisZeichen(sCSVData, sCSVDelimiter)
                If Left(sCSVData, 1) = sCSVDelimiter Then sTmp = sTmp & sCSVDelimiter
            Loop Until Left(sCSVData, 1) = sCSVSeparator Or sCSVData = vbNullString
            sCSVData = Mid(sCSVData, 2)
        Else 'Wert ohne Delimiter, bis Separator suchen:
            sTmp = GetBisZeichen(sCSVData, sCSVSeparator)
        End If
        
        sTmp = Replace(sTmp, "\t", vbTab)
        sTmp = Replace(sTmp, "\r", vbCr)
        sTmp = Replace(sTmp, "\n", vbCrLf)
        sTmp = Replace(sTmp, "\\", "\")
        
        lArrayCounter = lArrayCounter + 1
        If lArrayCounter > UBound(sarrDynArray()) Then ReDim Preserve sarrDynArray(1 To UBound(sarrDynArray()) + 1)
        sarrDynArray(lArrayCounter) = sTmp
    Loop
    
    CSV2Array = sarrDynArray()
    
End Function

Public Function GetBisZeichen(sTxt As String, sZeichen As String, Optional eCompare As VbCompareMethod) As String
        
    Dim lPos As Long
    
    lPos = InStr(1, sTxt, sZeichen, eCompare)
    If lPos > 0 Then
        GetBisZeichen = Left(sTxt, lPos - 1)
        sTxt = Mid(sTxt, lPos + Len(sZeichen))
    Else
        GetBisZeichen = sTxt
        sTxt = ""
    End If
    
End Function

Public Function GetArtikelFileVersion(sVersionLine As String) As String
    
    'extrahiert den Teilstring mit x.y.z
    Dim i As Long
    Dim sTmp As String
    
    For i = 1 To Len(sVersionLine)
        If Mid$(sVersionLine, i, 1) Like "[0-9.V" & vbCrLf & "]" Then sTmp = sTmp & Mid$(sVersionLine, i, 1)
    Next 'i
    
    If InStr(1, sTmp, "V", vbTextCompare) > 0 Then
        sTmp = Mid$(sTmp, InStr(1, sTmp, "V") + 1)
        If InStr(1, sTmp, vbCr, vbTextCompare) > 0 Then
            sTmp = Mid$(sTmp, 1, InStr(1, sTmp, vbCr) - 1)
        End If
        
        If InStr(1, sTmp, vbLf, vbTextCompare) > 0 Then
            sTmp = Mid$(sTmp, 1, InStr(1, sTmp, vbLf) - 1)
        End If
    End If
    GetArtikelFileVersion = sTmp
    
End Function

Private Function GetArtikelFileSeparator(sFirstLine As String) As String
    
    Dim lPos As Long
    
    lPos = InStr(1, sFirstLine, "Separator:", vbTextCompare)
    If lPos > 0 Then
        GetArtikelFileSeparator = Mid$(sFirstLine, lPos + 10, 1)
    End If
    
End Function

Private Function GetArtikelFileDelimiter(sFirstLine As String) As String
    
    Dim lPos As Long
    
    lPos = InStr(1, sFirstLine, "Delimiter:", vbTextCompare)
    If lPos > 0 Then
        GetArtikelFileDelimiter = Mid$(sFirstLine, lPos + 10, 1)
    End If
    
End Function

Public Function GetSystemUptime() As Double
    'liefert die Uptime in Sekunden
    GetSystemUptime = timeGetTime() / 1000
    
End Function

Public Function GetBOMVersion() As String
    
    With App
        GetBOMVersion = Trim$(CStr(.Major) & "." & CStr(.Minor) & "." & CStr(.Revision) & " " & gsBETASTRING)
    End With
    
End Function

Public Function GetJARVISVersion() As String
    
    GetJARVISVersion = GetFileVersion(App.Path & "\JARVIS-7.exe")
        
End Function

Public Sub CheckUpdate(ByRef oFrm As Form, Optional bQuietOnUpToDate As Boolean = False)
    
    On Error Resume Next
    
    Dim sBuffer As String
    Dim sTmp As String
    Dim sOldBOMVersion As String
    Dim sNewBOMVersion As String
    Dim sNewBetaVersion As String
    Dim sOldKeyVersion As String
    Dim sNewKeyVersion As String
    Dim sBOMDownloadUrl As String
    Dim sBetaDownloadUrl As String
    Dim sKeyDownloadUrl As String
    Dim sBasicUrl As String
    Dim sUpdaterParamBeta As String
    Dim sAvailableRessource As String
    Dim bIsBeta As Boolean
    Dim sMsg As String
    Dim bVersionNotOk As Boolean
    Dim lRet As VbMsgBoxResult
    Dim sOpenUrl As String
    Dim bRes As Boolean
    
    
    oFrm.PanelText oFrm.StatusBar1, 1, "", , , , True
    DoEvents
    
    If Not CheckInternetConnection Then
        oFrm.Ask_Online
        If Not IsOnline Then
            Exit Sub
        End If
    End If
    
    sTmp = gsBOMUrlSF & "BOM/info.php"
    sBuffer = ShortPost(sTmp, "", , , , True)
    
    If sBuffer = "" Then
        sTmp = gsBOMUrlHP & "/bominfo.ini"
        sBuffer = ShortPost(sTmp, "")
    End If
    
    sOldBOMVersion = GetBOMVersion()
    sOldKeyVersion = GetKeywordsFileVersion()
    
    sOpenUrl = gsBOMUrlHP
    
    If sBuffer = "" Then
    
        sMsg = gsarrLangTxt(257) _
            & vbCrLf & gsarrLangTxt(258) _
            & vbCrLf & vbCrLf & "URL: " & gsBOMUrlHP & vbCrLf & vbCrLf _
            & gsarrLangTxt(259)
        bVersionNotOk = True
    Else
        sTmp = gsTempPfad & "\bominfo.ini"
        Call SaveToFileAnsi(sBuffer, sTmp)
        Call INIGetValue(sTmp, "BOM", "VERSION", sNewBOMVersion)
        'Call INIGetValue(sTmp, "KEY", "VERSION", sNewKeyVersion)
        Call INIGetValue(sTmp, "BOM", "FILE", sBOMDownloadUrl)
        'Call INIGetValue(sTmp, "KEY", "FILE", sKeyDownloadUrl)
        Call INIGetValue(sTmp, "COMMON", "DOWNLOAD", sBasicUrl)
        'Call INIGetValue(sTmp, "BOM_Beta", "VERSION", sNewBetaVersion)
        'Call INIGetValue(sTmp, "BOM_Beta", "FILE", sBetaDownloadUrl)
        Call Kill(sTmp)
        
        sBOMDownloadUrl = sBasicUrl & "/" & sBOMDownloadUrl
        'sKeyDownloadUrl = sBasicUrl & "/" & sKeyDownloadUrl
        'sBetaDownloadUrl = sBasicUrl & "/" & sBetaDownloadUrl
        
        If sNewBOMVersion = "" Then
            sMsg = gsarrLangTxt(260) _
                & vbCrLf & gsarrLangTxt(258) _
                & vbCrLf & vbCrLf & "URL: " & gsBOMUrlHP & vbCrLf & vbCrLf _
                & gsarrLangTxt(259)
            bVersionNotOk = True
        Else
        
'            If gbCheckForUpdateBeta And VersionValue(sNewBetaVersion) > VersionValue(sNewBOMVersion) Then
'                sNewBOMVersion = sNewBetaVersion
'                sBOMDownloadUrl = sBetaDownloadUrl
'                bIsBeta = True
'                sUpdaterParamBeta = " /BETA"
'            End If
            
            If VersionValue(sOldBOMVersion) >= VersionValue(sNewBOMVersion) Then
                If VersionValue(sOldKeyVersion) >= VersionValue(sNewKeyVersion) Then
                    'alles ok
                    sMsg = gsarrLangTxt(261) & " V " & sOldBOMVersion & " " & gsarrLangTxt(263)
                    If bQuietOnUpToDate Then
                        If gbUsesModem And gbLastDialupWasManually Then oFrm.Ask_Offline
                        Exit Sub
                    End If
                Else
'                    'KEY nachladen ..
'                    bVersionNotOk = True
'                    sAvailableRessource = gsarrLangTxt(417) & " " & sNewKeyVersion
'                    sMsg = gsarrLangTxt(402) & vbCrLf & vbCrLf _
'                        & gsarrLangTxt(403) & ":" & vbTab & sOldKeyVersion & vbTab & vbCrLf _
'                        & gsarrLangTxt(404) & ":" & vbTab & sNewKeyVersion & vbCrLf & vbCrLf _
'                        & gsarrLangTxt(265)
'                    If sKeyDownloadUrl <> "" Then sOpenUrl = sKeyDownloadUrl
          
                End If
            Else
                'BOM nachladen ..
                bVersionNotOk = True
                sAvailableRessource = gsarrLangTxt(416) & " " & sNewBOMVersion & IIf(bIsBeta, " beta", "")
                sMsg = gsarrLangTxt(261) & " " & gsarrLangTxt(262) & vbCrLf & vbCrLf _
                    & gsarrLangTxt(400) & ":" & vbTab & sOldBOMVersion & vbTab & vbCrLf _
                    & gsarrLangTxt(401) & ":" & vbTab & sNewBOMVersion & IIf(bIsBeta, " beta", "") & vbCrLf & vbCrLf _
                    & gsarrLangTxt(265)
                If sBOMDownloadUrl <> "" Then sOpenUrl = sBOMDownloadUrl
                
            End If
        End If
    End If
    
    gbNewBOMVersionAvailable = CBool(Len(sAvailableRessource))
    
    If bVersionNotOk Then
  
        If bQuietOnUpToDate Then
            'bRes = ShowUpdateBox(sMsg)
            '...
            'frmCountdown.ZeigeCountdown(102, IIf(frmHaupt.WindowState = vbMinimized Or frmHaupt.Visible = False, 2, 1), 30, gsarrLangTxt(405), sMsg, gsarrLangTxt(406), gsarrLangTxt(407), gsarrLangTxt(408), gsarrLangTxt(409), False)
            bRes = ShowUpdateBox(oFrm, [ftUpdateBox], 102, IIf(oFrm.WindowState = vbMinimized Or oFrm.Visible = False, 2, 1), gsarrLangTxt(405), 30, sMsg, gsarrLangTxt(406), gsarrLangTxt(407), gsarrLangTxt(408), gsarrLangTxt(409), False)
            '...
            lRet = IIf(bRes, vbYes, vbNo)
            If Not bRes And gbNewBOMVersionAvailable Then
                oFrm.PanelText oFrm.StatusBar1, 1, gsarrLangTxt(415) & ": " & sAvailableRessource, False, vbYellow, , True
            End If
        Else
            lRet = MsgBox(sMsg, vbQuestion Or vbYesNo, gsarrLangTxt(43))
        End If
        
        If lRet = vbNo Then
            bVersionNotOk = False
        End If
    Else
        MsgBox sMsg, vbInformation, gsarrLangTxt(43)
    End If
    
    If bVersionNotOk Then
'        lRet = 0
'        If gbNewBOMVersionAvailable Then
'            'Updater starten (NoAsk, Backup On, startet BOM - beendet sich)
'            lRet = ExecuteDoc(oFrm.hWnd, "BOMUpdate.exe", "/MODE=AN /Backup=B /DONE=SQ" & sUpdaterParamBeta, True)
'        End If
'
'        If lRet Then
'            'Updater existiert und ist gestartet
'
'            'lieber nochmal Artikel sichern ..
'            Call WriteArtikelCsv2
'
'            'Subclassing entfernen
'            If Not InDevelopment Then
'                Call modSubclass.UnSubclass(frmDummy.hWnd)
'                If gbWheelUsed Then
'                    frmHaupt.MWheel1.DisableWheel
'                End If
'            End If
'
'            'Icon entfernen
'            oFrm.RemoveTrayIcon
'
'            'und wech, den Rest macht der Updater
'            End
'
'        Else
        
        'alte Methode
        Screen.MousePointer = vbHourglass
        Kill gsAppDataPath & "\DoUpdate.exe"
        
        GetPageToFile sOpenUrl, gsAppDataPath & "\DoUpdate.exe"
        
        Call WriteArtikelCsv2
        'Subclassing entfernen
        If Not InDevelopment Then
            Call modSubclass.UnSubclass(frmDummy.hWnd)
            If gbWheelUsed Then
                frmHaupt.MWheel1.DisableWheel
            End If
        End If

        'Icon entfernen
        oFrm.RemoveTrayIcon
        ShellExecute 0, "Open", gsAppDataPath & "\DoUpdate.exe", "/S", "", 1
        
       End
    End If
    
    If gbUsesModem And gbLastDialupWasManually Then oFrm.Ask_Offline
    
End Sub

Public Function GetTaskBarProps(ByVal sProperty As String) As Long
    
    Dim hRect As udtRectX
    Dim lHwnd As Long
    Dim lRes As Long
    
    sProperty = UCase(Trim$(sProperty))
    If sProperty = "" Then
        GetTaskBarProps = 0
        Exit Function
    End If
    
    lHwnd = FindWindow("Shell_traywnd", "")                  ' Handle TBarWindow ermitteln
    lRes = GetWindowRect(lHwnd, hRect)                       ' Fenster d. TaskBar
    With hRect
        Select Case sProperty
            Case "HANDLE"
                GetTaskBarProps = lHwnd
            Case "TOP"
                GetTaskBarProps = .rY1                 ' Pos. oben (inkl. Rand)
            Case "LEFT"
                GetTaskBarProps = .rX1                 ' Pos. links ( inkl.Rand)
            Case "HEIGHT"
                GetTaskBarProps = .rY2 - .rY1     ' Höhe (unten - oben)
            Case "WIDTH"
                GetTaskBarProps = .rX2 - .rX1     ' Breite (rechts - links)
            Case "ALIGN"
                If .rY1 < 1 Then
                    If .rX1 < 1 Then
                        If .rY2 > .rX2 Then GetTaskBarProps = 1 ' links
                        If .rX2 > .rY2 Then GetTaskBarProps = 3 ' oben
                    End If
                End If
                
                If .rY1 < 1 Then
                    If .rX1 > 0 Then GetTaskBarProps = 2            ' rechts
                End If
                
                If .rX1 < 1 Then
                    If .rY1 > 0 Then GetTaskBarProps = 4            ' unten
                End If
            Case Else
        End Select
    End With 'hRect
    
End Function

Public Function GetLongTempPath() As String

  Dim strBuffer As String
  Dim lngResult As Long
  Dim strTempPfad As String
  
  strBuffer = Space(255)
  lngResult = GetTempPath(255, strBuffer)
  
  If lngResult > 0 Then
    strTempPfad = Left(strBuffer, lngResult)
    GetLongTempPath = GetLongPath(strTempPfad)
  End If

End Function

Private Function GetLongPath(ByVal sShortName As String) As String

  Dim sLongName As String, sTemp As String, iSlashPos As Integer

  'Add \ to short name to prevent InStr from failing
  If Right(sShortName, 1) <> "\" Then sShortName = sShortName & "\"

  'Start from 4 to ignore the "[Drive Letter]:\" characters
  iSlashPos = InStr(4, sShortName, "\")

  'Pull out each string between \ character for conversion
  While iSlashPos
    sTemp = Dir(Left$(sShortName, iSlashPos - 1), _
      vbNormal Or vbHidden Or vbSystem Or vbDirectory)
    If sTemp = "" Then
      'Error 52 - Bad File Name or Number
      GetLongPath = ""
      Exit Function
    End If
    sLongName = sLongName & "\" & sTemp
    iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
  Wend

  'Prefix with the drive letter
  GetLongPath = Left$(sShortName, 2) & sLongName

End Function

Public Sub PlaySound(ByVal sPath As String)
    Call sndPlaySoundA(sPath, 1&)
End Sub

'Private Function DebugHack(bIsDebug As Boolean)
'  bIsDebug = True
'  DebugHack = True
'End Function

'Public Function IsIDE() As Boolean
'  Dim bIsDebug As Boolean
'
'  bIsDebug = False
'  Debug.Assert DebugHack(bIsDebug)
'  IsIDE = bIsDebug
'End Function

Public Function StripJavaScript(sHtmlCode As String) As String
    
    On Error GoTo ERROR_HANDLER
    Dim lPos1 As Long
    Dim lPos2 As Long
    Dim sTmp As String
    
    sTmp = sHtmlCode
    If gbStripJS Then
        lPos1 = InStr(1, sTmp, "<script", vbTextCompare)
        Do While lPos1 > 0
            lPos2 = InStr(lPos1, sTmp, "</script>", vbTextCompare)
            If lPos2 > lPos1 Then
                sTmp = Left(sTmp, lPos1 - 1) & Mid(sTmp, lPos2 + 9)
            Else
                Exit Do
            End If
            lPos1 = InStr(1, sTmp, "<script", vbTextCompare)
        Loop
    End If
    StripJavaScript = sTmp
    
  Exit Function
ERROR_HANDLER:
  StripJavaScript = sHtmlCode
  
End Function

Public Sub InitCurrencies()

  Set gcolWeNames = New Collection
  Set gcolWeValues = New Collection

  gcolWeNames.Add "AUD"
  gcolWeNames.Add "CAD"
  gcolWeNames.Add "CHF"
  gcolWeNames.Add "GBP"
  gcolWeNames.Add "USD"
  gcolWeValues.Add grWEDEFAULTAUD, "AUD"
  gcolWeValues.Add grWEDEFAULTCAD, "CAD"
  gcolWeValues.Add grWEDEFAULTCHF, "CHF"
  gcolWeValues.Add grWEDEFAULTGBP, "GBP"
  gcolWeValues.Add grWEDEFAULTUSD, "USD"

End Sub

Public Sub UpdateCurrencies()
    
    Dim lCounter As Long
    Dim sTmp As String
    Dim sKey As String
    Dim sBuffer As String
    Dim sUrl As String
    Dim sPostData As String
    Dim sReferer As String
    Dim lPos As Long
    Dim lPosStart As Long
    Dim lPosEnd As Long
    Dim vntLine As Variant
    Dim vntLines As Variant
    
    If Not CheckInternetConnection Then
        frmHaupt.Ask_Online
        If Not IsOnline Then
            Exit Sub
        End If
    End If
    
    Call InitCurrencies
    
    sUrl = "http://" & gsScript9 & gsScriptCommand9
    sPostData = gsCmdUpdateCurrencies
    sReferer = gsCmdUpdateCurrReferer
    
    sPostData = Replace(sPostData, "%D", Day(MyNow))
    sPostData = Replace(sPostData, "%M", Month(MyNow))
    sPostData = Replace(sPostData, "%Y", Year(MyNow))
    
'    sBuffer = ShortPost(sUrl, sPostData, sReferer)
    sBuffer = ShortPost(sUrl & "?" & sPostData, "", sReferer)
    sBuffer = StripComments(sBuffer)
    
    lPosStart = InStr(1, sBuffer, gsAnsCurrencyStart, vbTextCompare)
    lPosEnd = InStr(lPosStart + 1, sBuffer, gsAnsCurrencyEnd, vbTextCompare)
    If lPosStart > 0 And lPosEnd > 0 And lPosStart < lPosEnd Then
        
        sBuffer = Mid(sBuffer, lPosStart, lPosEnd - lPosStart)
        vntLines = Split(sBuffer, "</tr>")
        For Each vntLine In vntLines
            lPos = InStr(1, vntLine, gsAnsCurrency1)
            If lPos > 0 Then
                sKey = Mid(vntLine, lPos + Len(gsAnsCurrency1), 3)
            Else
                sKey = ""
            End If
            
            If ExistCollectionKey(gcolWeValues, sKey) Then
                sTmp = vntLine
                lPos = InStr(lPos + 1, sTmp, gsAnsCurrency2)
                If lPos > 0 Then
                    sTmp = Mid(sTmp, lPos + Len(gsAnsCurrency2))
                    lPos = InStr(1, sTmp, gsAnsCurrency3)
                    If lPos > 0 Then
                        sTmp = Mid(sTmp, lPos + Len(gsAnsCurrency3))
                        lPos = InStr(1, sTmp, gsAnsCurrency4)
                        If lPos > 0 Then
                            sTmp = Left(sTmp, lPos - 1)
                            If Val(sTmp) > 0 Then
                                gcolWeValues.Remove sKey
                                gcolWeValues.Add CDbl(Val(sTmp)), sKey
                                lCounter = lCounter + 1
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    If lCounter > 0 Then
        sTmp = gsarrLangTxt(410)
        sTmp = Replace(sTmp, "%OK%", lCounter)
        sTmp = Replace(sTmp, "%ALL%", gcolWeValues.Count)
        frmHaupt.PanelText frmHaupt.StatusBar1, 2, sTmp, True, vbGreen
        Call SaveCurrencies
    Else
        frmHaupt.PanelText frmHaupt.StatusBar1, 2, gsarrLangTxt(411), False, vbRed
    End If
    
    If gbUsesModem And gbLastDialupWasManually Then frmHaupt.Ask_Offline
    
End Sub

Private Sub SaveCurrencies()
    
    On Error Resume Next

    Dim sFile As String
    Dim vntWe As Variant
    
    sFile = gsAppDataPath & "\Settings.ini"
    
    Call INIDeleteSection(sFile, "Currency")
    
    For Each vntWe In gcolWeNames
        Call INISetValue(sFile, "Currency", vntWe, gcolWeValues(vntWe))
    Next
    
End Sub

Public Function ExistCollectionKey(inCollenction As Collection, sKey As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    Call inCollenction(sKey)
    ExistCollectionKey = True
    
Done:
On Error GoTo 0
Exit Function

ERROR_HANDLER:
Resume Done
End Function

Private Function UsrAccTest(ByVal sUAcsv As String) As String

    On Error GoTo errhdl
    
    Dim i As Integer, lErrNr As Long
    
    i = 1
    If giUserAnzahl > 0 Then
        Do
            If Trim$(sUAcsv) = gtarrUserArray(i).UaUser Then
                UsrAccTest = Trim$(sUAcsv)
                Exit Do
            Else
                UsrAccTest = ""
            End If
            i = i + 1
        Loop Until i = UBound(gtarrUserArray()) + 1
    Else
        UsrAccTest = ""
    End If
    
errhdl:

lErrNr = Err.Number
Err.Clear
If Not lErrNr = 0 And Not lErrNr = 20 Then
  Resume Next
End If

End Function

Public Function UsrAccToIndex(ByVal sUAccStr As String) As Integer

    Dim i As Integer
    
    UsrAccToIndex = 0
    For i = 1 To giUserAnzahl
        If Trim$(sUAccStr) = Trim$(gtarrUserArray(i).UaUser) Then
            UsrAccToIndex = i
            Exit For
        End If
    Next i
    
End Function

Public Function ItemToIndex(ByVal sItem As String) As Integer

    Dim i As Integer
    
    ItemToIndex = 0
    For i = 1 To giAktAnzArtikel
        If Trim$(sItem) = Trim$(gtarrArtikelArray(i).Artikel) Then
            ItemToIndex = i
            Exit For
        End If
    Next i
    
End Function

Public Function FindeBereich(sTxt As String, sVon As String, sVorBis As String, sBis As String, sRes As String, Optional lOffset As Long = 1) As Long
     
    Dim lBegin As Long
    Dim lVorEnde As Long
    Dim lEnde As Long
    
    lBegin = InStr(lOffset, sTxt, sVon)
    If lBegin > 0 Then
        lBegin = lBegin + Len(sVon)
        If sVorBis > "" Then
            lVorEnde = InStr(lBegin, sTxt, sVorBis)
        Else
            lVorEnde = lBegin
        End If
        
        If lVorEnde > 0 Then
            lVorEnde = lVorEnde + Len(sVorBis)
            lEnde = InStr(lVorEnde, sTxt, sBis)
            If lEnde > 0 Then
                sRes = Mid(sTxt, lBegin, lEnde - lBegin)
                FindeBereich = lEnde + Len(sBis)
            End If
        End If
    End If
    
    'If gbTest Then
        'Print #glFileNrDebug, Date2Str(MyNow) & ""
        'Print #glFileNrDebug, Date2Str(MyNow) & " FindeBereich:"
        'Print #glFileNrDebug, Date2Str(MyNow) & "   von     : " & sVon
        'Print #glFileNrDebug, Date2Str(MyNow) & "   vorbis  : " & sVorBis
        'Print #glFileNrDebug, Date2Str(MyNow) & "   bis     : " & sBis
        'Print #glFileNrDebug, Date2Str(MyNow) & "   Offset  : " & CStr(lOffset)
        'Print #glFileNrDebug, Date2Str(MyNow) & "   len(sTxt): " & CStr(Len(sTxt))
        'Print #glFileNrDebug, Date2Str(MyNow) & "   len(sRes): " & CStr(Len(sRes))
        'Print #glFileNrDebug, Date2Str(MyNow) & "   return  : " & CStr(FindeBereich)
        'Print #glFileNrDebug, Date2Str(MyNow) & ""
    'End If
 
End Function

Public Function GetMetaHttpEquivRefresh(ByVal sTxt As String) As String
    
    Dim i As Long
    Dim sTmp1 As String
    Dim sTmp2 As String
    
    i = 1
    Do While (i > 0)
        i = FindeBereich(sTxt, "<meta ", "", """>", sTmp2, i)
        sTmp1 = sTmp2 & """" 'letztes Anführungszeichen wieder dranbasteln
        If FindeBereich(sTmp1, "http-equiv=""", "", """", sTmp2) > 0 Then
            If LCase(Trim(sTmp2)) = "refresh" Then 'isses auch wirklich ne Weiterleitung?
                If FindeBereich(sTmp1, "content=""0", "", """", sTmp2) > 0 Then
                    sTmp1 = sTmp2 & """" 'nochmal dranbasteln
                    If FindeBereich(sTmp1, "http://", "", """", sTmp2) > 0 Then 'und die URL besorgen
                        GetMetaHttpEquivRefresh = HtmlZeichenConvert("http://" & sTmp2)
                        Exit Do
                    ElseIf FindeBereich(sTxt, "https://", "", """", sTmp2) > 0 Then
                        GetMetaHttpEquivRefresh = HtmlZeichenConvert("https://" & sTmp2)
                        Exit Do
                    End If
                End If
            End If
        End If
    Loop
    
End Function

Private Function StripComments(sTxt As String) As String
    
    Dim lPosStart As Long
    Dim lPosEnd As Long
    
    StripComments = sTxt
    
    Do While InStr(1, StripComments, "<!--") > 0
    
        lPosStart = InStr(1, StripComments, "<!--")
        lPosEnd = InStr(lPosStart, StripComments, "-->")
        
        If lPosEnd > lPosStart Then
            StripComments = Left(StripComments, lPosStart - 1) & Mid(StripComments, lPosEnd + 3)
        End If
        
        If lPosEnd = 0 Then Exit Do
        
    Loop
    
End Function

Private Sub BackupArtikelCsv()
    
    On Error Resume Next
    
    If Len(Dir(gsAppDataPath & "\Artikel.csv")) > 0 Then
        If FileLen(gsAppDataPath & "\Artikel.csv") > 10 Then
            If Len(Dir(gsAppDataPath & "\Artikel.bak")) > 0 Then
                Call Kill(gsAppDataPath & "\Artikel.bak")
            End If
            
            Name gsAppDataPath & "\Artikel.csv" As gsAppDataPath & "\Artikel.bak"
        End If
    End If
    
End Sub

Public Sub RestoreArtikelCsv()

    On Error Resume Next
    
    Dim fs As Long
    
    If Len(Dir(gsAppDataPath & "\Artikel.csv")) > 0 Then fs = FileLen(gsAppDataPath & "\Artikel.csv")
    If fs < 10 Then
        If Len(Dir(gsAppDataPath & "\Artikel.bak")) > 0 Then
            If FileLen(gsAppDataPath & "\Artikel.bak") > 10 Then
                If Len(Dir(gsAppDataPath & "\Artikel.csv")) > 0 Then Kill gsAppDataPath & "\Artikel.csv"
                Name gsAppDataPath & "\Artikel.bak" As gsAppDataPath & "\Artikel.csv"
            End If
        End If
    End If
    
End Sub

Public Function GetDateTimeString() As String

    Static Ticker As Byte
    
    Ticker = (Ticker + 1) Mod 10
    GetDateTimeString = Format(MyNow, "Short Date") & " " & Format(MyNow, "Long Time") & "." & Format((Timer - Int(Timer)) * 1000, "000") & CStr(Ticker)
    
End Function

Public Function ResolveItemUrl(ByVal sUrl As String) As Variant
    
    Dim sBuffer As String
    Dim vntArr As Variant
    Dim lPos As Long
    Dim sTmp As String
    
    ResolveItemUrl = Array()
    If Not LCase(sUrl) Like "http*://*" Then Exit Function
    
    sBuffer = ShortPost(sUrl)
    lPos = 1
    lPos = FindeBereich(sBuffer, gsAnsLinkStart, "", gsAnsLinkEnd, sTmp, lPos)
    vntArr = Array()
    Do While (lPos > 0)
        ReDim Preserve vntArr(UBound(vntArr) - LBound(vntArr) + 1)
        vntArr(UBound(vntArr)) = sTmp
        'DebugPrint sTmp
        lPos = FindeBereich(sBuffer, gsAnsLinkStart, "", gsAnsLinkEnd, sTmp, lPos)
    Loop
    ResolveItemUrl = vntArr
    Call DebugPrint("ResolveItemUrl: " & sUrl & " -> " & (UBound(vntArr) - LBound(vntArr) + 1), 3)

End Function

Public Function GetItemFromUrl(ByVal sUrl As String) As String
    
    On Error Resume Next
    
    Dim sTmp As String
    Dim iPos As Integer
    Dim v As Variant
    
    sTmp = sUrl

    For Each v In Array(gsAnsWatchItem, gsAnsWatchItem2, gsAnsWatchItem3)
        If iPos = 0 Then
            iPos = InStr(1, sTmp, v)
            If iPos > 0 Then iPos = iPos + Len(v)
        End If
    Next
    
    If iPos > 0 Then
        'scheint was angekommen zu sein ;-)
        GetItemFromUrl = GetNumericPart(Mid(sTmp, iPos))
    Else
        If IsNumeric(sTmp) Then
            GetItemFromUrl = GetNumericPart(sTmp)
        End If
    End If
    
    If Val(GetItemFromUrl) < 99999 Then GetItemFromUrl = ""
    
    If GetItemFromUrl = "" Then
        iPos = InStr(1, sTmp, "?")
        If iPos > 0 Then
            sTmp = Left(sTmp, iPos - 1)
            iPos = InStrRev(sTmp, "/")
            If iPos > 0 Then
                sTmp = Mid(sTmp, iPos + 1)
                If IsNumeric(sTmp) Then
                    GetItemFromUrl = GetNumericPart(sTmp)
                End If
            End If
        End If
    End If
    
    If Val(GetItemFromUrl) < 99999 Then GetItemFromUrl = ""
    
    Call DebugPrint("GetItemFromUrl: " & sUrl & " -> " & GetItemFromUrl, 3)
    
End Function

Public Sub DebugPrint(ByVal sTxt As String, Optional iRequiredDebugLevel As Integer = 1)
        
    On Error Resume Next
    
    If giDebugLevel >= iRequiredDebugLevel Then
        
        'Debug.Print Date2Str(MyNow) & "   " & sTxt
        
        With frmDebug.List1
            Call .AddItem(Date2Str(MyNow) & Space$(4) & sTxt)
            If .ListCount > 100 Then .RemoveItem 0
            .Selected(.ListCount - 1) = True
        End With
        
        Call OpenLogfile
        If gbTest Then Print #glFileNrDebug, Date2Str(MyNow) & Space$(3) & App.ThreadID & Space$(3) & sTxt
        Call CloseLogfile
        
    End If
    
    On Error GoTo 0
    
End Sub

Public Sub ShrinkLogfile()
    Const lBlockSize As Long = 4096&
    Static datLastCheck As Date
    
    Dim sTestFile As String
    Dim sTestFileTmp As String
    Dim iFileNrDebugTmp As Integer
    Dim lNewSize As Long
    Dim i As Integer
    Dim sBuffer As String
    
    sTestFile = gsAppDataPath & "\History.log"
    sTestFileTmp = gsAppDataPath & "\History.log.tmp"
    
    If DateDiff("n", datLastCheck, Now) < 5 Then Exit Sub
    
    datLastCheck = Now
    
    If glLogfileMaxSize <= 0 Then Exit Sub
    
    If gbTest Then
        On Error GoTo ERR_EXIT
        If FileLen(sTestFile) > glLogfileMaxSize Then
        
            gbTest = False
            
            If glLogfileShrinkPercent < 1 Then glLogfileShrinkPercent = 1
            If glLogfileShrinkPercent > 99 Then glLogfileShrinkPercent = 99
            
            glFileNrDebug = FreeFile()
            Open sTestFile For Binary Access Read As glFileNrDebug
                iFileNrDebugTmp = FreeFile()
                Open sTestFileTmp For Binary Access Write As iFileNrDebugTmp
                
                    lNewSize = glLogfileMaxSize * ((100 - glLogfileShrinkPercent) / 100)
                    lNewSize = Int(lNewSize / lBlockSize) * lBlockSize + lBlockSize
                    Seek #glFileNrDebug, FileLen(sTestFile) - lNewSize + 1
      
                    sBuffer = String(lBlockSize, " ")
                    
                    For i = 1 To lNewSize / lBlockSize
                        Get #glFileNrDebug, , sBuffer
                        If i = 1 Then
                            If InStr(1, sBuffer, vbCrLf) > 0 Then
                                sBuffer = Mid(sBuffer, InStr(1, sBuffer, vbCrLf) + 2)
                            End If
                            Put #iFileNrDebugTmp, , sBuffer
                            sBuffer = String(lBlockSize, " ")
                        Else
                            Put #iFileNrDebugTmp, , sBuffer
                        End If
                    Next i
                    
                Close iFileNrDebugTmp
            Close glFileNrDebug
            Call Kill(sTestFile)
            Name sTestFileTmp As sTestFile
        End If
ERR_EXIT:
        gbTest = True
    End If

End Sub

Public Sub CloseLogfile()
  
    On Error Resume Next
    Close #glFileNrDebug
    On Error GoTo 0
    
End Sub

Public Sub OpenLogfile()
    
    On Error GoTo ERROR_HANDLER
    
    glFileNrDebug = FreeFile()
    Open gsAppDataPath & "\History.log" For Append Lock Write As #glFileNrDebug
    
    Exit Sub
    
ERROR_HANDLER:

    If Err.Number = 70 Then
        Err.Clear
        Call Sleep(10)
        DoEvents
        Resume
    End If
    
  'Debug.Print "Error opening logfile: " & Err.Description
    
End Sub

Public Sub SetAppDataPath()
    
    On Error Resume Next
    
    '1) Gibts Daten im BOM-Verzeichnis und sind sie schreibbar?
    gsAppDataPath = App.Path
    Open gsAppDataPath & "\Settings.ini" For Input As #1
    Close #1
    If Err.Number = 0 Then ' es gibt eine Settings.ini im BOM-Verzeichnis
        Err.Clear
        
        Open gsAppDataPath & "\Settings.ini" For Append As #1
        Close #1
        If Err.Number = 0 Then ' und ich darf rein schreiben
            Exit Sub ' okay, nehm ich
        End If
    End If
    
    '2) Kann ich die Daten in Anwendungsdaten\BOM ablegen?
    'den Pfad zu Anwendungsdaten besorgen
    If modRegistry.GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AppData", gsAppDataPath) Then
        If gsAppDataPath > "" Then
            gsAppDataPath = gsAppDataPath & "\BOM"
            If Dir(gsAppDataPath, vbDirectory) = "" Then MkDir gsAppDataPath
            If Dir(gsAppDataPath, vbDirectory) > "" Then  ' es gibt (jetzt) Anwendungsdaten\BOM
                Err.Clear
                Open gsAppDataPath & "\Settings.ini" For Append As #1
                Close #1
                If Err.Number = 0 Then ' und ich darf rein schreiben
                    Exit Sub ' okay, nehm ich
                End If
            End If
        End If
    End If
    
    '3) Kann ich die Daten im BOM-Verzeichnis ablegen?
    gsAppDataPath = App.Path
    Err.Clear
    Open gsAppDataPath & "\Settings.ini" For Append As #1
    Close #1
    If Err.Number = 0 Then ' und ich darf rein schreiben
        Exit Sub ' okay, nehm ich
    End If
    
    '4) Kann ich die Daten im Temp-Verzeichnis ablegen?
    gsAppDataPath = GetLongTempPath()
    If gsAppDataPath > "" Then
        gsAppDataPath = gsAppDataPath & "\BOM"
        If Dir(gsAppDataPath, vbDirectory) = "" Then MkDir gsAppDataPath
        If Dir(gsAppDataPath, vbDirectory) > "" Then  ' es gibt (jetzt) Temp\BOM
            Err.Clear
            Open gsAppDataPath & "\Settings.ini" For Append As #1
            Close #1
            If Err.Number = 0 Then ' und ich darf rein schreiben
                Exit Sub ' okay, nehm ich
            End If
        End If
    End If
  
  MsgBox "Can't find a place to store application's data", vbCritical
  
  End  'MD-Marker
  
End Sub

Public Function DelSZ(ByVal sTmp As String) As String

    sTmp = Replace(sTmp, vbBack, " ", , , vbTextCompare)
    sTmp = Replace(sTmp, vbTab, " ", , , vbTextCompare)
    sTmp = Replace(sTmp, vbLf, " ", , , vbTextCompare)
    sTmp = Replace(sTmp, vbCr, " ", , , vbTextCompare)
    DelSZ = Trim$(sTmp)
    
End Function

Public Function MakeTempFile(Optional sPattern As String = "%d.tmp") As String
    
    On Error GoTo ERROR_HANDLER
    
    Dim sTmp As String
    Dim iFileNr As Integer
    Dim bOk As Boolean
    
    Randomize Timer
    
    iFileNr = FreeFile()
    
    Do While Not bOk
        
        sTmp = GetLongTempPath & "\" & sPattern
        sTmp = Replace(sTmp, "%d", Hex(Rnd(1) * 1000000000 + 1000000000))
        If Dir(sTmp) = "" Then
            Open sTmp For Binary Access Write Lock Read Write As iFileNr
                bOk = CBool(FileLen(sTmp) = 0)
                Put #iFileNr, , " "
            Close iFileNr
            
            Open sTmp For Output As iFileNr
            Close iFileNr
        End If
        
ERROR_OCCURED:
            
    Loop
    
    MakeTempFile = sTmp
Exit Function
    
ERROR_HANDLER:
    Err.Clear
    Resume ERROR_OCCURED
    
End Function

Public Sub SaveToFileAnsi(sTxt As String, sFileName As String)
    
    On Error Resume Next
    
    Dim iFileNr As Integer
    
    Call Kill(sFileName)
    'Create file for the page, bin cause better is dat
    iFileNr = FreeFile()
  Open sFileName For Binary Access Write As #iFileNr
  Put #iFileNr, , sTxt
  
errExit:
  Close #iFileNr

End Sub

Public Sub SaveToFile(ByRef Contents As String, ByRef FileName As String)
        Dim FileNumber As Integer
        Dim Buffer() As Byte
        
10      On Error GoTo Proc_Error
20      If FileExists(FileName) Then Kill FileName
        
30      If Len(Contents) > 0 Then
40        Buffer = ConvertToUTF8(Contents)
50        If Not IsDimmed(Buffer) Then
60          DebugPrint "SaveToFile: ConvertToUTF8 -> Returned empty byte array (DataLen=" & Len(Contents) & ")"
70          Exit Sub
80        Else
90          FileNumber = FreeFile
100         Open FileName For Binary Access Write As #FileNumber
110         Put #FileNumber, , Buffer
120         Close #FileNumber
130       End If
140     End If
150     Exit Sub
        
Proc_Error:
170     DebugPrint "error in sub SaveToFile: " & Err.Description & " Line: " & Erl
End Sub


Public Function IsDimmed(myArray As Variant) As Boolean
  On Error GoTo ErrHandler
  IsDimmed = UBound(myArray) >= LBound(myArray)

ErrHandler:
  ' obviously not dimensioned yet...
End Function

Public Function ConvertToUTF8(ByRef Source As String) As Byte()
  Dim Length As Long
  Dim Pointer As Long
  Dim Size As Long
  Dim Buffer() As Byte
  
  If Len(Source) > 0 Then
    Length = Len(Source)
    Pointer = StrPtr(Source)
    Size = WideCharToMultiByte(CP_UTF8, 0, Pointer, Length, 0, 0, 0, 0)
    If Size > 0 Then
      ReDim Buffer(0 To Size - 1)
      
      WideCharToMultiByte CP_UTF8, 0, Pointer, Length, VarPtr(Buffer(0)), Size, 0, 0
      ConvertToUTF8 = Buffer
    End If
  End If
  
End Function

Public Function FileExists(Datei As String) As Boolean
  
    If Datei = "" Then
        FileExists = False
        Exit Function
    End If
  
  On Error Resume Next
  FileExists = Dir$(Datei) <> ""
  FileExists = FileExists And Err = 0
  On Error GoTo 0
End Function

Public Function ReadFromFile(ByVal sFileName As String, Optional ByVal bBinary As Boolean = False) As String
    
    On Error Resume Next
    Dim iFileNr As Integer
    Dim sTmp As String
    Dim b() As Byte
    
    If Dir(sFileName) = "" Then Exit Function
    
    iFileNr = FreeFile()
    Open sFileName For Binary Access Read As #iFileNr
        If bBinary Then
            ReDim b(0 To FileLen(sFileName) - 1) As Byte
            
            Get #iFileNr, , b()
            ReadFromFile = b()
        Else
            sTmp = Space$(FileLen(sFileName))
            Get #iFileNr, , sTmp
            ReadFromFile = sTmp
        End If

errExit:
    Close #iFileNr
    
End Function

Public Sub SendCallBackData(ByVal lHwnd As Long, sData As String)
    
    Dim tCDS As COPYDATASTRUCT, b() As Byte, lR As Long
    
    b = StrConv(sData, vbFromUnicode)
    tCDS.dwData = 0
    tCDS.cbData = UBound(b()) + 1
    tCDS.lpData = VarPtr(b(0))
    
    'Give in if the existing app is not responding:
    lR = SendMessageTimeout(lHwnd, WM_COPYDATA, 0, tCDS, SMTO_NORMAL, 5000, lR)
    
End Sub

Public Function DumpAllUser() As String
    
    Dim sTmp As String
    Dim i As Integer
    
    For i = LBound(gtarrUserArray()) To UBound(gtarrUserArray())
        If gtarrUserArray(i).UaUser > "" Then
            sTmp = sTmp & gtarrUserArray(i).UaUser & vbCrLf
        End If
    Next i
    DumpAllUser = sTmp
    
End Function

Public Function DumpAllItems(sFromID As String) As String

    Dim sGlobalSeparator As String
    Dim sGlobalDelimiter As String
    Dim bGlobalUseIsoDate As Boolean
    Dim bGlobalUseUnixDate As Boolean
    Dim bGlobalSuppressHeader As Boolean
    
    sGlobalSeparator = gsSeparator
    sGlobalDelimiter = gsDelimiter
    bGlobalUseIsoDate = gbUseIsoDate
    bGlobalUseUnixDate = gbUseUnixDate
    bGlobalSuppressHeader = gbSuppressHeader
    
    gbUseIsoDate = False
    gbUseUnixDate = True
    gsSeparator = vbTab
    gsDelimiter = ""
    gbSuppressHeader = True
    
    If sFromID > "" Then DumpAllItems = CStr(GetChangeID()) & vbCrLf
    DumpAllItems = DumpAllItems & BuildArtikelCSV2(Val(sFromID))
    
    gsDelimiter = sGlobalDelimiter
    gsSeparator = sGlobalSeparator
    gbUseIsoDate = bGlobalUseIsoDate
    gbUseUnixDate = bGlobalUseUnixDate
    gbSuppressHeader = bGlobalSuppressHeader
    
End Function

'======================
' Send output to STDOUT
'======================
'
Public Sub Send(s As String)
    
    Dim llResult As Long
    Call WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), s, Len(s), llResult, ByVal 0&)
    
End Sub

' wir brauchen diese 'nachgebauten' Funktionen, weil die Originale unter WINE (Linux) buggy sind!
Public Function myDateSerial(ByVal iYear As Integer, ByVal iMonth As Integer, ByVal iDay As Integer) As Date
    
    myDateSerial = #1/1/2000#
    myDateSerial = DateAdd("yyyy", iYear - 2000, myDateSerial)
    myDateSerial = DateAdd("m", iMonth - 1, myDateSerial)
    myDateSerial = DateAdd("d", iDay - 1, myDateSerial)
    
End Function

Public Function myTimeSerial(ByVal iHour As Integer, ByVal iMinute As Integer, ByVal iSecond As Integer) As Date
    
    myTimeSerial = DateAdd("h", iHour, myTimeSerial)
    myTimeSerial = DateAdd("n", iMinute, myTimeSerial)
    myTimeSerial = DateAdd("s", iSecond, myTimeSerial)

End Function

Public Function GetRgbHexFromColor(lColor As Long) As String
    
    Dim r As String
    Dim g As String
    Dim b As String
    
    r = Hex(lColor And &HFF)
    g = Hex(Int(lColor / &H100) And &HFF)
    b = Hex(Int(lColor / &H10000) And &HFF)
    
    If Len(r) = 1 Then r = "0" & r
    If Len(g) = 1 Then g = "0" & g
    If Len(b) = 1 Then b = "0" & b
    
    GetRgbHexFromColor = r & g & b
    
End Function

Public Function GetColorFromRgbHex(sHexValue As String) As Long
    
    sHexValue = Left(Trim(sHexValue), 6)
    sHexValue = String(6 - Len(sHexValue), "0") & sHexValue
    GetColorFromRgbHex = RGB("&h" & Left(sHexValue, 2), "&h" & Mid(sHexValue, 3, 2), "&h" & Right(sHexValue, 2))
    
End Function

Public Function GetChangeID() As Long
    
    Static lCnt As Long
    
    If lCnt = 0 Then lCnt = Date2UnixDate(Now)
    lCnt = lCnt + 1
    GetChangeID = lCnt
    
End Function

Public Function MyLoadResPicture(ByVal vntID As Variant, ByVal iSize As Integer) As IPictureDisp
    
    On Error GoTo ERROR_HANDLER
    
    If gsIconSet > "" Then
        If Dir(App.Path & "\Icons\" & gsIconSet & "\Icon_" & CStr(vntID) & ".ico") > "" Then
            Set MyLoadResPicture = MyLoadPicture(App.Path & "\Icons\" & gsIconSet & "\Icon_" & CStr(vntID) & ".ico", iSize)
            Exit Function
        End If
    End If
    
    If Dir(gsAppDataPath & "\Icon_" & CStr(vntID) & ".ico") > "" Then
        Set MyLoadResPicture = MyLoadPicture(gsAppDataPath & "\Icon_" & CStr(vntID) & ".ico", iSize)
        Exit Function
    End If
    
    If Dir(App.Path & "\Icon_" & CStr(vntID) & ".ico") > "" Then
        Set MyLoadResPicture = MyLoadPicture(App.Path & "\Icon_" & CStr(vntID) & ".ico", iSize)
        Exit Function
    End If
    
ERROR_HANDLER:
    
    If Err.Number <> 0 Then Err.Clear
    On Error Resume Next
    Set MyLoadResPicture = LoadResPicture(CInt(vntID), vbResIcon)
    
End Function

Private Function MyLoadPicture(ByVal sFileName As String, ByVal iSize As Integer) As IPictureDisp
    
    Dim lHIcon As Long
    
    'Load an icon called Test.Ico from the directory:
    'If the icon contains more than one size of image,
    'set cx and cy to the width and height to load
    'the appropriate image in:
    lHIcon = LoadImage(App.hInstance, sFileName, IMAGE_ICON, CLng(iSize), CLng(iSize), LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS)
    'Set the picture to this icon:
    Set MyLoadPicture = IconToPicture(lHIcon)
    
End Function

Private Function IconToPicture(ByVal lHIcon As Long) As IPicture
       
    Dim oNewPic As Picture
    Dim tPicConv As PictDesc
    Dim IGuid As Guid
    
    
    If lHIcon <> 0 Then
        With tPicConv
            .cbSizeofStruct = Len(tPicConv)
            .picType = vbPicTypeIcon
            .hImage = lHIcon
        End With
        
        'Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        With IGuid
            .Data1 = &H7BF80980
            .Data2 = &HBF32
            .Data3 = &H101A
            .Data4(0) = &H8B
            .Data4(1) = &HBB
            .Data4(2) = &H0
            .Data4(3) = &HAA
            .Data4(4) = &H0
            .Data4(5) = &H30
            .Data4(6) = &HC
            .Data4(7) = &HAB
        End With
        Call OleCreatePictureIndirect(tPicConv, IGuid, True, oNewPic)
        
        Set IconToPicture = oNewPic
    End If
    
End Function

Public Function ShowUpdateBox(ByRef oFrm As Form, eFrmTyp As FrmTypEnum, lIconResNr As Long, lFrmPos, sFrmTitle As String, lCnt As Long, sTxt1 As String, sTxt2 As String, sTxt3 As String, sBtn1 As String, sBtn2 As String, bDefRetValue As Boolean) As Boolean
    
    If InStr(1, sTxt1, vbTab, vbBinaryCompare) Then
        sTxt1 = Replace(sTxt1, vbTab, " ", 1, -1, vbBinaryCompare)
    End If
    
    Call frmCountdown.InitFrm(eFrmTyp, lIconResNr, lFrmPos, sFrmTitle, lCnt, sTxt1, sTxt2, sTxt3, sBtn1, sBtn2, bDefRetValue)
    
    Load frmCountdown
    frmCountdown.Show vbModal, oFrm
    ShowUpdateBox = frmCountdown.FrmRetValue
    Unload frmCountdown

End Function

Public Function CDblSave(vVal As Variant, Optional fDefaultValue As Double = 0) As Double

  On Error GoTo ERROR_HANDLER
  CDblSave = CDbl(vVal)
  Exit Function
ERROR_HANDLER:
  CDblSave = fDefaultValue

End Function

Public Function FormLoaded(sFormName As String) As Boolean

    Dim oFrm As Form
    For Each oFrm In Forms
        If oFrm.Name = sFormName Then
            FormLoaded = True
            Exit For
        End If
    Next
    
End Function

Function getFile(ByVal sPfad As String) As String
 
Dim X As Long

 Do While InStr(X + 1, sPfad, "\") > 0
   X = X + 1
 Loop
   
 getFile = Mid(sPfad, X + 1)

End Function

Function getPath(ByVal sPfad As String) As String
 
Dim X As Long

 Do While InStr(X + 1, sPfad, "\") > 0
   X = X + 1
 Loop
   
 getPath = Left(sPfad, X)

End Function

Public Function GetScreenWidth() As Long
  GetScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN) * Screen.TwipsPerPixelX
End Function

Public Function GetScreenHeight() As Long
  GetScreenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN) * Screen.TwipsPerPixelY
End Function

Public Function GetFileVersion(ByVal FileName As String) As String
Dim nDummy As Long
Dim sBuffer()         As Byte
Dim nBufferLen        As Long
Dim lplpBuffer       As Long
Dim udtVerBuffer      As VS_FIXEDFILEINFO
Dim puLen     As Long
      
   nBufferLen = GetFileVersionInfoSize(FileName, nDummy)
   
   If nBufferLen > 0 Then
   
        ReDim sBuffer(nBufferLen) As Byte
        Call GetFileVersionInfo(FileName, 0&, nBufferLen, sBuffer(0))
        Call VerQueryValue(sBuffer(0), "\", lplpBuffer, puLen)
        Call CopyMemory(udtVerBuffer, ByVal lplpBuffer, Len(udtVerBuffer))
        
        GetFileVersion = udtVerBuffer.dwFileVersionMSh & "." & udtVerBuffer.dwFileVersionMSl & "." & udtVerBuffer.dwFileVersionLSl
  
    End If
End Function

Public Function GetMainVersion(ByVal FileName As String) As String
Dim nDummy As Long
Dim sBuffer()         As Byte
Dim nBufferLen        As Long
Dim lplpBuffer       As Long
Dim udtVerBuffer      As VS_FIXEDFILEINFO
Dim puLen     As Long
      
   nBufferLen = GetFileVersionInfoSize(FileName, nDummy)
   
   If nBufferLen > 0 Then
   
        ReDim sBuffer(nBufferLen) As Byte
        Call GetFileVersionInfo(FileName, 0&, nBufferLen, sBuffer(0))
        Call VerQueryValue(sBuffer(0), "\", lplpBuffer, puLen)
        Call CopyMemory(udtVerBuffer, ByVal lplpBuffer, Len(udtVerBuffer))
        
        GetMainVersion = udtVerBuffer.dwFileVersionMSh
  
    End If
End Function

Public Function maskString(stringToMask) As String
    
    maskString = stringToMask
    
    ' = durch ¥¥¥ ersetzen
    maskString = Replace(maskString, "=", Chr(165) & Chr(165) & Chr(165))
    ' Leerzeichen durch ±±± ersetzen
    maskString = Replace(maskString, " ", Chr(177) & Chr(177) & Chr(177))
    ' " durch ÐÐÐ ersetzen
    maskString = Replace(maskString, """", Chr(208) & Chr(208) & Chr(208))
    
End Function


Public Function GetSpecialFolderPath(ByVal FolderID As spfSpecialFolderConstants) As String

  Dim nItemList As ITEMIDLIST
  Dim nPath As String
  
  Const NOERROR = 0
  
  If SHGetSpecialFolderLocation(0, FolderID, nItemList) = NOERROR Then
    nPath = Space$(260)
    If SHGetPathFromIDList(nItemList.mkid.cb, nPath) <> 0 Then
      GetSpecialFolderPath = Left$(nPath, InStr(nPath, vbNullChar) - 1)
    End If
  End If
End Function

