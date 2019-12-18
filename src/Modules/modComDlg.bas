Attribute VB_Name = "modComDlg"
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
'********************************************************************************
'*      Module ComDlg                                                           *
'********************************************************************************
'* Die CommonDialoge über WindowsAPI, 3rdPartySource !!!                        *
'********************************************************************************

'ChooseColor structure and function declarations *************************

#If Win32 Then
   Type CHOOSECOLOR
      lStructSize As Long
      hwndOwner As Long
      hInstance As Long
      rgbResult As Long
      lpCustColors As Long
      flags As Long
      lCustData As Long
      lpfnHook As Long
      lpTemplateName As String
   End Type
   Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
#Else
   Type CHOOSECOLOR
      lStructSize As Long
      hwndOwner As Integer
      hInstance  As Integer
      rgbResult As Long
      lpCustColors As Long
      flags As Long
      lCustData As Long
      lpfnHook As Long
      lpTemplateName As Long
   End Type
   Declare Function CHOOSECOLOR Lib "COMMDLG.DLL" Alias "ChooseColor" (pChoosecolor As CHOOSECOLOR) As Integer
#End If

Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_RGBINIT = &H1
Public Const CC_SHOWHELP = &H8

' File Open/Save structures and declarations *****************************
#If Win32 Then
   Type OPENFILENAME
      lStructSize As Long         'Same
      hwndOwner As Long           'Was Integer
      hInstance As Long           'Was Integer
      lpstrFilter As String       'Was Long
      lpstrCustomFilter As String 'Was Long
      nMaxCustFilter As Long      'Same
      nFilterIndex As Long        'Same
      lpstrFile As String         'Was Long
      nMaxFile As Long            'Same
      lpstrFileTitle As String    'Was Long
      nMaxFileTitle As Long       'Same
      lpstrInitialDir As String   'Was Long
      lpstrTitle As String        'Was Long
      flags As Long               'Same
      nFileOffset As Integer      'Same
      nFileExtension As Integer   'Same
      lpstrDefExt As String       'Was Long
      lCustData As Long           'Same
      lpfnHook As Long            'Same
      lpTemplateName As String    'Was long
   End Type
   Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
   Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
   Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
#Else
   Type OPENFILENAME
      lStructSize As Long
      hwndOwner As Integer
      hInstance As Integer
      lpstrFilter As Long
      lpstrCustomFilter As Long
      nMaxCustFilter As Long
      nFilterIndex As Long
      lpstrFile As Long
      nMaxFile As Long
      lpstrFileTitle As Long
      nMaxFileTitle As Long
      lpstrInitialDir As Long
      lpstrTitle As Long
      flags As Long
      nFileOffset As Integer
      nFileExtension As Integer
      lpstrDefExt As Long
      lCustData As Long
      lpfnHook As Long
      lpTemplateName As Long
   End Type
   Declare Function GetOpenFileName Lib "COMMDLG.DLL" (pOpenfilename As OPENFILENAME) As Integer
   Declare Function GetSaveFileName Lib "COMMDLG.DLL" (pOpenfilename As OPENFILENAME) As Integer
   Declare Function GetFileTitle Lib "COMMDLG.DLL" (ByVal FName As String, ByVal Title As String, Size As Integer)
#End If

Public Const OFN_ALLOWMULTISELECT = &H200      'See Help Note for LFN Behavior
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000            'Windows 95 Only
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000          'Windows 95 Only
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000 'Windows 95 Only
Public Const OFN_NOLONGNAMES = &H40000         'Not Referenced in Help!
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10

'ChooseColor structure and function declarations *************************

'#If Win32 Then
'   Type CHOOSECOLOR
'      lStructSize As Long
'      hwndOwner As Long
'      hInstance As Long
'      rgbResult As Long
'      lpCustColors As Long
'      flags As Long
'      lCustData As Long
'      lpfnHook As Long
'      lpTemplateName As String
'   End Type
'   Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
'#Else
'   Type CHOOSECOLOR
'      lStructSize As Long
'      hwndOwner As Integer
'      hInstance  As Integer
'      rgbResult As Long
'      lpCustColors As Long
'      flags As Long
'      lCustData As Long
'      lpfnHook As Long
'      lpTemplateName As Long
'   End Type
'   Declare Function CHOOSECOLOR Lib "COMMDLG.DLL" Alias "ChooseColor" (pChoosecolor As CHOOSECOLOR) As Integer
'#End If
'
'Public Const CC_ENABLEHOOK = &H10
'Public Const CC_ENABLETEMPLATE = &H20
'Public Const CC_ENABLETEMPLATEHANDLE = &H40
'Public Const CC_FULLOPEN = &H2
'Public Const CC_PREVENTFULLOPEN = &H4
'Public Const CC_RGBINIT = &H1
'Public Const CC_SHOWHELP = &H8
'
'' FONT STUFF
'Global Const LF_FACESIZE = 32
'
'#If Win32 Then
'   Type LOGFONT
'      lfHeight As Long
'      lfWidth As Long
'      lfEscapement As Long
'      lfOrientation As Long
'      lfWeight As Long
'      lfItalic As Byte
'      lfUnderline As Byte
'      lfStrikeOut As Byte
'      lfCharSet As Byte
'      lfOutPrecision As Byte
'      lfClipPrecision As Byte
'      lfQuality As Byte
'      lfPitchAndFamily As Byte
'      lfFaceName(LF_FACESIZE) As Byte
'   End Type
'#Else
'   Type LOGFONT
'      lfHeight As Integer
'      lfWidth As Integer
'      lfEscapement As Integer
'      lfOrientation As Integer
'      lfWeight As Integer
'      lfItalic As String * 1
'      lfUnderline As String * 1
'      lfStrikeOut As String * 1
'      lfCharSet As String * 1
'      lfOutPrecision As String * 1
'      lfClipPrecision As String * 1
'      lfQuality As String * 1
'      lfPitchAndFamily As String * 1
'      lfFaceName As String * LF_FACESIZE
'   End Type
'#End If
'
'Global Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
'Public Const CCHDEVICENAME = 32
'Public Const CCHFORMNAME = 32
'
'Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'
''ChooseFont structure and function declarations *************************
'
'#If Win32 Then
'   Type ChooseFont
'      lStructSize As Long
'      hwndOwner As Long          '  caller's window handle
'      hdc As Long                '  printer DC/IC or NULL
'      lpLogFont As Long          '  ptr. to a LOGFONT struct
'      iPointSize As Long         '  10 * size in points of selected font
'      flags As Long              '  enum. type flags
'      rgbColors As Long          '  returned text color
'      lCustData As Long          '  data passed to hook fn.
'      lpfnHook As Long           '  ptr. to hook function
'      lpTemplateName As String   '  custom template name
'      hInstance As Long          '  instance handle of.EXE that
'                                 '    contains cust. dlg. template
'      lpszStyle As String        '  return the style field here
'                                 '  must be LF_FACESIZE or bigger
'      nFontType As Integer       '  same value reported to the EnumFonts
'                                 '    call back with the extra FONTTYPE_
'                                 '    bits added
'      MISSING_ALIGNMENT As Integer
'      nSizeMin As Long           '  minimum pt size allowed &
'      nSizeMax As Long           '  max pt size allowed if
'                                 '    CF_LIMITSIZE is used
'   End Type
'   Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
'   Declare Function VarPtr Lib "VB40032.dll" (lpVoid As Any) As Long 'Secret VB Runtime function
'#Else
'   Type ChooseFont
'      lStructSize As Long
'      hwndOwner As Integer
'      hdc As Integer
'      lpLogFont As Long
'      iPointSize As Integer
'      flags As Long
'      rgbColors As Long
'      lCustData As Long
'      lpfnHook As Long 'Integer
'      lpTemplateName As Long
'      hInstance  As Integer
'      lpszStyle As Long
'      nFontType As Integer
'      nSizeMin As Integer
'      nSizeMax As Integer
'   End Type
'   Declare Function ChooseFont Lib "COMMDLG.DLL" (pChoosefont As ChooseFont) As Integer
'#End If
'
'Public Const CF_ANSIONLY = &H400&
'Public Const CF_APPLY = &H200&
'Public Const CF_EFFECTS = &H100&
'Public Const CF_ENABLEHOOK = &H8&
'Public Const CF_ENABLETEMPLATE = &H10&
'Public Const CF_ENABLETEMPLATEHANDLE = &H20&
'Public Const CF_FIXEDPITCHONLY = &H4000&
'Public Const CF_FORCEFONTEXIST = &H10000
'Public Const CF_INITTOLOGFONTSTRUCT = &H40&
'Public Const CF_LIMITSIZE = &H2000&
'Public Const CF_NOFACESEL = &H80000
'Public Const CF_NOSCRIPTSEL = &H800000
'Public Const CF_NOSIMULATIONS = &H1000&
'Public Const CF_NOSIZESEL = &H200000
'Public Const CF_NOSTYLESEL = &H100000
'Public Const CF_NOVECTORFONTS = &H800&
'Public Const CF_NOVERTFONTS = &H1000000
'Public Const CF_OWNERDISPLAY = &H80
'Public Const CF_PRINTERFONTS = &H2
'Public Const CF_SCALABLEONLY = &H20000
'Public Const CF_SCREENFONTS = &H1
'Public Const CF_SCRIPTSONLY = CF_ANSIONLY
'Public Const CF_SELECTSCRIPT = &H400000
'Public Const CF_SHOWHELP = &H4&
'Public Const CF_TTONLY = &H40000
'Public Const CF_USESTYLE = &H80&
'Public Const CF_WYSIWYG = &H8000
'Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
'
'Global Const SIMULATED_FONTTYPE = &H8000
'Global Const PRINTER_FONTTYPE = &H4000
'Global Const SCREEN_FONTTYPE = &H2000
'Global Const BOLD_FONTTYPE = &H100
'Global Const ITALIC_FONTTYPE = &H200
'Global Const REGULAR_FONTTYPE = &H400
'
'Global Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1) 'WM_USER + 1
'
'Global Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
'Global Const SHAREVISTRING = "commdlg_ShareViolation"
'Global Const FILEOKSTRING = "commdlg_FileNameOK"
'Global Const COLOROKSTRING = "commdlg_ColorOK"
'Global Const SETRGBSTRING = "commdlg_SetRGBColor"
'Global Const FINDMSGSTRING = "commdlg_FindReplace"
'Global Const HELPMSGSTRING = "commdlg_help"
'
'Global Const CD_LBSELNOITEMS = -1
'Global Const CD_LBSELCHANGE = 0
'Global Const CD_LBSELSUB = 1
'Global Const CD_LBSELADD = 2
'
''Printer related structures and function declarations ********************
'
'#If Win32 Then
'   Type PRINTDLG
'      lStructSize As Long
'      hwndOwner As Long
'      hDevMode As Long
'      hDevNames As Long
'      hdc As Long
'      flags As Long
'      nFromPage As Integer
'      nToPage As Integer
'      nMinPage As Integer
'      nMaxPage As Integer
'      nCopies As Integer
'      hInstance As Long
'      lCustData As Long
'      lpfnPrintHook As Long
'      lpfnSetupHook As Long
'      lpPrintTemplateName As String
'      lpSetupTemplateName As String
'      hPrintTemplate As Long
'      hSetupTemplate As Long
'   End Type
'   Declare Function PRINTDLG Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG) As Long
'#Else
'   Type PRINTDLG
'      lStructSize As Long
'      hwndOwner As Integer
'      hDevMode As Integer
'      hDevNames As Integer
'      hdc As Integer
'      flags As Long
'      nFromPage As Integer
'      nToPage As Integer
'      nMinPage As Integer
'      nMaxPage As Integer
'      nCopies As Integer
'      hInstance As Integer
'      lCustData As Long
'      lpfnPrintHook As Long
'      lpfnSetupHook As Long
'      lpPrintTemplateName As Long
'      lpSetupTemplateName As Long
'      hPrintTemplate As Integer
'      hSetupTemplate As Integer
'   End Type
'   Declare Function PRINTDLG Lib "COMMDLG.DLL" Alias "PrintDlg" (pPrintdlg As PRINTDLG) As Integer
'#End If
'
'Global Const PD_ALLPAGES = &H0
'Global Const PD_SELECTION = &H1
'Global Const PD_PAGENUMS = &H2
'Global Const PD_NOSELECTION = &H4
'Global Const PD_NOPAGENUMS = &H8
'Global Const PD_COLLATE = &H10
'Global Const PD_PRINTTOFILE = &H20
'Global Const PD_PRINTSETUP = &H40
'Global Const PD_NOWARNING = &H80
'Global Const PD_RETURNDC = &H100
'Global Const PD_RETURNIC = &H200
'Global Const PD_RETURNDEFAULT = &H400
'Global Const PD_SHOWHELP = &H800
'Global Const PD_ENABLEPRINTHOOK = &H1000
'Global Const PD_ENABLESETUPHOOK = &H2000
'Global Const PD_ENABLEPRINTTEMPLATE = &H4000
'Global Const PD_ENABLESETUPTEMPLATE = &H8000
'Global Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
'Global Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
'Global Const PD_USEDEVMODECOPIES = &H40000
'Global Const PD_DISABLEPRINTTOFILE = &H80000
'Global Const PD_HIDEPRINTTOFILE = &H100000
'
'Type DEVNAMES                 'Same in Win16 and Win32
'    wDriverOffset As Integer
'    wDeviceOffset As Integer
'    wOutputOffset As Integer
'    wDefault As Integer
'End Type
'
'Global Const DN_DEFAULTPRN = &H1
'
''retrieves error value
'
'#If Win32 Then
'   Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
'#Else
'   Declare Function CommDlgExtendedError Lib "COMMDLG.DLL" () As Long
'#End If

'************************* end of Common Dialogs Declares ************

'GLOBAL MEMORY Stuff
Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Declare Sub MemoryCopy Lib "kernel32.dll" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal dwBytes As Long)

Global Const GMEM_MOVEABLE = &H2
Global Const GMEM_ZEROINIT = &H40
Global Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

'PRINTER stuff
'Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
'Type DEVMODE
'   dmDeviceName As String * CCHDEVICENAME
'   dmSpecVersion As Integer
'   dmDriverVersion As Integer
'   dmSize As Integer
'   dmDriverExtra As Integer
'   dmFields As Long
'   dmOrientation As Integer
'   dmPaperSize As Integer
'   dmPaperLength As Integer
'   dmPaperWidth As Integer
'   dmScale As Integer
'   dmCopies As Integer
'   dmDefaultSource As Integer
'   dmPrintQuality As Integer
'   dmColor As Integer
'   dmDuplex As Integer
'   dmYResolution As Integer
'   dmTTOption As Integer
'   dmCollate As Integer
'   dmFormName As String * CCHFORMNAME
'   dmUnusedPadding As Integer
'   dmBitsPerPel As Integer
'   dmPelsWidth As Long
'   dmPelsHeight As Long
'   dmDisplayFlags As Long
'   dmDisplayFrequency As Long
'End Type

'Function Dialog_Append$(F As Form, ByVal ftitle$, fpfad$, ByVal ffile$, ByVal ffilter$, ByVal fext$)
'
'Dialog_Append$ = DialogOpenSave$(F, ftitle$, fpfad$, ffile$, ffilter$, fext$, 2)
'
'End Function
'
'Function Dialog_Color%(F As Form, CurrColor&)
'
'Dim C As CHOOSECOLOR
'Dim Address As Long
'#If Win32 Then
'    Dim Memhandle As Long, wSize As Long
'#End If
'
'ReDim ClrArray(15) As Long    ' Holds Custom Colors
'
'wSize = Len(ClrArray(0)) * 16 ' Size of Memory Block
'
'Memhandle = GlobalAlloc(GHND, wSize)
'If Memhandle = 0 Then Dialog_Color% = False: Exit Function
'Address = GlobalLock(Memhandle)
'
'For i& = 0 To UBound(ClrArray): ClrArray(i&) = &HFFFFFF: Next
'
'Call MemoryCopy(ByVal Address, ClrArray(0), wSize)
'
'C.lStructSize = Len(C)
'C.hwndOwner = F.hwnd
'C.lpCustColors = Address
'C.rgbResult = CurrColor&
'C.flags = CC_RGBINIT Or CC_FULLOPEN
'
'result = CHOOSECOLOR(C)
'
'Call MemoryCopy(ClrArray(0), ByVal Address, wSize)
'
'ok = GlobalUnlock(Memhandle)    'Free The Memory
'ok = GlobalFree(Memhandle)
'
'If result = 0 Then Dialog_Color% = False: Exit Function
'
'CurrColor& = C.rgbResult
'
'Dialog_Color% = True
'
'End Function
'
'Function Dialog_Font%(F As Form, ffont$, fsize%, fstyle$, fFlags&)
'
'Dim a As ChooseFont
'Dim LF As LOGFONT
'Dim Address As Long
'Dim Memhandle As Long
'Dim FaceNameString As String
'
'#If Win32 Then
'    Dim result As Long
'#End If
'
'LF.lfHeight = fsize% / (72 / GetDeviceCaps(F.hdc, LOGPIXELSY)) * -1
'If InStr(fstyle$, "b") Then LF.lfWeight = 700 Else lFont.lfWeight = 300
'#If Win16 Then
'   If InStr(fstyle$, "i") Then LF.lfItalic = Chr$(1)
'   If InStr(fstyle$, "s") Then LF.lfStrikeOut = Chr$(1)
'   If InStr(fstyle$, "u") Then LF.lfUnderline = Chr$(1)
'   LF.lfFaceName = ffont$ & Chr$(0)
'#Else
'   If InStr(fstyle$, "i") Then LF.lfItalic = 1
'   If InStr(fstyle$, "s") Then LF.lfStrikeOut = 1
'   If InStr(fstyle$, "u") Then LF.lfUnderline = 1
'   Call MemoryCopy(LF.lfFaceName(0), ByVal ffont$, LF_FACESIZE)
'#End If
'
'#If Win32 Then
'    a.lpLogFont = VarPtr(LF)
'#Else
'    Memhandle = GlobalAlloc(GHND, Len(LF))
'    If Memhandle = 0 Then Dialog_Font% = False: Exit Function
'    Address = GlobalLock(Memhandle)
'    a.lpLogFont = Address
'    Call MemoryCopy(ByVal Address, LF, Len(LF))
'#End If
'
'a.hdc = F.hdc
'a.lStructSize = Len(a)
'a.hwndOwner = F.hwnd
'a.flags = fFlags& Or CF_INITTOLOGFONTSTRUCT 'CF_SCREENFONTS Or CF_EFFECTS
'a.nFontType = SCREEN_FONTTYPE
'a.rgbColors = 0&
'
'result = ChooseFont(a)
'
'#If Win16 Then
'    If result <> 0 Then Call MemoryCopy(LF, ByVal Address, Len(LF))
'    ok = GlobalUnlock(Memhandle)
'    ok = GlobalFree(Memhandle)
'    FaceNameString = LF.lfFaceName 'Convert to string for portability below - PMC
'#Else
'    FaceNameString = StrConv(LF.lfFaceName, vbUnicode) 'Convert to string from byte array
'#End If
'If result = 0 Then Dialog_Font% = False: Exit Function
'
'fsize% = Abs(LF.lfHeight * (72 / GetDeviceCaps(F.hdc, LOGPIXELSY)))
'ffont$ = Left$(FaceNameString, InStr(FaceNameString, Chr$(0)) - 1)
'
'fstyle$ = ""
'If LF.lfWeight >= 500 Then fstyle$ = fstyle$ + "b"
'#If Win16 Then
'    If Asc(LF.lfItalic) Then fstyle$ = fstyle$ + "i"
'    If Asc(LF.lfStrikeOut) Then fstyle$ = fstyle$ + "s"
'    If Asc(LF.lfUnderline) Then fstyle$ = fstyle$ + "u"
'#Else
'    If LF.lfItalic Then fstyle$ = fstyle$ + "i"
'    If LF.lfStrikeOut Then fstyle$ = fstyle$ + "s"
'    If LF.lfUnderline Then fstyle$ = fstyle$ + "u"
'#End If
'
''.. = A.rgbColors
'
'Dialog_Font% = True
'
'End Function


Function Dialog_Open$(F As Form, ByVal ftitle$, fpfad$, ByVal ffile$, ByVal ffilter$, ByVal fext$)

Dialog_Open$ = DialogOpenSave$(F, ftitle$, fpfad$, ffile$, ffilter$, fext$, 0)

End Function

'Function Dialog_Printer%(F As Form, fhdc&, fmark%, fmin%, fmax%, fab%, fbis%)
'
'Dim Address As Long
'Dim p As PRINTDLG
'Dim D As DEVMODE
'#If Win32 Then
'    Dim Memhandle As Long, wSize As Long
'#End If
'
'flag& = PD_HIDEPRINTTOFILE Or PD_USEDEVMODECOPIES Or PD_RETURNDC
'If fmark% <> 0 Then flag& = flag& Or PD_SELECTION Else flag& = flag& Or PD_NOSELECTION
'If fmax% > 1 And fmark% = 0 And (fab% > fmin% Or fbis% < fmax%) Then flag& = flag& Or PD_PAGENUMS
'If fmin% = fmax% Then flag& = flag& Or PD_NOPAGENUMS
'If fmark% < 0 Then flag& = flag& Or PD_PRINTSETUP
'
'p.lStructSize = Len(p)
'p.hwndOwner = hwnd
'p.flags = flag& 'PD_RETURNIC
'p.nFromPage = fab%
'p.nToPage = fbis%
'p.nMinPage = fmin%
'p.nMaxPage = fmax%
'p.nCopies = 1
'result = PRINTDLG(p)
'
'If result = 0 Then Dialog_Printer% = False: Exit Function
'
'If p.hdc <> 0 Then ok = DeleteDC(p.hdc)
'If p.hDevNames <> 0 Then ok = GlobalFree(p.hDevNames)
'
'Address = GlobalLock(p.hDevMode)
'Call MemoryCopy(D, ByVal Address, Len(D))
'ok = GlobalUnlock(p.hDevMode)
'ok = GlobalFree(p.hDevMode)
'
'fhdc& = p.hdc
'fmark% = p.flags And PD_SELECTION
'If p.flags And PD_PAGENUMS Then
'   fab% = p.nFromPage
'   fbis% = p.nToPage
'Else
'   fbis% = fmax%
'   fab% = fmin%
'End If
'Dialog_Printer% = True
'
'End Function
'
'
'Sub Dialog_PrinterSetup(F As Form, fhdc%)
'
'Dim Address As Long
'Dim p As PRINTDLG
'Dim D As DEVMODE
'#If Win32 Then
'    Dim Memhandle As Long, wSize As Long
'#End If
'
'p.lStructSize = Len(p)
'p.hwndOwner = F.hwnd
'p.flags = PD_PRINTSETUP
'result = PRINTDLG(p)
'
'If result = 0 Then Exit Sub
'
'If p.hdc <> 0 Then ok = DeleteDC(p.hdc)
'If p.hDevNames <> 0 Then ok = GlobalFree(p.hDevNames)
'
'Address = GlobalLock(p.hDevMode)
'Call MemoryCopy(D, ByVal Address, Len(D))
'
'ok = GlobalUnlock(p.hDevMode)
'ok = GlobalFree(p.hDevMode)
'
'End Sub

Function Dialog_Save$(F As Form, ByVal ftitle$, fpfad$, ByVal ffile$, ByVal ffilter$, ByVal fext$)

Dialog_Save$ = DialogOpenSave$(F, ftitle$, fpfad$, ffile$, ffilter$, fext$, 1)

End Function

Function DialogOpenSave$(F As Form, ByVal ftitle$, fpfad$, ByVal ffile$, ByVal ffilter$, ByVal fext$, SaveFile%)

Dim O As OPENFILENAME, iOk As Integer
Dim wSize As Long
#If Win32 Then
    Dim Memhandle As Long
#Else
    Dim Memhandle As Integer, Address As Long
#End If

szFile$ = ffile$ + String$(128 - Len(ffile$), 0)
szFilter$ = ffilter$ + "||"
Do While InStr(szFilter$, "|")
   szFilter$ = Left$(szFilter$, InStr(szFilter$, "|") - 1) + Chr$(0) + Mid$(szFilter$, InStr(szFilter$, "|") + 1)
Loop

#If Win16 Then
    wSize = Len(szFile$) + Len(szFilter$) + Len(ftitle$) + Len(fpfad$) + Len(fext$) + 3
    Memhandle = GlobalAlloc(GHND, wSize)
    If Memhandle = 0 Then DialogOpenSave$ = "": Exit Function
    Address = GlobalLock(Memhandle)
    Call MemoryCopy(ByVal Address, ByVal (szFile$ + szFilter$ + ftitle$ + Chr$(0) + fpfad$ + Chr$(0) + fext$ + Chr$(0)), wSize)
#End If

O.lStructSize = Len(O)
O.hwndOwner = F.hWnd
O.nFilterIndex = 1
O.nMaxFile = Len(szFile$)

Select Case SaveFile%
Case 0: O.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
Case 1: O.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT
Case 2: O.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
End Select

#If Win32 Then
   O.lpstrFile = szFile$
   O.lpstrFilter = szFilter$
   O.lpstrTitle = ftitle$
   O.lpstrDefExt = fext$
   O.lpstrInitialDir = fpfad$
#Else
   O.lpstrFile = Address
   O.lpstrFilter = Address + Len(szFile$)
   O.lpstrTitle = Address + Len(szFile$) + Len(szFilter$)
   O.lpstrInitialDir = Address + Len(szFile$) + Len(szFilter$) + Len(ftitle$) + 1
   O.lpstrDefExt = Address + Len(szFile$) + Len(szFilter$) + Len(ftitle$) + 1 + Len(fpfad$) + 1
#End If

If SaveFile% = 0 Then result = GetOpenFileName(O) Else result = GetSaveFileName(O)

#If Win16 Then
    If result <> 0 Then Call MemoryCopy(ByVal szFile$, ByVal Address, Len(szFile$))
    iOk = GlobalUnlock(Memhandle)    'Free The Memory
    iOk = GlobalFree(Memhandle)
    File$ = Left$(szFile$, InStr(szFile$, Chr$(0)) - 1)
#Else
    File$ = Left$(O.lpstrFile, InStr(O.lpstrFile, Chr$(0)) - 1)
#End If

If result = 0 Then DialogOpenSave$ = "": Exit Function

fpfad$ = Left$(File$, O.nFileOffset)
DialogOpenSave$ = Right$(File$, Len(File$) - O.nFileOffset)

End Function

Function Dialog_Color%(F As Form, CurrColor&)

Dim c As CHOOSECOLOR, lOk As Long, i As Long
Dim Address As Long
#If Win32 Then
    Dim Memhandle As Long, wSize As Long
#End If

ReDim ClrArray(15) As Long    ' Holds Custom Colors

wSize = Len(ClrArray(0)) * 16 ' Size of Memory Block

Memhandle = GlobalAlloc(GHND, wSize)
If Memhandle = 0 Then Dialog_Color% = False: Exit Function
Address = GlobalLock(Memhandle)

For i = 0 To UBound(ClrArray): ClrArray(i&) = &HFFFFFF: Next

Call MemoryCopy(ByVal Address, ClrArray(0), wSize)

c.lStructSize = Len(c)
c.hwndOwner = F.hWnd
c.lpCustColors = Address
c.rgbResult = CurrColor&
c.flags = CC_RGBINIT Or CC_FULLOPEN

result = CHOOSECOLOR(c)

Call MemoryCopy(ClrArray(0), ByVal Address, wSize)

lOk = GlobalUnlock(Memhandle)    'Free The Memory
lOk = GlobalFree(Memhandle)

If result = 0 Then Dialog_Color% = False: Exit Function

CurrColor& = c.rgbResult

Dialog_Color% = True

End Function

