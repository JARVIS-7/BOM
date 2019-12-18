Attribute VB_Name = "modWinInetCalls"
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
' Zugriffe über WININET statt über INET- Control
' ersetzt altes "ShortPost"
' die Specs müssten mal aufgeräumt werden ..
' Flags von der M$- Seite


'WinINet - dll

Private Declare Function InternetOpen Lib "wininet.dll" _
         Alias "InternetOpenA" _
            (ByVal lpszCallerName As String, _
             ByVal dwAccessType As Long, _
             ByVal lpszProxyName As String, _
             ByVal lpszProxyBypass As String, _
             ByVal dwFlags As Long) As Long
             
'Private Declare Function InternetOpenUrl Lib "wininet.dll" _
        Alias "InternetOpenUrlA" _
           (ByVal hInet As Long, _
            ByVal lpszUrl As String, _
            ByVal lpszHeaders As String, _
            ByVal dwHeadersLength As Long, _
            ByVal dwFlags As Long, _
            ByVal dwContext As Long) As Long
           

Private Declare Function InternetConnect Lib "wininet.dll" _
            Alias "InternetConnectA" _
            (ByVal hInternetSession As Long, _
             ByVal lpszServerName As String, _
             ByVal nProxyPort As Integer, _
             ByVal lpszUsername As String, _
             ByVal lpszPassword As String, _
             ByVal dwService As Long, _
             ByVal dwFlags As Long, _
             ByVal dwContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
            (ByVal hFile As Long, _
             ByVal bBuffer As Any, _
             ByVal lNumBytesToRead As Long, _
             lNumberOfBytesRead As Long) As Integer

Private Declare Function HttpOpenRequest Lib "wininet.dll" _
            Alias "HttpOpenRequestA" _
            (ByVal hInternetSession As Long, _
             ByVal lpszVerb As String, _
             ByVal lpszObjectName As String, _
             ByVal lpszVersion As String, _
             ByVal lpszReferer As String, _
             ByVal lpszAcceptTypes As Long, _
             ByVal dwFlags As Long, _
             ByVal dwContext As Long) As Long

Private Declare Function HttpSendRequest Lib "wininet.dll" _
            Alias "HttpSendRequestA" _
            (ByVal hHttpRequest As Long, _
             ByVal sHeaders As String, _
             ByVal lHeadersLength As Long, _
             ByVal sOptional As String, _
             ByVal lOptionalLength As Long) As Boolean

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
            (ByVal hInternetHandle As Long) As Boolean

Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" _
             Alias "HttpAddRequestHeadersA" _
             (ByVal hHttpRequest As Long, _
             ByVal sHeaders As String, _
             ByVal lHeadersLength As Long, _
             ByVal lModifiers As Long) As Integer

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" _
             Alias "InternetGetLastResponseInfoA" _
             (ByRef lpdwError As Long, _
             ByVal lpszBuffer As String, _
             ByRef lpdwBufferLength As Long) As Boolean

Private Declare Function HttpQueryInfo Lib "wininet.dll" _
             Alias "HttpQueryInfoA" _
             (ByVal hHttpRequest As Long, _
             ByVal lInfoLevel As Long, _
             ByVal sBuffer As String, _
             ByRef lBufferLength As Long, _
             ByRef lIndex As Long) As Boolean

    ' Internet Errors
'    Const INTERNET_ERROR_BASE = 12000
'    Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
'    Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
'    Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
'    Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
'    Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
'    Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
'    Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
'    Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
'    Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
'    Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
'    Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
'    Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
'    Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
'    Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
'    Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
'    Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
'    Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
'    Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
'    Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
'    Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
'    Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
'    Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
'    Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
'    Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
'    Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
'    Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
'    Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
'    Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
'    Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
'    Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
'    Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
'    Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
'    Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
'    Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)
'    Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
'    Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
'    Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
'    Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
'    Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
'    Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
'    Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
'    Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
'    Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
'    Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
'    Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
'    Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
'    Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
'    Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
'    Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
'    Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
'    Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)
'    ' FTP API errors
'    Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110)
'    Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111)
'    Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112)
'    ' gopher API errors
'    Const ERROR_GOPHER_PROTOCOL_ERROR = (INTERNET_ERROR_BASE + 130)
'    Const ERROR_GOPHER_NOT_FILE = (INTERNET_ERROR_BASE + 131)
'    Const ERROR_GOPHER_DATA_ERROR = (INTERNET_ERROR_BASE + 132)
'    Const ERROR_GOPHER_END_OF_DATA = (INTERNET_ERROR_BASE + 133)
'    Const ERROR_GOPHER_INVALID_LOCATOR = (INTERNET_ERROR_BASE + 134)
'    Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = (INTERNET_ERROR_BASE + 135)
'    Const ERROR_GOPHER_NOT_GOPHER_PLUS = (INTERNET_ERROR_BASE + 136)
'    Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = (INTERNET_ERROR_BASE + 137)
'    Const ERROR_GOPHER_UNKNOWN_LOCATOR = (INTERNET_ERROR_BASE + 138)
'    ' HTTP API errors
'    Const ERROR_HTTP_HEADER_NOT_FOUND = (INTERNET_ERROR_BASE + 150)
'    Const ERROR_HTTP_DOWNLEVEL_SERVER = (INTERNET_ERROR_BASE + 151)
'    Const ERROR_HTTP_INVALID_SERVER_RESPONSE = (INTERNET_ERROR_BASE + 152)
'    Const ERROR_HTTP_INVALID_HEADER = (INTERNET_ERROR_BASE + 153)
'    Const ERROR_HTTP_INVALID_QUERY_REQUEST = (INTERNET_ERROR_BASE + 154)
'    Const ERROR_HTTP_HEADER_ALREADY_EXISTS = (INTERNET_ERROR_BASE + 155)
'    Const ERROR_HTTP_REDIRECT_FAILED = (INTERNET_ERROR_BASE + 156)
'    Const ERROR_HTTP_NOT_REDIRECTED = (INTERNET_ERROR_BASE + 160)
'    Const ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 161)
'    Const ERROR_HTTP_COOKIE_DECLINED = (INTERNET_ERROR_BASE + 162)
'    Const ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 168)
'    ' additional Internet API error codes
'    Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157)
'    Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158)
'    Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159)
'    Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163)
'    Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164)
'    Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165)
'    Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166)
'    Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167)
'    Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169)
'    Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170)
'    ' InternetAutodial specific errors
'    Const ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = (INTERNET_ERROR_BASE + 171)
'    Const INTERNET_ERROR_LAST = ERROR_INTERNET_FAILED_DUETOSECURITYCHECK
    '
    ' flags common to open functions (not In
    '     ternetOpen()):
    '
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000    ' retrieve the original item
'    '
'    ' flags for InternetOpenUrl():
'    '
Private Const INTERNET_FLAG_RAW_DATA As Long = &H40000000   ' FTP/gopher find: receive the item as raw (structured) data
Private Const INTERNET_FLAG_EXISTING_CONNECT As Long = &H20000000    ' FTP: use existing InternetConnect handle For server If possible
'    '
'    ' flags for InternetOpen():
'    '
'    Const INTERNET_FLAG_ASYNC = &H10000000 ' this request is asynchronous (where supported)
'    '
'    ' protocol-specific flags:
'    '
Private Const INTERNET_FLAG_PASSIVE As Long = &H8000000    ' used For FTP connections
'    '
'    ' additional cache flags
'    '
'Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000    ' don't write this item To the cache
''    Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
'Private Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000    ' make this item persistent in cache
'Private Const INTERNET_FLAG_FROM_CACHE = &H1000000    ' use offline semantics
''    Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE
''    '
''    ' additional flags
''    '
Private Const INTERNET_FLAG_SECURE As Long = &H800000   ' use PCT/SSL If applicable (HTTP)
'Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000    ' use keep-alive semantics
Private Const INTERNET_FLAG_NO_AUTO_REDIRECT As Long = &H200000   ' don't handle redirections automatically
'Private Const INTERNET_FLAG_READ_PREFETCH = &H100000   ' Do background read prefetch
Private Const INTERNET_FLAG_NO_COOKIES As Long = &H80000    ' no automatic cookie handling
'Private Const INTERNET_FLAG_NO_AUTH = &H40000    ' no automatic authentication handling
'Private Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000    ' return cache file if net request fails
'    '
'    ' Security Ignore Flags, Allow HttpOpenR
'    '     equest to overide
'    ' Secure Channel (SSL/PCT) failures of t
'    '     he following types.
'    '
'Private Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000&    ' ex: https:// to http://
'Private Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000&    ' ex: http:// to https://
Private Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID As Long = &H2000&   ' expired X509 Cert.
Private Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID As Long = &H1000&   ' bad common name in X509 Cert.
'    '
'    ' more caching flags
'    '
'Private Const INTERNET_FLAG_RESYNCHRONIZE = &H800&    ' asking wininet To update an item If it is newer
'Private Const INTERNET_FLAG_HYPERLINK = &H400&    ' asking wininet To Do hyperlinking semantic which works right For scripts
'Private Const INTERNET_FLAG_NO_UI = &H200&    ' no cookie popup
'Private Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100&    ' asking wininet To add "pragma: no-cache"
'Private Const INTERNET_FLAG_CACHE_ASYNC = &H80&    ' ok To perform lazy cache-write
'Private Const INTERNET_FLAG_FORMS_SUBMIT = &H40&    ' this is a forms submit
'Private Const INTERNET_FLAG_NEED_FILE = &H10&    ' need a file For this request
'    Const INTERNET_FLAG_MUST_CACHE_REQUEST = INTERNET_FLAG_NEED_FILE
    '
    ' flags for FTP
    '
'Private Const INTERNET_FLAG_TRANSFER_ASCII = &H1
'Private Const INTERNET_FLAG_TRANSFER_BINARY = &H2
    '
    ' flags field masks
    '
'    Const SECURITY_INTERNET_MASK = INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or INTERNET_FLAG_IGNORE_CERT_DATE_INVALID Or INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS Or INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP
'    Const INTERNET_FLAGS_MASK = INTERNET_FLAG_RELOAD Or INTERNET_FLAG_RAW_DATA Or INTERNET_FLAG_EXISTING_CONNECT Or INTERNET_FLAG_PASSIVE Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_MAKE_PERSISTENT Or INTERNET_FLAG_FROM_CACHE Or INTERNET_FLAG_SECURE Or INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_AUTO_REDIRECT Or INTERNET_FLAG_READ_PREFETCH Or INTERNET_FLAG_NO_COOKIES Or INTERNET_FLAG_NO_AUTH Or INTERNET_FLAG_CACHE_IF_NET_FAIL Or SECURITY_INTERNET_MASK Or INTERNET_FLAG_RESYNCHRONIZE Or INTERNET_FLAG_HYPERLINK Or INTERNET_FLAG_NO_UI Or INTERNET_FLAG_PRAGMA_NOCACHE Or INTERNET_FLAG_CACHE_ASYNC Or INTERNET_FLAG_FORMS_SUBMIT Or INTERNET_FLAG_NEED_FILE Or INTERNET_FLAG_TRANSFER_BINARY Or INTERNET_FLAG_TRANSFER_ASCII  'Or INTERNET_FLAG_ASYNC
'    Const INTERNET_ERROR_MASK_INSERT_CDROM = &H1
'    Const INTERNET_OPTIONS_MASK = (Not INTERNET_FLAGS_MASK)
    '
    ' common per-API flags (new APIs)
    '
'    Const WININET_API_FLAG_ASYNC = &H1 ' force async operation
'    Const WININET_API_FLAG_SYNC = &H4 ' force sync operation
'    Const WININET_API_FLAG_USE_CONTEXT = &H8 ' use value supplied in dwContext (even If 0)
    '
    ' INTERNET_NO_CALLBACK - if this value i
    '     s presented as the dwContext parameter
    ' then no call-backs will be made for th
    '     at API
    '
    'Const INTERNET_NO_CALLBACK = 0
    
    ' Proxy etc beim Open
Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_OPEN_TYPE_DIRECT As Long = 1
Private Const INTERNET_OPEN_TYPE_PROXY As Long = 3
Private Const INTERNET_DEFAULT_HTTP_PORT As Long = 80
Private Const INTERNET_DEFAULT_HTTPS_PORT As Long = 443

    'Type of service to access.
Private Const INTERNET_SERVICE_HTTP As Long = 3
'
Private Const HTTP_ADDREQ_FLAG_ADD As Long = &H20000000
Private Const HTTP_ADDREQ_FLAG_REPLACE As Long = &H80000000
Private Const HTTP_QUERY_RAW_HEADERS_CRLF As Long = 22
'
Private Const INTERNET_OPTION_CONNECT_TIMEOUT As Long = 2
Private Const INTERNET_OPTION_SEND_TIMEOUT As Long = 5
Private Const INTERNET_OPTION_RECEIVE_TIMEOUT As Long = 6

 'Typ für lpBuffer in InternetSetOption
Private Type INTERNET_CONNECTED_INFO
    dwConnectedState As Long
    dwFlags As Long
End Type

'Private Type TESTSTRUCT
'    dwTimeout As Long
'End Type

' Benötigte API-Konstante
Private Const INTERNET_STATE_CONNECTED            As Long = &H1
Private Const INTERNET_STATE_DISCONNECTED_BY_USER As Long = &H10
Private Const INTERNET_OPTION_CONNECTED_STATE     As Long = &H32
Private Const ISO_FORCE_DISCONNECTED              As Long = &H1

' Internet-Optionen setzen:
Private Declare Function InternetSetOption _
  Lib "wininet.dll" Alias "InternetSetOptionA" ( _
  ByVal hInternet As Long, _
  ByVal Options As Long, _
  ByRef lpBuffer As Any, _
  ByVal BufferLength As Long _
  ) As Long

' Internet-Optionen auslesen:
Private Declare Function InternetQueryOption _
  Lib "wininet.dll" Alias "InternetQueryOptionA" ( _
  ByVal hInternet As Long, _
  ByVal Options As Long, _
  ByRef lpBuffer As Long, _
  ByRef BufferLength As Long _
  ) As Long
                 
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                            (ByVal dwFlags As Long, _
                            lpSource As Any, _
                            ByVal dwMessageId As Long, _
                            ByVal dwLanguageId As Long, _
                            ByVal lpBuffer As String, _
                            ByVal nSize As Long, _
                            Arguments As Long) As Long

'Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Public Function WinInetReadEx(sUrl As String, sPostData As String, sReferer As String, sEBayUser As String) As String
    
    Dim hInternetOpen As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim bRet As Boolean
    Dim bErr As Boolean
    Dim bDoLoop As Boolean
    Dim lNumberOfBytesRead  As Long
    Dim lpszPostData As String
    Dim lPostDataLen As Long
    Dim sServerHeader As String
    Dim lPos1 As Long
    Dim lPos2 As Long
    Dim sErrDescription As String
    Dim sServer As String
    Dim sUrlTmp As String
    Dim sHeader As String
    Dim sCookies As String
    Dim bReadBuffer() As Byte
    Dim sVerb As String
    Dim lProxyFlag As Long
    Dim sProxy As String
    Dim sProxyUser As String
    Dim sProxyPass As String
    Dim lPort As Long
    Dim lFlags As Long
    Dim lTimeOut As Long
    Dim lBufferFillSize  As Long
    
    ReDim bReadBuffer(0 To 10239) As Byte
    
    sServer = GetServer(sUrl)
    sUrlTmp = GetUrl(sUrl)
    sHeader = "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & _
        "Accept-Language: " & gsBrowserLanguage & vbCrLf & _
        "Connection: Keep-Alive" & vbCrLf
    
    If Not gbUseIECookies And sEBayUser <> "anonymous" Then
        sCookies = goCookieHandler.GetCookieHeader(sUrl, sEBayUser)
        If sCookies > "" Then
            sHeader = sHeader & "Cookie: " & sCookies & vbCrLf
        End If
    End If
    
    If Len(sPostData) > 0 Then
        sVerb = "POST"
        sHeader = sHeader & "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
            "Content-Length: " & Len(sPostData) & vbCrLf
    Else
        sVerb = "GET"
        sReferer = ""
    End If
    
    If gbUseProxy Then
        lProxyFlag = INTERNET_OPEN_TYPE_PROXY
        sProxy = gsProxyName & ":" & CStr(giProxyPort)
    ElseIf gbUseDirectConnect Then
        lProxyFlag = INTERNET_OPEN_TYPE_DIRECT
        sProxy = vbNullString
    Else
        lProxyFlag = INTERNET_OPEN_TYPE_PRECONFIG
        sProxy = vbNullString
    End If
    
    If gbUseProxyAuthentication Then
        sProxyUser = gsProxyUser
        sProxyPass = gsProxyPass
    Else
        sProxyUser = vbNullString
        sProxyPass = vbNullString
    End If
    
    'Build the required making sure that we ignore any local caches
    
    lPort = GetPort(sUrl)
    lFlags = INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_AUTO_REDIRECT ' wir machen die Redirects jetzt selbst, sonst gibts Ärger wegen HTTPS->HTTP und WinXP-SP2
    If Not gbUseIECookies Or sEBayUser = "anonymous" Then lFlags = lFlags Or INTERNET_FLAG_NO_COOKIES ' wir machen unsere Kekse jetzt auch selbst, so sind wir unabhängig vom IE und es können mehrere Benutzer gleichzeitig angemeldet sein!
    
    If sUrl Like "https://*" Then
        If lPort = 0 Then lPort = INTERNET_DEFAULT_HTTPS_PORT
        lFlags = lFlags Or INTERNET_FLAG_SECURE Or INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or INTERNET_FLAG_IGNORE_CERT_DATE_INVALID 'Or INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP
    Else
        If lPort = 0 Then lPort = INTERNET_DEFAULT_HTTP_PORT
        'lFlags = lFlags Or INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS
    End If
    
    hInternetOpen = 0
    hInternetConnect = 0
    hHttpOpenRequest = 0
    
    On Error GoTo ErrorHandler
    
    frmHaupt.SetStatus "InternetOpen"
    DoEvents
    
    'Open a HTTP connection
    hInternetOpen = InternetOpen(gsBrowserIdString, lProxyFlag, sProxy, vbNullString, 0)
    
    If hInternetOpen <> 0 Then
    
        frmHaupt.SetStatus "InternetConnect"
        DoEvents
        
        'set timeouts
        lTimeOut = glHttpTimeOut
        If lTimeOut > 0 Then
            Call InternetSetOption(hInternetOpen, INTERNET_OPTION_CONNECT_TIMEOUT, lTimeOut, LenB(lTimeOut))
            Call InternetSetOption(hInternetOpen, INTERNET_OPTION_RECEIVE_TIMEOUT, lTimeOut, LenB(lTimeOut))
            Call InternetSetOption(hInternetOpen, INTERNET_OPTION_SEND_TIMEOUT, lTimeOut, LenB(lTimeOut))
        End If
        
        'connect to the remote HTTP server
        hInternetConnect = InternetConnect(hInternetOpen, _
            sServer, lPort, sProxyUser, sProxyPass, INTERNET_SERVICE_HTTP, 0, 0)
        
        If hInternetConnect <> 0 Then
            
            frmHaupt.SetStatus "HttpOpenRequest"
            DoEvents
            
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, _
                sVerb, sUrlTmp, "HTTP/1.0", sReferer, 0, lFlags, 0)
            
            If hHttpOpenRequest <> 0 Then
           
                'build the request header
                
                bRet = HttpAddRequestHeaders(hHttpOpenRequest, _
                    sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
                
                If (bRet = False) Then GoTo ErrorHandler
  
                'build the request body
                
                lpszPostData = sPostData
                lPostDataLen = Len(lpszPostData)
                
                frmHaupt.SetStatus "HttpSendRequest"
                DoEvents
                
                'fire the request off to the remote server
                bRet = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, lpszPostData, lPostDataLen)
                
                If (bRet = False) Then GoTo ErrorHandler
                
                
                sServerHeader = Space(10000)
                bRet = HttpQueryInfo(hHttpOpenRequest, HTTP_QUERY_RAW_HEADERS_CRLF, sServerHeader, Len(sServerHeader), 0)
                
                frmHaupt.SetStatus "InternetReadFile"
                'Dieses DoEvents muss auskommentiert bleiben, _
                sonst ist diese Funktion nicht reentrant ! lg 14.09.2003
                'DoEvents
                
                'read the response
                bDoLoop = True
                While bDoLoop
                    bDoLoop = InternetReadFile(hHttpOpenRequest, _
                        ByVal VarPtr(bReadBuffer(lBufferFillSize)), UBound(bReadBuffer) - lBufferFillSize + 1, lNumberOfBytesRead)
                        
                        lBufferFillSize = lBufferFillSize + lNumberOfBytesRead
                        If UBound(bReadBuffer) - lBufferFillSize < 1024 Then
                            ReDim Preserve bReadBuffer(0 To UBound(bReadBuffer) + 10240) As Byte
                        End If
                        
                        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                Wend

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                If lBufferFillSize > 0 Then
                    ReDim Preserve bReadBuffer(0 To lBufferFillSize - 1) As Byte
                    
                    If InStr(1, sServerHeader, "charset=utf-8", vbTextCompare) > 0 Then
                        WinInetReadEx = ByteArray2String(Decode_UTF8(bReadBuffer))
                        WinInetReadEx = Replace(WinInetReadEx, "charset=utf-8", "charset=ISO-8859-1", , , vbTextCompare)
                    Else
                        WinInetReadEx = StrConv(bReadBuffer, vbUnicode)
                    End If
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                frmHaupt.SetStatus "HttpCloseRequest"
                DoEvents
                bRet = InternetCloseHandle(hHttpOpenRequest)
            Else
                bErr = True
                frmHaupt.SetStatus "HttpOpenRequest failed", True
                DoEvents
            End If
            
            If Not bErr Then
                frmHaupt.SetStatus "InternetDisconnect"
                DoEvents
            End If
            bRet = InternetCloseHandle(hInternetConnect)
        Else
            bErr = True
            frmHaupt.SetStatus "InternetConnect failed", True
            DoEvents
        End If
        
        If Not bErr Then
            frmHaupt.SetStatus "InternetClose"
            DoEvents
        End If
        bRet = InternetCloseHandle(hInternetOpen)
    Else
        bErr = True
        frmHaupt.SetStatus "InternetOpen failed", True
        DoEvents
    End If
    
    If Not bErr Then
        frmHaupt.SetStatus "Idle", True
        DoEvents
    End If
 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If sServerHeader Like "HTTP* 3## *" Then
        lPos1 = InStr(1, sServerHeader, vbCrLf & "Location: ", vbTextCompare)
        If lPos1 > 0 Then
            lPos1 = lPos1 + 12
            lPos2 = InStr(lPos1, sServerHeader, vbCrLf)
            If lPos2 > 0 Then
                WinInetReadEx = Trim(Mid(sServerHeader, lPos1, lPos2 - lPos1))
                If Not LCase(WinInetReadEx) Like "http*" Then
                  If Not Left(WinInetReadEx, 1) = "/" Then
                      WinInetReadEx = GetUrlDirectory(sUrlTmp) & WinInetReadEx
                  End If
                  WinInetReadEx = "http://" & sServer & WinInetReadEx
                End If
                WinInetReadEx = "Redirect:" & WinInetReadEx
            End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not gbUseIECookies Then Call goCookieHandler.ExtractCookies(sServerHeader, sEBayUser)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
  Exit Function
  
ErrorHandler:
    
    If (Err.LastDllError < 12000) Then
        sErrDescription = Err.LastDllError
    Else
        If (Err.LastDllError = 12003) Then
            sErrDescription = GetLastResponse()
        Else
            sErrDescription = GetWinInetErrDesc(Err.LastDllError)
        End If
    End If
    
    frmHaupt.SetStatus sErrDescription, True
    DoEvents
    
End Function

Private Function GetWinInetErrDesc(lError As Long) As String
    
    Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800
    Dim dwLength As Long
    Dim strBuffer As String * 257
    Dim lModule As Long
    
    lModule = GetModuleHandle("wininet.dll")
    
    dwLength = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
        ByVal lModule, lError, 0&, ByVal strBuffer, 256&, 0&)

    
    If dwLength > 0 Then
        GetWinInetErrDesc = Left(strBuffer, dwLength - 2)
    End If
End Function

Private Function GetLastResponse() As String
'This function retrieves last server response.
    Dim lError As Long
    Dim strBuffer As String
    Dim lBufferSize As Long
    Dim retVal As Long

    retVal = InternetGetLastResponseInfo(lError, _
                                         strBuffer, _
                                         lBufferSize)
    strBuffer = String(lBufferSize + 1, 0)
    retVal = InternetGetLastResponseInfo(lError, _
                                         strBuffer, _
                                         lBufferSize)
    GetLastResponse = strBuffer
End Function

Private Function GetPort(ByVal sUrl As String) As Long
    
    Dim lPos As Long
    Dim sTmp As String
    
    lPos = InStr(1, sUrl, "//")
    
    If lPos > 0 Then lPos = lPos + 2 Else lPos = 1
    sUrl = Mid(sUrl, lPos, Len(sUrl) - lPos + 1)
    
    lPos = InStr(1, sUrl, "/")
    If lPos = 0 Then lPos = Len(sUrl) + 1
    sTmp = Left$(sUrl, lPos - 1)
    
    If sTmp Like "*:?*" Then GetPort = Val(Mid(sTmp, InStr(1, sTmp, ":") + 1))

End Function

Private Function GetServer(ByVal sUrl As String) As String
    
    Dim lPos As Long
    Dim sTmp As String
    
    lPos = InStr(1, sUrl, "//")
    
    If lPos > 0 Then lPos = lPos + 2 Else lPos = 1
    sUrl = Mid(sUrl, lPos, Len(sUrl) - lPos + 1)
    
    lPos = InStr(1, sUrl, "/")
    If lPos = 0 Then lPos = Len(sUrl) + 1
    sTmp = Left$(sUrl, lPos - 1)
    
    If sTmp Like "*:?*" Then sTmp = Left(sTmp, InStr(1, sTmp, ":") - 1)
    
    GetServer = sTmp

End Function

Private Function GetUrl(ByVal sUrl As String) As String
    
    Dim lPos As Long
    
    lPos = InStr(1, sUrl, "//")
    
    If lPos > 0 Then lPos = lPos + 2 Else lPos = 1
    sUrl = Mid(sUrl, lPos, Len(sUrl) - lPos + 1)
    
    lPos = InStr(1, sUrl, "/")
    If lPos = 0 Then
        GetUrl = "/"
    Else
        GetUrl = Mid(sUrl, lPos)
    End If
    
End Function

Private Function GetUrlDirectory(ByVal sUrl As String) As String
    
    Dim lPos As Long
    
    lPos = InStrRev(sUrl, "?")
    If lPos > 0 Then
        sUrl = Left(sUrl, lPos - 1)
    End If
    
    lPos = InStrRev(sUrl, "/")
    If lPos > 0 Then
        sUrl = Left(sUrl, lPos)
    End If

    GetUrlDirectory = sUrl
    
End Function

Public Function CheckIEStatus() As Boolean

'  Dim tmp As Variant
'  If modRegistry.GetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "GlobalUserOffline", tmp) Then
'    If tmp Then MsgBox gsarrLangTxt(448) & vbCrLf & gsarrLangTxt(449), vbInformation
'  End If
  If IeGetOfflineMode() Then Call IeSetOfflineMode(False)

End Function

Private Function IeGetOfflineMode() As Boolean
    
    'Fragt ab, ob sich das System momentan im Offline-Modus
    'befindet (True) oder nicht (False).
    Dim lState As Long
    
    'InternetQueryOption aufrufen
    Call InternetQueryOption(0, _
        INTERNET_OPTION_CONNECTED_STATE, lState, Len(lState))
        
    'Gesetztes Bit INTERNET_STATE_CONNECTED auswerten
    IeGetOfflineMode = CBool((lState And INTERNET_STATE_CONNECTED) = 0)
    
End Function

Private Sub IeSetOfflineMode(ByVal bOffline As Boolean, Optional ByVal bForce As Boolean = True)
    
    'Legt den Online-/Offline-Status des Systems fest.
    Dim ICI As INTERNET_CONNECTED_INFO
    
    'Offline-/Online-Flags setzen
    If bOffline Then
        'Flag für den Offline-Modus setzen
        ICI.dwConnectedState = INTERNET_STATE_DISCONNECTED_BY_USER
        If bForce Then ICI.dwFlags = ISO_FORCE_DISCONNECTED
    Else
        'Flag für den Online-Modus setzen
        ICI.dwConnectedState = INTERNET_STATE_CONNECTED
    End If
    
    'Festgelegten Zustand erwirken
    Call InternetSetOption(0, _
        INTERNET_OPTION_CONNECTED_STATE, ICI, LenB(ICI))
    
End Sub


