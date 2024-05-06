Attribute VB_Name = "modWinInet"
Option Explicit

Public gWLastDllError As Long ' para las funciones

Public Type FTP_t
   Agent                As String
   User                 As String
   Passw                As String
   Url                  As String
   
   hInternetSession     As Long
   bPassive             As Boolean
   hInternetConnect     As Long
   hHttpOpenRequest     As Long
   Context              As Long
   
   DllErr               As Long
   Rc                   As Long
   bNotFound            As Boolean
   bExist               As Boolean
End Type

Public Type Proxy_t
   bProxy   As Boolean
   Host     As String
   Port     As Integer
   User     As String
   Passw    As String
   FHost    As String
End Type
Public gProxy As Proxy_t

Public Type Web_t
   CntTimeOut  As Long     ' milliseconds
   SndTimeOut  As Long     ' milliseconds
   RcvTimeOut  As Long     ' milliseconds
End Type
Public gWeb As Web_t


Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)

#If MAX_PATH = "" Then
Public Const MAX_PATH = 260 ' en SDK50.bas
#End If

Public Const NO_ERROR = 0
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_OFFLINE = &H1000

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME_T
        ftLastAccessTime As FILETIME_T
        ftLastWriteTime As FILETIME_T
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Public Const ERROR_NO_MORE_FILES = 18

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
    
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpGetFileSize Lib "wininet.dll" _
(ByVal hFtpFile As Long, lpSize As Long) As Boolean

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean
' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function GetHostByName Lib "WSock32.dll" (ByVal HostName As String) As Long

' User agent constant.
Public Const scUserAgent = "vb wininet"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

Public Const INTERNET_OPTION_IGNORE_OFFLINE = 77


' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Declare Function HttpOpenRequest2 Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' Sends the specified request to the HTTP server.
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' InternetErrorDlg
Public Declare Function InternetErrorDlg Lib "wininet.dll" _
(ByVal hWnd As Long, ByVal hInternet As Long, ByVal dwError As Long, ByVal dwFlags As Long, ByVal lppvData As Long) As Long

' InternetErrorDlg constants
Public Const FLAGS_ERROR_UI_FILTER_FOR_ERRORS = &H1
Public Const FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS = &H2
Public Const FLAGS_ERROR_UI_FLAGS_GENERATE_DATA = &H4
Public Const FLAGS_ERROR_UI_FLAGS_NO_UI = &H8
Public Const FLAGS_ERROR_UI_SERIALIZE_DIALOGS = &H10

Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_STATUS_DENIED = 401
Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Type INTERNET_BUFFERS
    dwStructSize As Long        ' used for API versioning. Set to sizeof(INTERNET_BUFFERS)
    Next As Long                ' INTERNET_BUFFERS chain of buffers
    lpcszHeader As Long       ' pointer to headers (may be NULL)
    dwHeadersLength As Long     ' length of headers if not NULL
    dwHeadersTotal As Long      ' size of headers if not enough buffer
    lpvBuffer As Long           ' pointer to data buffer (may be NULL)
    dwBufferLength As Long      ' length of data buffer if not NULL
    dwBufferTotal As Long       ' total size of chunk, or content-length if not chunked
    dwOffsetLow As Long         ' used for read-ranges (only used in HttpSendRequest2)
    dwOffsetHigh As Long
End Type

Public Declare Function HttpSendRequestEx Lib "wininet.dll" Alias "HttpSendRequestExA" _
(ByVal hHttpRequest As Long, lpBuffersIn As INTERNET_BUFFERS, ByVal lpBuffersOut As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Public Declare Function HttpEndRequest Lib "wininet.dll" Alias "HttpEndRequestA" _
(ByVal hHttpRequest As Long, ByVal lpBuffersOut As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long


Public Declare Function InternetWriteFile Lib "wininet.dll" _
        (ByVal hFile As Long, ByVal sBuffer As String, _
        ByVal lNumberOfBytesToRead As Long, _
        lNumberOfBytesRead As Long) As Integer

Public Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFileName As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Boolean

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40
Public Const INTERNET_OPTION_DATAFILE_NAME = 33

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

' Internet Errors
Public Const INTERNET_ERROR_BASE = 12000

Public Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Public Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Public Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Public Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Public Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Public Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Public Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Public Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Public Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Public Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Public Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Public Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Public Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Public Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Public Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Public Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Public Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Public Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Public Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Public Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Public Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Public Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Public Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Public Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Public Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Public Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Public Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Public Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Public Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Public Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Public Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Public Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)

Public Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Public Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Public Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Public Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Public Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Public Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Public Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Public Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Public Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Public Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Public Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Public Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Public Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Public Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Public Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Public Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Public Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)

' FTP API errors

Public Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110)
Public Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111)
Public Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112)

' gopher API errors

Public Const ERROR_GOPHER_PROTOCOL_ERROR = (INTERNET_ERROR_BASE + 130)
Public Const ERROR_GOPHER_NOT_FILE = (INTERNET_ERROR_BASE + 131)
Public Const ERROR_GOPHER_DATA_ERROR = (INTERNET_ERROR_BASE + 132)
Public Const ERROR_GOPHER_END_OF_DATA = (INTERNET_ERROR_BASE + 133)
Public Const ERROR_GOPHER_INVALID_LOCATOR = (INTERNET_ERROR_BASE + 134)
Public Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = (INTERNET_ERROR_BASE + 135)
Public Const ERROR_GOPHER_NOT_GOPHER_PLUS = (INTERNET_ERROR_BASE + 136)
Public Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = (INTERNET_ERROR_BASE + 137)
Public Const ERROR_GOPHER_UNKNOWN_LOCATOR = (INTERNET_ERROR_BASE + 138)

' HTTP API errors

Public Const ERROR_HTTP_HEADER_NOT_FOUND = (INTERNET_ERROR_BASE + 150)
Public Const ERROR_HTTP_DOWNLEVEL_SERVER = (INTERNET_ERROR_BASE + 151)
Public Const ERROR_HTTP_INVALID_SERVER_RESPONSE = (INTERNET_ERROR_BASE + 152)
Public Const ERROR_HTTP_INVALID_HEADER = (INTERNET_ERROR_BASE + 153)
Public Const ERROR_HTTP_INVALID_QUERY_REQUEST = (INTERNET_ERROR_BASE + 154)
Public Const ERROR_HTTP_HEADER_ALREADY_EXISTS = (INTERNET_ERROR_BASE + 155)
Public Const ERROR_HTTP_REDIRECT_FAILED = (INTERNET_ERROR_BASE + 156)
Public Const ERROR_HTTP_NOT_REDIRECTED = (INTERNET_ERROR_BASE + 160)
Public Const ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 161)
Public Const ERROR_HTTP_COOKIE_DECLINED = (INTERNET_ERROR_BASE + 162)
Public Const ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 168)

' additional Internet API error codes

Public Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157)
Public Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158)
Public Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159)
Public Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163)
Public Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164)
Public Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165)

Public Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166)
Public Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167)
Public Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169)
Public Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170)

' InternetAutodial specific errors

Public Const ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = (INTERNET_ERROR_BASE + 171)

Public Const INTERNET_ERROR_LAST = ERROR_INTERNET_FAILED_DUETOSECURITYCHECK

'
' flags common to open functions (not InternetOpen()):
'

Public Const INTERNET_FLAG_RELOAD = &H80000000             ' retrieve the original item

'
' flags for InternetOpenUrl():
'

Public Const INTERNET_FLAG_RAW_DATA = &H40000000           ' FTP/gopher find: receive the item as raw (structured) data
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000   ' FTP: use existing InternetConnect handle for server if possible

'
' flags for InternetOpen():
'

Public Const INTERNET_FLAG_ASYNC = &H10000000              ' this request is asynchronous (where supported)

'
' protocol-specific flags:
'

Public Const INTERNET_FLAG_PASSIVE = &H8000000             ' used for FTP connections

'
' additional cache flags
'

Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000      ' don't write this item to the cache
Public Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000     ' make this item persistent in cache
Public Const INTERNET_FLAG_FROM_CACHE = &H1000000          ' use offline semantics
Public Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE

'
' additional flags
'

Public Const INTERNET_FLAG_SECURE = &H800000               ' use PCT/SSL if applicable (HTTP)
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000      ' use keep-alive semantics
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000     ' don't handle redirections automatically
Public Const INTERNET_FLAG_READ_PREFETCH = &H100000        ' do background read prefetch
Public Const INTERNET_FLAG_NO_COOKIES = &H80000            ' no automatic cookie handling
Public Const INTERNET_FLAG_NO_AUTH = &H40000               ' no automatic authentication handling
Public Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000     ' return cache file if net request fails

'
' Security Ignore Flags, Allow HttpOpenRequest to overide
'  Secure Channel (SSL/PCT) failures of the following types.
'

Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000       ' ex: https:// to http://
Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000      ' ex: http:// to https://
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000      ' expired X509 Cert.
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000        ' bad common name in X509 Cert.

'
' more caching flags
'

Public Const INTERNET_FLAG_RESYNCHRONIZE = &H800           ' asking wininet to update an item if it is newer
Public Const INTERNET_FLAG_HYPERLINK = &H400               ' asking wininet to do hyperlinking semantic which works right for scripts
Public Const INTERNET_FLAG_NO_UI = &H200                   ' no cookie popup
Public Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100          ' asking wininet to add "pragma: no-cache"
Public Const INTERNET_FLAG_CACHE_ASYNC = &H80              ' ok to perform lazy cache-write
Public Const INTERNET_FLAG_FORMS_SUBMIT = &H40             ' this is a forms submit
Public Const INTERNET_FLAG_NEED_FILE = &H10                ' need a file for this request
Public Const INTERNET_FLAG_MUST_CACHE_REQUEST = INTERNET_FLAG_NEED_FILE

'
' flags for FTP
'

Public Const INTERNET_FLAG_TRANSFER_ASCII = FTP_TRANSFER_TYPE_ASCII       ' = &H00000001
Public Const INTERNET_FLAG_TRANSFER_BINARY = FTP_TRANSFER_TYPE_BINARY     ' = &H00000002

'
' flags field masks
'

Public Const SECURITY_INTERNET_MASK = INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
                                 INTERNET_FLAG_IGNORE_CERT_DATE_INVALID Or _
                                 INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS Or _
                                 INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP

Public Const INTERNET_FLAGS_MASK = INTERNET_FLAG_RELOAD Or _
                                 INTERNET_FLAG_RAW_DATA Or _
                                 INTERNET_FLAG_EXISTING_CONNECT Or _
                                 INTERNET_FLAG_ASYNC Or _
                                 INTERNET_FLAG_PASSIVE Or _
                                 INTERNET_FLAG_NO_CACHE_WRITE Or _
                                 INTERNET_FLAG_MAKE_PERSISTENT Or _
                                 INTERNET_FLAG_FROM_CACHE Or _
                                 INTERNET_FLAG_SECURE Or _
                                 INTERNET_FLAG_KEEP_CONNECTION Or _
                                 INTERNET_FLAG_NO_AUTO_REDIRECT Or _
                                 INTERNET_FLAG_READ_PREFETCH Or _
                                 INTERNET_FLAG_NO_COOKIES Or _
                                 INTERNET_FLAG_NO_AUTH Or _
                                 INTERNET_FLAG_CACHE_IF_NET_FAIL Or _
                                 SECURITY_INTERNET_MASK Or _
                                 INTERNET_FLAG_RESYNCHRONIZE Or _
                                 INTERNET_FLAG_HYPERLINK Or _
                                 INTERNET_FLAG_NO_UI Or _
                                 INTERNET_FLAG_PRAGMA_NOCACHE Or _
                                 INTERNET_FLAG_CACHE_ASYNC Or _
                                 INTERNET_FLAG_FORMS_SUBMIT Or _
                                 INTERNET_FLAG_NEED_FILE Or _
                                 INTERNET_FLAG_TRANSFER_BINARY Or _
                                 INTERNET_FLAG_TRANSFER_ASCII
                                

Public Const INTERNET_ERROR_MASK_INSERT_CDROM = &H1

Public Const INTERNET_OPTIONS_MASK = (Not INTERNET_FLAGS_MASK)

'
' common per-API flags (new APIs)
'

Public Const WININET_API_FLAG_ASYNC = &H1                  ' force async operation
Public Const WININET_API_FLAG_SYNC = &H4                   ' force sync operation
Public Const WININET_API_FLAG_USE_CONTEXT = &H8            ' use value supplied in dwContext (even if 0)

'
' INTERNET_NO_CALLBACK - if this value is presented as the dwContext parameter
' then no call-backs will be made for that API
'

Public Const INTERNET_NO_CALLBACK = 0

Public Declare Function InternetDial Lib "wininet.dll" _
(ByVal hWnd As Long, ByVal sConnectoid As String, _
 ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long

Public Declare Function InternetHangUp Lib "wininet.dll" _
(ByVal dwConnection As Long, ByVal dwReserved As Long) As Long

Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
(ByVal hInternetSession As Long, ByVal sUrl As String, _
ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

Declare Function InternetSetStatusCallback Lib "WinINet" Alias "InternetSetStatusCallbackW" (hFtp As Long, ByVal lpfnInternetCallback As Long) As Long

Public Declare Function CoCreateGuid Lib "OLE32" (g As GUID) As Long
Public Declare Function StringFromGUID2 Lib "OLE32" (g As GUID, ByVal lpsz As Long, ByVal cchMax As Long) As Long

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public TOKEN_SII As String

Public Const SII_LOGIN = 1
Public Const SII_LOGOUT = 2
Public Const SII_RETENCIONES = 3
Public Const SII_COMPRAS = 4
Public Const SII_VENTAS = 5

Public Const URL_SII_LOGIN = "https://zeusr.sii.cl/cgi_AUT2000/CAutInicio.cgi"
Public Const URL_SII_LOGOUT = "https://zeusr.sii.cl/cgi_AUT2000/autTermino.cgi"
Public Const URL_SII_COMPRA = "https://www4.sii.cl/consdcvinternetui/services/data/facadeService/getDetalleCompraExport"
Public Const URL_SII_VENTA = "https://www4.sii.cl/consdcvinternetui/services/data/facadeService/getDetalleVentaExport"
Public Const URL_SII_RETENCIONES = "https://loa.sii.cl/cgi_IMT/TMBCOC_InformeMensualBheRec.cgi"

Public Function GetIErr(ByVal lErrorCode As Long) As String
   Dim ErrStr As String
   Dim Rc As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' From       : MSDN
' Name       : GetErr (Formerly TranslateErrorCode)
' Purpose    : Provides message for DLL error codes
' Parameters : The DLL error code
' Return val : String containing message
' Algorithm : Selects the appropriate string
''''''''''''''''''''''''''''''''''''''''''''''''''''''

   Dim sBuffer As String
   Dim nBuffer As Long

   Select Case lErrorCode
      Case 12001: ErrStr = "No more handles could be generated at this time"
      Case 12002: ErrStr = "The request has timed out."
      Case 12003:
         'extended error. Retrieve the details using
         'the InternetGetLastResponseInfo API.
         
         sBuffer = Space$(MAX_PATH)
         nBuffer = Len(sBuffer)
         
         If InternetGetLastResponseInfo(Rc, _
                                        sBuffer, _
                                        nBuffer) Then
            ErrStr = Trim0(sBuffer)
         Else
            ErrStr = "Extended error returned from server."
         End If
         
      Case 12004: ErrStr = "An internal error has occurred."
      Case 12005: ErrStr = _
         "URL is invalid."
      Case 12006: ErrStr = _
         "URL scheme could not be recognized, or is not supported."
      Case 12007: ErrStr = _
         "Server name could not be resolved. Check your Firewall access."
      Case 12008: ErrStr = _
         "Requested protocol could not be located."
      Case 12009: ErrStr = _
         "Request to InternetQueryOption or InternetSetOption" & _
         " specified an invalid option value."
      Case 12010: ErrStr = _
         "Length of an option supplied to InternetQueryOption or" & _
         " InternetSetOption is incorrect for the type of" & _
         " option specified."
      Case 12011: ErrStr = _
         "Request option can not be set, only queried. "
      Case 12012: ErrStr = _
         "Win32 Internet support is being shutdown or unloaded."
      Case 12013: ErrStr = _
         "Request to connect and login to an FTP server could not" & _
         " be completed because the supplied user name is incorrect."
      Case 12014: ErrStr = _
         "Request to connect and login to an FTP server could not" & _
         " be completed because the supplied password is incorrect. "
      Case 12015: ErrStr = _
         "Request to connect to and login to an FTP server failed."
      Case 12016: ErrStr = _
         "Requested operation is invalid. "
      Case 12017: ErrStr = _
         "Operation was canceled, usually because the handle on" & _
         " which the request was operating was closed before the" & _
         " operation completed."
      Case 12018: ErrStr = _
         "Type of handle supplied is incorrect for this operation."
      Case 12019: ErrStr = _
         "Requested operation can not be carried out because the" & _
         " handle supplied is not in the correct state."
      Case 12020: ErrStr = _
         "Request can not be made via a proxy."
      Case 12021: ErrStr = _
         "Required registry value could not be located. "
      Case 12022: ErrStr = _
         "Required registry value was located but is an incorrect" & _
         " type or has an invalid value."
      Case 12023: ErrStr = _
         "Direct network access cannot be made at this time. "
      Case 12024: ErrStr = _
         "Asynchronous request could not be made because a zero" & _
         " context value was supplied."
      Case 12025: ErrStr = _
         "Asynchronous request could not be made because a" & _
         " callback function has not been set."
      Case 12026: ErrStr = _
         "Required operation could not be completed because" & _
         " one or more requests are pending."
      Case 12027: ErrStr = _
         "Format of the request is invalid."
      Case 12028: ErrStr = _
         "Requested item could not be located."
      Case 12029: ErrStr = _
         "Attempt to connect to the server failed."
      Case 12030: ErrStr = _
         "Connection with the server has been terminated."
      Case 12031: ErrStr = _
         "Connection with the server has been reset."
      Case 12036: ErrStr = _
         "Request failed because the handle already exists."
      Case Else: ErrStr = _
         "Error details not available."
   End Select

   GetIErr = lErrorCode & ": " & ErrStr
   
End Function

Function webEncode(ByVal Buf As String) As String
   Dim i As Integer, oBuf As String, a As Integer

   oBuf = ""
   For i = 1 To Len(Buf)
      If InStr("&?/%# =" & vbCrLf, Mid(Buf, i, 1)) Then
         a = Asc(Mid(Buf, i, 1))
         oBuf = oBuf & "%" & Right("00" & Hex(a), 2)
      Else
         oBuf = oBuf & Mid(Buf, i, 1)
     End If
   Next i
   
   webEncode = oBuf

End Function


' Lee el contenido de una página web
' Host del tipo www.sii.cl sin http://
' TimeOut en milisegundos
Public Function FwWebReadPage(ByVal Host As String, ByVal Path As String, Optional ByVal TimeOut As Long = -1) As String
   Dim hInternetSession As Long, hInternetConnect As Long, hHttpOpenRequest As Long, OpenFlags As Long
   Dim bDoLoop As Boolean
   Dim sReadBuffer As String, lNumberOfBytesRead As Long
   Dim Buf As String
   Dim Rc As Long, Dt As Long, i As Integer, j As Integer, k As Integer
   Dim Port As Long

   FwWebReadPage = ""
   gWLastDllError = 0
   
   hInternetSession = InternetOpen("FairVer" & Format(Now, "hhnnss"), INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
   If hInternetSession = 0 Then
      gWLastDllError = ERR.LastDllError
      Debug.Print "InternetOpen, iErr=" & GetIErr(gWLastDllError)
      Call AddLog("ReadPage: InternetOpen, iErr=" & GetIErr(gWLastDllError) & ", " & GetLastSystemError(gWLastDllError))
      Exit Function
   End If
   
   DoEvents
   
   Rc = 5000
   Rc = InternetSetOption(hInternetSession, INTERNET_OPTION_CONNECT_TIMEOUT, Rc, Len(Rc))

   Rc = 1
   Rc = InternetSetOption(hInternetSession, INTERNET_OPTION_IGNORE_OFFLINE, Rc, Len(Rc))

   DoEvents
   
   OpenFlags = INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION
   
   If StrComp(Left(Host, 8), "https://", vbTextCompare) = 0 Then
      OpenFlags = OpenFlags Or INTERNET_FLAG_SECURE
      Port = INTERNET_DEFAULT_HTTPS_PORT
      Host = Mid(Host, 9)
   Else
      Port = INTERNET_DEFAULT_HTTP_PORT
      
      If StrComp(Left(Host, 7), "http://", vbTextCompare) = 0 Then
         Host = Mid(Host, 8)
      End If
   End If
      
   If Path = "" Then
      i = InStr(Host, "/")
      If i > 0 Then
         Path = Mid(Host, i)
         Host = Left(Host, i - 1)
      End If
   End If
   
   Call AddDebug("ReadPage: Port=" & Port & ", OpenFlags=&H" & Hex(OpenFlags) & ", [" & Host & "] [" & Path & "]")
   
   hInternetConnect = InternetConnect(hInternetSession, Host, Port, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
   If hInternetConnect = 0 Then
      gWLastDllError = ERR.LastDllError
      Debug.Print "InternetConnect, iErr=" & GetIErr(gWLastDllError) & ", " & Host
      Call AddLog("ReadPage: InternetConnect, iErr=" & GetIErr(gWLastDllError) & ", " & GetLastSystemError(gWLastDllError) & ", " & Host)
      Call InternetCloseHandle(hInternetSession)
      Exit Function
   End If
   
   DoEvents

   hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", Path, "HTTP/1.0", vbNullString, 0, OpenFlags, 0)
   If hHttpOpenRequest = 0 Then
      gWLastDllError = ERR.LastDllError
      Debug.Print "HttpOpenRequest, iErr=" & GetIErr(gWLastDllError) & ", " & Host
      Call AddLog("ReadPage: HttpOpenRequest, iErr=" & GetIErr(gWLastDllError) & ", " & GetLastSystemError(gWLastDllError) & ", " & Host)
      Call InternetCloseHandle(hInternetConnect)
      Call InternetCloseHandle(hInternetSession)
      Exit Function
   End If

   If TimeOut = -1 Then
      TimeOut = 15 * 1000 ' 30 segundos
   End If

   If TimeOut Then
      Rc = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, TimeOut, Len(TimeOut))
      
      Rc = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SEND_TIMEOUT, TimeOut, Len(TimeOut))
   End If
   
   DoEvents

   Buf = ""
   Rc = 0
   Rc = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, Buf, Rc)
   If Rc = 0 Then
      gWLastDllError = ERR.LastDllError
      Debug.Print "HttpSendRequest Err=" & gWLastDllError & ", " & GetLastSystemError(gWLastDllError) & ", " & Host
      Call AddLog("ReadPage: HttpSendRequest, iErr=" & GetIErr(gWLastDllError) & ", " & GetLastSystemError(gWLastDllError) & ", " & Host)
      Call InternetCloseHandle(hHttpOpenRequest)
      Call InternetCloseHandle(hInternetConnect)
      Call InternetCloseHandle(hInternetSession)
      Debug.Print "**** Verificar que no esté Offline"
      Exit Function
   End If
   
   DoEvents
   
   Buf = ""
   If Rc Then
      Do
         sReadBuffer = Space(1024)
         bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
         Buf = Buf & Left(sReadBuffer, lNumberOfBytesRead)
         
         DoEvents
      Loop While lNumberOfBytesRead > 0
   End If

   Call InternetCloseHandle(hHttpOpenRequest)
   Call InternetCloseHandle(hInternetConnect)
   Call InternetCloseHandle(hInternetSession)
              
   FwWebReadPage = Buf
           
End Function

Public Function FwWebReadPagePost(ByVal Host As String, ByVal Path As String, Optional ByVal TimeOut As Long = -1) As String
   Dim hInternetSession As Long, hInternetConnect As Long, hHttpOpenRequest As Long, OpenFlags As Long
   Dim bDoLoop As Boolean
   Dim sReadBuffer As String, lNumberOfBytesRead As Long
   Dim Buf As String
   Dim Rc As Long, Dt As Long, i As Integer, j As Integer, k As Integer
   Dim Port As Long

   FwWebReadPagePost = ""
   gWLastDllError = 0
   
   hInternetSession = InternetOpen("FairVer" & Format(Now, "hhnnss"), INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
   If hInternetSession = 0 Then
      gWLastDllError = ERR.LastDllError
      Debug.Print "InternetOpen, iErr=" & GetIErr(gWLastDllError)
      Call AddLog("ReadPage: InternetOpen, iErr=" & GetIErr(gWLastDllError) & ", " & GetLastSystemError(gWLastDllError))
      Exit Function
   End If
   
   DoEvents
   
   Rc = 5000
   Rc = InternetSetOption(hInternetSession, INTERNET_OPTION_CONNECT_TIMEOUT, Rc, Len(Rc))

   Rc = 1
   Rc = InternetSetOption(hInternetSession, INTERNET_OPTION_IGNORE_OFFLINE, Rc, Len(Rc))

   DoEvents
   
   OpenFlags = INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION
   
   If StrComp(Left(Host, 8), "https://", vbTextCompare) = 0 Then
      OpenFlags = OpenFlags Or INTERNET_FLAG_SECURE
      Port = INTERNET_DEFAULT_HTTPS_PORT
      Host = Mid(Host, 9)
   Else
      Port = INTERNET_DEFAULT_HTTP_PORT
      
      If StrComp(Left(Host, 7), "http://", vbTextCompare) = 0 Then
         Host = Mid(Host, 8)
      End If
   End If
      
   If Path = "" Then
      i = InStr(Host, "/")
      If i > 0 Then
         Path = Mid(Host, i)
         Host = Left(Host, i - 1)
      End If
   End If
   
   Call AddDebug("ReadPage: Port=" & Port & ", OpenFlags=&H" & Hex(OpenFlags) & ", [" & Host & "] [" & Path & "]")
   
   hInternetConnect = InternetConnect(hInternetSession, Host, Port, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
   If hInternetConnect = 0 Then
      gWLastDllError = ERR.LastDllError
      Debug.Print "InternetConnect, iErr=" & GetIErr(gWLastDllError) & ", " & Host
      Call AddLog("ReadPage: InternetConnect, iErr=" & GetIErr(gWLastDllError) & ", " & GetLastSystemError(gWLastDllError) & ", " & Host)
      Call InternetCloseHandle(hInternetSession)
      Exit Function
   End If
   
   DoEvents

   hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "POST", Path, "HTTP/1.0", vbNullString, 0, OpenFlags, 0)
   If hHttpOpenRequest = 0 Then
      gWLastDllError = ERR.LastDllError
      Debug.Print "HttpOpenRequest, iErr=" & GetIErr(gWLastDllError) & ", " & Host
      Call AddLog("ReadPage: HttpOpenRequest, iErr=" & GetIErr(gWLastDllError) & ", " & GetLastSystemError(gWLastDllError) & ", " & Host)
      Call InternetCloseHandle(hInternetConnect)
      Call InternetCloseHandle(hInternetSession)
      Exit Function
   End If

   If TimeOut = -1 Then
      TimeOut = 15 * 2000 ' 30 segundos
   End If

   If TimeOut Then
      Rc = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, TimeOut, Len(TimeOut))
      'Rc = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, 0, 0)
      
      Rc = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SEND_TIMEOUT, 0, Len(TimeOut))
      'Rc = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SEND_TIMEOUT, 0, 0)
   End If
   
   DoEvents

   Buf = ""
   Rc = 0
   Rc = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, Buf, Rc)
   If Rc = 0 Then
      gWLastDllError = ERR.LastDllError
      Debug.Print "HttpSendRequest Err=" & gWLastDllError & ", " & GetLastSystemError(gWLastDllError) & ", " & Host
      Call AddLog("ReadPage: HttpSendRequest, iErr=" & GetIErr(gWLastDllError) & ", " & GetLastSystemError(gWLastDllError) & ", " & Host)
      Call InternetCloseHandle(hHttpOpenRequest)
      Call InternetCloseHandle(hInternetConnect)
      Call InternetCloseHandle(hInternetSession)
      Debug.Print "**** Verificar que no esté Offline"
      Exit Function
   End If
   
   DoEvents
   
   Buf = ""
   If Rc Then
      Do
         sReadBuffer = Space(1024)
         bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
         Buf = Buf & Left(sReadBuffer, lNumberOfBytesRead)
         
         DoEvents
      Loop While lNumberOfBytesRead > 0
   End If

   Call InternetCloseHandle(hHttpOpenRequest)
   Call InternetCloseHandle(hInternetConnect)
   Call InternetCloseHandle(hInternetSession)
              
   FwWebReadPagePost = Buf
           
End Function

Public Function FwOpenFTP(Ftp As FTP_t, Optional ByVal bMsg As Boolean = 1) As Long
   Dim Flags As Long

   Ftp.DllErr = 0
      
   Ftp.hInternetSession = InternetOpen(Ftp.Agent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
   
   If Ftp.hInternetSession = 0 Then
      Ftp.DllErr = ERR.LastDllError
      If bMsg Then
         MsgBox1 "Error en InternetOpen: " & GetIErr(Ftp.DllErr), vbExclamation
      Else
         Call AddLog("InternetOpen: Error " & GetIErr(Ftp.DllErr))
      End If
      FwOpenFTP = Ftp.DllErr
      Exit Function
   End If
   
   Flags = INTERNET_FLAG_EXISTING_CONNECT
   If Ftp.bPassive Then
      Flags = Flags Or INTERNET_FLAG_PASSIVE
   End If
   
   Ftp.hInternetConnect = InternetConnect(Ftp.hInternetSession, Ftp.Url, INTERNET_DEFAULT_FTP_PORT, Ftp.User, Ftp.Passw, INTERNET_SERVICE_FTP, Flags, 0)

   If Ftp.hInternetConnect = 0 Then
      Ftp.DllErr = ERR.LastDllError
      If bMsg Then
         MsgBox1 "Error en InternetConnect: " & GetIErr(Ftp.DllErr), vbExclamation
      Else
         Call AddLog("InternetConnect: Error " & GetIErr(Ftp.DllErr))
      End If
      FwOpenFTP = Ftp.DllErr
      Exit Function
   End If

   FwOpenFTP = 0
   
End Function

Public Function FwCloseFTP(Ftp As FTP_t) As Long

   If Ftp.hInternetConnect Then
      If InternetCloseHandle(Ftp.hInternetConnect) = 0 Then
         FwCloseFTP = ERR.LastDllError
      Else
         Ftp.hInternetConnect = 0
      End If
   End If
   
   If Ftp.hInternetSession And Ftp.hInternetConnect = 0 Then
      If InternetCloseHandle(Ftp.hInternetSession) = 0 Then
         FwCloseFTP = ERR.LastDllError
      Else
         Ftp.hInternetSession = 0
      End If
   End If

End Function

Public Function FwFtpDelFile(Ftp As FTP_t, ByVal fname As String, Optional ByVal bMsg As Boolean = 1) As Long
   Dim Rc As Long, Buf As String

   Ftp.DllErr = 0
   Rc = FtpDeleteFile(Ftp.hInternetConnect, fname)

   If Rc <> 1 Then
      Ftp.DllErr = ERR.LastDllError
      
      Buf = GetIErr(Ftp.DllErr)
      
      If bMsg Then
         MsgBox1 "Error FTP " & Buf, vbExclamation
      End If
      
      Ftp.bNotFound = (InStr(Buf, " 550 ") > 0)  ' no encontrado
   End If

   FwFtpDelFile = Ftp.DllErr
   
End Function

Public Function FwFtpOpenFile(Ftp As FTP_t, ByVal fname As String, Optional ByVal bMsg As Boolean = 1) As Long
   Dim hFile As Long, Msg As String

   Ftp.DllErr = 0
   Ftp.bNotFound = 0
   
   hFile = FtpOpenFile(Ftp.hInternetConnect, fname, GENERIC_READ, FTP_TRANSFER_TYPE_BINARY, Ftp.Context)

   If hFile = 0 Then
      
      Ftp.DllErr = ERR.LastDllError
      Msg = GetIErr(Ftp.DllErr)
      
      Ftp.bNotFound = (InStr(Msg, " 550 ") > 0)  ' no encontrado
      
      If bMsg Then
         MsgBox1 "Error FTP " & Msg, vbExclamation
      End If
      
   End If

   FwFtpOpenFile = hFile

End Function

Public Function FwFtpPutFile(Ftp As FTP_t, ByVal LocFName As String, ByVal RemFName As String, Optional ByVal bMsg As Boolean = 1) As Long
   Dim Rc As Long
   Dim Msg As String
   
   If Ftp.hInternetConnect <= 0 Then
      If bMsg Then
         MsgBox1 "FwFtpPutFile: falta el Open FTP.", vbExclamation
      Else
         Call AddLog("FwFtpPutFile: falta el Open FTP.")
      End If
   End If
   
   Ftp.DllErr = 0
   Ftp.bNotFound = 0
   Ftp.bExist = 0

   ERR.Clear
   Rc = FtpPutFile(Ftp.hInternetConnect, LocFName, RemFName, FTP_TRANSFER_TYPE_BINARY, Ftp.Context)

   If Rc <> 1 Then ' OK ?
      
      Ftp.DllErr = ERR.LastDllError
      Msg = GetIErr(Ftp.DllErr)
      
      Ftp.bExist = (InStr(Msg, " 550") > 0)
      
      If bMsg Then
         MsgBox1 "Error FTP: " & Msg, vbExclamation
      Else
         Call AddLog("FwFtpPutFile: Error " & Msg & ", '" & LocFName & "' => '" & RemFName & "'.")
      End If
   End If

   FwFtpPutFile = Ftp.DllErr

End Function

Public Function FwFtpMkDir(Ftp As FTP_t, ByVal Folder As String, Optional ByVal bMsg As Boolean = 1) As Long
   Dim Rc As Long, Msg As String
      
   Ftp.DllErr = 0
   Ftp.bNotFound = 0
   Ftp.bExist = 0

   Rc = FtpCreateDirectory(Ftp.hInternetConnect, Folder)

   If Rc <> 1 Then ' OK ?
      
      Ftp.DllErr = ERR.LastDllError
      Msg = GetIErr(Ftp.DllErr)
      
      Ftp.bExist = (InStr(Msg, " 550") > 0)
      
      If bMsg Then
         MsgBox1 "Error FTP " & Msg, vbExclamation
      End If
   End If

   FwFtpMkDir = Ftp.DllErr

End Function

Public Function FwFtpSetOption(Ftp As FTP_t, ByVal lOption As Long, ByVal sBuffer As String) As Integer

   FwFtpSetOption = InternetSetOptionStr(Ftp.hInternetSession, lOption, sBuffer, Len(sBuffer))

End Function

Public Function FwPostPageOld(ByVal Url As String, ByVal Params As String, Optional ContentType As String = "") As String
Dim hOpen As Long, hConnection As Long, hFile As Long, nRead As Long, Rc As Long
Dim File As String, sHeader As String
Dim xmlStr As String, tmpBuf As String * 2048
Dim bDoLoop As Boolean, i As Integer, Port As Integer
Dim nErr As Long, strErr As String, CharSet As String, ReqFlags As Long

FwPostPageOld = ""

' XmlBody = "<?xml version=""1.0"" encoding=""UTF-8""?>"
' XmlBody = XmlBody & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:ns=""urn:wsident""><SOAP-ENV:Body SOAP-ENV:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/""><ns:wsident><IdentReq><wsUsuario>Autentia</wsUsuario><wsClave>@ut3nti4.</wsClave><Pais></Pais><Rut>1-8</Rut><NroAudit>AAAAA</NroAudit><Sistema>SIS</Sistema><Institucion></Institucion><Clave></Clave></IdentReq></ns:wsident></SOAP-ENV:Body></SOAP-ENV:Envelope>"

If StrComp(Left(Url, 8), "https://", vbTextCompare) = 0 Then
Port = INTERNET_DEFAULT_HTTPS_PORT
ReqFlags = INTERNET_FLAG_SECURE
Url = Mid(Url, 9)
ElseIf StrComp(Left(Url, 7), "http://", vbTextCompare) = 0 Then
Port = INTERNET_DEFAULT_HTTP_PORT
Url = Mid(Url, 8)
ReqFlags = 0
End If

'<<<<<<< .mine
' 'Crea los objetos DOMDocument y XMLHTTP
'
' MsgBox "MSXML2.XMLHTTP"
' Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
'=======
i = InStr(Url, "/")
If i > 0 Then
File = Mid(Url, i)
Url = Left(Url, i - 1)
Else
File = ""
End If
'>>>>>>> .r536

'<<<<<<< .mine
' MsgBox "MSXML2.DOMDocument"
' Set objDom = CreateObject("MSXML2.DOMDocument")
'
'
' 'Carga el XML
' objDom.async = False
' objDom.loadXML XmlBody
'=======
i = InStr(Url, ":")
If i > 0 Then
Port = Val(Mid(Url, i + 1))
Url = Left(Url, i - 1)
End If
'>>>>>>> .r536

If Url = "" Then
Exit Function
End If

If gProxy.bProxy Then ' 26 oct 2011: pam: se agrega soporte para Proxy
hOpen = InternetOpen("WebProx", INTERNET_OPEN_TYPE_PROXY, gProxy.FHost, vbNullString, 0)
Else
hOpen = InternetOpen("", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
End If
If hOpen = 0 Then
nErr = ERR.LastDllError
strErr = "Error " & nErr & " en InternetOpen. " & GetLastSystemError(nErr)
Exit Function
End If



Rc = 1
Rc = InternetSetOption(hOpen, INTERNET_OPTION_IGNORE_OFFLINE, Rc, Len(Rc))



If gWeb.CntTimeOut > 0 Then
Rc = InternetSetOption(hOpen, INTERNET_OPTION_CONNECT_TIMEOUT, gWeb.CntTimeOut, Len(gWeb.CntTimeOut))
End If



hConnection = InternetConnect(hOpen, Url, Port, gProxy.User, gProxy.Passw, INTERNET_SERVICE_HTTP, 0, 0)
If hConnection = 0 Then
nErr = ERR.LastDllError
strErr = "Error " & nErr & " en InternetConnect. " & GetLastSystemError(nErr)
Call InternetCloseHandle(hOpen)
Exit Function
End If

' Uncomment to use http GET request:
' File = File & xml
' hFile = HttpOpenRequest(hConnection, "GET", File, "HTTP/1.0", vbNullString, 0, 0, 0)

' Using http POST request:
hFile = HttpOpenRequest(hConnection, "POST", File, vbNullString, vbNullString, 0, ReqFlags, 1)
If hFile = 0 Then
nErr = ERR.LastDllError
strErr = "Error " & nErr & " en HttpOpenRequest. " & GetLastSystemError(nErr)
Call InternetCloseHandle(hConnection)
Call InternetCloseHandle(hOpen)
Exit Function
End If
'MsgBox "Error: " & Err.Number & ", " & Err.Description



' sHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf



If ContentType = "" Then
ContentType = "application/x-www-form-urlencoded"
End If

' ContentType = "text/html"
CharSet = "utf-8"



sHeader = "Content-Type: " & ContentType & "; charset=" & CharSet & vbCrLf
Call HttpAddRequestHeaders(hFile, sHeader, Len(sHeader), 0)

sHeader = "POST: " & File & vbCrLf
Call HttpAddRequestHeaders(hFile, sHeader, Len(sHeader), 0)

If gWeb.SndTimeOut > 0 Then
Rc = InternetSetOption(hFile, INTERNET_OPTION_SEND_TIMEOUT, gWeb.SndTimeOut, Len(gWeb.SndTimeOut))
End If

If gWeb.RcvTimeOut > 0 Then
Rc = InternetSetOption(hFile, INTERNET_OPTION_RECEIVE_TIMEOUT, gWeb.RcvTimeOut, Len(gWeb.RcvTimeOut))
End If

' end of code to use POST request
If HttpSendRequest(hFile, vbNullString, 0, Params, Len(Params)) = 0 Then
nErr = ERR.LastDllError
strErr = GetLastSystemError(nErr)
If nErr <> ERR.LastDllError Then
strErr = "Error " & nErr & " en HttpSendRequest. " & strErr
End If
Call InternetCloseHandle(hFile)
Call InternetCloseHandle(hConnection)
Call InternetCloseHandle(hOpen)
Exit Function



End If


bDoLoop = True
xmlStr = ""
Do While bDoLoop
tmpBuf = ""
bDoLoop = InternetReadFile(hFile, tmpBuf, Len(tmpBuf), nRead)
If bDoLoop = False Then
nErr = ERR.LastDllError
strErr = "Error " & nErr & " en InternetReadFile. " & GetLastSystemError(nErr)
Exit Do
ElseIf nRead > 0 Then
xmlStr = xmlStr & Left(tmpBuf, nRead)
Else
Exit Do
End If
Loop

FwPostPageOld = xmlStr

' CheckResponse xmlStr ' xmlStr is the string that should be returned but is always empty...
If hFile <> 0 Then Call InternetCloseHandle(hFile)
If hConnection <> 0 Then Call InternetCloseHandle(hConnection)
If hOpen <> 0 Then Call InternetCloseHandle(hOpen)

End Function ' PostWebservice


' Invoca un web service pero sin usar los objetos MSXML2.DOMDocument o MSXML2.XMLHTTP
' gProxy: para poner datos del proxy
' gWeb: para configurar timeouts
Public Function FwPostPage(ByVal Url As String, ByVal Params As String, Optional ContentType As String = "", Optional ApiKey As String = "") As String
   Dim hOpen As Long, hConnection As Long, hFile As Long, nRead As Long, Rc As Long
   Dim File As String, sHeader As String
   Dim xmlStr As String, tmpBuf As String * 2048
   Dim bDoLoop As Boolean, i As Integer, Port As Integer
   Dim nErr As Long, strErr As String, CharSet As String, ReqFlags As Long
   
   FwPostPage = ""
   
'   XmlBody = "<?xml version=""1.0"" encoding=""UTF-8""?>"
'   XmlBody = XmlBody & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:ns=""urn:wsident""><SOAP-ENV:Body SOAP-ENV:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/""><ns:wsident><IdentReq><wsUsuario>Autentia</wsUsuario><wsClave>@ut3nti4.</wsClave><Pais></Pais><Rut>1-8</Rut><NroAudit>AAAAA</NroAudit><Sistema>SIS</Sistema><Institucion></Institucion><Clave></Clave></IdentReq></ns:wsident></SOAP-ENV:Body></SOAP-ENV:Envelope>"
      
   If StrComp(Left(Url, 8), "https://", vbTextCompare) = 0 Then
      Port = INTERNET_DEFAULT_HTTPS_PORT
      ReqFlags = INTERNET_FLAG_SECURE
      Url = Mid(Url, 9)
   ElseIf StrComp(Left(Url, 7), "http://", vbTextCompare) = 0 Then
      Port = INTERNET_DEFAULT_HTTP_PORT
      Url = Mid(Url, 8)
      ReqFlags = 0
   End If
   
'<<<<<<< .mine
'   'Crea los objetos DOMDocument y XMLHTTP
'
'   MsgBox "MSXML2.XMLHTTP"
'   Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
'=======
   i = InStr(Url, "/")
   If i > 0 Then
      File = Mid(Url, i)
      Url = Left(Url, i - 1)
   Else
      File = ""
   End If
'>>>>>>> .r536
   
'<<<<<<< .mine
'   MsgBox "MSXML2.DOMDocument"
'   Set objDom = CreateObject("MSXML2.DOMDocument")
'
'
'   'Carga el XML
'   objDom.async = False
'   objDom.loadXML XmlBody
'=======
   i = InStr(Url, ":")
   If i > 0 Then
      Port = Val(Mid(Url, i + 1))
      Url = Left(Url, i - 1)
   End If
'>>>>>>> .r536
   
   If Url = "" Then
      Exit Function
   End If
   
   If gProxy.bProxy Then  ' 26 oct 2011: pam: se agrega soporte para Proxy
      hOpen = InternetOpen("WebProx", INTERNET_OPEN_TYPE_PROXY, gProxy.FHost, vbNullString, 0)
   Else
      hOpen = InternetOpen("", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
   End If
   If hOpen = 0 Then
      nErr = ERR.LastDllError
      strErr = "Error " & nErr & " en InternetOpen. " & GetLastSystemError(nErr)
      Exit Function
   End If

   Rc = 1
   Rc = InternetSetOption(hOpen, INTERNET_OPTION_IGNORE_OFFLINE, Rc, Len(Rc))

   If gWeb.CntTimeOut > 0 Then
      Rc = InternetSetOption(hOpen, INTERNET_OPTION_CONNECT_TIMEOUT, gWeb.CntTimeOut, Len(gWeb.CntTimeOut))
   End If

   hConnection = InternetConnect(hOpen, Url, Port, gProxy.User, gProxy.Passw, INTERNET_SERVICE_HTTP, 0, 0)
   If hConnection = 0 Then
      nErr = ERR.LastDllError
      strErr = "Error " & nErr & " en InternetConnect. " & GetLastSystemError(nErr)
      Call InternetCloseHandle(hOpen)
      Exit Function
   End If
   
   ' Uncomment to use http GET request:
   ' File = File & xml
   ' hFile = HttpOpenRequest(hConnection, "GET", File, "HTTP/1.0", vbNullString, 0, 0, 0)
                  
   ' Using http POST request:
   'hFile = HttpOpenRequest(hConnection, "POST", File, "HTTP/1.0", vbNullString, 0, ReqFlags, 1)
   hFile = HttpOpenRequest(hConnection, "POST", File, vbNullString, vbNullString, 0, ReqFlags, 1)
   'hFile = HttpOpenRequest(hConnection, "POST", "HTTP/1.0", vbNullString, vbNullString, 0, ReqFlags, 1)
   If hFile = 0 Then
      nErr = ERR.LastDllError
      strErr = "Error " & nErr & " en HttpOpenRequest. " & GetLastSystemError(nErr)
      Call InternetCloseHandle(hConnection)
      Call InternetCloseHandle(hOpen)
      Exit Function
   End If
   'MsgBox "Error: " & Err.Number & ", " & Err.Description

'   sHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf

   If ApiKey = "" Then
      If W.InDesign Then
        ApiKey = API_KEY_TECNOBACKDESA
      Else
        ApiKey = API_KEY_TECNOBACK
      End If
      sHeader = "x-api-key: " & ApiKey
      Call HttpAddRequestHeaders(hFile, sHeader, Len(sHeader), 0)
   End If
   

   If ContentType = "" Then
      'ContentType = "application/x-www-form-urlencoded"
      ContentType = "application/json"
   End If
   
   ContentType = "application/json"
   CharSet = "utf-8"

   'sHeader = "Content-Type: " & ContentType & "; charset=" & CharSet & vbCrLf
   sHeader = "Content-Type: " & ContentType '& "; charset=" & CharSet & vbCrLf
   Call HttpAddRequestHeaders(hFile, sHeader, Len(sHeader), 0)
   
   sHeader = "POST: " & File & vbCrLf
   'sHeader = "POST: https://" & Url & File & vbCrLf
   Call HttpAddRequestHeaders(hFile, sHeader, Len(sHeader), 0)
   
   If gWeb.SndTimeOut > 0 Then
      Rc = InternetSetOption(hFile, INTERNET_OPTION_SEND_TIMEOUT, gWeb.SndTimeOut, Len(gWeb.SndTimeOut))
   End If
      
   If gWeb.RcvTimeOut > 0 Then
      Rc = InternetSetOption(hFile, INTERNET_OPTION_RECEIVE_TIMEOUT, gWeb.RcvTimeOut, Len(gWeb.RcvTimeOut))
   End If
   
'   Dim myMSXML As Object
'   Set myMSXML = CreateObject("Microsoft.XmlHttp")
'myMSXML.Open "POST", "https://" & Url & File, False
'myMSXML.setRequestHeader "x-api-key", ApiKey
'myMSXML.setRequestHeader "Content-Type", "application/json"
'myMSXML.send "{ " & Mid(Params, 2) & " } "
'MsgBox myMSXML.responseText
      
   ' end of code to use POST request
   Params = "{ " & Mid(Params, 2) & " } "
'   Params = Replace(Replace(Params, ":", "="), ",", "&")
'   Params = "usuario=77049060-K&clave=ca248bdc97d6c37fc195968eaa43dc7b"
'   sHeader = "x-api-key: " & ApiKey & vbLf & ", Content-Type: " & ContentType
   If HttpSendRequest(hFile, vbNullString, 0, Params, Len(Params)) = 0 Then
      nErr = ERR.LastDllError
      strErr = GetLastSystemError(nErr)
      If nErr <> ERR.LastDllError Then
         strErr = "Error " & nErr & " en HttpSendRequest. " & strErr
      End If
      Call InternetCloseHandle(hFile)
      Call InternetCloseHandle(hConnection)
      Call InternetCloseHandle(hOpen)
      Exit Function

   End If
      
           
   bDoLoop = True
   xmlStr = ""
   Do While bDoLoop
      tmpBuf = ""
      bDoLoop = InternetReadFile(hFile, tmpBuf, Len(tmpBuf), nRead)
      If bDoLoop = False Then
         nErr = ERR.LastDllError
         strErr = "Error " & nErr & " en InternetReadFile. " & GetLastSystemError(nErr)
         Exit Do
      ElseIf nRead > 0 Then
         xmlStr = xmlStr & Left(tmpBuf, nRead)
      Else
         Exit Do
      End If
   Loop
   
   FwPostPage = xmlStr
            
   ' CheckResponse xmlStr ' xmlStr is the string that should be returned but is always empty...
   If hFile <> 0 Then Call InternetCloseHandle(hFile)
   If hConnection <> 0 Then Call InternetCloseHandle(hConnection)
   If hOpen <> 0 Then Call InternetCloseHandle(hOpen)
   
End Function   ' PostWebservice

Public Function FwPostPageSII2(Url As String, Params As String, Content As String, orden As Integer) As String

   Dim myMSXML As Object
   Set myMSXML = CreateObject("Microsoft.XmlHttp")
   Dim i As Integer
   i = InStr(Url, "/")
    'myMSXML.Open "POST", "https://" & Url & File, False
    myMSXML.Open "POST", Url & Mid(Url, i), False
    'myMSXML.setRequestHeader "x-api-key", ApiKey
    'myMSXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    myMSXML.setRequestHeader "Content-Type", Content
    myMSXML.Send Params
    'MsgBox myMSXML.responseText
    If orden = SII_LOGIN Then
        TOKEN_SII = Mid(myMSXML.GetAllResponseHeaders, InStr(1, myMSXML.GetAllResponseHeaders, "TOKEN=", vbTextCompare) + 6, 13)
    End If
    FwPostPageSII2 = myMSXML.ResponseText

End Function

' Obtiene el mensaje que viene en hexa desde url_default
Public Function GetMensajeSII(ByVal Url As String) As String
   Dim i As Integer, j As Integer, HMsg As String
   
   i = InStr(Url, "id=" & Chr(34) & "titulo" & Chr(34) & ">S")
   If i > 0 Then
      i = i + 12
      j = InStr(i, Url, "</p>") ' por si viene algo después
      
      If j > 0 Then
         HMsg = Mid(Url, i, j - i)
      Else
         HMsg = Mid(Url, i)
      End If
      
      GetMensajeSII = HMsg ' Trim(Hex2Str(HMsg))
   End If

End Function

' Invoca un web service pero sin usar los objetos MSXML2.DOMDocument o MSXML2.XMLHTTP
Public Function FwWebSaveFile(ByVal Url As String, ByVal Params As String, ByVal fname As String) As Long
   Dim hOpen As Long, hConnection As Long, hFile As Long, nRead As Long, Rc As Long
   Dim File As String, sHeader As String
   Dim tmpBuf As String * 4100, FSize As Long, Buf As String, auxBuf As String
   Dim bDoLoop As Boolean, i As Integer, Port As Integer
   Dim nErr As Long, strErr As String, OpenFlags As Long
   
   FwWebSaveFile = -1
   gWLastDllError = 0

'   XmlBody = "<?xml version=""1.0"" encoding=""UTF-8""?>"
'   XmlBody = XmlBody & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:ns=""urn:wsident""><SOAP-ENV:Body SOAP-ENV:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/""><ns:wsident><IdentReq><wsUsuario>Autentia</wsUsuario><wsClave>@ut3nti4.</wsClave><Pais></Pais><Rut>1-8</Rut><NroAudit>AAAAA</NroAudit><Sistema>SIS</Sistema><Institucion></Institucion><Clave></Clave></IdentReq></ns:wsident></SOAP-ENV:Body></SOAP-ENV:Envelope>"
   
   OpenFlags = INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION  ' 19 mar 2021: se agrega INTERNET_FLAG_RELOAD para que siempre lo baje
      
   If StrComp(Left(Url, 8), "https://", vbTextCompare) = 0 Then
      OpenFlags = OpenFlags Or INTERNET_FLAG_SECURE
      Port = INTERNET_DEFAULT_HTTPS_PORT
      Url = Mid(Url, 9)
   ElseIf StrComp(Left(Url, 7), "http://", vbTextCompare) = 0 Then
      Port = INTERNET_DEFAULT_HTTP_PORT
      Url = Mid(Url, 8)
   End If
   
'<<<<<<< .mine
'   'Crea los objetos DOMDocument y XMLHTTP
'
'   MsgBox "MSXML2.XMLHTTP"
'   Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
'=======
   i = InStr(Url, "/")
   If i > 0 Then
      File = Mid(Url, i)
      Url = Left(Url, i - 1)
   Else
      File = ""
   End If
'>>>>>>> .r536
   
'<<<<<<< .mine
'   MsgBox "MSXML2.DOMDocument"
'   Set objDom = CreateObject("MSXML2.DOMDocument")
'
'
'   'Carga el XML
'   objDom.async = False
'   objDom.loadXML XmlBody
'=======
   i = InStr(Url, ":")
   If i > 0 Then
      Port = Val(Mid(Url, i + 1))
      Url = Left(Url, i - 1)
   End If
'>>>>>>> .r536
   
   If gProxy.bProxy Then  ' 26 oct 2011: pam: se agrega soporte para Proxy
      hOpen = InternetOpen("WebProx", INTERNET_OPEN_TYPE_PROXY, gProxy.FHost, vbNullString, 0)
   Else
      hOpen = InternetOpen("", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
   End If
   If hOpen = 0 Then
      nErr = ERR.LastDllError
      gWLastDllError = nErr
      strErr = "Error " & nErr & " en InternetOpen. " & GetLastSystemError(nErr)
      Call AddLog("FwWebSaveFile: " & strErr)
      Exit Function
   End If

   Rc = 1
   Rc = InternetSetOption(hOpen, INTERNET_OPTION_IGNORE_OFFLINE, Rc, Len(Rc))

   If gWeb.CntTimeOut > 0 Then
      Rc = InternetSetOption(hOpen, INTERNET_OPTION_CONNECT_TIMEOUT, gWeb.CntTimeOut, Len(gWeb.CntTimeOut))
   End If

   hConnection = InternetConnect(hOpen, Url, Port, gProxy.User, gProxy.Passw, INTERNET_SERVICE_HTTP, 0, 0)
   If hConnection = 0 Then
      nErr = ERR.LastDllError
      gWLastDllError = nErr
      strErr = "Error " & nErr & " en InternetConnect. " & GetLastSystemError(nErr)
      Call AddLog("FwWebSaveFile: " & strErr)
      Call InternetCloseHandle(hOpen)
      Exit Function
   End If
   
   ' Uncomment to use http GET request:
   ' File = File & xml
   ' hFile = HttpOpenRequest(hConnection, "GET", File, "HTTP/1.0", vbNullString, 0, 0, 0)
                  
   ' Using http POST request:
   hFile = HttpOpenRequest(hConnection, "POST", File, vbNullString, vbNullString, 0, OpenFlags, 1)
   If hFile = 0 Then
      nErr = ERR.LastDllError
      gWLastDllError = nErr
      strErr = "Error " & nErr & " en HttpOpenRequest. " & GetLastSystemError(nErr)
      Call AddLog("FwWebSaveFile: " & strErr)
      Call InternetCloseHandle(hConnection)
      Call InternetCloseHandle(hOpen)
      Exit Function
   End If
   'MsgBox "Error: " & Err.Number & ", " & Err.Description

'   sHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
   sHeader = "Content-Type: application/x-www-form-urlencoded; charset=utf-8" & vbCrLf
   Call HttpAddRequestHeaders(hFile, sHeader, Len(sHeader), 0)
   
   sHeader = "POST: " & File & vbCrLf
   Call HttpAddRequestHeaders(hFile, sHeader, Len(sHeader), 0)
   
   If gWeb.SndTimeOut > 0 Then
      Rc = InternetSetOption(hFile, INTERNET_OPTION_SEND_TIMEOUT, gWeb.SndTimeOut, Len(gWeb.SndTimeOut))
   End If
   
   ' end of code to use POST request
   If HttpSendRequest(hFile, vbNullString, 0, Params, Len(Params)) = 0 Then
      nErr = ERR.LastDllError
      strErr = GetLastSystemError(nErr)
      gWLastDllError = nErr
      If nErr <> ERR.LastDllError Then
         strErr = "Error " & nErr & " en HttpSendRequest. " & strErr
         Debug.Print strErr

      End If
      Call AddLog("FwWebSaveFile: " & strErr)
      Call InternetCloseHandle(hFile)
      Call InternetCloseHandle(hConnection)
      Call InternetCloseHandle(hOpen)
      Exit Function

   End If
      
   If gWeb.RcvTimeOut > 0 Then
      Rc = InternetSetOption(hFile, INTERNET_OPTION_RECEIVE_TIMEOUT, gWeb.RcvTimeOut, Len(gWeb.RcvTimeOut))
   End If
           
   Dim Fd As Long
   ERR.Clear
   Fd = FreeFile()
   Open fname For Binary Access Write As #Fd
   If ERR.Number Then
      strErr = "Error " & ERR.Number & ", " & ERR.Description & " al abrir " & fname
      Call AddLog("FwWebSaveFile: " & strErr)
      Call InternetCloseHandle(hFile)
      Call InternetCloseHandle(hConnection)
      Call InternetCloseHandle(hOpen)
      Exit Function
   End If
           
   bDoLoop = True
   FSize = 0
   i = 0
   Do While bDoLoop
      bDoLoop = InternetReadFile(hFile, tmpBuf, Len(tmpBuf) - 4, nRead)
      If bDoLoop = False Then
         nErr = ERR.LastDllError
         gWLastDllError = nErr
         strErr = "Error " & nErr & " en InternetReadFile. " & GetLastSystemError(nErr)
         Call AddLog("FwWebSaveFile: " & strErr)
         FSize = -nErr
         Exit Do
      ElseIf nRead > 0 Then
         Buf = Left(tmpBuf, nRead) ' se copia a otra variable porque de lo contrario el Put le pone 4 bytes extra al inicio
         Put #Fd, , Buf
         FSize = FSize + nRead
'         xmlStr = xmlStr & Left(tmpBuf, nRead)

         If i = 0 Then
            If InStr(Buf, "HTTP Error 404 ") > 0 Then
               FSize = -404
               Exit Do
            ElseIf Right(fname, 4) = ".exe" And Left(tmpBuf, 2) <> "MZ" Then
               FSize = -404
               Exit Do
            End If
            
            auxBuf = Trim(FwGetXmlTag(Buf, "title"))
            If StrComp(auxBuf, "No se encuentra la página", vbTextCompare) = 0 Or StrComp(auxBuf, "Page not found", vbTextCompare) = 0 Then
               FSize = -20
               Exit Do
            End If
            
         End If
         i = i + 1
      Else
         Exit Do
      End If
   Loop
   
   Close #Fd
   FwWebSaveFile = FSize
      
   ' CheckResponse xmlStr ' xmlStr is the string that should be returned but is always empty...
   If hFile <> 0 Then Call InternetCloseHandle(hFile)
   If hConnection <> 0 Then Call InternetCloseHandle(hConnection)
   If hOpen <> 0 Then Call InternetCloseHandle(hOpen)
   
End Function   ' FwWebSaveFile


' Download a file from Internet and save it to a local file
'
' it works with HTTP and FTP, but you must explicitly include
' the protocol name in the URL, as in
'    CopyURLToFile "http://www.vb2themax.com/default.asp", "C:\vb2themax.htm"

Public Function FwCopyURLToFile(ByVal Url As String, ByVal Filename As String)
    Dim hInternetSession As Long
    Dim hUrl As Long
    Dim FileNum As Integer
    Dim OK As Boolean
    Dim NumberOfBytesRead As Long
    Dim Buffer As String
    Dim fileIsOpen As Boolean

    On Error GoTo ErrorHandler

    ' check obvious syntax errors
    If Len(Url) = 0 Or Len(Filename) = 0 Then ERR.Raise 5

    ' open an Internet session, and retrieve its handle
    hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, _
        vbNullString, vbNullString, 0)
    If hInternetSession = 0 Then ERR.Raise vbObjectError + 1000, , _
        "An error occurred calling InternetOpen function"

    ' open the file and retrieve its handle
    hUrl = InternetOpenUrl(hInternetSession, Url, vbNullString, 0, _
        INTERNET_FLAG_EXISTING_CONNECT, 0)
    If hUrl = 0 Then ERR.Raise vbObjectError + 1000, , _
        "An error occurred calling InternetOpenUrl function"

    ' ensure that there is no local file
    On Error Resume Next
    Kill Filename

    On Error GoTo ErrorHandler
    
    ' open the local file
    FileNum = FreeFile()
    Open Filename For Binary As FileNum
    fileIsOpen = True

    ' prepare the receiving buffer
    Buffer = Space(4096)
    
    Do
        ' read a chunk of the file - returns True if no error
        OK = InternetReadFile(hUrl, Buffer, Len(Buffer), NumberOfBytesRead)

        ' exit if error or no more data
        If NumberOfBytesRead = 0 Or Not OK Then Exit Do
        
        ' save the data to the local file
        Put #FileNum, , Left$(Buffer, NumberOfBytesRead)
    Loop
    
    ' flow into the error handler

ErrorHandler:
    ' close the local file, if necessary
    If fileIsOpen Then Close #FileNum
    ' close internet handles, if necessary
    If hUrl Then InternetCloseHandle hUrl
    If hInternetSession Then InternetCloseHandle hInternetSession
    
    ' report the error to the client, if there is one
    If ERR Then ERR.Raise ERR.Number, , ERR.Description
    
    FwCopyURLToFile = ERR.Number
End Function

Public Function GetIPFromHostName(ByVal sHostName As String) As String

   Dim nBytes As Long
   Dim ptrHosent As Long
   Dim ptrName As Long
   Dim ptrAddress As Long
   Dim ptrIPAddress As Long
   Dim sAddress As String
   
   sAddress = Space(4)
   ptrHosent = GetHostByName(sHostName & vbNullChar)
   
   If ptrHosent <> 0 Then
     ptrAddress = ptrHosent + 12
     CopyMemory ptrAddress, ByVal ptrAddress, 4            'get the IP address
     CopyMemory ptrIPAddress, ByVal ptrAddress, 4
     CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4
     GetIPFromHostName = IPToText(sAddress)
   End If
   
End Function

Private Function IPToText(ByVal IPAddress As String) As String

   IPToText = CStr(Asc(IPAddress)) & "." & _
             CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
             CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
             CStr(Asc(Mid$(IPAddress, 4, 1)))
                 
End Function
