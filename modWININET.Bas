Attribute VB_Name = "modWININET"
Option Explicit
Public hOpen As Long
Public hConnection As Long


Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2


Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)

Public Const MAX_PATH = 260
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


Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Const ERROR_NO_MORE_FILES = 18

Public Declare Function InternetTimeToSystemTime Lib "wininet.dll" (ByVal lpszTime As String, ByRef pst As SYSTEMTIME, ByVal dwReserved As Long) As Long

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
    
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Public Const scUserAgent = "vb wininet"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H1
Public Const INTERNET_FLAG_PASSIVE = &H8000000

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003
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

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Public Const INTERNET_FLAG_MULTIPART = &H200000

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
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

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

'################## Internet Auto Dial ###########
Public Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal hWndParent As Long) As Long
Public Const INTERNET_AUTODIAL_FORCE_ONLINE = 1&
Public Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2&
Public Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4&
Public Enum AutoDialsFlags
    ADF_FORCE_ONLINE = INTERNET_AUTODIAL_FORCE_ONLINE
    ADF_FORCE_UNATTENDED = INTERNET_AUTODIAL_FORCE_UNATTENDED
End Enum
'################################################

'##### USED TO CHANGED FILE DATE/TIME ######
Public Declare Function SetFileTimeWrite Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, ByVal MullP As Long, ByVal NullP2 As Long, lpLastWriteTime As FileTime) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FileTime) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FileTime, lpFileTime As FileTime) As Long
Public Declare Function FileTimeToSystemTime& Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME)
'###########################################

'

'###### USED IN DETECTING INTERNET CONNECTION
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpSFlags As Long, ByVal dwReserved As Long) As Long
Public Const INTERNET_CONNECTION_LAN As Long = &H2
Public Const INTERNET_CONNECTION_MODEM As Long = &H1
'#############################


'###### USED IN DISCONNECTING NETWORK
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function RASHangUp Lib "RasApi32.dll" Alias "RasHangUpA" (ByVal hRascon As Long) As Long
Public Declare Function RASEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnection As Long) As Long
Public Declare Function RASGetConnectionStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRascon As Long, lpStatus As Any) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Type RASConnType
    dwSize As Long
    hRascon As Long
End Type

Public Type RASConnStatusType
    dwSize As Long
    RASConnState As Long
    dwError As Long
End Type

Public Const SW_SHOWNORMAL = 1
Public Const HKEY_DYN_DATA = &H80000006
Public Const ERROR_SUCCESS = 0
Public Const KEY_QUERY_VALUE = &H1
Public Const REG_DWORD = 4
'##########################################



Public Function AutoDial(hwndParentWindow As Long, lOption As AutoDialsFlags, Optional bFailIfSecurityCheck = True) As Boolean
    Dim lFlags As Long
    Dim lRetValue As Long
    
    On Local Error GoTo ErrHandle
    
    If bFailIfSecurityCheck Then lFlags = lOption Or INTERNET_AUTODIAL_FAILIFSECURITYCHECK
    
    lRetValue = InternetAutodial(lFlags, hwndParentWindow)
    
    If lRetValue <> 0 Then AutoDial = True
ErrHandle:
End Function

Private Function CalcFDate(FTime As SYSTEMTIME) As String
    Dim newDate As String
    
    On Local Error Resume Next
    
    With FTime
      newDate = Format$(.wDay, "00") & "/" & Format$(.wMonth, "00") & "/" & Format$(.wYear, "0000")
    End With
    newDate = Format$(newDate, "dd/mmm/yyyy")
    CalcFDate = newDate
End Function

Public Function CalcFTime(FTime As SYSTEMTIME) As String
    Dim newDate As String
    
    On Local Error Resume Next
    
    With FTime
        newDate = Format$(.wHour, "00") & ":" & Format$(.wMinute, "00") & ":" & Format$(.wSecond, "00")
    End With
    CalcFTime = newDate
End Function

Public Function ModifyFileStamp(Filename As String, newDate As Date, newTime As Date) As Integer
    Dim x As Long
    Dim Handle As Long
    Dim System_Time As SYSTEMTIME
    Dim File_Time As FileTime
    Dim Local_Time As FileTime
    Dim TimeStamp As Variant
    Dim intYear As Integer
    Dim intMonth As Integer
    Dim intDay As Integer
    Dim intHour As Integer
    Dim intMinute As Integer
    Dim intSecond As Integer
    
    On Local Error Resume Next
    
    intYear = Format$(newDate, "yyyy")
    intMonth = Format$(newDate, "mm")
    intDay = Format$(newDate, "dd")
    intHour = Format$(newTime, "hh")
    intMinute = Format$(newTime, "nn")
    intSecond = Format$(newTime, "ss")
    
    TimeStamp = DateSerial(intYear, intMonth, intDay) + TimeSerial(intHour, intMinute, intSecond)
  
    System_Time.wYear = Format$(CDate(TimeStamp), "yyyy") 'year(TimeStamp)
    System_Time.wMonth = Month(TimeStamp)
    System_Time.wDay = Day(TimeStamp)
    System_Time.wDayOfWeek = Weekday(TimeStamp) - 1
    System_Time.wHour = Hour(TimeStamp)
    System_Time.wMinute = Minute(TimeStamp)
    System_Time.wSecond = Second(TimeStamp)
    System_Time.wMilliseconds = 0

    'convert the system time to a file time
    x = SystemTimeToFileTime(System_Time, Local_Time)

    'convert local file time to file time based on UTC
    x = LocalFileTimeToFileTime(Local_Time, File_Time)

    'open the file so we can get a file handle to the file
    Handle = CreateFile(Filename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    'now change the file time and date stamp
    ModifyFileStamp = SetFileTimeWrite(Handle, ByVal 0&, ByVal 0&, File_Time)
    CloseHandle Handle

End Function

'Status ONLINE
Public Function Online() As Boolean
    'If you are online it will return True, otherwise False
    Online = InternetGetConnectedState(0&, 0&)
End Function

'Status LAN
Public Function ViaLAN() As Boolean
    Dim SFlags As Long
    'return the flags associated with the connection
    Call InternetGetConnectedState(SFlags, 0&)

    'True if the Sflags has a LAN connection
    ViaLAN = SFlags And INTERNET_CONNECTION_LAN
End Function

'Status MODEM
Public Function ViaModem() As Boolean
    Dim SFlags As Long
    'return the flags associated with the connection
    Call InternetGetConnectedState(SFlags, 0&)

    'True if the Sflags has a modem connection
    ViaModem = SFlags And INTERNET_CONNECTION_MODEM
End Function

'Hangup
Public Sub Hangup()
    Dim hRASConn As Long
    
    On Local Error Resume Next
    
    hRASConn = GetRASConnHandle()
    
    If hRASConn <> 0 Then Call RASHangUp(hRASConn)
        
End Sub

'Hangup
Private Function GetRASConnHandle() As Long
    Dim udtRASConn(255) As RASConnType
    Dim udtStatus As RASConnStatusType
    Dim lRet As Long

    On Local Error Resume Next
    
    udtRASConn(0).dwSize = 412
    
    lRet = RASEnumConnections(udtRASConn(0), udtRASConn(0).dwSize * 256, 0)
    
    If lRet <> 0 Then
        GetRASConnHandle = 0
    Else
        GetRASConnHandle = udtRASConn(0).hRascon
    End If
    
End Function


