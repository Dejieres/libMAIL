Option Compare Database
Option Explicit
Option Private Module

' Copyright 2009-2013 Denis SCHEIDT
' Ce programme est distribué sous Licence LGPL

'    This file is part of libMAIL

'    libMAIL is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Lesser General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    libMAIL is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Lesser General Public License for more details.

'    You should have received a copy of the GNU Lesser General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.


' Constantes Winsock
Public Const PF_INET = 2                        ' Internet
Public Const SOCK_STREAM = 1                    ' Connecté
Public Const SOCK_DGRAM = 2                     ' Non Connecté
Public Const IPPROTO_IP = 0                     ' Protocole par défaut (TCP si SOCK_STREAM, UDP si SOCK_DGRAM)
Public Const IPPROTO_TCP = 6                    ' TCP
Public Const IPPROTO_UDP = 17                   ' UDP
Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1
Public Const MSG_OOB = &H1                      ' Envoie des données urgentes (Out Of Band)
Public Const MSG_PEEK = &H2                     ' Ne retire pas les données de la queue de réception
Public Const MSG_DONTROUTE = &H4                ' Les données ne routeront pas

Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYS_STATUS_LEN = 128
Public Const WSAEWOULDBLOCK = 10035

Public Const FD_READ = &H1                      ' Des données normales sont arrivées
Public Const FD_WRITE = &H2                     ' Le tampon de sortie est vide
Public Const FD_OOB = &H4                       ' Des données spéciales sont arrivées
Public Const FD_ACCEPT = &H8                    ' Un appel de l'extérieur vient d'arriver
Public Const FD_CONNECT = &H10                  ' La connection est terminée
Public Const FD_CLOSE = &H20                    ' La connection vient de se fermer
Public Const INADDR_NONE                As Long = &HFFFFFFFF
Public Const INADDR_ANY                 As Long = &H0&

Public Const FIOBION                    As Long = &H8004667E

Public Const FD_SETSIZE = 64

Public Const SO_SNDBUF As Long = &H1001&        ' Send buffer size.
Public Const SO_RCVBUF As Long = &H1002&        ' Receive buffer size.
Public Const SOL_SOCKET As Long = 65535         ' Options for socket level.




' Constantes TimeZone
Public Const TIME_ID_ZONE_INVALID       As Long = 0  ' Impossible de déterminer heure d'été
Public Const TIME_ZONE_STANDARD         As Long = 1  ' Heure standard
Public Const TIME_ZONE_DAYLIGHT         As Long = 2  ' Heure d'été

' Constantes OpenFileName
Public Const OFN_ALLOWMULTISELECT       As Long = &H200
Public Const OFN_CREATEPROMPT           As Long = &H2000
Public Const OFN_ENABLEHOOK             As Long = &H20
Public Const OFN_ENABLETEMPLATE         As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE   As Long = &H80
Public Const OFN_EXPLORER               As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT     As Long = &H400
Public Const OFN_FILEMUSTEXIST          As Long = &H1000
Public Const OFN_HIDEREADONLY           As Long = &H4
Public Const OFN_LONGNAMES              As Long = &H200000
Public Const OFN_NOCHANGEDIR            As Long = &H8
Public Const OFN_NODEREFERENCELINKS     As Long = &H100000
Public Const OFN_NOLONGNAMES            As Long = &H40000
Public Const OFN_NONETWORKBUTTON        As Long = &H20000
Public Const OFN_NOREADONLYRETURN       As Long = &H8000&
Public Const OFN_NOTESTFILECREATE       As Long = &H10000
Public Const OFN_NOVALIDATE             As Long = &H100
Public Const OFN_OVERWRITEPROMPT        As Long = &H2
Public Const OFN_PATHMUSTEXIST          As Long = &H800
Public Const OFN_READONLY               As Long = &H1
Public Const OFN_SHAREAWARE             As Long = &H4000
Public Const OFN_SHAREFALLTHROUGH       As Long = 2
Public Const OFN_SHAREWARN              As Long = 0
Public Const OFN_SHARENOWARN            As Long = 1
Public Const OFN_SHOWHELP               As Long = &H10
Public Const OFN_ENABLESIZING           As Long = &H800000
Public Const OFS_MAXPATHNAME            As Long = 260

' Constantes icônes de notification
Public Const NIM_ADD                    As Long = &H0
Public Const NIM_MODIFY                 As Long = &H1
Public Const NIM_DELETE                 As Long = &H2
Public Const NIM_SETVERSION             As Long = &H4

Public Const NIF_MESSAGE                As Long = &H1
Public Const NIF_ICON                   As Long = &H2
Public Const NIF_TIP                    As Long = &H4
Public Const NIF_INFO                   As Long = &H10

Public Const NIIF_NONE                  As Long = &H0
Public Const NIIF_WARNING               As Long = &H2
Public Const NIIF_ERROR                 As Long = &H3
Public Const NIIF_INFO                  As Long = &H1
Public Const NIIF_GUID                  As Long = &H4

Public Const NOTITYICON_VERSION         As Long = &H3

Public Const WM_MOUSEMOVE               As Long = &H200
Public Const WM_LBUTTONDOWN             As Long = &H201
Public Const WM_LBUTTONUP               As Long = &H202
Public Const WM_LBUTTONDBLCLK           As Long = &H203
Public Const WM_RBUTTONDOWN             As Long = &H204
Public Const WM_RBUTTONUP               As Long = &H205
Public Const WM_RBUTTONDBLCLK           As Long = &H206

Public Const BITSPIXEL                  As Long = 12&
Public Const LOGPIXELSX                 As Long = 88&
Public Const LOGPIXELSY                 As Long = 90&
Public Const HWND_DESKTOP               As Long = 0&

Public Const QS_HOTKEY                  As Long = &H80
Public Const QS_KEY                     As Long = &H1
Public Const QS_MOUSEBUTTON             As Long = &H4
Public Const QS_MOUSEMOVE               As Long = &H2
Public Const QS_PAINT                   As Long = &H20
Public Const QS_POSTMESSAGE             As Long = &H8
Public Const QS_SENDMESSAGE             As Long = &H40
Public Const QS_TIMER                   As Long = &H10
Public Const QS_ALLINPUT                As Long = &HFF
Public Const QS_MOUSE                   As Long = &H6
Public Const QS_INPUT                   As Long = &H7
Public Const QS_ALLEVENTS               As Long = &HBF

Public Const PM_REMOVE = &H1


' Constantes pour SpecialFolders
Public Const CSIDL_DESKTOP              As Long = &H0&      ' Bureau.
Public Const CSIDL_MYDOCUMENTS          As Long = &H5&      ' Mes Documents
Public Const CSIDL_PROGRAMS_FILES       As Long = &H26&     ' Program Files.



Public Type IN_ADDR
    S_addr As Long
End Type

Public Type SOCK_ADDR
    sin_family                              As Integer
    sin_port                                As Integer
    sin_addr                                As IN_ADDR
    sin_zero(0 To 7)                        As Byte
End Type

Public Type WSA_DATA
    wVersion                               As Integer
    wHighVersion                           As Integer
    strDescription(WSADESCRIPTION_LEN + 1) As Byte
    strSystemStatus(WSASYS_STATUS_LEN + 1) As Byte
    iMaxSockets                            As Integer
    iMaxUdpDg                              As Integer
    lpVendorInfo                           As Long
End Type

Public Type fd_set
    fd_count                As Long
#If Vba7 Then
    fd_array(FD_SETSIZE)    As LongPtr
#Else
    fd_array(FD_SETSIZE)    As Long
#End If
End Type

Public Type timeval
    tv_sec                  As Long
    tv_usec                 As Long
End Type

' Structures de données utilisée par TimeZone
' SYSTEMTIME doit être définie avant TIME_ZONE_INFORMATION
Public Type SYSTEMTIME
    wYear                   As Integer
    wMonth                  As Integer
    wDayOfWeek              As Integer
    wDay                    As Integer
    wHour                   As Integer
    wMinute                 As Integer
    wSecond                 As Integer
    wMillisecond            As Integer
End Type

Public Type TIME_ZONE_INFORMATION
    Bias                    As Long             ' Décalage 'normal' (-60)
    StandardName(0 To 31)   As Integer          ' Nom de la zone
    StandardDate            As SYSTEMTIME       ' Date retour heure d'hiver (dernier dimanche d'octobre)
    StandardBias            As Long             ' Décalage heure d'hiver (0)
    DaylightName(0 To 31)   As Integer          ' Nom de la zone (Paris...)
    DaylightDate            As SYSTEMTIME       ' Date passage heure d'été (dernier dimanche de mars)
    DaylightBias            As Long             ' Décalage heure été (-60)
End Type

' Structure de données pour la boîte de dialogue 'Ouvrir'
Public Type OPENFILENAME
    lStructSize             As Long
#If Vba7 Then
    hWndOwner               As LongPtr
    hInstance               As LongPtr
#Else
    hWndOwner               As Long
    hInstance               As Long
#End If
    sFilter                 As String
    sCustomFilter           As String
    lMaxCustFilter          As Long
    lFilterIndex            As Long
    sFile                   As String
    lMaxFile                As Long
    sFileTitle              As String
    lMaxFileTitle           As Long
    sInitialDir             As String
    sDialogTitle            As String
    lFlags                  As Long
    iFileOffset             As Integer
    iFileExtension          As Integer
    sDefFileExt             As String
    lCustData               As Long
#If Vba7 Then
    lfnHook                 As LongPtr
#Else
    lfnHook                 As Long
#End If
    sTemplateName           As String
#If Vba7 Then
    pvReserved              As LongPtr
#Else
    pvReserved              As Long
#End If
    dwReserved              As Long
    lFlagsEx                As Long
End Type

' Structures de données utilisées pour les icônes de notification
Public Type BITMAPINFOHEADER
    lSize               As Long
    lWidth              As Long
    lHeight             As Long
    iPlanes             As Integer
    lBitCount           As Integer
    lCompression        As Long
    lSizeImage          As Long
    lXPelsPerMeter      As Long
    lYPelsPerMeter      As Long
    lClrUsed            As Long
    lClrImportant       As Long
End Type

Public Type RGBQUAD
    bBleu               As Byte
    bVert               As Byte
    bRouge              As Byte
    bReserve            As Byte
End Type

Public Type NOTIFYICONDATA
    cbSize              As Long                 ' Taille de la structure
#If Vba7 Then
    hwnd                As LongPtr              ' Handle de fenêtre
#Else
    hwnd                As Long                 ' Handle de fenêtre
#End If
    uId                 As Long
    uFlags              As Long
    uCallBackMessage    As Long
#If Vba7 Then
    hIcon               As LongPtr
#Else
    hIcon               As Long
#End If
    szTip               As String * 128         ' 64 avant W2K
    dwState             As Long                 ' La suite est valide à partir de W2K
    dwStateMask         As Long
    szInfo              As String * 256
    uTimeOut            As Long                 ' Délai en ms
    szInfoTitle         As String * 64
    dwInfoFlags         As Long
End Type

' Structure de données pour la version de l'OS
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type


' Structure pour DoEvents rapide
Type POINTAPI
        x               As Long
        y               As Long
End Type

Type MSG
    hwnd                As Long
    Message             As Long
    wParam              As Long
    lParam              As Long
    lTime               As Long
    pt                  As POINTAPI
End Type


' Structures pour SpecialFolders
Type SHITEMID
    SHItem              As Long
    itemID()            As Byte
End Type

Type ITEMIDLIST
    shellID             As SHITEMID
End Type

#If Vba7 Then
    Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32" () As Long
    Declare PtrSafe Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare PtrSafe Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

    Declare PtrSafe Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As LongPtr, ByVal Index As Long) As Long

    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As LongPtr)
    Declare PtrSafe Function GetComputerNameEx Lib "kernel32" Alias "GetComputerNameExA" (ByVal NameType As Long, ByVal lpBuffer As String, nSize As Long) As Long
    Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
    Declare PtrSafe Function GetUserDefaultLangID Lib "kernel32" () As Integer
    Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Declare PtrSafe Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (x As Currency) As Long
    Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (x As Currency) As Long
    Declare PtrSafe Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long

    Declare PtrSafe Function WNetGetUser Lib "Mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

    Declare PtrSafe Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwnd As LongPtr, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
    Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long

    Declare PtrSafe Function CreateIcon Lib "user32" (ByVal hInstance As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Byte, ByVal nBitsPixel As Byte, lpANDbits As Byte, lpXORbits As Byte) As LongPtr
    Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
    Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As LongPtr
    Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
    Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long

    Declare PtrSafe Function bind Lib "wsock32" (ByVal sock As Long, addr As SOCK_ADDR, ByVal namelen As Long) As Long
    Declare PtrSafe Function closesocket Lib "wsock32" (ByVal sock As Long) As Long
    Declare PtrSafe Function connect Lib "wsock32" (ByVal sock As Long, Name As SOCK_ADDR, ByVal namelen As Integer) As Long
    Declare PtrSafe Function gethostbyname Lib "wsock32" (ByVal hostname As String) As LongPtr
    Declare PtrSafe Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Declare PtrSafe Function htons Lib "wsock32" (ByVal hostshort As Integer) As Integer
    Declare PtrSafe Function inet_addr Lib "wsock32" (ByVal cp As String) As Long
    Declare PtrSafe Function ioctlsocket Lib "wsock32" (ByVal sock As LongPtr, ByVal cmd As Long, argp As Long) As Long
    Declare PtrSafe Function listen Lib "wsock32" (ByVal sock As Long, ByVal backlog As Integer) As Integer
    Declare PtrSafe Function recv Lib "wsock32" (ByVal sock As Long, buffer As Any, ByVal Length As Long, ByVal flags As Long) As Long
    Declare PtrSafe Function send Lib "wsock32" (ByVal sock As Long, buffer As Any, ByVal Length As Long, ByVal flags As Long) As Long
    Declare PtrSafe Function socket Lib "wsock32" (ByVal afinet As Integer, ByVal socktype As Integer, ByVal protocol As Integer) As Long
    Declare PtrSafe Function WSAAsyncSelect Lib "wsock32" (ByVal sock As Long, ByVal hwnd As Long, ByVal wMsg As Integer, ByVal lEvent As Long) As Integer
    Declare PtrSafe Function WSACleanup Lib "wsock32" () As Long
    Declare PtrSafe Function WSAGetLastError Lib "wsock32" () As Long
    Declare PtrSafe Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Integer, wsData As WSA_DATA) As Long
    Declare PtrSafe Function WSSelect Lib "wsock32" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
#Else
    Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
    Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

    Declare Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long

    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)
    Declare Function GetComputerNameEx Lib "kernel32" Alias "GetComputerNameExA" (ByVal NameType As Long, ByVal lpBuffer As String, nSize As Long) As Long
    Declare Function GetLastError Lib "kernel32" () As Long
    Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
    Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Declare Function QueryPerformanceCounter Lib "kernel32" (x As Currency) As Boolean
    Declare Function QueryPerformanceFrequency Lib "kernel32" (x As Currency) As Boolean
    Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long

    Declare Function WNetGetUser Lib "Mpr" Alias "WNetGetUserA" (ByVal lpName As Any, ByVal lpUserName As String, lpnLength As Long) As Long

    Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwnd As Long, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
    Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long

    Declare Function CreateIcon Lib "user32" (ByVal hInstance As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Byte, ByVal nBitsPixel As Byte, lpANDbits As Byte, lpXORbits As Byte) As Long
    Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
    Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
    Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long

    Declare Function bind Lib "wsock32" (ByVal sock As Long, addr As SOCK_ADDR, ByVal namelen As Long) As Long
    Declare Function closesocket Lib "wsock32" (ByVal sock As Long) As Long
    Declare Function connect Lib "wsock32" (ByVal sock As Long, Name As SOCK_ADDR, ByVal namelen As Integer) As Long
    Declare Function gethostbyname Lib "wsock32" (ByVal hostname As String) As Long
    Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Declare Function htons Lib "wsock32" (ByVal hostshort As Integer) As Integer
    Declare Function inet_addr Lib "wsock32" (ByVal cp As String) As Long
    Declare Function ioctlsocket Lib "wsock32" (ByVal sock As Long, ByVal cmd As Long, argp As Long) As Long
    Declare Function listen Lib "wsock32" (ByVal sock As Long, ByVal backlog As Integer) As Integer
    Declare Function recv Lib "wsock32" (ByVal sock As Long, buffer As Any, ByVal Length As Long, ByVal flags As Long) As Long
    Declare Function send Lib "wsock32" (ByVal sock As Long, buffer As Any, ByVal Length As Long, ByVal flags As Long) As Long
    Declare Function socket Lib "wsock32" (ByVal afinet As Integer, ByVal socktype As Integer, ByVal protocol As Integer) As Long
    Declare Function WSAAsyncSelect Lib "wsock32" (ByVal sock As Long, ByVal hwnd As Long, ByVal wMsg As Integer, ByVal lEvent As Long) As Integer
    Declare Function WSACleanup Lib "wsock32" () As Long
    Declare Function WSAGetLastError Lib "wsock32" () As Long
    Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Integer, wsData As WSA_DATA) As Long
    Declare Function WSSelect Lib "wsock32" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
#End If