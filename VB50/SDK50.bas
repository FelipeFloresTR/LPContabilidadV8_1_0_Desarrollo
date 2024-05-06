Attribute VB_Name = "SDK50"
Option Explicit

' Err.Number
' https://msdn.microsoft.com/en-us/library/aa264975(v=VS.60).aspx
Public Const ERR_FILENOTFND As Integer = 53     ' File Not Found
Public Const ERR_FILEOPEN As Integer = 55       ' File already open
Public Const ERR_FILEEXIST As Integer = 58      ' File already exist
Public Const ERR_DISKFULL As Integer = 61       ' Disk full
Public Const ERR_TOMANYFILES As Integer = 67    ' Too many files
Public Const ERR_PATHFILE As Integer = 75       ' Path/File access error.
Public Const ERR_PATHNOFND As Integer = 76      ' Path not found

Public Const ERR_EXIST As Integer = ERR_PATHFILE   ' Directorio ya existe

Type RECT_T
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI_T
   x As Long
   Y As Long
End Type

Public Type FILETIME_T
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_MORE_DATA = 234&
Public Const ERROR_FILE_NOT_FOUND = 2&

'Grid
'ColAlignment,FixedAlignment Properties
Global Const GRID_ALIGNLEFT = 0
Global Const GRID_ALIGNRIGHT = 1
Global Const GRID_ALIGNCENTER = 2

Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Declare Function GetDialogBaseUnits Lib "USER32" () As Long

Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer

Public Declare Function SetCursorPos Lib "USER32" (ByVal x As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI_T) As Long
Declare Function GetCursorPos2 Lib "USER32" (lpPoint As Long) As Long
Declare Function LoadCursor Lib "USER32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Declare Function SetCursor Lib "USER32" (ByVal hCursor As Long) As Long
Global Const IDC_ARROW = 32512
Global Const IDC_CROSS = 32515
Global Const IDC_WAIT = 32514

Declare Function SetCapture Lib "USER32" (ByVal hWnd As Long) As Long
Declare Function ReleaseCapture Lib "USER32" () As Long

Declare Function GetFocus Lib "USER32" () As Long
Declare Function ApiSetFocus Lib "USER32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Declare Sub GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT_T)

Declare Function ScreenToClient Lib "USER32" (ByVal hWnd As Long, lpPoint As POINTAPI_T) As Long
Declare Function ClientToScreen Lib "USER32" (ByVal hWnd As Long, lpPoint As POINTAPI_T) As Long
Declare Function GetClientRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT_T) As Long

Declare Function SetParent Lib "USER32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_BROADCAST = &HFFFF&

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOACTIVATE = &H10


Public Declare Function ShowWindow Lib "USER32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5
Public Const SW_SHOWNA = 8
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWDEFAULT = 10

Declare Function DrawText Lib "USER32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT_T, ByVal wFormat As Long) As Long
' DrawText() Format Flags
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000

Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function FillRect Lib "USER32" (ByVal hDC As Long, lpRect As RECT_T, ByVal hBrush As Long) As Long
Declare Function InvertRect Lib "USER32" (ByVal hDC As Long, lpRect As RECT_T) As Long
Declare Function DrawFocusRect Lib "USER32" (ByVal hDC As Long, lpRect As RECT_T) As Long

Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Declare Function DeletePen Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Public Const PS_SOLID = 0
Public Const PS_DOT = 2

Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Const WHITE_BRUSH = 0
Public Const BLACK_BRUSH = 4
Public Const NULL_BRUSH = 5
Public Const NULL_PEN = 8

Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Declare Function DeleteBrush Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long

Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Const HS_BDIAGONAL = 3
Public Const HS_BDIAGONAL1 = 7
Public Const HS_CROSS = 4
Public Const HS_DENSE1 = 9
Public Const HS_DENSE2 = 10
Public Const HS_DENSE3 = 11
Public Const HS_DENSE4 = 12
Public Const HS_DENSE5 = 13
Public Const HS_DENSE6 = 14
Public Const HS_DENSE7 = 15
Public Const HS_DENSE8 = 16
Public Const HS_DIAGCROSS = 5
Public Const HS_DITHEREDBKCLR = 24
Public Const HS_FDIAGONAL = 2
Public Const HS_FDIAGONAL1 = 6
Public Const HS_HALFTONE = 18
Public Const HS_HORIZONTAL = 0
Public Const HS_NOSHADE = 17
Public Const HS_SOLID = 8
Public Const HS_SOLIDBKCLR = 23
Public Const HS_SOLIDCLR = 19
Public Const HS_SOLIDTEXTCLR = 21
Public Const HS_VERTICAL = 1

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Global Const LOGPIXELSX = 88
Global Const LOGPIXELSY = 90
Global Const HORZRES = 8
Global Const VERTRES = 10

Declare Function InvalidateRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT_T, ByVal bErase As Long) As Long
Declare Function InvalidateRect2 Lib "USER32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Declare Function UpdateWindow Lib "USER32" (ByVal hWnd As Long) As Long

Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageS Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)

Public Const WS_EX_APPWINDOW = &H40000

Public Const LBS_NOSEL = &H4000     ' Para Listbox

'Public Const WM_SETREDRAW = &HB
'Public Const WM_COMMAND = &H111
'Public Const WM_MOUSEMOVE = &H200
'Public Const WM_LBUTTONDOWN = &H201
'Public Const WM_LBUTTONUP = &H202
'Public Const WM_LBUTTONDBLCLK = &H203
'Public Const WM_RBUTTONDOWN = &H204
'Public Const WM_RBUTTONUP = &H205
'Public Const WM_RBUTTONDBLCLK = &H206
'Public Const WM_GETTEXT = &HD

Enum WM_Messages
   WM_SETREDRAW = &HB
   WM_COMMAND = &H111
   WM_MOUSEMOVE = &H200
   WM_LBUTTONDOWN = &H201
   WM_LBUTTONUP = &H202
   WM_LBUTTONDBLCLK = &H203
   WM_RBUTTONDOWN = &H204
   WM_RBUTTONUP = &H205
   WM_RBUTTONDBLCLK = &H206
   WM_GETTEXT = &HD
   WM_USER = &H400
End Enum


Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)

Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR

' EDIT Messages
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_GETLINECOUNT = &HBA

'End Enum

Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2

Public Const WM_WININICHANGE = &H1A
Public Const WM_SETTINGCHANGE = WM_WININICHANGE

' Ctes para VM_VSCROLL y VM_HSCROLL
Public Const SB_LINEDOWN = 1
Public Const SB_LINEUP = 0
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGEUP = 2
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5

Declare Function BeginPaint Lib "USER32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT_T) As Long
Declare Function EndPaint Lib "USER32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT_T) As Long
Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Type PAINTSTRUCT_T
   hDC As Long
   fErase As Long
   rcPaint As RECT_T
   fRestore As Long
   fIncUpdate As Long
   rgbReserved As Byte
End Type

Declare Function GetBoundsRect Lib "gdi32" (ByVal hDC As Long, lprcBounds As RECT_T, ByVal Flags As Long) As Long
Declare Function SetBoundsRect Lib "gdi32" (ByVal hDC As Long, lprcBounds As RECT_T, ByVal Flags As Long) As Long
Public Const DCB_ACCUMULATE = &H2
Public Const DCB_DISABLE = &H8
Public Const DCB_ENABLE = &H4
Public Const DCB_RESET = &H1

Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI_T) As Long

Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

Declare Function GetSysColor Lib "USER32" (ByVal nIndex As Long) As Long
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15

Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
Public Const SM_CYCAPTION = 4
Public Const SM_CYMENU = 15
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYHSCROLL = 3

Declare Function SetScrollPos Lib "USER32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Declare Function ShowScrollBar Lib "USER32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Declare Function SetScrollRange Lib "USER32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Public Const SB_HORZ = 0
Public Const SB_BOTH = 3
Public Const SB_VERT = 1

Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Public Const MM_ANISOTROPIC = 8

Type SIZE_T
   cx As Long
   cy As Long
End Type

Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE_T) As Long
Declare Function SetWindowExtEx2 Lib "gdi32" Alias "SetWindowExtEx" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByVal lpSize As Long) As Long
Declare Function SetViewportExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE_T) As Long
Declare Function SetViewportExtEx2 Lib "gdi32" Alias "SetViewportExtEx" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByVal lpSize As Long) As Long
Declare Function ScaleWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As SIZE_T) As Long
Declare Function ScaleWindowExtEx2 Lib "gdi32" Alias "ScaleWindowExtEx" (ByVal hDC As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, ByVal lpSize As Long) As Long

Public Declare Function MsgBeep Lib "USER32" Alias "MessageBeep" (ByVal wType As VbMsgBoxStyle) As Byte

Declare Function GetMenu Lib "USER32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "USER32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function TrackPopupMenu Lib "USER32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, pRect As Long) As Long
Public Declare Function HiliteMenuItem Lib "USER32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Public Declare Function GetMenuString Lib "USER32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetMenuItemInfo Lib "USER32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO_T) As Long
Public Declare Function GetMenuItemCount Lib "USER32" (ByVal hMenu As Long) As Long
Public Declare Function SetMenuDefaultItem Lib "USER32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long

Public Const MF_BYCOMMAND = &H0&
Public Const MF_HILITE = &H80&
Public Const MF_UNHILITE = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_STRING = &H0&
Public Const MIIM_ID = &H2&
Public Const MIIM_TYPE = &H10&

Public Type MENUITEMINFO_T
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String ' * 21
    cch As Long
End Type

Public Declare Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function EnumThreadWindows Lib "USER32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "USER32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function EraseSection Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As String) As Long
Declare Function FlushProfile Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Zero As Long, ByVal Zero As Long, ByVal Zero As Long, ByVal lplFileName As String) As Long
Declare Function EraseEntry Lib "kernel32" Alias "WritePrivateProfileString" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String) As Long

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
        (ByVal dwMessage As Long, lpData As NOTIFYICONDATA_T) As Integer

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function ShellExecute2 Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Long, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function SIIReg Lib "SII.dll" Alias "GetCompras" (ByVal rut As String, ByVal dv As String, ByVal password As String, ByVal year As String, ByVal month As String) As String


Declare Function SetForegroundWindow Lib "USER32" (ByVal hWnd As Long) As Long

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1

Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "USER32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Type VERINFO_t               'Version FIXEDFILEINFO
    'There is data in the following two dwords, but it is for Windows internal
    '   use and we should ignore it
    Ignore(1 To 8) As Byte
    'Signature As Long
    'StrucVersion As Long
    FileVerPart2 As Integer
    FileVerPart1 As Integer
    FileVerPart4 As Integer
    FileVerPart3 As Integer
    ProductVerPart2 As Integer
    ProductVerPart1 As Integer
    ProductVerPart4 As Integer
    ProductVerPart3 As Integer
    FileFlagsMask As Long 'VersionFileFlags
    FileFlags As Long 'VersionFileFlags
    FileOS As Long 'VersionOperatingSystemTypes
    FileType As Long
    FileSubtype As Long 'VersionFileSubTypes
    'I've never seen any data in the following two dwords, so I'll ignore them
    Ignored(1 To 8) As Byte 'DateHighPart As Long, DateLowPart As Long
End Type

Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, ByVal lpData As String) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'Public Declare Function VerQueryValue Lib "version.dll" (pBlock As Any, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, puLen As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal lpBuf As Long, ByVal szReceive As String, lpBufPtr As Long, lLen As Long) As Long
Public Declare Function VerQueryValue2 Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As String, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, puLen As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Public Declare Function HtmlHelpTopic Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As String) As Long


' HH API const
'------------------------------------------------------------------------------
Public Const HH_DISPLAY_TOPIC = &H0         ' select last opened tab, [display a specified topic]
Public Const HH_DISPLAY_TOC = &H1           ' select contents tab, [display a specified topic]
Public Const HH_DISPLAY_INDEX = &H2         ' select index tab and searches for a keyword
Public Const HH_DISPLAY_SEARCH = &H3        ' select search tab (perform a search is fixed here)
      
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or
  
Public Const HH_HELP_CONTEXT = &HF          ' display mapped numeric value in dwData
     
Public Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to WinHelp
Public Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to WinHelp

' HH_DISPLAY_SEARCH Command Related Structures and Constants
'-----------------------------------------------------------
Public Const HH_FTS_DEFAULT_PROXIMITY = -1
Public Const HH_MAX_TABS = 19               ' maximum number of tabs

Public Declare Function SetMenuItemBitmaps Lib "USER32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

' Este módulo permite registrar el producto en el Registry

Public Const MAX_KEY_LENGTH = 255
Public Const MAX_VALUE_NAME = 16383

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExS Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueExS Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long  ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME_T) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME_T) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2

Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const SYNCHRONIZE = &H100000
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' Constants that will be used in the API functions
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&

' Declare the needed API functions
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal bsName As String, ByVal Buff As String, ByVal ch As Long) As Long

Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal plpFilePart As Long) As Long

Declare Function GetUniversalPath Lib "Mpr" Alias "WNetGetUniversalNameA" (ByVal lpLocalPath As String, ByVal dwInfoLevel As Long, ByVal lpBuffer As String, lpBufferSize As Long) As Long
Private Const UNIVERSAL_NAME_INFO_LEVEL As Integer = 1   ' The function returns a simple string with the UNC name.
Private Const REMOTE_NAME_INFO_LEVEL As Integer = 2      ' The function returns a tuple based in the Win32 REMOTE_NAME_INFO data structure.

Declare Function AllocConsole Lib "kernel32" () As Long
Declare Function GetConsoleWindow Lib "kernel32" () As Long
Declare Function FreeConsole Lib "kernel32" () As Long
Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, ByVal lpReserved As Long) As Long
Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long

Private Declare Function fCreateShellLink Lib "VB5STKIT.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const VER_NT_WORKSTATION = 1
Private Const VER_NT_DOMAIN_CONTROLLER = 2
Private Const VER_NT_SERVER = 3

Public Type OSVERSIONINFOEX_T
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
    wServicePackMajor As Integer 'win2000 only
    wServicePackMinor As Integer 'win2000 only
    wSuiteMask As Integer 'win2000 only
    wProductType As Byte 'win2000 only
    wReserved As Byte
End Type

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX_T) As Long

'* Para Service

Public Const SERVICE_WIN32_OWN_PROCESS = &H10&

Public Const SERVICE_RUNNING = &H4&

Public Const SC_MANAGER_ALL_ACCESS = &HF003F

Public Type SERVICE_STATUS_T
   dwServiceType As Long
   dwCurrentState As Long
   dwControlsAccepted As Long
   dwWin32ExitCode As Long
   dwServiceSpecificExitCode As Long
   dwCheckPoint As Long
   dwWaitHint As Long
End Type

Public Declare Function SetServiceStatus Lib "advapi32.dll" (ByVal hServiceStatus As Long, lpServiceStatus As SERVICE_STATUS_T) As Long
Public Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenSCManager2 Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long

Public Declare Function OemToChar Lib "USER32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Declare Function CharToOem Lib "USER32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Declare Function CharToOemBuf Lib "USER32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Public Declare Function OemToCharBuf Lib "USER32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long

Public Const CP_UTF7 = 65000  ' UTF-7 translation
Public Const CP_UTF8 = 65001  ' UTF-8 translation
Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Declare Function WideCharToMultiByte1 Lib "kernel32" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'****************** Para NotifyIconData ***************
Type NOTIFYICONDATA_T
    cbSize           As Long
    hWnd             As Long        'handle of window that receives notification message
    uID              As Long        'app-defined identifier of taskbar icon
    uFlags           As Long        'flag settings
    uCallbackMessage As Long        'app-defined message identifier
    hIcon            As Long        'handle to an icon
    szTip            As String * 64 'tool text display message
End Type

'define Notify Icon Action
Enum NotifyIcon_Action
    NIA_ADD = &H0
    NIA_MODIFY = &H1
    NIA_DELETE = &H2
End Enum

'define Notify Icon Data
Enum NotifyIcon_DATA
    NID_MESSAGE = &H1
    NID_ICON = &H2
    NID_TIP = &H4
End Enum

'************** NotifyIconData *****************

'********** Para CreateProcess ******************
Public Type STARTUPINFO
   Cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

' dwFlags
Public Const STARTF_USESHOWWINDOW = &H1

Public Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Public Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&



'************** FIN CreateProcess *********************

'*************** NetBios *******************

Private Const NCBNAMSZ = 16    '  absolute length of a net name
Private Const NCBRESET = &H32  ' NCB RESET
Private Const NCBASTAT = &H33  ' NCB ADAPTER STATUS
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const HEAP_ZERO_MEMORY As Long = &H8

Private Type ADAPTER_STATUS_t
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type

Private Type NET_CONTROL_BLOCK_t  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte 'Reserved, must be 0
   ncb_event      As Long
End Type

Public Type NAME_BUFFER_t
   Name           As String * NCBNAMSZ
   name_num       As Integer
   name_flags     As Integer
End Type

Public Type ASTAT_t
   adapt          As ADAPTER_STATUS_t
   NameBuff(30)   As NAME_BUFFER_t
End Type

Public Type ADAPTER_STATUS_t_old
   adapter_address As String * 6
   rev_major As Integer
   reserved0 As Integer
   adapter_type As Integer
   rev_minor As Integer
   duration As Integer
   frmr_recv As Integer
   frmr_xmit As Integer
   iframe_recv_err As Integer
   xmit_aborts As Integer
   xmit_success As Long
   recv_success As Long
   iframe_xmit_err As Integer
   recv_buff_unavail As Integer
   t1_timeouts As Integer
   ti_timeouts As Integer
   Reserved1 As Long
   free_ncbs As Integer
   max_cfg_ncbs As Integer
   max_ncbs As Integer
   xmit_buf_unavail As Integer
   max_dgram_size As Integer
   pending_sess As Integer
   max_cfg_sess As Integer
   max_sess As Integer
   max_sess_pkt_size As Integer
   name_count As Integer
End Type

Public Type NCB_t_old
   ncb_command As Integer
   ncb_retcode As Integer
   ncb_lsn As Integer
   ncb_num As Integer
   ncb_buffer As String
   ncb_length As Integer
   ncb_callname As String * NCBNAMSZ
   ncb_name As String * NCBNAMSZ
   ncb_rto As Integer
   ncb_sto As Integer
   ncb_post As Long
   ncb_lana_num As Integer
   ncb_cmd_cplt As Integer
   ncb_reserve(10) As Byte  ' Reserved, must be 0
   ncb_event As Long
End Type

Private Declare Function Netbios Lib "netapi32.dll" (pncb As NET_CONTROL_BLOCK_t) As Byte

Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" _
  (ByVal hHeap As Long, _
   ByVal dwFlags As Long, _
   ByVal dwBytes As Long) As Long
     
Private Declare Function HeapFree Lib "kernel32" _
  (ByVal hHeap As Long, _
   ByVal dwFlags As Long, _
   lpMem As Any) As Long

' *** Punteros en VB
' Para usar con las funciones: VarPtr, StrPtr y ObjPtr
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'*********** PARA MAC Address

Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260 ' 256 + 4
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132 ' 128 + 4
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 6
Public Const MIB_IF_TYPE_TOKENRING = 9
Public Const MIB_IF_TYPE_FDDI = 15
Public Const MIB_IF_TYPE_PPP = 23
Public Const MIB_IF_TYPE_LOOPBACK = 24
Public Const MIB_IF_TYPE_SLIP = 28

Type IP_ADDR_STRING
   Next As Long
   IPAddress As String * 16
   IpMask As String * 16
   Context As Long
End Type

Type IP_ADAPTER_INFO_t
   Next As Long
   ComboIndex As Long
   AdapterName As String * MAX_ADAPTER_NAME_LENGTH
   Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
   AddressLength As Long
   Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
   Index As Long
   Type As Long
   DhcpEnabled As Long
   CurrentIpAddress As Long
   IpAddressList As IP_ADDR_STRING
   GatewayList As IP_ADDR_STRING
   DhcpServer As IP_ADDR_STRING
   HaveWins As Byte
   PrimaryWinsServer As IP_ADDR_STRING
   SecondaryWinsServer As IP_ADDR_STRING
   LeaseObtained As Long
   LeaseExpires As Long
   xxx   As String * 3  ' para completar los 640
End Type

Declare Function GetAdaptersInfo Lib "Iphlpapi.dll" (AdapterInfo As Any, pOutBufLen As Long) As Long

Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal Border As Long) As Long

' **** Para Open File *****

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME_T) As Long

Public Type OPENFILENAME_T
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

'***** Para Browse for Folder ****************

Public Type BrowseInfo_t
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
Public Const BIF_BROWSEINCLUDEFILES As Long = &H4000

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo_t) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Type SYSTEMTIME_t
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        WDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME_t)

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SETSELECTION As Long = (WM_USER + 102)
Public gBrowseForFolderDefault As String


' ********** para ShellAndWait

Public Enum ShellAndWaitResult_e
    Success = 0
    Failure = 1
    TimeOut = 2
    InvalidParameter = 3
    SysWaitAbandoned = 4
    UserWaitAbandoned = 5
    UserBreak = 6
End Enum

Public Enum ActionOnBreak_e
    IgnoreBreak = 0
    AbandonWait = 1
    PromptUser = 2
End Enum

Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Private Const STATUS_WAIT_0 As Long = &H0
Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)
Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
Private Const WAIT_TIMEOUT As Long = 258&
Private Const WAIT_FAILED As Long = &HFFFFFFFF
Private Const WAIT_INFINITE = -1&

' *****************************


Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function

' Funcion para asignar el directorio por defecto
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

    ' Look for BFFM_INITIALIZED
    If uMsg = BFFM_INITIALIZED Then
        Call SendMessageS(hWnd, BFFM_SETSELECTION, True, gBrowseForFolderDefault)
    End If
    
    BrowseCallbackProc = 0
End Function

Public Function BrowseForFile(hwndOwner As Long, ByVal sPrompt As String, Optional ByVal DefFolder As String = "") As String
   Dim iNull As Integer
   Dim lpIDList As Long
   Dim lResult As Long
   Dim sPath As String
   Dim udtBI As BrowseInfo_t
      
   'initialise variables
   With udtBI
      .hwndOwner = hwndOwner
      .lpszTitle = lstrcat(sPrompt, "")
      .ulFlags = BIF_BROWSEINCLUDEFILES
      
      If DefFolder <> "" Then
         gBrowseForFolderDefault = DefFolder
         .lpfnCallback = FARPROC(AddressOf BrowseCallbackProc)
      End If
      
   End With
   
   'Call the browse for folder API
   lpIDList = SHBrowseForFolder(udtBI)
    
   'get the resulting string path
   If lpIDList Then
      sPath = String(MAX_PATH, 0)
      lResult = SHGetPathFromIDList(lpIDList, sPath)
      Call CoTaskMemFree(lpIDList)
      iNull = InStr(sPath, vbNullChar)
      If iNull Then
         sPath = Left(sPath, iNull - 1)
      Else
         sPath = ""
      End If
   Else
      sPath = ""
   End If
   
   'If cancel was pressed, sPath = ""
   BrowseForFile = sPath

End Function

Public Function BrowseForFolder(hwndOwner As Long, ByVal sPrompt As String, Optional ByVal DefFolder As String = "") As String
   Dim iNull As Integer
   Dim lpIDList As Long
   Dim lResult As Long
   Dim sPath As String
   Dim udtBI As BrowseInfo_t
      
   'initialise variables
   With udtBI
      .hwndOwner = hwndOwner
      .lpszTitle = lstrcat(sPrompt, "")
      .ulFlags = BIF_RETURNONLYFSDIRS
      
      If DefFolder <> "" Then
         gBrowseForFolderDefault = DefFolder
         .lpfnCallback = FARPROC(AddressOf BrowseCallbackProc)
      End If
      
   End With
   
   'Call the browse for folder API
   lpIDList = SHBrowseForFolder(udtBI)
    
   'get the resulting string path
   If lpIDList Then
      sPath = String(MAX_PATH, 0)
      lResult = SHGetPathFromIDList(lpIDList, sPath)
      Call CoTaskMemFree(lpIDList)
      iNull = InStr(sPath, vbNullChar)
      If iNull Then
         sPath = Left(sPath, iNull - 1)
      Else
         sPath = ""
      End If
   Else
      sPath = ""
   End If
   
   'If cancel was pressed, sPath = ""
   BrowseForFolder = sPath

End Function

Sub MoveTo(ByVal hDC As Long, ByVal x As Integer, ByVal Y As Integer)
   Dim AuxPt As POINTAPI_T
   
   Call MoveToEx(hDC, x, Y, AuxPt)
End Sub
Sub DestroyPen(hPen As Long)
   Call DeleteObject(hPen)
   hPen = 0
End Sub
Sub DestroyBrush(hBrush As Long)
   Call DeleteObject(hBrush)
   hBrush = 0
End Sub

Sub xFillRect(ByVal hDC As Long, lpRect As RECT_T, ByVal hBrush As Long)
   Dim Rct As RECT_T
   
   Rct.Left = lpRect.Left
   Rct.Top = lpRect.Top
   
   Rct.Right = lpRect.Right
   Rct.Bottom = lpRect.Bottom

   Call FillRect(hDC, Rct, hBrush)
   
End Sub

Function CreateBrush(Color As Long)
   Dim Col As Long

   If Color And &H80000000 Then
      Col = GetSysColor(Color And &HFFFFFF)
   Else
      Col = Color
      
   End If
   
   CreateBrush = CreateSolidBrush(Col)

End Function

Public Function GetUserName() As String
   Dim Buf As String * 51
   Dim Rc As Long
   
   Rc = GetUserNameA(Buf, 50)
   'GetUserName = Left(Buf, StrLen(Buf))
   Rc = InStr(Buf, Chr(0))
   GetUserName = Left(Buf, Rc - 1)

End Function
Public Function GetComputerName() As String
   Dim Buf As String * 51
   Dim Rc As Long
   
   Rc = GetComputerNameA(Buf, 50)
   'GetComputerName = Left(Buf, StrLen(Buf))
   Rc = InStr(Buf, Chr(0))
   GetComputerName = Left(Buf, Rc - 1)
   
End Function

Public Function DelRegValue(ByVal hKey As Long, ByVal Path As String, ByVal SubKey As String) As Long
   Dim Key As Long, Rc As Long

   DelRegValue = 0

   Rc = RegOpenKeyEx(hKey, Path, 0, KEY_SET_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   ' Eliminamos el valor
   DelRegValue = RegDeleteValue(Key, SubKey)
   
   Rc = RegCloseKey(Key)
   Key = 0

End Function
' Ej: hKey = HKEY_LOCAL_MACHINE
Public Function QryRegValue(ByVal hKey As Long, ByVal Path As String, ByVal SubKey As String, Optional ByVal Def As String = "") As String
   Dim Key As Long, Rc As Long, BufLen As Long, DataType As Long
   Dim Buf1 As String, lValue As Long

   QryRegValue = ""

   Rc = RegOpenKeyEx(hKey, Path, 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   ' Obtenemos el tipo y largo antes de consultar
   BufLen = 0
   DataType = 0
   Rc = RegQueryValueEx(Key, SubKey, 0&, DataType, ByVal 0&, BufLen)
         
   Select Case DataType
      Case REG_SZ, REG_BINARY, REG_EXPAND_SZ:       ' String data
         If BufLen > 0 Then ' 4 nov 2020: en el equipo de un cliente de Contabilidad (Rodolfo Friz) se cae cuando BufLen=0.
            Buf1 = Space(BufLen + 10)
            BufLen = Len(Buf1) - 5
            Rc = RegQueryValueExS(Key, SubKey, 0, REG_SZ, Buf1, BufLen)
         
            If Rc = 0 Then
               QryRegValue = Left(Buf1, BufLen - 1)
            Else
               QryRegValue = Def
            End If
         End If
      
      Case REG_DWORD:    ' numeric data
         Rc = RegQueryValueEx(Key, SubKey, 0&, DataType, lValue, 4&)      ' 4& = 4-byte word (long integer)
         If Rc = 0 Then
            QryRegValue = str(lValue)
         Else
            QryRegValue = Def
         End If
            
   End Select
   
   Rc = RegCloseKey(Key)
   Key = 0

End Function

Public Function SetRegValue(ByVal hKey As Long, ByVal Path As String, ByVal SubKey As String, Value As Variant) As Long
   Dim Key As Long, Rc As Long, BufLen As Long, DataType As Long
   Dim Buf1 As String * 301
   Dim StrValue As String, lValue As Long

   Rc = RegOpenKeyEx(hKey, Path, 0, KEY_SET_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   If IsNumeric(Value) Then
      DataType = REG_DWORD
   Else
      DataType = REG_SZ
   End If

   Select Case DataType
      Case REG_SZ:       ' String data
         StrValue = Trim(Value) & Chr(0)     ' null terminated
         SetRegValue = RegSetValueExS(Key, SubKey, 0&, DataType, StrValue, Len(StrValue))
                                   
      Case REG_DWORD:    ' numeric data
         lValue = CLng(Value)
         SetRegValue = RegSetValueEx(Key, SubKey, 0&, DataType, lValue, 4&)                                            ' 4& = 4-byte word (long integer)
   End Select
   
   Rc = RegCloseKey(Key)

End Function
Public Function CreateRegKey(ByVal hKey As Long, ByVal Path As String) As Long
   Dim Key As Long, Rc As Long

   CreateRegKey = RegCreateKey(hKey, Path, Key)

   Rc = RegCloseKey(Key)

End Function

'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
Public Function RegKeyRead(i_RegKey As String) As String
   Dim myWS As Object

   On Error Resume Next
  'access Windows scripting
   Set myWS = CreateObject("WScript.Shell")
   If Not myWS Is Nothing Then
  'read key from registry
      RegKeyRead = myWS.RegRead(i_RegKey)
   End If
End Function

Private Function MakeMacAddress(b() As Byte, sDelim As String) As String

   Dim Cnt As Long
   Dim Buff As String
   
   On Local Error GoTo MakeMac_error
 
  'so far, MAC addresses are
  'exactly 6 segments in size (0-5)
   If UBound(b) = 5 Then
   
     'concatenate the first five values
     'together and separate with the
     'delimiter char
      For Cnt = 0 To 4
         Buff = Buff & Right$("00" & Hex(b(Cnt)), 2) & sDelim
      Next
      
     'and append the last value
      Buff = Buff & Right$("00" & Hex(b(5)), 2)
         
   End If  'UBound(b)
   
   MakeMacAddress = Buff
   
MakeMac_exit:
   Exit Function
   
MakeMac_error:
   MakeMacAddress = "(error building MAC address)"
   Resume MakeMac_exit
   
End Function



Public Function GetVersionInfo() As String
    Dim myOS As OSVERSIONINFOEX_T
    Dim bExInfo As Boolean
    Dim sOS As String

    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
    'try win2000 version
    If GetVersionEx(myOS) = 0 Then
        'if fails
        myOS.dwOSVersionInfoSize = 148 'ignore reserved data
        If GetVersionEx(myOS) = 0 Then
            GetVersionInfo = "Microsoft Windows (Unknown)"
            Exit Function
        End If
    Else
        bExInfo = True
    End If
    
    With myOS
        'is version 4
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            'nt platform
            Select Case .dwMajorVersion
            Case 3, 4
                sOS = "Microsoft Windows NT"
            Case 5
                sOS = "Microsoft Windows 2000/XP"
            End Select
            If bExInfo Then
                'workstation/server?
                If .wProductType = VER_NT_SERVER Then
                    sOS = sOS & " Server"
                ElseIf .wProductType = VER_NT_DOMAIN_CONTROLLER Then
                    sOS = sOS & " Domain Controller"
                ElseIf .wProductType = VER_NT_WORKSTATION Then
                    sOS = sOS & " Workstation"
                End If
            End If
            
            'get version/build no
            sOS = sOS & " Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & StripTerminator(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
            
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            'get minor version info
            If .dwMinorVersion = 0 Then
                sOS = "Microsoft Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                sOS = "Microsoft Windows 98"
            ElseIf .dwMinorVersion = 90 Then
                sOS = "Microsoft Windows Millenium"
            Else
                sOS = "Microsoft Windows 9?"
            End If
            'get version/build no
            sOS = sOS & "Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & StripTerminator(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
        End If
    End With
    GetVersionInfo = sOS
End Function
Private Function StripTerminator(sString As String) As String
    StripTerminator = Left$(sString, InStr(sString, Chr$(0)) - 1)
End Function
' Hace que las extesiones de archivos sean visibles
Public Sub RegUnhideFileExt()
   Dim Rc As Long, hKey As Long
   
   Rc = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", 0, KEY_SET_VALUE, hKey)
   If Rc <> ERROR_SUCCESS Then
      Exit Sub
   End If
   
   Rc = 0 ' No ocultar
   Rc = RegSetValueEx(hKey, "HideFileExt", 0, REG_DWORD, Rc, Len(Rc))
   'Debug.Print "RegUnhideFileExt: " & Rc

   Rc = RegCloseKey(hKey)

End Sub

Public Function GetFileVersion1(ByVal fname As String) As String
   Dim bInfo() As Byte, iLen As Long, Rc As Long, lTemp As Long
   Dim Value As String, vLen As Long
   Dim i As Integer, j As Integer, k As Integer, l As Integer
   Dim hKey As Long, Buff As String * 255
   Dim lpBuffer As Long, VerInfo As VERINFO_t
   Const sEXE As String * 1 = "\"
   Dim Buf As String

   On Error Resume Next
   
   'Call AddLog("GetFileVersion: a GetFileVersionInfoSize(" & FName & ")")
   
   GetFileVersion1 = ""
   iLen = GetFileVersionInfoSize(fname & vbNullChar, lTemp)
   If iLen <= 0 Then
      GetFileVersion1 = ""
      Exit Function
   End If

   ReDim bInfo(iLen + 1000)
   
   'Call AddLog("GetFileVersion: a GetFileVersionInfo iLen=" & iLen)
   Rc = GetFileVersionInfo(fname & vbNullChar, lTemp, iLen, VarPtr(bInfo(0)))
   If Rc = 0 Then
      GetFileVersion1 = ""
      Exit Function
   End If
   
   For i = 0 To iLen
      If bInfo(i) Then
         Debug.Print i & ") " & bInfo(i) & " [" & Chr(bInfo(i)) & "]"
      End If
   Next i
      
   Rc = VerQueryValue(VarPtr(bInfo(0)), sEXE & vbNullChar, lpBuffer, iLen)
   If Rc = 0 Then
      GetFileVersion1 = ""
      Exit Function
   End If
   
   CopyMemory VerInfo, ByVal lpBuffer, iLen
   
   
   'Call AddLog("GetFileVersion: a sacar caracteres nulos. iLen=" & iLen)
   
   Value = ""
   For i = 1 To Len(Buf) / 2
      Value = Value & Mid(Buf, i * 2 + 1, 1)
   Next i
   
   Buf = Value
   
   'Call AddLog("GetFileVersion: a leer info Buf=[" & Buf & "]")
   
   i = InStr(Buf, "Comments")
   If i Then
      Buf = Mid(Buf, i - 2)
   Else
      i = InStr(Buf, "CompanyName")
      If i Then
         Buf = Mid(Buf, i - 2)
      End If
   End If
   
   k = 1
   l = 0
   Value = ""
   
   Do
      i = InStr(k, Buf, Chr(1), vbBinaryCompare)
                  
      If i <= 0 Then
         Exit Do
      End If
   
      If Asc(Mid(Buf, i + 1, 1)) = 1 Then  ' largo = a marca
         i = i + 1
      End If
   
      vLen = Asc(Mid(Buf, i - 1, 1))
      j = InStr(i, Buf, Chr(0), vbBinaryCompare)
         
      l = InStr(j + 1, Buf, Chr(1), vbBinaryCompare)
      
      If l > 0 And l - i < 10 Then  ' pura basura...
         k = l + 1
      Else
      
         'Call AddLog("GetFileVersion: i=" & i & ", j=" & j & ", l=" & l & ", vLen=" & vLen)
         
         If l <> 0 And vLen > l - j - 3 Then
            vLen = l - j - 3
            'Call AddLog("GetFileVersion: vLen=" & vLen)
         End If
         
         If vLen <= 0 Then
            If Value <> "" Then
               Value = Value & vbCrLf
            End If
            k = i + 1
         Else
            
            Value = Value & Mid(Buf, i + 1, j - i - 1) & ": " & vbTab & ReplaceStr(Mid(Buf, j + 1, vLen), vbCrLf, " ") & vbCrLf
         
            Debug.Print Value
         
            k = j + vLen
         End If
         
      End If
      
   Loop
   
   Value = Trim(ReplaceStr(Value, Chr(0), ""))
   For i = 1 To 5
      Value = ReplaceStr(Value, vbCrLf & vbCrLf, vbCrLf)
   Next i
   
   GetFileVersion1 = Value
   'Call AddLog("GetFileVersion: saliendo")

End Function


'   Dim Fso As Scripting.FileSystemObject
'
'   Set Fso = New Scripting.FileSystemObject
'
'   Fn = gConfig.NewVerPath & "\" & App.EXEName & ".exe_"
'
'   If Fso.FileExists(Fn) Then
'      Ver = Fso.GetFileVersion(Fn)
'   End If

Public Function GetFileVersion(ByVal fname As String) As String
   Dim Buf As String, iLen As Long, Rc As Long
   Dim Value As String, vLen As Long
   Dim i As Integer, j As Integer, k As Integer, l As Integer
   Dim hKey As Long, Buff As String * 255

   On Error Resume Next
   
   'Call AddLog("GetFileVersion: a GetFileVersionInfoSize(" & FName & ")")
   
   GetFileVersion = ""
   iLen = GetFileVersionInfoSize(fname, App.hInstance)
   If iLen <= 0 Then
      GetFileVersion = ""
      Exit Function
   End If

   Buf = String(iLen + 100, Chr(0))
   
   'Call AddLog("GetFileVersion: a GetFileVersionInfo iLen=" & iLen)
   Rc = GetFileVersionInfo(fname, App.hInstance, iLen, Buf)
   If Rc = 0 Then
      GetFileVersion = ""
      Exit Function
   End If
      
   'Call AddLog("GetFileVersion: a sacar caracteres nulos. iLen=" & iLen)
   
   Value = ""
   For i = 1 To Len(Buf) / 2
      Value = Value & Mid(Buf, i * 2 + 1, 1)
   Next i
   
   Buf = Value
   
   'Call AddLog("GetFileVersion: a leer info Buf=[" & Buf & "]")
   
   i = InStr(Buf, "Comments")
   If i Then
      Buf = Mid(Buf, i - 2)
   Else
      i = InStr(Buf, "CompanyName")
      If i Then
         Buf = Mid(Buf, i - 2)
      End If
   End If
   
   k = 1
   l = 0
   Value = ""
   
   Do
      i = InStr(k, Buf, Chr(1), vbBinaryCompare)
                  
      If i <= 0 Then
         Exit Do
      End If
   
      If Asc(Mid(Buf, i + 1, 1)) = 1 Then  ' largo = a marca
         i = i + 1
      End If
   
      vLen = Asc(Mid(Buf, i - 1, 1))
      j = InStr(i, Buf, Chr(0), vbBinaryCompare)
         
      l = InStr(j + 1, Buf, Chr(1), vbBinaryCompare)
      
      If l > 0 And l - i < 10 Then  ' pura basura...
         k = l + 1
      Else
      
         'Call AddLog("GetFileVersion: i=" & i & ", j=" & j & ", l=" & l & ", vLen=" & vLen)
         
         If l <> 0 And vLen > l - j - 3 Then
            vLen = l - j - 3
            'Call AddLog("GetFileVersion: vLen=" & vLen)
         End If
         
         If vLen <= 0 Then
            If Value <> "" Then
               Value = Value & vbCrLf
            End If
            k = i + 1
         Else
            
            Value = Value & Mid(Buf, i + 1, j - i - 1) & ": " & vbTab & ReplaceStr(Mid(Buf, j + 1, vLen), vbCrLf, " ") & vbCrLf
         
            Debug.Print Value
         
            k = j + vLen
         End If
         
      End If
      
   Loop
   
   Value = Trim(ReplaceStr(Value, Chr(0), ""))
   For i = 1 To 5
      Value = ReplaceStr(Value, vbCrLf & vbCrLf, vbCrLf)
   Next i
   
   GetFileVersion = Value
   'Call AddLog("GetFileVersion: saliendo")

End Function

Public Function GetMac() As String
   Dim Mac As String, n As Integer

   Mac = GetMacAddress_new(n)
   If Mac = "" Then
      Mac = GetMacAddress2() ' 8 ene 2016: se pone esta funcion antes
      If Mac = "" Then
         Mac = GetMacAddress()
      End If
   End If

   GetMac = Mac
   
End Function

Public Function GetMacAddress() As String
   Dim AdapterInfo As IP_ADAPTER_INFO_t  ' Allocate information
   Dim dwBufLen As Long, Rc As Long
   Dim b As Byte, Mac As String, i As Integer
   Dim Buf As String
   
   On Error Resume Next
   
   dwBufLen = Len(AdapterInfo)            ' Save memory size of buffer
   
   Rc = GetAdaptersInfo(AdapterInfo, dwBufLen)    ' Call GetAdapterInfo
                                                    
   If Rc = 0 Then
      Mac = ""
      If AdapterInfo.Type = MIB_IF_TYPE_ETHERNET Then
         For i = 0 To 5  '  Bytes 0 through 5 inclusive
            b = AdapterInfo.Address(i)
            Mac = Mac & "-" & Right("00" & Hex(b), 2)
         Next i
      End If
      GetMacAddress = Mid(Mac, 2)
   Else
      GetMacAddress = ""
   End If

                                             
End Function


Public Function GetMacAddress_new(nAd As Integer) As String
   Dim AdapterInfo As IP_ADAPTER_INFO_t  ' Allocate information
   Dim InfoSize As Long, Rc As Long, pAdapt As Long
   Dim b As Byte, Mac As String, i As Integer, nTot As Integer
   Dim Buf As String, Descr As String
   Dim AdapterInfoBuffer() As Byte

   On Error Resume Next
   GetMacAddress_new = ""

   InfoSize = 0
   Rc = GetAdaptersInfo(ByVal 0&, InfoSize)
   If Rc <> 0 And InfoSize = 0 Then
      Exit Function
   End If

   If InfoSize < Len(AdapterInfo) Then ' por siaca
      Exit Function
   End If

   nTot = Int(InfoSize / Len(AdapterInfo))

   ReDim AdapterInfoBuffer(InfoSize + 100) As Byte

   Rc = GetAdaptersInfo(AdapterInfoBuffer(0), InfoSize)    ' Call GetAdapterInfo
   If Rc <> 0 Then
      GetMacAddress_new = ""
      Exit Function
   End If
   
   CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)
   pAdapt = AdapterInfo.Next

   Mac = ""
   nAd = 0

   Do
      nAd = nAd + 1
      If AdapterInfo.Type = MIB_IF_TYPE_ETHERNET Then
         Descr = Left(AdapterInfo.Description, StrLen(AdapterInfo.Description))
         
         ' Si no es virtual o bluetooth o vpn
         If InStr(1, Descr, "virtual", vbTextCompare) = 0 And InStr(1, Descr, "bluetooth", vbTextCompare) = 0 And InStr(1, Descr, "vpn", vbTextCompare) = 0 And InStr(1, Descr, "TAP-Win", vbTextCompare) = 0 Then
            
            Mac = ""
            For i = 0 To AdapterInfo.AddressLength - 1 '  Bytes 0 through 5 inclusive
               b = AdapterInfo.Address(i)
               Mac = Mac & "-" & Right("00" & Hex(b), 2)
            Next i
            
            Mac = Mid(Mac, 2)
            
            Call AddDebug("1589: PC: " & GetComputerName() & ", MAC: " & Mac & ", " & Descr)
            
            Exit Do
         End If
      End If
      
      ' controlamos por dos lados
      If pAdapt <= 0 Or nAd >= nTot Then
         Exit Do
      End If
      
      CopyMemory AdapterInfo, ByVal pAdapt, Len(AdapterInfo)
      pAdapt = AdapterInfo.Next

   Loop

   GetMacAddress_new = Mac
                                             
End Function

' 8 ene 2016:
Public Function GetMacAddress2() As String
   Dim Devices As Object
   Dim Device As Object
   Dim temp As Variant
   Dim Info As String, Mac As String
   
   ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa394216(v=vs.85).aspx#properties
   
   Set Devices = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapter")
   For Each Device In Devices
      If Device.PhysicalAdapter = True And Device.AdapterTypeID = 0 And Device.ConfigManagerErrorCode = 0 _
      And StrComp(Left(Device.PNPDeviceID, 4), "PCI\") = 0 Then
   
         For Each temp In Device.Properties_
            If StrComp(temp.Name, "MACAddress") = 0 Then
               If IsNull(temp) Then
                  Info = "NULL"
               Else
                  Info = CStr(temp)
                  
                  If Device.NetEnabled Then ' privilegiamos la conectada
                     Mac = Info
                  ElseIf Mac = "" Then
                     Mac = Info
                  End If
                  
               End If
               
'               Info = Info & " - " & Device.NetEnabled & " - " & Device.PNPDeviceID & " - " & Device.AdapterType & " - " & Device.Name
'               List1.AddItem Info
               Exit For
            End If
         Next temp
      End If
   Next Device
   
   GetMacAddress2 = Replace(Mac, ":", "-")
   
End Function


Public Function GetIPAddress() As String
   Dim AdapterInfo As IP_ADAPTER_INFO_t  ' Allocate information
   Dim InfoSize As Long, Rc As Long, pAdapt As Long
   Dim b As Byte, IP As String, IP_Conn As String, i As Integer, nTot As Integer
   Dim Buf As String, nAd As Integer
   Dim AdapterInfoBuffer() As Byte

   On Error Resume Next
   GetIPAddress = ""

   InfoSize = 0
   Rc = GetAdaptersInfo(ByVal 0&, InfoSize)
   If Rc <> 0 And InfoSize = 0 Then
      Exit Function
   End If

   If InfoSize < Len(AdapterInfo) Then ' por siaca
      Exit Function
   End If

   nTot = Int(InfoSize / Len(AdapterInfo))

   ReDim AdapterInfoBuffer(InfoSize + 100) As Byte

   Rc = GetAdaptersInfo(AdapterInfoBuffer(0), InfoSize)    ' Call GetAdapterInfo
   If Rc <> 0 Then
      GetIPAddress = ""
      Exit Function
   End If
   
   CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)
   pAdapt = AdapterInfo.Next

   IP = ""
   nAd = 0

   Do
      nAd = nAd + 1
      If AdapterInfo.Type = MIB_IF_TYPE_ETHERNET Then
         
         If InStr(1, AdapterInfo.Description, "virtual", vbTextCompare) = 0 And InStr(1, AdapterInfo.Description, "bluetooth", vbTextCompare) = 0 Then
            IP = Trim0(AdapterInfo.IpAddressList.IPAddress)
            If IP <> "0.0.0.0" Then
               Exit Do
            End If
         End If
      End If
      
      ' controlamos por dos lados
      If pAdapt <= 0 Or nAd >= nTot Then
         Exit Do
      End If
      
      CopyMemory AdapterInfo, ByVal pAdapt, Len(AdapterInfo)
      pAdapt = AdapterInfo.Next

   Loop

   GetIPAddress = IP
                                             
End Function

' Returns an array with the local IP addresses (as strings).
' Author: Christian d'Heureuse, www.source-code.biz
Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim Rc As Long
   Rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If Rc <> 0 Then
      ERR.Raise vbObjectError, , "GetIpAddrTable failed with return value " & Rc
   End If
   
   Dim NrOfEntries As Integer
   NrOfEntries = Buf(1) * 256 + Buf(0)
   
   If NrOfEntries = 0 Then
      GetIpAddrTable = Array()
      Exit Function
   End If
   
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, s As String: s = ""
      For j = 0 To 3
         s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j)
      Next
      IpAddrs(i) = s
   Next i
      
   GetIpAddrTable = IpAddrs
End Function

Public Sub AddTrayIcon(Frm As Form, hWnd As Long, Icondata As NOTIFYICONDATA_T, ByVal InfoText As String, Optional ByVal Oper As NotifyIcon_Action = NotifyIcon_Action.NIA_ADD)
   Dim Rc As Long
        
   With Icondata
      'define size of structure
      .cbSize = Len(Icondata)
      'define the form's handle
      .hWnd = hWnd
      
      'specify NULL for the id
      .uID = vbNull
      .uID = hWnd
      
      'no callback procedure used
      .uCallbackMessage = WM_Messages.WM_LBUTTONDOWN
      .uCallbackMessage = WM_Messages.WM_MOUSEMOVE
      'specify the Icon to use -- forms
      
      'Me.Icon = LoadPicture("c:\Windows\bmps\honey.bmp")
      'Me.Icon = LoadPicture("p:\icons\arte\pint05.ico")
      .hIcon = Frm.Icon
      'specify the TIP to use -- forms
      .szTip = InfoText & vbNullChar
      
      'set flags to indicate program is
      'specifying Callback, Icon, and Tip
      .uFlags = NotifyIcon_DATA.NID_MESSAGE Or _
                NotifyIcon_DATA.NID_ICON Or _
                NotifyIcon_DATA.NID_TIP
                
   End With
    
   'add to the SysTray
   ' Rc = Shell_NotifyIcon(NotifyIcon_Action.NIA_ADD, Icondata)
   Rc = Shell_NotifyIcon(Oper, Icondata)
   
End Sub

' La ventana debe tener ShowInTaskBar = false
Public Sub ToTaskBar(ByVal Frm As Form)

   If Frm.ShowInTaskbar Then
      Debug.Print "*** ERROR: ShowInTaskbar debe estar en False ***"
   End If

   ' show this form on the taskbar
   Call SetWindowLong(Frm.hWnd, GWL_EXSTYLE, (GetWindowLong(Frm.hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW))

End Sub

Public Function CapsLockOn() As Boolean
    Dim xState As Integer
    xState = GetKeyState(vbKeyCapital)
    CapsLockOn = (xState = 1 Or xState = -127)
End Function
Public Function GetLastSystemError(ByVal LastDllErr As Long) As String

   Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
   Dim sError As String  '\\ Preinitilise a string buffer to put any error message into
   Dim lErrMsg As Long
   
   sError = Space(1000)
   
   lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, LastDllErr, 0, sError, Len(sError), 0)
   
   GetLastSystemError = Trim(Left(sError, lErrMsg))

End Function
' Ojo: la hora puede ser diferente por el cambio de hora
' Retorna hasta milisegundos
Public Function SystemTime() As Double
   Dim sTm As SYSTEMTIME_t, Tm As Double
    
   On Error Resume Next
   Tm = Now
   Call GetSystemTime(sTm)

   If ERR.LastDllError = 0 Then
      SystemTime = Int(Tm) + Hour(Tm) / 24 + sTm.wMinute / (24 * 60) + sTm.wSecond / (24# * 60# * 60#) + sTm.wMilliseconds / (24# * 60# * 60# * 1000#)
   Else
      SystemTime = Tm
   End If

End Function
' Cuenta la cantidad de veces que se está ejecutando un proceso (xxx.exe)
' independiente del usuario
Public Function ProcessCount(ByVal Process As String) As Long
   Dim objWMIService, colProcesses
   
   Set objWMIService = GetObject("winmgmts:")
   Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name='" & Process & "'")
   ProcessCount = colProcesses.Count
   
   Set colProcesses = Nothing
   Set objWMIService = Nothing
   
End Function
Public Sub PGSetBackColor(ByVal ProgressBarHwnd As Long, ByVal RGBValue As Long)
    Call SendMessage(ProgressBarHwnd, SB_SETBKCOLOR, 0, RGBValue)
End Sub
 
Public Sub PGSetBarColor(ByVal ProgressBarHwnd As Long, ByVal RGBValue As Long)
    Call SendMessage(ProgressBarHwnd, PBM_SETBARCOLOR, 0, RGBValue)
End Sub
'
'Public Function ShellAndWait(ShellCommand As String, _
'                    TimeOutMs As Long, _
'                    ShellWindowState As VbAppWinStyle, _
'                    BreakKey As ActionOnBreak_e) As ShellAndWaitResult_e
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' ShellAndWait
''
'' This function calls Shell and passes to it the command text in ShellCommand. The function
'' then waits for TimeOutMs (in milliseconds) to expire.
''
''   Parameters:
''       ShellCommand
''           is the command text to pass to the Shell function.
''
''       TimeOutMs
''           is the number of milliseconds to wait for the shell'd program to wait. If the
''           shell'd program terminates before TimeOutMs has expired, the function returns
''           ShellAndWaitResult.Success = 0. If TimeOutMs expires before the shell'd program
''           terminates, the return value is ShellAndWaitResult.TimeOut = 2.
''
''       ShellWindowState
''           is an item in VbAppWinStyle specifying the window state for the shell'd program.
''
''       BreakKey
''           is an item in ActionOnBreak indicating how to handle the application's cancel key
''           (Ctrl Break). If BreakKey is ActionOnBreak.AbandonWait and the user cancels, the
''           wait is abandoned and the result is ShellAndWaitResult.UserWaitAbandoned = 5.
''           If BreakKey is ActionOnBreak.IgnoreBreak, the cancel key is ignored. If
''           BreakKey is ActionOnBreak.PromptUser, the user is given a ?Continue? message. If the
''           user selects "do not continue", the function returns ShellAndWaitResult.UserBreak = 6.
''           If the user selects "continue", the wait is continued.
''
''   Return values:
''            ShellAndWaitResult.Success = 0
''               indicates the the process completed successfully.
''            ShellAndWaitResult.Failure = 1
''               indicates that the Wait operation failed due to a Windows error.
''            ShellAndWaitResult.TimeOut = 2
''               indicates that the TimeOutMs interval timed out the Wait.
''            ShellAndWaitResult.InvalidParameter = 3
''               indicates that an invalid value was passed to the procedure.
''            ShellAndWaitResult.SysWaitAbandoned = 4
''               indicates that the system abandoned the wait.
''            ShellAndWaitResult.UserWaitAbandoned = 5
''               indicates that the user abandoned the wait via the cancel key (Ctrl+Break).
''               This happens only if BreakKey is set to ActionOnBreak.AbandonWait.
''            ShellAndWaitResult.UserBreak = 6
''               indicates that the user broke out of the wait after being prompted with
''               a ?Continue message. This happens only if BreakKey is set to
''               ActionOnBreak.PromptUser.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Dim TaskID As Long
'Dim ProcHandle As Long
'Dim WaitRes As Long
'Dim Ms As Long
'Dim MsgRes As VbMsgBoxResult
'Dim SaveCancelKey As XlEnableCancelKey
'Dim ElapsedTime As Long
'Dim Quit As Boolean
'Const ERR_BREAK_KEY = 18
'Const DEFAULT_POLL_INTERVAL = 500
'
'If Trim(ShellCommand) = vbNullString Then
'    ShellAndWait = ShellAndWaitResult_e.InvalidParameter
'    Exit Function
'End If
'
'If TimeOutMs < 0 Then
'    ShellAndWait = ShellAndWaitResult_e.InvalidParameter
'    Exit Function
'ElseIf TimeOutMs = 0 Then
'    Ms = WAIT_INFINITE
'Else
'    Ms = TimeOutMs
'End If
'
'Select Case BreakKey
'    Case AbandonWait, IgnoreBreak, PromptUser
'        ' valid
'    Case Else
'        ShellAndWait = ShellAndWaitResult_e.InvalidParameter
'        Exit Function
'End Select
'
'Select Case ShellWindowState
'    Case vbHide, vbMaximizedFocus, vbMinimizedFocus, vbMinimizedNoFocus, vbNormalFocus, vbNormalNoFocus
'        ' valid
'    Case Else
'        ShellAndWait = ShellAndWaitResult_e.InvalidParameter
'        Exit Function
'End Select
'
'On Error Resume Next
'Err.Clear
'TaskID = Shell(ShellCommand, ShellWindowState)
'If (Err.Number <> 0) Or (TaskID = 0) Then
'    ShellAndWait = ShellAndWaitResult_e.Failure
'    Exit Function
'End If
'
'ProcHandle = OpenProcess(SYNCHRONIZE, False, TaskID)
'If ProcHandle = 0 Then
'    ShellAndWait = ShellAndWaitResult_e.Failure
'    Exit Function
'End If
'
'On Error GoTo ErrH:
'SaveCancelKey = Application.EnableCancelKey
'Application.EnableCancelKey = xlErrorHandler
'WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
'Do Until WaitRes = WAIT_OBJECT_0
'    DoEvents
'    Select Case WaitRes
'        Case WAIT_ABANDONED
'            ' Windows abandoned the wait
'            ShellAndWait = ShellAndWaitResult_e.SysWaitAbandoned
'            Exit Do
'        Case WAIT_OBJECT_0
'            ' Successful completion
'            ShellAndWait = ShellAndWaitResult_e.Success
'            Exit Do
'        Case WAIT_FAILED
'            ' attach failed
'            ShellAndWait = ShellAndWaitResult_e.Failure
'            Exit Do
'        Case WAIT_TIMEOUT
'            ' Wait timed out. Here, this time out is on DEFAULT_POLL_INTERVAL.
'            ' See if ElapsedTime is greater than the user specified wait
'            ' time out. If we have exceed that, get out with a TimeOut status.
'            ' Otherwise, reissue as wait and continue.
'            ElapsedTime = ElapsedTime + DEFAULT_POLL_INTERVAL
'            If Ms > 0 Then
'                ' user specified timeout
'                If ElapsedTime > Ms Then
'                    ShellAndWait = ShellAndWaitResult_e.TimeOut
'                    Exit Do
'                Else
'                    ' user defined timeout has not expired.
'                End If
'            Else
'                ' infinite wait -- do nothing
'            End If
'            ' reissue the Wait on ProcHandle
'            WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
'
'        Case Else
'            ' unknown result, assume failure
'            ShellAndWait = ShellAndWaitResult_e.Failure
'            Exit Do
'            Quit = True
'    End Select
'Loop
'
'CloseHandle ProcHandle
'Application.EnableCancelKey = SaveCancelKey
'Exit Function
'
'ErrH:
'Debug.Print "ErrH: Cancel: " & Application.EnableCancelKey
'If Err.Number = ERR_BREAK_KEY Then
'    If BreakKey = ActionOnBreak_e.AbandonWait Then
'        CloseHandle ProcHandle
'        ShellAndWait = ShellAndWaitResult_e.UserWaitAbandoned
'        Application.EnableCancelKey = SaveCancelKey
'        Exit Function
'    ElseIf BreakKey = ActionOnBreak_e.IgnoreBreak Then
'        Err.Clear
'        Resume
'    ElseIf BreakKey = ActionOnBreak_e.PromptUser Then
'        MsgRes = MsgBox("User Process Break." & vbCrLf & _
'            "Continue to wait?", vbYesNo)
'        If MsgRes = vbNo Then
'            CloseHandle ProcHandle
'            ShellAndWait = ShellAndWaitResult_e.UserBreak
'            Application.EnableCancelKey = SaveCancelKey
'        Else
'            Err.Clear
'            Resume Next
'        End If
'    Else
'        CloseHandle ProcHandle
'        Application.EnableCancelKey = SaveCancelKey
'        ShellAndWait = ShellAndWaitResult_e.Failure
'    End If
'Else
'    ' some other error. assume failure
'    CloseHandle ProcHandle
'    ShellAndWait = ShellAndWaitResult_e.Failure
'End If
'
'Application.EnableCancelKey = SaveCancelKey
'
'End Function
'
'

Public Function Oem2Ansi(ByVal Oem As String)
   Dim Ansi As String, l As Integer
   
   l = Len(Oem)
   Ansi = Space(l + 2)

   Call OemToCharBuf(Oem, Ansi, l)
   Oem2Ansi = Left(Ansi, l)

End Function
Public Function Ansi2Oem(ByVal Ansi As String)
   Dim Oem As String, l As Integer
   
   l = Len(Ansi)
   Oem = Space(l + 2)

   Call CharToOemBuf(Ansi, Oem, l)
   Ansi2Oem = Left(Oem, l)

End Function

Function ConvertFromUTF7(ByRef InString As String) As String
   Dim i As Long
   Dim l As Long
   Dim temp As String

'   If Errorhandling Then On Error Resume Next

   l = Len(InString)
   temp = String(l * 2, 0)
   i = MultiByteToWideChar(CP_UTF7, 0, InString & Chr(0), -1, temp, l)
   If i > 0 Then
      ConvertFromUTF7 = StrConv(Left(temp, (i - 1) * 2), vbFromUnicode)
      i = InStr(ConvertFromUTF7, Chr(0))
      If i Then
         ConvertFromUTF7 = Left(ConvertFromUTF7, i - 1)
      End If
   Else
      ConvertFromUTF7 = InString
   End If
   
End Function

Function ConvertToUTF7(ByRef InString As String) As String
   Dim i As Long
   Dim l As Long
   Dim temp As String
   Dim temp2 As String
   
   If InString = "" Then
      Exit Function
   End If
   
'   If Errorhandling Then On Error Resume Next
   
   temp = StrConv(InString, vbUnicode)
   l = Len(temp) * 4
   temp2 = String(l * 4, 0)
   temp = temp & Chr(0)
   i = WideCharToMultiByte1(CP_UTF7, 0, temp, -1, temp2, l, ByVal 0, 0)
   If i > 0 Then
      ConvertToUTF7 = temp2
      i = InStr(ConvertToUTF7, Chr(0))
      If i Then
         ConvertToUTF7 = Left(ConvertToUTF7, i - 1)
      End If
   Else
      ConvertToUTF7 = InString
   End If
   
End Function

Function Utf8Ansi(ByVal sIn As String)
   Dim i As Long, x As Integer, Y As Integer
   Dim o As String, c As String * 1
   Dim l As Long
   Dim a As Integer

   i = 1
   o = ""
   l = Len(sIn)
   Do While i <= l
      c = Mid(sIn, i, 1)
      x = Asc(c)
      
      If i < l Then
         Y = Asc(Mid(sIn, i + 1, 1))
      Else
         Y = 0
      End If
      
      If (x And &HC0) = &HC0 And (Y And &HC0) = &H80 Then ' bits  11000000  y  10000000
         a = ((x And &H3F) * 64) Or (Y And &H3F)
         ' 7 feb 2020: se validan que sean caracteres ANSI
         If (a < 128 Or a > 255) Or a = 129 Or a = 141 Or a = 143 Or a = 144 Or a = 157 Then  ' no es válido en ANSI
            o = o & "_"
         Else
            o = o & Chr(a)       ' << 6 => *64  = 2^6
         End If
         i = i + 2
      Else
         o = o & c
         i = i + 1
      End If
      
   Loop

   Utf8Ansi = o

End Function


Function Ansi2UTF8_2(ByVal InString As String) As String
   Dim l As Integer, c As Integer, i As Integer, Out As String, x As Integer, n As Integer
   
   l = Len(InString)
   Out = Space(l * 2)
   n = 0
   
   For i = 1 To l
      c = Asc(Mid(InString, i, 1))
      If c < 128 Then
         n = n + 1
         Mid(Out, n, 1) = Mid(InString, i, 1)
      Else
         x = &HC0 Or Int(c / (2 ^ 6))
         n = n + 1
         Mid(Out, n, 1) = Chr(x)
         x = &H80 Or (c And &H3F)
         n = n + 1
         Mid(Out, n, 1) = Chr(x)
      End If
   Next i
   
   Ansi2UTF8_2 = Left(Out, n)
   
'char* iso_latin_1_to_utf8(char* buffer, char* end, unsigned char c) {
'    if (c < 128) {
'        if (buffer == end) { throw std::runtime_error("out of space"); }
'        *buffer++ = c;
'    }
'    else {
'        if (end - buffer < 2) { throw std::runtime_error("out of space"); }
'        *buffer++ = 0xC0 | (c >> 6);
'        *buffer++ = 0x80 | (c & 0x3f);
'    }
'    return buffer;
'}

End Function

' *** Ojo: en algunos Windows no funciona, mejor usar Ansi2Utf8_2
Function Ansi2Utf8(ByVal InString As String) As String
   Dim i As Long
   Dim l As Long
   Dim temp As String
   Dim temp2 As String
   
   If InString = "" Then
      Exit Function
   End If
   
'   If Errorhandling Then On Error Resume Next
   
   temp = StrConv(InString, vbUnicode)
   l = Len(temp) * 4
   temp2 = String(l * 4, 0)
   temp = temp & Chr(0)
   
'   l = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(temp), -1, vbNull, 0&, 0&, 0&) ' solo obtiene el largo
   
   i = WideCharToMultiByte(CP_UTF8, 0, temp, -1, temp2, l, ByVal 0, 0)
   If i > 0 Then
      Ansi2Utf8 = temp2
      i = InStr(Ansi2Utf8, Chr(0))
      If i Then
         Ansi2Utf8 = Left(Ansi2Utf8, i - 1)
      End If
   Else
      Ansi2Utf8 = InString
   End If
   
End Function

Function Utf82Ansi(ByVal InString As String) As String
   Dim i As Long
   Dim l As Long
   Dim temp As String

'   If Errorhandling Then On Error Resume Next

   l = Len(InString)
   temp = String(l * 2, 0)
   i = MultiByteToWideChar(CP_UTF8, 0, InString & Chr(0), -1, temp, l)
   If i > 0 Then
      Utf82Ansi = StrConv(Left(temp, (i - 1) * 2), vbFromUnicode)
      i = InStr(Utf82Ansi, Chr(0))
      If i Then
         Utf82Ansi = Left(Utf82Ansi, i - 1)
      End If
   Else
      Utf82Ansi = InString
   End If
   
End Function

Public Function writeStdout(ByVal Msg As String) As Long
   Dim lRes As Long, bRc As Boolean

   lRes = -1
   bRc = WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), Msg, Len(Msg), lRes, ByVal 0&)
   If bRc Then
      writeStdout = lRes
   Else
      lRes = ERR.LastDllError()
   
      MsgBox GetLastSystemError(lRes)
      
   
      writeStdout = -lRes
   End If
End Function

Public Function GetPID()
'   Dim Pid As Long
'   Call GetWindowThreadProcessId(hWnd, Pid)
   
   GetPID = GetCurrentProcessId()
End Function
' 8 ene 2018: se agrega función, Convierte algo como  z:\hola  a \\server\serv\hola
Public Function GetFullPath(ByVal Path As String, FullPath As String) As Long
   Dim Buf As String, szBuf As Long, Rc As Long, i As Integer, Out As String
   
   szBuf = 256
   Buf = Space(szBuf)
   
   On Error Resume Next
   
   Rc = GetUniversalPath(Path, UNIVERSAL_NAME_INFO_LEVEL, Buf, szBuf)
'   Rc = GetUniversalPath(Path, REMOTE_NAME_INFO_LEVEL, Buf, szBuf)

   If ERR.Number And Rc = 0 Then
      Rc = ERR.Number
   End If
   
   If Rc Then
      FullPath = ""
      GetFullPath = Rc
      Exit Function
   End If
   
   Rc = InStr(6, Buf, Chr(0), vbBinaryCompare)
   
   If Rc > 8 Then
      FullPath = Mid(Buf, 5, Rc - 5)
   Else
      FullPath = ""
      GetFullPath = -1
      Exit Function
   End If
   
End Function
