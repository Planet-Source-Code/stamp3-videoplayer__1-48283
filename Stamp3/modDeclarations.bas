Attribute VB_Name = "modDeclarations"
Option Explicit

'constants
Public Const VolumeMax = 10000
Public Const StateClosed = 0
Public Const StatePlaying = 1
Public Const StateStopped = 2
Public Const StatePaused = 3
Public Const TimerInterval = 50
Public Const Filter = "Video Files" & vbNullChar & "*.avi;*.mpg;*.mpeg;*.vob;*.divx;*.asf;*.m2v;*.wmv"
Public Const FilterExt = " avi mpg mpeg vob divx asf m2v wmv stn "

Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SPI_SETPOWEROFFACTIVE = 86
Public Const SPI_GETPOWEROFFACTIVE = 84

Public Const WM_SETTEXT = &HC

Public Const ModeRandom = 0
Public Const ModeLoop = 1
Public Const ModeRepeat = 2
Public Const ModeNone = 3

'types
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type MOUSESTAT
    X As Long
    Y As Long
    Button As Long
End Type

Public Type Stamp3ConfigData
    Signature As String * 6
    Volume As Long
    CurrentSkinFile As String
    PlaylistWidth As Integer
    AlwaysOnTop As Boolean
    PreventScreenSaver As Boolean
    PlaybackMode As Byte
    Zoom As Single
    VideoZoomOn As Boolean
End Type


Public Type Stamp3Playlist
    File As String
    Title As String
End Type

Public Type Stamp3Data
    'app spec
    AppPath As String
    
    'skin
    Skin As Stamp3SkinData
    
    'gui
    ControlsVisible As Boolean
    PlaylistVisible As Boolean
    gLastTempState As String
    
    'options
    ScreenSaverState As Boolean
    PowerOffState As Boolean
    
    'required for playback
    mTitle As String
    mStatus As String
    mLength As Long
    mFile As String
    mPosition As Long
    mState As Byte
    
    'video spec
    vHwnd As Long
    vWidth As Long
    vHeight As Long
    vAspectRatio As Double
    
    'config
    Config As Stamp3ConfigData
    
    'playlist
    pList() As Stamp3Playlist
    pItems As Integer
    pCurrent As Integer
    
    'playback objects
    oMediaControl As IMediaControl
    oMediaPosition As IMediaPosition
    oMediaEvent As IMediaEvent
    oVideo As IBasicVideo
    oVideoWindow As IVideoWindow
    oAudio As IBasicAudio

    'counters
    cCDT As Integer
    cPCDT As Integer
    cTDT As Integer

End Type

'variables
Public Stamp As Stamp3Data
Public SkinNames() As String
Public SkinFiles() As String
Public Skins As Integer

'declarations
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Const HTBOTTOMRIGHT = 17
Public Const HTCAPTION = 2

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const MAX_PATH = 260&
Public Const REG_SZ = 1
Public Const SHCNE_ASSOCCHANGED = &H8000000
Public Const SHCNF_IDLIST = &H0&

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long

Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const IDLE_PRIORITY_CLASS = &H40
