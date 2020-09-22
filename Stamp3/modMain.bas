Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    If App.PrevInstance = True Then
        Dim whandle As Long
        whandle = FindWindowEx(0, 0, "ThunderRT6FormDC", "stamp3 res hwnd")
        whandle = FindWindowEx(whandle, 0, "ThunderRT6TextBox", vbNullString)
        SendMessage whandle, WM_SETTEXT, ByVal 0, ByVal Command$
        End
    End If
    
    Load frmRes
    
    Load frmMain
    LoadStamp3
    frmMain.Show
    
    If Command$ <> vbNullString Then
        nParseCommand Command, False
    Else
        If Dir(Stamp.AppPath & "current.stn") <> vbNullString Then nPlaylist_Load False, False, Stamp.AppPath & "current.stn"
    End If
End Sub

Public Sub LoadStamp3()
    Skin_SetUpNames
    Stamp.AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    Stamp.vHwnd = frmMain.VideoDisplay.hwnd
    frmMain.picFullscreenTitle.Visible = False
    nResetCurrentFileData
    Stamp.vHeight = 250
    Stamp.vWidth = 250
    LoadConfig
    frmRes.mnuPriorityNormal.Checked = True
    Timer_Set
End Sub

Public Sub UnloadStamp3()
    frmMain.Hide
    SaveConfig
    nPlaylist_Save Stamp.AppPath & "current.stn"
    nPlaylist_Close
    ResetScreenSaverActive
    Timer_Kill
    Reset
    GUI_ShowCursor True
    End
End Sub

Public Sub SaveConfig()
    Dim FileHandle As Integer
    FileHandle = FreeFile
    If Dir(Stamp.AppPath & "config.bin") <> vbNullString Then Kill Stamp.AppPath & "config.bin"
    Open Stamp.AppPath & "config.bin" For Binary As FileHandle
    Put #1, 1, Stamp.Config
    Close FileHandle
End Sub

Public Sub ResetConfig()
    With Stamp.Config
        .Signature = "Stamp3"
        .Volume = VolumeMax
        .CurrentSkinFile = "#"
        .PlaylistWidth = 4000
        .AlwaysOnTop = False
        .PreventScreenSaver = True
        .PlaybackMode = ModeNone
        .Zoom = 1
        .VideoZoomOn = False
    End With
    ApplyConfig
    SaveConfig
End Sub

Public Sub LoadConfig()
    Dim FileHandle As Integer
    FileHandle = FreeFile
    On Error GoTo ErrOut
    Open Stamp.AppPath & "config.bin" For Binary As FileHandle
    Get #1, 1, Stamp.Config
    Close FileHandle
    If Stamp.Config.Signature <> "Stamp3" Then GoTo ErrOut
    
    ApplyConfig
    
    Exit Sub
    
ErrOut:
    Close FileHandle
    ResetConfig
End Sub

Public Sub ApplyConfig()
    Skin_Load
    BackupScreenSaverActive
    SetScreenSaverPrevent
    SetAlwaysOnTop
    nPlayback_SetMode Stamp.Config.PlaybackMode
    frmMain.SetVideoDisplay
End Sub

Public Function GetFullFileName(FileName As String) As String
    Dim ret As Long
    Dim buffer As String
    Dim tmp As String
    buffer = String(4096, 0)
    ret = GetFullPathName(FileName, 4096, buffer, tmp)
    GetFullFileName = Left(buffer, ret)
End Function

