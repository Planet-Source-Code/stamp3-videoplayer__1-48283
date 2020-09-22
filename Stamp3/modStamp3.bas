Attribute VB_Name = "modStamp3"
Option Explicit

Public Sub nPlaylist_QueryFile(MultiSelect As Boolean)
    Dim Filelist() As String
    Dim Files As Long
    Dim a As Long
    
    'asks for files
    If MultiSelect = True Then
        'ask for one or more files
        Files = ShowMultipleOpen(Filter, frmMain.hwnd, "Stamp3 - Add files to playlist", Filelist)
        'if cancel exit
        If Files = 0 Then Exit Sub
        'add to playlistlist
        For a = 0 To Files - 1
            nPlaylist_Add Filelist(a), True
        Next
    Else
        ReDim Filelist(0)
        Files = 1
        'ask for one file only
        Filelist(0) = ShowOpen(Filter, frmMain.hwnd, "Stamp3 - Open file")
        'if cancel exit
        If Filelist(0) = vbNullString Then Exit Sub
        'clear playlist
        nPlaylist_Clear
        'add to playlistlist
        nPlaylist_Add Filelist(0), True
    End If
    
End Sub

Public Sub nPlaylist_Remove()
    Dim Index As Integer
    Dim a As Integer, b As Integer
    
    'if no file is selected, beep and exit
    If frmMain.Playlist.ListIndex = -1 Then
        Beep
        Exit Sub
    End If
    
    'set index
    Index = frmMain.Playlist.ListIndex + 1
    'remove from list
    frmMain.Playlist.RemoveItem Index - 1
    
    'if only one file clear the playlist
    If Stamp.pItems = 1 Then
        nPlaylist_Clear
        Exit Sub
    End If
    
    'if index is not the last move files one step up
    If Index <> Stamp.pItems Then
        For a = Index To Stamp.pItems - 1
            Stamp.pList(a).File = Stamp.pList(a + 1).File
            Stamp.pList(a).Title = Stamp.pList(a + 1).Title
        Next
    End If
    
    'redim playlist
    Stamp.pItems = Stamp.pItems - 1
    ReDim Preserve Stamp.pList(1 To Stamp.pItems)
    
    If Index = Stamp.pCurrent Then
        '// MAKE SHIT //
        nPlaylist_Close
    End If
    
End Sub

Public Sub nPlaylist_Add(File As String, Autoplay As Boolean)
    'add file to list
    Stamp.pItems = Stamp.pItems + 1
    ReDim Preserve Stamp.pList(1 To Stamp.pItems)
    Stamp.pList(Stamp.pItems).File = File
    Stamp.pList(Stamp.pItems).Title = GUI_MakeTitleString(File)
    
    'add to gui
    frmMain.Playlist.AddItem Stamp.pList(Stamp.pItems).Title
    
    'play if emnty list and autoplay is enabled
    If Stamp.pItems = 1 Then
        If Autoplay = True Then
            nPlaylist_Play 1, True
        Else
            Stamp.pCurrent = 1
        End If
    End If
    nPlaylist_SelectCurrent
End Sub

Public Sub nPlaylist_Play(Index As Integer, SetDisplay As Boolean)
    'if playlist is emnty, beep and exit
    If Stamp.pItems = 0 Then
        Beep
        Exit Sub
    End If
    
    'destroy objects
    nPlaylist_Destroy
    
    'set current file to new and assign filename
    Stamp.pCurrent = Index
    Stamp.mFile = Stamp.pList(Index).File
    
    'prepare play
    On Error GoTo ErrOut:
    Set Stamp.oMediaControl = New FilgraphManager
    Set Stamp.oMediaPosition = Stamp.oMediaControl
    Set Stamp.oMediaEvent = Stamp.oMediaControl
    Set Stamp.oVideo = Stamp.oMediaControl
    Set Stamp.oAudio = Stamp.oMediaControl
    Set Stamp.oVideoWindow = Stamp.oMediaControl
    Stamp.oMediaControl.RenderFile Stamp.mFile
    Stamp.oVideoWindow.WindowStyle = &H6000000
    Stamp.oVideoWindow.Owner = frmMain.VideoDisplay.hwnd
    
    'update volume
    nMedia_SetVolume
    
    'populate main struct
    Stamp.mLength = Int(Stamp.oMediaPosition.Duration * 1000)
    Stamp.mTitle = Stamp.pList(Index).Title
    Stamp.vHeight = Stamp.oVideo.SourceHeight
    Stamp.vWidth = Stamp.oVideo.SourceWidth
    Stamp.vAspectRatio = Stamp.vWidth / Stamp.vHeight
    Stamp.mState = StatePlaying
    Stamp.mPosition = 0
    Stamp.mStatus = GUI_FormatTime(0) & "/" & GUI_FormatTime(Stamp.mLength)
    
    'update gui
    frmMain.PicLogo.Visible = False
    frmMain.VideoDisplay.Backcolor = vbBlack
    nPlaylist_SelectCurrent
    If frmMain.WindowState <> vbMaximized Then Stamp.PlaylistVisible = False
    GUI_SetInfo
    'do not resize window to fit stamp.oVideo if in fullscreen
    If frmMain.WindowState <> vbMaximized And SetDisplay = True Then
        'fit window to stamp.oVideo
        frmMain.SetVideoDisplay
    Else
        'update window
        frmMain.ResizeWindow
    End If
    
    'run playback
    Stamp.oMediaControl.Run
    
    If frmMain.WindowState = vbMaximized Then
        frmMain.picFullscreenTitle.Cls
        frmMain.picFullscreenTitle.Visible = True
        frmMain.picFullscreenTitle.Height = frmMain.picFullscreenTitle.TextHeight(Stamp.mTitle)
        frmMain.picFullscreenTitle.Print Stamp.mTitle
        Stamp.cTDT = 30
    End If
    
    Exit Sub
    
ErrOut:
    MsgBox "Could not open file for playback. File might be damaged or unsupported!", vbCritical
    If Stamp.pItems = 1 And Stamp.Config.PlaybackMode = ModeRepeat Or Stamp.Config.PlaybackMode = ModeLoop Then nPlaylist_Close Else nPlaylist_Next
End Sub

Public Sub nPlaylist_Next()
    Select Case Stamp.Config.PlaybackMode
    Case ModeLoop
        'plays the same file
        nPlaylist_Play Stamp.pCurrent, True
    Case ModeRandom
        'plays a random file
        nPlaylist_Play Int(Rnd * Stamp.pItems) + 1, True
    Case Else
        If Stamp.Config.PlaybackMode = ModeRepeat And Stamp.pCurrent = Stamp.pItems Then
            'if eop then play 1. file
            nPlaylist_Play 1, True
        ElseIf Stamp.Config.PlaybackMode <> ModeRepeat And Stamp.pCurrent = Stamp.pItems Then
            'if eop and no repeat
            If Stamp.mState = StatePlaying And Stamp.mPosition = Stamp.mLength And Stamp.pCurrent = Stamp.pItems Then
                'stop if last file in playlist
                nPlaylist_Close
            Else
                'beep if otherwise
                Beep
            End If
        Else
            'if not eop play next file in playlist
            nPlaylist_Play Stamp.pCurrent + 1, True
        End If
    End Select
End Sub

Public Sub nPlaylist_Previous()
    Select Case Stamp.Config.PlaybackMode
    Case ModeLoop
        'plays the same file
        nPlaylist_Play Stamp.pCurrent, True
        Exit Sub
    Case ModeRandom
        'plays a random file
        nPlaylist_Play Int(Rnd * Stamp.pItems) + 1, True
        Exit Sub
    Case Else
        If Stamp.Config.PlaybackMode = ModeRepeat And Stamp.pCurrent = 1 Then
            'if 1. file then play last file
            nPlaylist_Play Stamp.pItems, True
        ElseIf Stamp.Config.PlaybackMode <> ModeRepeat And Stamp.pCurrent = 1 Then
            'if eop and no repeat
            Beep
        Else
            'else play previous file
            nPlaylist_Play Stamp.pCurrent - 1, True
        End If
    End Select
End Sub

Public Sub nPlaylist_Stop()
    'if already stopped, beep and exit
    If Stamp.mState = StateStopped Then
        Beep
        Exit Sub
    End If
    
    'destroys the objects
    nPlaylist_Destroy
    
    'updates the struct
    Stamp.mState = StateStopped
    Stamp.mPosition = 0
    Stamp.mStatus = GUI_FormatTime(Stamp.mPosition)
    
    'updates gui
    frmMain.PicLogo.Visible = True
    frmMain.VideoDisplay.Backcolor = Stamp.Skin.VideoDisplayColor
    GUI_SetInfo
    GUI_SetProgressPos
End Sub

Public Sub nPlaylist_Close()
    nPlaylist_Destroy
    nResetCurrentFileData
    frmMain.PicLogo.Visible = True
    frmMain.VideoDisplay.Backcolor = Stamp.Skin.VideoDisplayColor
    GUI_SetInfo
    GUI_SetProgressPos
End Sub

Public Sub nPlaylist_Destroy()
    Set Stamp.oMediaControl = Nothing
    Set Stamp.oMediaPosition = Nothing
    Set Stamp.oMediaEvent = Nothing
    Set Stamp.oVideo = Nothing
    Set Stamp.oVideoWindow = Nothing
    Set Stamp.oAudio = Nothing
End Sub

Public Sub nPlaylist_SelectCurrent()
    If Stamp.pItems <> 0 Then frmMain.Playlist.ListIndex = Stamp.pCurrent - 1
End Sub

Public Sub nPlaylist_Pause()
    If Stamp.mState = StatePaused Then
        'if paused then resume
        Stamp.oMediaControl.Run
        Stamp.mState = StatePlaying
    ElseIf Stamp.mState = StatePlaying Then
        'if playing then pause
        Stamp.oMediaControl.Pause
        Stamp.mState = StatePaused
        'tells the timer to "blink" the time
        Stamp.cCDT = 0
        Stamp.cPCDT = 0
    Else
        'if no file or stopped, beep
        Beep
    End If
End Sub

Public Sub nPlaylist_Clear()
    'erases playlist
    Stamp.pCurrent = 0
    Stamp.pItems = 0
    Erase Stamp.pList
    frmMain.Playlist.Clear
    
    'destroys objects
    nPlaylist_Destroy
    'resets struct
    nResetCurrentFileData
    
    'updates gui
    frmMain.PicLogo.Visible = True
    frmMain.VideoDisplay.Backcolor = Stamp.Skin.VideoDisplayColor
    GUI_SetInfo
    GUI_SetProgressPos
    
End Sub

Public Sub nResetCurrentFileData()
    'resets data required for playback
    Stamp.mFile = vbNullString
    Stamp.mLength = 0
    Stamp.mPosition = 0
    Stamp.mState = StateStopped
    Stamp.mTitle = "Stamp3"
    Stamp.mStatus = vbNullString
End Sub

Public Sub nMedia_Seek(Time As Long)
    If Stamp.mState = StateStopped Then
        'if no stamp.oVideo playing beep and exit
        Beep
    Else
        'sets time
        Stamp.mPosition = Time
        GUI_SetTempStatus "Seeking to " & GUI_FormatTime(Time)
        'seeks
        Stamp.oMediaPosition.CurrentPosition = Time / 1000
    End If
    GUI_SetInfo
    GUI_SetProgressPos
End Sub

Public Function nMedia_GetPosition() As Long
    'gets current position in seconds and return in ms
    nMedia_GetPosition = Int(Stamp.oMediaPosition.CurrentPosition * 1000)
End Function

Public Sub nMedia_SetVolume()
    'makes sure its valid
    If Stamp.Config.Volume > VolumeMax Then
        Stamp.Config.Volume = VolumeMax
    ElseIf Stamp.Config.Volume < 0 Then
        Stamp.Config.Volume = 0
    End If
    
    'if no stamp.oAudio object then stamp.oVideo is without sound
    On Local Error Resume Next
    Stamp.oAudio.Volume = -VolumeMax + (Stamp.Config.Volume / 10000) ^ (1 / 3) * 10000
End Sub

Public Sub nPlaylist_Save(Optional FileName As String)
    Dim FileHandle As Integer
    Dim a As Integer
    
    'if no filename specified then ask
    If FileName = vbNullString Then FileName = ShowSave("List files (*.stn)" & vbNullChar & "*.stn" & vbNullChar, frmMain.hwnd, "Stamp3 - Save playlist")
    'if cancel exit
    If FileName = vbNullString Then Exit Sub
    FileHandle = FreeFile
    Open FileName For Output As FileHandle
    For a = 1 To Stamp.pItems
        Print #1, Stamp.pList(a).File
    Next
    Close FileHandle
End Sub

Public Sub nPlaylist_Load(Add As Boolean, Autoplay As Boolean, Optional FileName As String)
    Dim FileHandle As Integer
    Dim StrData As String
    
    'if no filename specified then ask
    If FileName = vbNullString Then FileName = ShowOpen("List files (*.stn)" & vbNullChar & "*.stn" & vbNullChar, frmMain.hwnd, "Stamp3 - Load playlist")
    'if user cancel exit
    If FileName = vbNullString Then Exit Sub
    
    FileHandle = FreeFile
    Open FileName For Input As FileHandle
    'clear list if not adding
    If Add = False Then nPlaylist_Clear
    Do Until EOF(FileHandle)
        Line Input #1, StrData
        'adds to playlist
        nPlaylist_Add StrData, Autoplay
    Loop
    Close FileHandle
End Sub

Public Sub nPlaylist_Toggle()
    'toggles visible state of playlist
    If frmMain.Playlist.Visible = True Then
        nPlaylist_Show False
    Else
        nPlaylist_Show True
    End If
End Sub

Public Sub nPlaylist_Show(Visible As Boolean)
    'sets state
    Stamp.PlaylistVisible = Visible
    'update gui
    frmRes.mnuShowPlaylist.Checked = Visible
    frmMain.ResizeWindow
    nPlaylist_SelectCurrent
End Sub

Public Sub nPlaylist_AddDir()
    Dim Dir As String
    Dir = BrowseForFolder(frmMain.hwnd)
    If Dir = vbNullString Then Exit Sub
    frmRes.FDB1.Search Dir, "*.*"
End Sub

Public Sub nParseCommand(Command As String, AddList As Boolean)
    Dim tmp As Variant, prm As String
    Dim FileName As String
    
    'if no command then skip
    If Command <> vbNullString Then
        tmp = InStr(1, Command, " ")
        If tmp = 0 Then tmp = Command Else tmp = LCase(Left(Command, tmp - 1))
        Select Case LCase(tmp)
        Case "/refreshskin"
            Skin_Extract
            Skin_Load
        Case "/testskin"
            Stamp.Config.CurrentSkinFile = Mid(Command, 10)
            Skin_Extract
            Skin_Load
        Case "/fullscreen"
            GUI_SetDisplayMode True
            prm = Mid(Command, Len(tmp) + 2)
            If Left(prm, 1) <> """" Then
                'if not use of quotes get full name
                FileName = GetFullFileName(prm)
            Else
                FileName = Mid(prm, 2, Len(prm) - 2)
            End If
            If Right(FileName, 4) = ".stn" Then
                nPlaylist_Load AddList, True, FileName
            Else
                nPlaylist_Add FileName, True
            End If
        Case Else
            If Left(Command, 1) <> """" Then
                'if not use of quotes get full name
                FileName = GetFullFileName(Command)
            Else
                FileName = Mid(Command, 2, Len(Command) - 2)
            End If
            If Right(FileName, 4) = ".stn" Then
                nPlaylist_Load AddList, True, FileName
            Else
                nPlaylist_Add FileName, True
            End If
        End Select
    End If
    
    'show window
    'SetWindowPos frmMain.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    frmMain.Show
    
End Sub

Public Sub nPlayback_ToggleMode(NewMode As Byte)
    If Stamp.Config.PlaybackMode = NewMode Then nPlayback_SetMode ModeNone Else nPlayback_SetMode NewMode
End Sub

Public Sub nPlayback_SetMode(Mode As Byte)
    Stamp.Config.PlaybackMode = Mode
    Select Case Mode
    Case ModeNone
        frmRes.mnuLoop.Checked = False
        frmRes.mnuPlaybackRandom.Checked = False
        frmRes.mnuPlaybackRepeat.Checked = False
    Case ModeRandom
        frmRes.mnuLoop.Checked = False
        frmRes.mnuPlaybackRandom.Checked = True
        frmRes.mnuPlaybackRepeat.Checked = False
    Case ModeRepeat
        frmRes.mnuLoop.Checked = False
        frmRes.mnuPlaybackRandom.Checked = False
        frmRes.mnuPlaybackRepeat.Checked = True
    Case ModeLoop
        frmRes.mnuLoop.Checked = True
        frmRes.mnuPlaybackRandom.Checked = False
        frmRes.mnuPlaybackRepeat.Checked = False
    End Select
End Sub

