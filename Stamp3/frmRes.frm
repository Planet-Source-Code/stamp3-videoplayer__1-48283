VERSION 5.00
Begin VB.Form frmRes 
   Caption         =   "stamp3 res hwnd"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin Stamp3.FDB FDB1 
      Left            =   960
      Top             =   720
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About ""Stamp3""..."
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open...	O"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add...	+"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save..."
      End
      Begin VB.Menu Sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSkinBrowser 
         Caption         =   "Skin browser..."
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Begin VB.Menu nuPriority 
            Caption         =   "Priority"
            Begin VB.Menu mnuPriorityRT 
               Caption         =   "Realtime"
            End
            Begin VB.Menu mnuPriorityHigh 
               Caption         =   "High"
            End
            Begin VB.Menu mnuPriorityNormal 
               Caption         =   "Normal"
            End
            Begin VB.Menu mnuPriorityIdle 
               Caption         =   "Idle"
            End
         End
         Begin VB.Menu mnuFA 
            Caption         =   "File Associations..."
         End
         Begin VB.Menu mnuScreensaverToggle 
            Caption         =   "Prevent screensaver	T"
         End
         Begin VB.Menu mnuAlwaysOnTop 
            Caption         =   "Always on top	A"
         End
         Begin VB.Menu mnuResetconfig 
            Caption         =   "Reset Config"
         End
      End
      Begin VB.Menu Sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "Playback"
         Begin VB.Menu mnuPlaybackPlay 
            Caption         =   "Play	X"
         End
         Begin VB.Menu mnuPlaybackStop 
            Caption         =   "Stop	C"
         End
         Begin VB.Menu mnuPlaybackPause 
            Caption         =   "Pause	V"
         End
         Begin VB.Menu mnuPlaybackPrevious 
            Caption         =   "Previous	Z"
         End
         Begin VB.Menu mnuPlaybackNext 
            Caption         =   "Next	B"
         End
         Begin VB.Menu Sep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlaybackRepeat 
            Caption         =   "Repeat	R"
         End
         Begin VB.Menu mnuPlaybackRandom 
            Caption         =   "Random	E"
         End
         Begin VB.Menu mnuLoop 
            Caption         =   "Loop	L"
         End
      End
      Begin VB.Menu mnuPlaylist 
         Caption         =   "Playlist"
         Begin VB.Menu mnuPlaylistShow 
            Caption         =   "Playlist..."
         End
         Begin VB.Menu Sep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlaylistAdd 
            Caption         =   "Add...	+"
         End
         Begin VB.Menu mnuPlaylistRemove 
            Caption         =   "Remove	-"
         End
         Begin VB.Menu mnuPlaylistClear 
            Caption         =   "Clear"
         End
         Begin VB.Menu Sep7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlaylistLoad 
            Caption         =   "Load..."
         End
         Begin VB.Menu mnuPlaylistSave 
            Caption         =   "Save..."
         End
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuVideoZoom 
            Caption         =   "Video Zoom"
            Begin VB.Menu mnuVideoZoomOn 
               Caption         =   "Enabled	U"
            End
            Begin VB.Menu mnuVideoZoomReset 
               Caption         =   "Reset"
            End
         End
         Begin VB.Menu Sep11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuZoomFullscreen 
            Caption         =   "Fullscreen	5"
         End
         Begin VB.Menu Sep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuZoom400 
            Caption         =   "400%	4"
         End
         Begin VB.Menu mnuZoom200 
            Caption         =   "200%	3"
         End
         Begin VB.Menu mnuZoom100 
            Caption         =   "100%	2"
         End
         Begin VB.Menu mnuZoom50 
            Caption         =   "50%	1"
         End
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullscreen 
         Caption         =   "Fullscreen	CTRL+ENTER"
      End
      Begin VB.Menu mnuControls 
         Caption         =   "Controls...	W"
      End
      Begin VB.Menu mnuShowPlaylist 
         Caption         =   "Playlist...	P"
      End
      Begin VB.Menu Sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FDB1_FileFound(FileName As String, Path As String, MaxPath As Long, HiSize As Long, LoSize As Long)
    Dim tmp As Variant
    tmp = InStrRev(FileName, ".")
    If tmp = 0 Then Exit Sub
    tmp = LCase(Mid(FileName, tmp + 1))
    If tmp = "stn" Then
        nPlaylist_Load Path & FileName, True
    ElseIf InStr(1, FilterExt, " " & tmp & " ") <> 0 Then
        nPlaylist_Add Path & FileName, False 'Playlist_Add Path & FileName
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
    frmMain.Enabled = False
End Sub

Private Sub mnuAdd_Click()
    'Playlist_QueryFile
    nPlaylist_QueryFile True
End Sub

Private Sub mnuAlwaysOnTop_Click()
    GUI_ToggleAlwaysOnTop
End Sub

Private Sub mnuControls_Click()
    frmMain.ToggleControls
End Sub

Private Sub mnuExit_Click()
    UnloadStamp3
End Sub

Private Sub mnuFA_Click()
    frmFileAss.Show
End Sub

Private Sub mnuFullscreen_Click()
    GUI_ToggleScreenMode
End Sub

Private Sub mnuLoad_Click()
    nPlaylist_Load False, True
End Sub

Private Sub mnuLoop_Click()
    nPlayback_ToggleMode ModeLoop
End Sub

Private Sub mnuOpen_Click()
    nPlaylist_QueryFile False
End Sub

Private Sub mnuPlaybackNext_Click()
    nPlaylist_Next
    'Playlist_Play stamp.pcurrent + 1
End Sub

Private Sub mnuPlaybackPause_Click()
    nPlaylist_Pause
    'Media_Pause
End Sub

Private Sub mnuPlaybackPlay_Click()
    'Media_Play False
    nPlaylist_Play Stamp.pCurrent, False
End Sub

Private Sub mnuPlaybackPrevious_Click()
    nPlaylist_Previous
    'Playlist_Play stamp.pcurrent - 1
End Sub

Private Sub mnuPlaybackRandom_Click()
    nPlayback_ToggleMode ModeRandom
End Sub

Private Sub mnuPlaybackRepeat_Click()
    nPlayback_ToggleMode ModeRepeat
End Sub

Private Sub mnuPlaybackStop_Click()
    'Media_Stop
    nPlaylist_Stop
End Sub

Private Sub mnuPlaylistAdd_Click()
    nPlaylist_QueryFile True
    'Playlist_QueryFile
End Sub

Private Sub mnuPlaylistClear_Click()
    nPlaylist_Clear
    'Playlist_Clear
End Sub

Private Sub mnuPlaylistLoad_Click()
    nPlaylist_Load False, True
End Sub

Private Sub mnuPlaylistRemove_Click()
    nPlaylist_Remove
    'If frmMain.Playlist.ListIndex <> -1 Then Playlist_Remove
End Sub

Private Sub mnuPlaylistSave_Click()
    nPlaylist_Save
End Sub

Private Sub mnuPlaylistShow_Click()
     nPlaylist_Show True
End Sub

Private Sub mnuPriorityHigh_Click()
    TogglePriority HIGH_PRIORITY_CLASS
End Sub

Private Sub mnuPriorityIdle_Click()
    TogglePriority IDLE_PRIORITY_CLASS
End Sub

Private Sub mnuPriorityNormal_Click()
    TogglePriority NORMAL_PRIORITY_CLASS
End Sub

Private Sub mnuPriorityRT_Click()
    TogglePriority REALTIME_PRIORITY_CLASS
End Sub

Private Sub mnuResetconfig_Click()
    ResetConfig
End Sub

Private Sub mnuSave_Click()
    nPlaylist_Save
End Sub

Private Sub mnuScreensaverToggle_Click()
    ToggleScreenSaverActive
End Sub

Private Sub mnuShowPlaylist_Click()
    nPlaylist_Toggle
End Sub

Private Sub mnuSkinBrowser_Click()
    GUI_ShowSkinBrowser
End Sub

Private Sub mnuVideoZoomOn_Click()
    GUI_ToggleVideoZoom
End Sub

Private Sub mnuVideoZoomReset_Click()
    GUI_ResetVideoZoom
End Sub

Private Sub mnuZoom100_Click()
    frmMain.SetVideoDisplay
End Sub

Private Sub mnuZoom200_Click()
    frmMain.SetVideoDisplay 2
End Sub

Private Sub mnuZoom400_Click()
    frmMain.SetVideoDisplay 4
End Sub

Private Sub mnuZoom50_Click()
    frmMain.SetVideoDisplay 0.5
End Sub

Private Sub mnuZoomFullscreen_Click()
    GUI_SetDisplayMode True
End Sub

Private Sub txtLink_Change()
    nParseCommand txtLink, True
End Sub
