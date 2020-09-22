VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Stamp3"
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   4590
   Begin VB.PictureBox picFullscreenTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   32
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   2640
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   30
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2640
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   29
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   28
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2640
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   26
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   25
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox conProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1320
      ScaleHeight     =   225
      ScaleWidth      =   975
      TabIndex        =   15
      Top             =   3360
      Width           =   975
      Begin VB.PictureBox PicButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox ProgressRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox ProgressLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.Image ProgressMiddle 
         Height          =   225
         Left            =   480
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox conVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   975
      TabIndex        =   11
      Top             =   3120
      Width           =   975
      Begin VB.PictureBox PicButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox VolumeRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox VolumeLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.Image VolumeMiddle 
         Height          =   255
         Left            =   480
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picPlaylistbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3000
      ScaleHeight     =   855
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   615
      Begin VB.PictureBox PicButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox PicButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox PicButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox PicButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox PicButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox PicButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.ListBox Playlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      IntegralHeight  =   0   'False
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox VideoDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
      Begin VB.Image PicLogo 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox FullscreenControls 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   2895
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblVolume2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   855
      End
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   2640
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   31
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblVolume 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lbl2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Image vbrdTop 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   255
   End
   Begin VB.Image vbrdLeft 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   600
      Stretch         =   -1  'True
      Top             =   720
      Width           =   255
   End
   Begin VB.Image vbrdBottom 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   600
      Stretch         =   -1  'True
      Top             =   960
      Width           =   255
   End
   Begin VB.Image vbrdTopRight 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   360
      Top             =   720
      Width           =   255
   End
   Begin VB.Image vbrdRight 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   360
      Stretch         =   -1  'True
      Top             =   960
      Width           =   255
   End
   Begin VB.Image vbrdBottomLeft 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   840
      Top             =   960
      Width           =   255
   End
   Begin VB.Image vbrdBottomRight 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   840
      Top             =   720
      Width           =   255
   End
   Begin VB.Image vbrdTopLeft 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      Top             =   960
      Width           =   255
   End
   Begin VB.Image Titlebar 
      Height          =   255
      Index           =   0
      Left            =   840
      Top             =   360
      Width           =   735
   End
   Begin VB.Image Titlebar 
      Height          =   255
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Image Titlebar 
      Height          =   255
      Index           =   2
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Image brdTopLeft 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1320
      Top             =   960
      Width           =   255
   End
   Begin VB.Image brdTopRight 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1800
      Top             =   720
      Width           =   255
   End
   Begin VB.Image brdRight 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   720
      Width           =   255
   End
   Begin VB.Image brdLeft 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   720
      Width           =   255
   End
   Begin VB.Image brdBottom 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   960
      Width           =   255
   End
   Begin VB.Image brdBottomLeft 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1560
      Top             =   960
      Width           =   255
   End
   Begin VB.Image brdBottomRight 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1560
      Top             =   720
      Width           =   255
   End
   Begin VB.Image brdTop 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pSliding As Boolean
Dim pStartX As Single
Dim plResize As Boolean

Dim NoResize As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 19
        GUI_ShowSkinBrowser
    Case 27
        If WindowState = vbMaximized Then
            GUI_SetDisplayMode False
        ElseIf Playlist.Visible = True Then
            nPlaylist_Toggle
        End If
    Case vbKeyU, Asc("u")
        GUI_ToggleVideoZoom
    Case vbKeySpace
        nPlaylist_Pause
    Case vbKey1
        SetVideoDisplay 0.5
    Case vbKey2
        SetVideoDisplay 1
    Case vbKey3
        SetVideoDisplay 2
    Case vbKey4
        SetVideoDisplay 4
    Case 10
        GUI_ToggleScreenMode
    Case 13
        nPlaylist_Play Stamp.pCurrent, True
    Case vbKeyR, Asc("r")
        nPlayback_ToggleMode ModeRepeat
    Case vbKeyE, Asc("e")
        nPlayback_ToggleMode ModeRandom
    Case vbKeyL, Asc("l")
        nPlayback_ToggleMode ModeLoop
    Case vbKeyP, 112
        nPlaylist_Toggle
    Case 43
        nPlaylist_QueryFile True
    Case 45
        nPlaylist_Remove
    Case vbKeyZ, 122
        nPlaylist_Previous
    Case vbKeyX, 120
        nPlaylist_Play Stamp.pCurrent, False
    Case vbKeyC, 99
        nPlaylist_Stop
    Case vbKeyV, 118
        nPlaylist_Pause
    Case vbKeyB, 98
        nPlaylist_Next
    Case vbKeyN, 110
        Stamp.Config.Volume = Stamp.Config.Volume - 1000
        nMedia_SetVolume
        GUI_SetVolumePos
    Case vbKeyM, 109
        Stamp.Config.Volume = Stamp.Config.Volume + 1000
        nMedia_SetVolume
        GUI_SetVolumePos
    Case vbKeyO, 111
        nPlaylist_QueryFile False
    Case vbKeyA, 97
        GUI_ToggleAlwaysOnTop
    Case vbKeyT, 116
        ToggleScreenSaverActive
    Case vbKey5
        GUI_SetDisplayMode True
    Case 24
        UnloadStamp3
    Case vbKeyQ, 113
        nPlaylist_AddDir
    Case vbKeyW, Asc("w")
        ToggleControls
    'Case vbKeyS, Asc("s")
    '    nMedia_Seek Stamp.mPosition + 56000
    'Case vbKey0
    '    nMedia_Seek Stamp.mPosition + 50000
    End Select
End Sub

Private Sub Form_Load()
    If Command = vbNullString Then
        Left = Screen.Width / 2 - Width / 2
        Top = Screen.Height / 2 - Height / 2
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadStamp3
End Sub

Private Sub FullscreenControls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub lbl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub lblVolume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub lbl4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub lbl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub lblVolume2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub PicButton_Click(Index As Integer)
    Select Case Index
    Case BtnExit
        UnloadStamp3
    Case BtnOpen
        nPlaylist_QueryFile False
    Case BtnPause
        nPlaylist_Pause
    Case BtnPlay
        nPlaylist_Play Stamp.pCurrent, True
    Case BtnStop
        nPlaylist_Stop
    Case BtnMinimize
        WindowState = vbMinimized
    Case BtnMenu
        PopupMenu frmRes.mnu, , PicButton(BtnMenu).Left, PicButton(BtnMenu).Top + PicButton(BtnMenu).Height
    Case BtnPlaylistAdd
        nPlaylist_QueryFile True
    Case BtnPlaylistClear
        nPlaylist_Clear
    Case BtnPlaylistRemove
        nPlaylist_Remove
    Case BtnPlaylistAddDir
         nPlaylist_AddDir
    Case BtnPlaylistSave
        nPlaylist_Save
    Case BtnPlaylistLoad
        nPlaylist_Load False, True
    End Select
End Sub

Private Sub PicButton_DblClick(Index As Integer)
    If Index = BtnMenu Then UnloadStamp3
End Sub

Private Sub PicButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Byte
    
    Select Case Index
    Case BtnResize
        ReleaseCapture
        SendMessage Me.hwnd, &HA1, HTBOTTOMRIGHT, 0
    Case BtnProgressSlider
        If Button = vbKeyLButton Then
            pSliding = True
            pStartX = X
        End If
    Case BtnVolumeSlider
        If Button = vbKeyLButton Then
            pSliding = True
            pStartX = X
        End If
    End Select
    
    For a = 0 To Buttons
        If Index <> a Then PicButton(a).Picture = Stamp.Skin.Pic_Button(a)
    Next
    PicButton(Index).Picture = Stamp.Skin.Pic_ButtonD(Index)
End Sub

Private Sub PicButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Byte
        
    Select Case Index
    Case BtnProgressSlider
        If Button = 0 Then GUI_SetTempStatus "Progress control"
        If pSliding Then
            If PicButton(BtnProgressSlider).Left + X - pStartX + PicButton(BtnProgressSlider).Width > conProgress.Width Then
                PicButton(BtnProgressSlider).Left = conProgress.Width - PicButton(BtnProgressSlider).Width
            ElseIf PicButton(BtnProgressSlider).Left + X - pStartX < 0 Then
                PicButton(BtnProgressSlider).Left = 0
            Else
                PicButton(BtnProgressSlider).Left = PicButton(BtnProgressSlider).Left + X - pStartX
            End If
            GUI_ScrollingProgress
        End If
    Case BtnVolumeSlider
        If Button = 0 Then GUI_SetTempStatus "Volume control"
        If pSliding Then
            If PicButton(BtnVolumeSlider).Left + X - pStartX + PicButton(BtnVolumeSlider).Width > conVolume.Width Then
                PicButton(BtnVolumeSlider).Left = conVolume.Width - PicButton(BtnVolumeSlider).Width
            ElseIf PicButton(BtnVolumeSlider).Left + X - pStartX < 0 Then
                PicButton(BtnVolumeSlider).Left = 0
            Else
                PicButton(BtnVolumeSlider).Left = PicButton(BtnVolumeSlider).Left + X - pStartX
            End If
            GUI_ScrollingVolume
        End If
    Case BtnExit
        GUI_SetTempStatus "Click to exit"
    Case BtnMinimize
        GUI_SetTempStatus "Click to minimize"
    Case BtnMenu
        GUI_SetTempStatus "Click to show menu"
    Case BtnPlay
        GUI_SetTempStatus "Click to start playback"
    Case BtnStop
        GUI_SetTempStatus "Click to stop playback"
    Case BtnPause
        If Stamp.mState = StatePaused Then GUI_SetTempStatus "Click to resume playback" Else GUI_SetTempStatus "Click to pause playback"
    Case BtnOpen
        GUI_SetTempStatus "Click to open file"
    Case BtnResize
        GUI_SetTempStatus "Drag to resize window"
    Case BtnPlaylistAdd
        GUI_SetTempStatus "Click to add file to playlist"
    Case BtnPlaylistAddDir
        GUI_SetTempStatus "Click to directory to playlist"
    Case BtnPlaylistRemove
        GUI_SetTempStatus "Click to remove file from playlist"
    Case BtnPlaylistClear
        GUI_SetTempStatus "Click to clear playlist"
    Case BtnPlaylistLoad
        GUI_SetTempStatus "Click to load playlist"
    Case BtnPlaylistSave
        GUI_SetTempStatus "Click to save playlist"
    End Select
    
    For a = 0 To Buttons
        If Index <> a Then PicButton(a).Picture = Stamp.Skin.Pic_Button(a)
    Next
    If Button = 0 Then PicButton(Index).Picture = Stamp.Skin.Pic_ButtonH(Index)
End Sub

Private Sub PicButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case BtnProgressSlider
        pSliding = False
        GUI_ReleasedScrollProgress
    Case BtnVolumeSlider
        pSliding = False
        GUI_ReleasedScrollVolume
    End Select
    GUI_ResetButtonGraphics
End Sub

Private Sub picPlaylistbuttons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub Playlist_DblClick()
    If Playlist.ListIndex = -1 Then Exit Sub
    nPlaylist_Play Playlist.ListIndex + 1, True
End Sub

Private Sub Playlist_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Playlist_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Playlist_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Playlist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub Playlist_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Integer
    
    For a = 1 To Data.Files.Count
        If LCase(Right(Data.Files(a), 4)) = ".stn" Then nPlaylist_Load Data.Files(a), True Else nPlaylist_Add Data.Files(a), True
    Next
    
    frmMain.SetFocus
End Sub

Private Sub ProgressLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    If Stamp.mState = StateStopped Then
        Beep
        Exit Sub
    End If
    If Button = vbKeyLButton Then
        tmp = (X - (PicButton(BtnProgressSlider).Width / 2)) / ((frmMain.conProgress.Width - frmMain.PicButton(BtnProgressSlider).Width) / Stamp.mLength)
        If tmp > Stamp.mLength Then
            tmp = Stamp.mLength
        ElseIf tmp < 0 Then
            tmp = 0
        End If
        nMedia_Seek tmp
    End If
End Sub

Private Sub ProgressLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    GUI_ResetGraphics
    If Stamp.mState = StatePlaying Or Stamp.mState = StatePaused Then
        tmp = (X - (PicButton(BtnProgressSlider).Width / 2)) / ((frmMain.conProgress.Width - frmMain.PicButton(BtnProgressSlider).Width) / Stamp.mLength)
        If tmp > Stamp.mLength Then
            tmp = Stamp.mLength
        ElseIf tmp < 0 Then
            tmp = 0
        End If
        GUI_SetTempStatus "Seek to " & GUI_FormatTime(tmp)
    End If
End Sub

Private Sub ProgressMiddle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    If Stamp.mState = StateStopped Then
        Beep
        Exit Sub
    End If
    If Button = vbKeyLButton Then
        tmp = (X + ProgressLeft.Width - (PicButton(BtnProgressSlider).Width / 2)) / ((frmMain.conProgress.Width - frmMain.PicButton(BtnProgressSlider).Width) / Stamp.mLength)
        If tmp > Stamp.mLength Then
            tmp = Stamp.mLength
        ElseIf tmp < 0 Then
            tmp = 0
        End If
        nMedia_Seek tmp
    End If
End Sub

Private Sub ProgressMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    GUI_ResetGraphics
    If Stamp.mState = StatePlaying Or Stamp.mState = StatePaused Then
        tmp = (X + ProgressLeft.Width - (PicButton(BtnProgressSlider).Width / 2)) / ((frmMain.conProgress.Width - frmMain.PicButton(BtnProgressSlider).Width) / Stamp.mLength)
        If tmp > Stamp.mLength Then
            tmp = Stamp.mLength
        ElseIf tmp < 0 Then
            tmp = 0
        End If
        GUI_SetTempStatus "Seek to " & GUI_FormatTime(tmp)
    End If
End Sub

Private Sub ProgressRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    If Stamp.mState = StateStopped Then
        Beep
        Exit Sub
    End If
    If Button = vbKeyLButton Then
        tmp = (X + ProgressLeft.Width + ProgressMiddle.Width - (PicButton(BtnProgressSlider).Width / 2)) / ((frmMain.conProgress.Width - frmMain.PicButton(BtnProgressSlider).Width) / Stamp.mLength)
        If tmp > Stamp.mLength Then
            tmp = Stamp.mLength
        ElseIf tmp < 0 Then
            tmp = 0
        End If
        nMedia_Seek tmp
    End If
End Sub

Private Sub ProgressRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    GUI_ResetGraphics
    If Stamp.mState = StatePlaying Or Stamp.mState = StatePaused Then
        tmp = (X + ProgressLeft.Width + ProgressMiddle.Width - (PicButton(BtnProgressSlider).Width / 2)) / ((frmMain.conProgress.Width - frmMain.PicButton(BtnProgressSlider).Width) / Stamp.mLength)
        If tmp > Stamp.mLength Then
            tmp = Stamp.mLength
        ElseIf tmp < 0 Then
            tmp = 0
        End If
        GUI_SetTempStatus "Seek to " & GUI_FormatTime(tmp)
    End If
End Sub

Private Sub Titlebar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyLButton Then
        ReleaseCapture
        SendMessage Me.hwnd, &HA1, HTCAPTION, 0
    End If
End Sub

Private Sub Titlebar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Byte
    For a = 0 To Buttons
        PicButton(a).Picture = Stamp.Skin.Pic_Button(a)
    Next
    GUI_SetTempStatus "Drag to move window"
End Sub

Private Sub Form_Resize()
    ResizeWindow
End Sub

Private Sub vbrdLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyLButton And WindowState = vbMaximized And Playlist.Visible = True Then plResize = True
End Sub

Private Sub vbrdLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyLButton And WindowState = vbMaximized And Playlist.Visible = True And plResize = True Then
        vbrdLeft.Left = vbrdLeft.Left + X + 30
        Stamp.Config.PlaylistWidth = frmMain.Width - vbrdLeft.Left + vbrdLeft.Width
        ResizeMaximizedWindow
    End If
End Sub

Private Sub vbrdLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    plResize = False
End Sub

Private Sub VolumeLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub VolumeMiddle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    If Button = vbKeyLButton Then
        tmp = (X + VolumeLeft.Width - (PicButton(BtnVolumeSlider).Width / 2)) / ((frmMain.conVolume.Width - frmMain.PicButton(BtnVolumeSlider).Width) / VolumeMax)
        If tmp > VolumeMax Then
            tmp = VolumeMax
        ElseIf tmp < 0 Then
            tmp = 0
        End If
        Stamp.Config.Volume = tmp
        nMedia_SetVolume
        GUI_SetVolumePos
    End If
End Sub

Private Sub VolumeMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Private Sub VolumeRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GUI_ResetGraphics
End Sub

Public Sub ResizeWindow()
    
    Select Case WindowState
    Case vbNormal
        GUI_ControlsVisibleMIN
        ResizeNormalWindow
        frmRes.mnuControls.Enabled = False
    Case vbMaximized
        GUI_ControlsVisibleMAX
        ResizeMaximizedWindow
        frmRes.mnuControls.Enabled = True
    Case Else
        Exit Sub
    End Select
    
End Sub

Public Sub SetVideoDisplay(Optional Zoom As Single)
    If Stamp.mState = StateClosed Then
        Beep
        Exit Sub
    End If
    NoResize = True
    If frmMain.WindowState <> vbNormal Then frmMain.WindowState = vbNormal
    If Zoom = 0 Then Zoom = 1
    Width = Sqr(Zoom) * Stamp.vWidth * 15 + Stamp.Skin.BorderWidth * 4 + 90
    Height = 420 + Stamp.Skin.BorderWidth * 4 + PicButton(BtnExit).Height + Sqr(Zoom) * Stamp.vHeight * 15 + Stamp.Skin.DisplayLineHeight * 2 + Stamp.Skin.Con_ProgressHeight + PicButton(BtnPlay).Height
    NoResize = False
    ResizeWindow
End Sub

Public Sub ResizeNormalWindow()
    If NoResize = True Then Exit Sub
    
    frmRes.mnuFullscreen.Checked = False
    
    If Width < 3000 Then Width = 3000
    If Height < 4000 Then Height = 4000
    
    brdTopLeft.Top = 0
    brdTopLeft.Left = 0
    
    brdTopRight.Top = 0
    brdTopRight.Left = Width - brdTopRight.Width
    
    brdTop.Width = brdTopRight.Left - brdTopLeft.Width
    brdTop.Height = Stamp.Skin.BorderWidth
    brdTop.Top = 0
    brdTop.Left = brdTopLeft.Width
    
    brdBottomLeft.Top = Height - brdBottomLeft.Height
    brdBottomLeft.Left = 0
    
    brdBottomRight.Top = Height - brdBottomRight.Height
    brdBottomRight.Left = Width - brdBottomRight.Width
    
    brdBottom.Width = brdBottomRight.Left - brdBottomLeft.Width
    brdBottom.Height = Stamp.Skin.BorderWidth
    brdBottom.Top = Height - Stamp.Skin.BorderWidth
    brdBottom.Left = brdBottomLeft.Width
    
    brdLeft.Width = Stamp.Skin.BorderWidth
    brdLeft.Top = brdTopLeft.Top + brdTopLeft.Height
    brdLeft.Left = 0
    brdLeft.Height = brdBottomLeft.Top - brdTopLeft.Height
    
    brdRight.Width = Stamp.Skin.BorderWidth
    brdRight.Top = brdLeft.Top
    brdRight.Left = brdTopRight.Left
    brdRight.Height = brdBottomRight.Top - brdTopRight.Height
    
    PicButton(BtnMenu).Left = brdTopLeft.Width
    PicButton(BtnMenu).Top = brdTop.Height
    PicButton(BtnExit).Top = PicButton(BtnMenu).Top
    PicButton(BtnMinimize).Top = PicButton(BtnMenu).Top
    PicButton(BtnExit).Left = brdRight.Left - PicButton(BtnExit).Width
    PicButton(BtnMinimize).Left = PicButton(BtnExit).Left - PicButton(BtnMinimize).Width
    
    Titlebar(2).Top = brdTop.Height
    Titlebar(2).Height = 16 * 15
    Titlebar(2).Width = (PicButton(BtnMinimize).Left - PicButton(BtnMenu).Left - PicButton(BtnMenu).Width - Titlebar(0).Width) / 2
    Titlebar(2).Left = PicButton(BtnMinimize).Left - Titlebar(2).Width
    
    Titlebar(0).Top = brdTop.Height
    Titlebar(0).Height = 15 * 16
    Titlebar(0).Left = Titlebar(2).Left - Titlebar(0).Width
    
    Titlebar(1).Top = brdTop.Height
    Titlebar(1).Height = 16 * 15
    Titlebar(1).Width = Titlebar(0).Left - PicButton(BtnMenu).Left - PicButton(BtnMenu).Width
    Titlebar(1).Left = Titlebar(0).Left - Titlebar(1).Width
    
    PicButton(BtnPlay).Top = brdBottom.Top - 120 - PicButton(BtnPlay).Height
    PicButton(BtnPlay).Left = brdRight.Width + 120
    PicButton(BtnStop).Left = PicButton(BtnPlay).Left + PicButton(BtnPlay).Width + 120
    PicButton(BtnStop).Top = PicButton(BtnPlay).Top
    PicButton(BtnPause).Left = PicButton(BtnStop).Left + PicButton(BtnStop).Width + 120
    PicButton(BtnPause).Top = PicButton(BtnPlay).Top
    PicButton(BtnOpen).Top = PicButton(BtnPlay).Top
    PicButton(BtnOpen).Left = brdRight.Left - 120 - PicButton(BtnOpen).Width
    
    conProgress.Height = Stamp.Skin.Con_ProgressHeight
    conProgress.Left = brdLeft.Width + 90
    conProgress.Top = PicButton(BtnPlay).Top - 120 - conProgress.Height
    conProgress.Width = brdRight.Left - brdLeft.Width - 180
    
    conVolume.Height = Stamp.Skin.DisplayLineHeight
    conVolume.Width = Stamp.Skin.Con_VolumeWidth
    conVolume.Top = conProgress.Top - 90 - conVolume.Height
    conVolume.Left = brdRight.Left - 90 - conVolume.Width
    
    lbl1.Height = Stamp.Skin.DisplayLineHeight
    lbl2.Height = Stamp.Skin.DisplayLineHeight
    lblVolume.Height = Stamp.Skin.DisplayLineHeight
    lbl2.Top = conVolume.Top
    lbl2.Left = conProgress.Left
    lbl1.Top = conVolume.Top - lbl1.Height
    lbl1.Left = lbl2.Left
    lblVolume.Top = lbl1.Top
    lblVolume.Left = conVolume.Left
    lblVolume.Width = conVolume.Width
    lbl1.Width = lblVolume.Left - 90 - lbl1.Left
    lbl2.Width = lbl1.Width
    
    'misc
    VideoDisplay.Left = brdLeft.Width + 90
    VideoDisplay.Width = brdRight.Left - brdLeft.Width - 180
    VideoDisplay.Top = PicButton(BtnExit).Top + PicButton(BtnExit).Height + 90
    VideoDisplay.Height = lbl1.Top - 90 - VideoDisplay.Top
    PicButton(BtnResize).Top = brdBottom.Top - PicButton(BtnResize).Height
    PicButton(BtnResize).Left = brdRight.Left - PicButton(BtnResize).Width
    
    'videoframe
    vbrdTop.Height = Stamp.Skin.BorderWidth
    vbrdTop.Top = VideoDisplay.Top - vbrdTop.Height
    vbrdTop.Left = VideoDisplay.Left
    vbrdTop.Width = VideoDisplay.Width
    
    vbrdTopLeft.Height = Stamp.Skin.BorderWidth
    vbrdTopLeft.Width = Stamp.Skin.BorderWidth
    vbrdTopLeft.Top = vbrdTop.Top
    vbrdTopLeft.Left = VideoDisplay.Left - vbrdTopLeft.Width
    
    vbrdTopRight.Height = Stamp.Skin.BorderWidth
    vbrdTopRight.Width = Stamp.Skin.BorderWidth
    vbrdTopRight.Top = vbrdTop.Top
    vbrdTopRight.Left = VideoDisplay.Left + VideoDisplay.Width
    
    vbrdRight.Width = Stamp.Skin.BorderWidth
    vbrdRight.Left = vbrdTopRight.Left
    vbrdRight.Top = VideoDisplay.Top
    vbrdRight.Height = VideoDisplay.Height
    
    vbrdLeft.Width = Stamp.Skin.BorderWidth
    vbrdLeft.Left = vbrdTopLeft.Left
    vbrdLeft.Top = VideoDisplay.Top
    vbrdLeft.Height = VideoDisplay.Height
    
    vbrdBottom.Height = Stamp.Skin.BorderWidth
    vbrdBottom.Left = VideoDisplay.Left
    vbrdBottom.Width = VideoDisplay.Width
    vbrdBottom.Top = VideoDisplay.Top + VideoDisplay.Height
    
    vbrdBottomRight.Height = Stamp.Skin.BorderWidth
    vbrdBottomRight.Width = Stamp.Skin.BorderWidth
    vbrdBottomRight.Top = vbrdBottom.Top
    vbrdBottomRight.Left = vbrdRight.Left
    
    vbrdBottomLeft.Height = Stamp.Skin.BorderWidth
    vbrdBottomLeft.Width = Stamp.Skin.BorderWidth
    vbrdBottomLeft.Top = vbrdBottom.Top
    vbrdBottomLeft.Left = vbrdLeft.Left
    
    'progressbars
    ProgressLeft.Left = 0
    ProgressLeft.Top = 0
    ProgressLeft.Height = conProgress.Height
    
    ProgressRight.Left = conProgress.Width - ProgressRight.Width
    ProgressRight.Top = 0
    ProgressRight.Height = conProgress.Height
    
    ProgressMiddle.Left = ProgressLeft.Width
    ProgressMiddle.Height = conProgress.Height
    ProgressMiddle.Width = ProgressRight.Left - ProgressMiddle.Left
    ProgressMiddle.Top = 0
    
    VolumeLeft.Left = 0
    VolumeLeft.Top = 0
    VolumeLeft.Height = conVolume.Height
    
    VolumeRight.Left = conVolume.Width - VolumeRight.Width
    VolumeRight.Top = 0
    VolumeRight.Height = conVolume.Height
    
    VolumeMiddle.Left = VolumeLeft.Width
    VolumeMiddle.Height = conVolume.Height
    VolumeMiddle.Width = VolumeRight.Left - VolumeMiddle.Left
    VolumeMiddle.Top = 0
    
    Playlist.Top = VideoDisplay.Top
    Playlist.Left = VideoDisplay.Left
    Playlist.Width = VideoDisplay.Width
    
    PicLogo.Left = VideoDisplay.Width / 2 - PicLogo.Width / 2
    PicLogo.Top = VideoDisplay.Height / 2 - PicLogo.Height / 2
    
    picPlaylistbuttons.Left = Playlist.Left
    picPlaylistbuttons.Height = PicButton(BtnPlaylistAdd).Height + 90
    picPlaylistbuttons.Top = VideoDisplay.Top + VideoDisplay.Height - picPlaylistbuttons.Height
    picPlaylistbuttons.Width = VideoDisplay.Width
    Playlist.Height = picPlaylistbuttons.Top - VideoDisplay.Top
    
    GUI_ResizeVideodisplay
    GUI_SetProgressPos
    vbrdLeft.MousePointer = 0
End Sub

Public Sub ResizeMaximizedWindow()
    frmRes.mnuFullscreen.Checked = True
    If Stamp.Config.PlaylistWidth > frmMain.Width / 2 Then Stamp.Config.PlaylistWidth = frmMain.Width / 2
    If Stamp.Config.PlaylistWidth < Stamp.Skin.PlaylistWidthMin Then Stamp.Config.PlaylistWidth = Stamp.Skin.PlaylistWidthMin
    
    If Playlist.Visible = True Then
        VideoDisplay.Width = frmMain.Width - Stamp.Config.PlaylistWidth - Stamp.Skin.BorderWidth
    Else
        VideoDisplay.Width = frmMain.Width
    End If
    
    VideoDisplay.Top = 0
    VideoDisplay.Left = 0
    
    FullscreenControls.Height = 120 + PicButton(BtnPlay).Height + 120 + conProgress.Height + 90 + lbl1.Height + lbl2.Height + 90 + 90
    
    If Stamp.ControlsVisible = True Then
        VideoDisplay.Height = frmMain.Height - FullscreenControls.Height
    Else
        VideoDisplay.Height = frmMain.Height
    End If
    
    vbrdTop.Top = VideoDisplay.Height
    vbrdTop.Left = 0
    vbrdTop.Width = VideoDisplay.Width
    vbrdTop.Height = Stamp.Skin.BorderWidth
    FullscreenControls.Top = vbrdTop.Top + vbrdTop.Height
    FullscreenControls.Width = VideoDisplay.Width
    FullscreenControls.Left = 0
    FullscreenControls.Height = 1500
    
    Playlist.Top = 0
    picPlaylistbuttons.Height = PicButton(BtnPlaylistAdd).Height + 90
    Playlist.Left = frmMain.Width - Stamp.Config.PlaylistWidth
    Playlist.Width = Stamp.Config.PlaylistWidth
    Playlist.Height = frmMain.Height - picPlaylistbuttons.Height
    
    picPlaylistbuttons.Left = Playlist.Left
    picPlaylistbuttons.Top = Playlist.Height
    picPlaylistbuttons.Width = Playlist.Width
    
    vbrdLeft.Top = 0
    vbrdLeft.Width = Stamp.Skin.BorderWidth
    vbrdLeft.Left = Playlist.Left - vbrdLeft.Width
    vbrdLeft.Height = frmMain.Height
    vbrdLeft.MousePointer = 9
    PicLogo.Left = VideoDisplay.Width / 2 - PicLogo.Width / 2
    PicLogo.Top = VideoDisplay.Height / 2 - PicLogo.Height / 2
    
    PicButton(BtnPlay).Top = frmMain.Height - 120 - PicButton(BtnPlay).Height
    PicButton(BtnPlay).Left = 120
    PicButton(BtnStop).Left = PicButton(BtnPlay).Left + PicButton(BtnPlay).Width + 120
    PicButton(BtnStop).Top = PicButton(BtnPlay).Top
    PicButton(BtnPause).Left = PicButton(BtnStop).Left + PicButton(BtnStop).Width + 120
    PicButton(BtnPause).Top = PicButton(BtnPlay).Top
    PicButton(BtnOpen).Top = PicButton(BtnPlay).Top
    PicButton(BtnOpen).Left = VideoDisplay.Width - 120 - PicButton(BtnOpen).Width
    
    conProgress.Left = 90
    conProgress.Height = Stamp.Skin.Con_ProgressHeight
    conProgress.Top = PicButton(BtnPlay).Top - 120 - conProgress.Height
    conProgress.Width = VideoDisplay.Width - 180
    
    conVolume.Height = Stamp.Skin.DisplayLineHeight
    conVolume.Width = Stamp.Skin.Con_VolumeWidth
    conVolume.Top = conProgress.Top - 90 - conVolume.Height
    conVolume.Left = VideoDisplay.Width - 90 - conVolume.Width
    
    lbl3.Height = Stamp.Skin.DisplayLineHeight
    lbl4.Height = Stamp.Skin.DisplayLineHeight
    lblVolume2.Height = Stamp.Skin.DisplayLineHeight
    lbl4.Left = 90
    lbl3.Top = 90
    lbl4.Top = lbl3.Top + lbl3.Height
    lbl3.Left = lbl4.Left
    lblVolume2.Top = lbl3.Top
    lblVolume2.Left = conVolume.Left
    lblVolume2.Width = conVolume.Width
    lbl3.Width = lblVolume2.Left - 90 - lbl3.Left
    lbl4.Width = lbl3.Width
    
    ProgressLeft.Left = 0
    ProgressLeft.Top = 0
    ProgressLeft.Height = conProgress.Height
    
    ProgressRight.Left = conProgress.Width - ProgressRight.Width
    ProgressRight.Top = 0
    ProgressRight.Height = conProgress.Height
    
    ProgressMiddle.Left = ProgressLeft.Width
    ProgressMiddle.Height = conProgress.Height
    ProgressMiddle.Width = ProgressRight.Left - ProgressMiddle.Left
    ProgressMiddle.Top = 0
    
    VolumeLeft.Left = 0
    VolumeLeft.Top = 0
    VolumeLeft.Height = conVolume.Height
    
    VolumeRight.Left = conVolume.Width - VolumeRight.Width
    VolumeRight.Top = 0
    VolumeRight.Height = conVolume.Height
    
    VolumeMiddle.Left = VolumeLeft.Width
    VolumeMiddle.Height = conVolume.Height
    VolumeMiddle.Width = VolumeRight.Left - VolumeMiddle.Left
    VolumeMiddle.Top = 0
    
    picFullscreenTitle.Top = 0
    picFullscreenTitle.Left = 0
    picFullscreenTitle.Width = frmMain.VideoDisplay.Width

    GUI_ResizeVideodisplay
    GUI_SetProgressPos
End Sub

Public Sub ToggleControls()
    If WindowState <> vbMaximized Then Exit Sub
    If Stamp.ControlsVisible = True Then
        ShowControls False
    Else
        ShowControls True
    End If
    ResizeWindow
End Sub

Public Sub ShowControls(Visible As Boolean)
    Dim a As Integer
    Stamp.ControlsVisible = Visible
    For a = 0 To frmMain.Controls.Count - 1
        Select Case LCase(frmMain.Controls(a).Name)
        Case "vbrdtop", "lbl4", "lbl3", "fullscreencontrols", "conprogress", "convolume", "lblvolume2", "progressmiddle", "progressleft", "progressright", "volumemiddle", "volumeright", "volumeleft"
            frmMain.Controls(a).Visible = Visible
        End Select
        If a < 10 Then PicButton(a).Visible = Visible
    Next
    frmRes.mnuControls.Checked = Visible
    frmMain.ResizeWindow
End Sub
