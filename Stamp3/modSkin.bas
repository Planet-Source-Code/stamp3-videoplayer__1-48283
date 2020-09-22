Attribute VB_Name = "modSkin"
Option Explicit

Public Const BtnPlay = 0
Public Const BtnStop = 1
Public Const BtnPause = 2
Public Const BtnOpen = 3
Public Const BtnMenu = 4
Public Const BtnExit = 5
Public Const BtnMinimize = 6
Public Const BtnResize = 7
Public Const BtnVolumeSlider = 8
Public Const BtnProgressSlider = 9
Public Const BtnPlaylistAdd = 10
Public Const BtnPlaylistAddDir = 11
Public Const BtnPlaylistRemove = 12
Public Const BtnPlaylistClear = 13
Public Const BtnPlaylistLoad = 14
Public Const BtnPlaylistSave = 15
Public Const Buttons = 15

Public Type Stamp3SkinData
    Pic_Button(Buttons) As IPictureDisp
    Pic_ButtonH(Buttons) As IPictureDisp
    Pic_ButtonD(Buttons) As IPictureDisp
    
    Con_ProgressHeight As Integer
    Con_VolumeWidth As Integer
    DisplayLineHeight As Integer
    BorderWidth As Integer
    Font As String
    FontSize As Integer
    FontColor As Long
    VideoDisplayColor As Long
    PlaylistWidthMin As Integer
End Type

Public Type Stamp3SkinRAW
    Signature As String * 10
    Name As String
    Author As String
    BorderWidth As Integer
    ProgressBarHeight As Integer
    VolumeWidth As Integer
    DisplayLineHeight As Integer
    Backcolor As Long
    Font As String
    FontSize As Integer
    FontColor As Long
    VideoDisplayColor As Long
    CommentLine1 As String
    CommentLine2 As String
    CommentLine3 As String
End Type

Public Type Stamp3SkinBitmapRAW
    ID As Byte
    Size As Long
End Type

Public SkinFile(72)

Public Sub Skin_Load()
    Dim SknRAW As Stamp3SkinRAW
    Dim FileHandle As Integer
    Dim a As Byte
    
    If Stamp.Config.CurrentSkinFile = "#" Then
        Skin_LoadInternal
        Exit Sub
    End If
    
    If Dir(Stamp.AppPath & "currentskin\skin.bin") = vbNullString Then GoTo ErrOut
    On Local Error GoTo ErrOut
    
    FileHandle = FreeFile
    Open Stamp.AppPath & "currentskin\skin.bin" For Binary As FileHandle
    Get #1, 1, SknRAW
    Close FileHandle
    
    If SknRAW.Signature <> "Stamp3Skin" Then GoTo ErrOut
    
    Stamp.Skin.BorderWidth = SknRAW.BorderWidth * 15
    Stamp.Skin.Con_ProgressHeight = SknRAW.ProgressBarHeight * 15
    Stamp.Skin.Con_VolumeWidth = SknRAW.VolumeWidth * 15
    Stamp.Skin.DisplayLineHeight = SknRAW.DisplayLineHeight * 15
    Stamp.Skin.Font = SknRAW.Font
    Stamp.Skin.FontColor = SknRAW.FontColor
    Stamp.Skin.FontSize = SknRAW.FontSize
    Stamp.Skin.VideoDisplayColor = SknRAW.VideoDisplayColor
    
    With Stamp.Skin
        Set .Pic_Button(BtnPlay) = LoadPicture(Stamp.AppPath & "currentskin\play.bmp")
        Set .Pic_ButtonD(BtnPlay) = LoadPicture(Stamp.AppPath & "currentskin\play_d.bmp")
        Set .Pic_ButtonH(BtnPlay) = LoadPicture(Stamp.AppPath & "currentskin\play_h.bmp")
        Set .Pic_Button(BtnStop) = LoadPicture(Stamp.AppPath & "currentskin\stop.bmp")
        Set .Pic_ButtonD(BtnStop) = LoadPicture(Stamp.AppPath & "currentskin\stop_d.bmp")
        Set .Pic_ButtonH(BtnStop) = LoadPicture(Stamp.AppPath & "currentskin\stop_h.bmp")
        Set .Pic_Button(BtnPause) = LoadPicture(Stamp.AppPath & "currentskin\pause.bmp")
        Set .Pic_ButtonD(BtnPause) = LoadPicture(Stamp.AppPath & "currentskin\pause_d.bmp")
        Set .Pic_ButtonH(BtnPause) = LoadPicture(Stamp.AppPath & "currentskin\pause_h.bmp")
        Set .Pic_Button(BtnOpen) = LoadPicture(Stamp.AppPath & "currentskin\open.bmp")
        Set .Pic_ButtonD(BtnOpen) = LoadPicture(Stamp.AppPath & "currentskin\open_d.bmp")
        Set .Pic_ButtonH(BtnOpen) = LoadPicture(Stamp.AppPath & "currentskin\open_h.bmp")
        Set .Pic_Button(BtnMenu) = LoadPicture(Stamp.AppPath & "currentskin\menu.bmp")
        Set .Pic_ButtonD(BtnMenu) = LoadPicture(Stamp.AppPath & "currentskin\menu_d.bmp")
        Set .Pic_ButtonH(BtnMenu) = LoadPicture(Stamp.AppPath & "currentskin\menu_h.bmp")
        Set .Pic_Button(BtnExit) = LoadPicture(Stamp.AppPath & "currentskin\exit.bmp")
        Set .Pic_ButtonD(BtnExit) = LoadPicture(Stamp.AppPath & "currentskin\exit_d.bmp")
        Set .Pic_ButtonH(BtnExit) = LoadPicture(Stamp.AppPath & "currentskin\exit_h.bmp")
        Set .Pic_Button(BtnMinimize) = LoadPicture(Stamp.AppPath & "currentskin\minimize.bmp")
        Set .Pic_ButtonD(BtnMinimize) = LoadPicture(Stamp.AppPath & "currentskin\minimize_d.bmp")
        Set .Pic_ButtonH(BtnMinimize) = LoadPicture(Stamp.AppPath & "currentskin\minimize_h.bmp")
        Set .Pic_Button(BtnResize) = LoadPicture(Stamp.AppPath & "currentskin\resize.bmp")
        Set .Pic_ButtonD(BtnResize) = LoadPicture(Stamp.AppPath & "currentskin\resize_d.bmp")
        Set .Pic_ButtonH(BtnResize) = LoadPicture(Stamp.AppPath & "currentskin\resize_h.bmp")
        Set .Pic_Button(BtnProgressSlider) = LoadPicture(Stamp.AppPath & "currentskin\progress_slider.bmp")
        Set .Pic_ButtonD(BtnProgressSlider) = LoadPicture(Stamp.AppPath & "currentskin\progress_slider_d.bmp")
        Set .Pic_ButtonH(BtnProgressSlider) = LoadPicture(Stamp.AppPath & "currentskin\progress_slider_h.bmp")
        Set .Pic_Button(BtnVolumeSlider) = LoadPicture(Stamp.AppPath & "currentskin\volume_slider.bmp")
        Set .Pic_ButtonD(BtnVolumeSlider) = LoadPicture(Stamp.AppPath & "currentskin\volume_slider_d.bmp")
        Set .Pic_ButtonH(BtnVolumeSlider) = LoadPicture(Stamp.AppPath & "currentskin\volume_slider_h.bmp")
        
        Set .Pic_Button(BtnPlaylistAdd) = LoadPicture(Stamp.AppPath & "currentskin\playlist_add.bmp")
        Set .Pic_ButtonD(BtnPlaylistAdd) = LoadPicture(Stamp.AppPath & "currentskin\playlist_add_d.bmp")
        Set .Pic_ButtonH(BtnPlaylistAdd) = LoadPicture(Stamp.AppPath & "currentskin\playlist_add_h.bmp")
        Set .Pic_Button(BtnPlaylistRemove) = LoadPicture(Stamp.AppPath & "currentskin\playlist_remove.bmp")
        Set .Pic_ButtonD(BtnPlaylistRemove) = LoadPicture(Stamp.AppPath & "currentskin\playlist_remove_d.bmp")
        Set .Pic_ButtonH(BtnPlaylistRemove) = LoadPicture(Stamp.AppPath & "currentskin\playlist_remove_h.bmp")
        Set .Pic_Button(BtnPlaylistSave) = LoadPicture(Stamp.AppPath & "currentskin\playlist_save.bmp")
        Set .Pic_ButtonD(BtnPlaylistSave) = LoadPicture(Stamp.AppPath & "currentskin\playlist_save_d.bmp")
        Set .Pic_ButtonH(BtnPlaylistSave) = LoadPicture(Stamp.AppPath & "currentskin\playlist_save_h.bmp")
        Set .Pic_Button(BtnPlaylistLoad) = LoadPicture(Stamp.AppPath & "currentskin\playlist_load.bmp")
        Set .Pic_ButtonD(BtnPlaylistLoad) = LoadPicture(Stamp.AppPath & "currentskin\playlist_load_d.bmp")
        Set .Pic_ButtonH(BtnPlaylistLoad) = LoadPicture(Stamp.AppPath & "currentskin\playlist_load_h.bmp")
        Set .Pic_Button(BtnPlaylistAddDir) = LoadPicture(Stamp.AppPath & "currentskin\playlist_adddir.bmp")
        Set .Pic_ButtonD(BtnPlaylistAddDir) = LoadPicture(Stamp.AppPath & "currentskin\playlist_adddir_d.bmp")
        Set .Pic_ButtonH(BtnPlaylistAddDir) = LoadPicture(Stamp.AppPath & "currentskin\playlist_adddir_h.bmp")
        Set .Pic_Button(BtnPlaylistClear) = LoadPicture(Stamp.AppPath & "currentskin\playlist_clear.bmp")
        Set .Pic_ButtonD(BtnPlaylistClear) = LoadPicture(Stamp.AppPath & "currentskin\playlist_clear_d.bmp")
        Set .Pic_ButtonH(BtnPlaylistClear) = LoadPicture(Stamp.AppPath & "currentskin\playlist_clear_h.bmp")
    End With
    
    With frmMain
        .vbrdTop.Picture = LoadPicture(Stamp.AppPath & "currentskin\vtop.bmp")
        .vbrdTopRight.Picture = LoadPicture(Stamp.AppPath & "currentskin\vtopright.bmp")
        .vbrdTopLeft.Picture = LoadPicture(Stamp.AppPath & "currentskin\vtopleft.bmp")
        .vbrdBottom.Picture = LoadPicture(Stamp.AppPath & "currentskin\vbottom.bmp")
        .vbrdBottomRight.Picture = LoadPicture(Stamp.AppPath & "currentskin\vbottomright.bmp")
        .vbrdBottomLeft.Picture = LoadPicture(Stamp.AppPath & "currentskin\vbottomleft.bmp")
        .vbrdRight.Picture = LoadPicture(Stamp.AppPath & "currentskin\vright.bmp")
        .vbrdLeft.Picture = LoadPicture(Stamp.AppPath & "currentskin\vleft.bmp")
        .Titlebar(1).Picture = LoadPicture(Stamp.AppPath & "currentskin\titlebarstretch.bmp")
        .Titlebar(2).Picture = LoadPicture(Stamp.AppPath & "currentskin\titlebarstretch.bmp")
        .Titlebar(0).Picture = LoadPicture(Stamp.AppPath & "currentskin\titlebar.bmp")
        .PicLogo.Picture = LoadPicture(Stamp.AppPath & "currentskin\logo.bmp")
        .brdTop.Picture = LoadPicture(Stamp.AppPath & "currentskin\btop.bmp")
        .brdTopRight.Picture = LoadPicture(Stamp.AppPath & "currentskin\btopright.bmp")
        .brdTopLeft.Picture = LoadPicture(Stamp.AppPath & "currentskin\btopleft.bmp")
        .brdBottom.Picture = LoadPicture(Stamp.AppPath & "currentskin\bbottom.bmp")
        .brdBottomRight.Picture = LoadPicture(Stamp.AppPath & "currentskin\bbottomright.bmp")
        .brdBottomLeft.Picture = LoadPicture(Stamp.AppPath & "currentskin\bbottomleft.bmp")
        .brdRight.Picture = LoadPicture(Stamp.AppPath & "currentskin\bright.bmp")
        .brdLeft.Picture = LoadPicture(Stamp.AppPath & "currentskin\bleft.bmp")
        .ProgressLeft.Picture = LoadPicture(Stamp.AppPath & "currentskin\progress_left.bmp")
        .ProgressMiddle.Picture = LoadPicture(Stamp.AppPath & "currentskin\progress_middle.bmp")
        .ProgressRight.Picture = LoadPicture(Stamp.AppPath & "currentskin\progress_right.bmp")
        .VolumeLeft.Picture = LoadPicture(Stamp.AppPath & "currentskin\volume_left.bmp")
        .VolumeRight.Picture = LoadPicture(Stamp.AppPath & "currentskin\volume_right.bmp")
        .VolumeMiddle.Picture = LoadPicture(Stamp.AppPath & "currentskin\volume_middle.bmp")
        .Backcolor = SknRAW.Backcolor
        .lbl1.Font.Name = SknRAW.Font
        .lbl2.Font.Name = SknRAW.Font
        .lblVolume.Font.Name = SknRAW.Font
        .lbl1.Font.Size = SknRAW.FontSize
        .lbl2.Font.Size = SknRAW.FontSize
        .lblVolume.Font.Size = SknRAW.FontSize
        .lbl1.ForeColor = SknRAW.FontColor
        .lbl2.ForeColor = SknRAW.FontColor
        .lbl3.Font.Name = .lbl1.Font.Name
        .lbl3.Font.Size = .lbl1.Font.Size
        .lbl3.ForeColor = .lbl2.ForeColor
        .lbl4.Font.Name = .lbl2.Font.Name
        .lbl4.Font.Size = .lbl2.Font.Size
        .lbl4.ForeColor = .lbl2.ForeColor
        .lblVolume.ForeColor = SknRAW.FontColor
        .lblVolume2.ForeColor = SknRAW.FontColor
        .picFullscreenTitle.Font = SknRAW.Font
        If Stamp.mState <> StatePlaying Then .VideoDisplay.Backcolor = Stamp.Skin.VideoDisplayColor Else .VideoDisplay.Backcolor = vbBlack
        .picPlaylistbuttons.Backcolor = .Backcolor
        .FullscreenControls.Backcolor = .Backcolor
        Stamp.Skin.PlaylistWidthMin = 180
        For a = 10 To 15
            .PicButton(a).Top = 45
            If a = 10 Then .PicButton(a).Left = 90 Else .PicButton(a).Left = .PicButton(a - 1).Left + .PicButton(a - 1).Width + 90
            Stamp.Skin.PlaylistWidthMin = Stamp.Skin.PlaylistWidthMin + .PicButton(a).Width + 90
        Next
    End With
    
    GUI_ResetGraphics
    frmMain.ResizeWindow
    GUI_SetInfo
    GUI_SetProgressPos
    GUI_SetVolumePos
    Exit Sub
    
ErrOut:
    Beep
    Skin_Extract
End Sub

Public Sub Skin_Extract()
    Dim SknRAW As Stamp3SkinRAW
    Dim SknBMRAW As Stamp3SkinBitmapRAW
    Dim StrData As String
    Dim InPutHandle As Integer, OutputHandle As Integer
    Dim a As Long
    
    On Local Error GoTo ErrOut
    
    If Stamp.Config.CurrentSkinFile = "#" Then
        Skin_LoadInternal
    Else
        If Dir(Stamp.AppPath & "CurrentSkin", vbDirectory) = vbNullString Then MkDir Stamp.AppPath & "CurrentSkin"
        If Dir(Stamp.AppPath & "CurrentSkin\*.*") <> vbNullString Then Kill Stamp.AppPath & "CurrentSkin\*.*"
        
        InPutHandle = FreeFile
        Open Stamp.Config.CurrentSkinFile For Binary As InPutHandle
        
        OutputHandle = FreeFile
        Open Stamp.AppPath & "CurrentSkin\skin.bin" For Binary As OutputHandle
        Get InPutHandle, 1, SknRAW
        Put OutputHandle, 1, SknRAW
        Close OutputHandle
        If SknRAW.Signature <> "Stamp3Skin" Then GoTo ErrOut
        
        For a = 0 To UBound(SkinFile)
            Get InPutHandle, , SknBMRAW
            OutputHandle = FreeFile
            Open Stamp.AppPath & "currentskin\" & SkinFile(SknBMRAW.ID) & ".bmp" For Binary As OutputHandle
            StrData = String(SknBMRAW.Size, 0)
            Get InPutHandle, , StrData
            StrData = DecodeString(StrData)
            Put OutputHandle, 1, StrData
            Close OutputHandle
        Next
        
        Close InPutHandle
    End If
    
    Skin_Load
    Exit Sub
    
ErrOut:
    Close InPutHandle
    Close OutputHandle
    Beep
    Skin_LoadInternal
End Sub

Public Sub Skin_LoadInternal()
    Dim a As Byte
    Stamp.Skin.BorderWidth = 45
    Stamp.Skin.Con_ProgressHeight = 225
    Stamp.Skin.Con_VolumeWidth = 855
    Stamp.Skin.DisplayLineHeight = 195
    Stamp.Skin.Font = "Verdana"
    Stamp.Skin.FontColor = vbWhite
    Stamp.Skin.FontSize = 8
    Stamp.Skin.VideoDisplayColor = vbBlack
    
    With Stamp.Skin
        Set .Pic_Button(BtnPlay) = LoadResPicture("play", vbResBitmap)
        Set .Pic_ButtonD(BtnPlay) = LoadResPicture("play_d", vbResBitmap)
        Set .Pic_ButtonH(BtnPlay) = LoadResPicture("play_h", vbResBitmap)
        Set .Pic_Button(BtnStop) = LoadResPicture("stop", vbResBitmap)
        Set .Pic_ButtonD(BtnStop) = LoadResPicture("stop_d", vbResBitmap)
        Set .Pic_ButtonH(BtnStop) = LoadResPicture("stop_h", vbResBitmap)
        Set .Pic_Button(BtnPause) = LoadResPicture("pause", vbResBitmap)
        Set .Pic_ButtonD(BtnPause) = LoadResPicture("pause_d", vbResBitmap)
        Set .Pic_ButtonH(BtnPause) = LoadResPicture("pause_h", vbResBitmap)
        Set .Pic_Button(BtnOpen) = LoadResPicture("open", vbResBitmap)
        Set .Pic_ButtonD(BtnOpen) = LoadResPicture("open_d", vbResBitmap)
        Set .Pic_ButtonH(BtnOpen) = LoadResPicture("open_h", vbResBitmap)
        Set .Pic_Button(BtnMenu) = LoadResPicture("menu", vbResBitmap)
        Set .Pic_ButtonD(BtnMenu) = LoadResPicture("menu_d", vbResBitmap)
        Set .Pic_ButtonH(BtnMenu) = LoadResPicture("menu_h", vbResBitmap)
        Set .Pic_Button(BtnExit) = LoadResPicture("exit", vbResBitmap)
        Set .Pic_ButtonD(BtnExit) = LoadResPicture("exit_d", vbResBitmap)
        Set .Pic_ButtonH(BtnExit) = LoadResPicture("exit_h", vbResBitmap)
        Set .Pic_Button(BtnMinimize) = LoadResPicture("minimize", vbResBitmap)
        Set .Pic_ButtonD(BtnMinimize) = LoadResPicture("minimize_d", vbResBitmap)
        Set .Pic_ButtonH(BtnMinimize) = LoadResPicture("minimize_h", vbResBitmap)
        Set .Pic_Button(BtnResize) = LoadResPicture("resize", vbResBitmap)
        Set .Pic_ButtonD(BtnResize) = LoadResPicture("resize_d", vbResBitmap)
        Set .Pic_ButtonH(BtnResize) = LoadResPicture("resize_h", vbResBitmap)
        Set .Pic_Button(BtnProgressSlider) = LoadResPicture("progress_slider", vbResBitmap)
        Set .Pic_ButtonD(BtnProgressSlider) = LoadResPicture("progress_slider_d", vbResBitmap)
        Set .Pic_ButtonH(BtnProgressSlider) = LoadResPicture("progress_slider_h", vbResBitmap)
        Set .Pic_Button(BtnVolumeSlider) = LoadResPicture("volume_slider", vbResBitmap)
        Set .Pic_ButtonD(BtnVolumeSlider) = LoadResPicture("volume_slider_d", vbResBitmap)
        Set .Pic_ButtonH(BtnVolumeSlider) = LoadResPicture("volume_slider_h", vbResBitmap)
        
        Set .Pic_Button(BtnPlaylistAdd) = LoadResPicture("playlist_add", vbResBitmap)
        Set .Pic_ButtonD(BtnPlaylistAdd) = LoadResPicture("playlist_add_d", vbResBitmap)
        Set .Pic_ButtonH(BtnPlaylistAdd) = LoadResPicture("playlist_add_h", vbResBitmap)
        Set .Pic_Button(BtnPlaylistRemove) = LoadResPicture("playlist_remove", vbResBitmap)
        Set .Pic_ButtonD(BtnPlaylistRemove) = LoadResPicture("playlist_remove_d", vbResBitmap)
        Set .Pic_ButtonH(BtnPlaylistRemove) = LoadResPicture("playlist_remove_h", vbResBitmap)
        Set .Pic_Button(BtnPlaylistSave) = LoadResPicture("playlist_save", vbResBitmap)
        Set .Pic_ButtonD(BtnPlaylistSave) = LoadResPicture("playlist_save_d", vbResBitmap)
        Set .Pic_ButtonH(BtnPlaylistSave) = LoadResPicture("playlist_save_h", vbResBitmap)
        Set .Pic_Button(BtnPlaylistLoad) = LoadResPicture("playlist_load", vbResBitmap)
        Set .Pic_ButtonD(BtnPlaylistLoad) = LoadResPicture("playlist_load_d", vbResBitmap)
        Set .Pic_ButtonH(BtnPlaylistLoad) = LoadResPicture("playlist_load_h", vbResBitmap)
        Set .Pic_Button(BtnPlaylistAddDir) = LoadResPicture("playlist_adddir", vbResBitmap)
        Set .Pic_ButtonD(BtnPlaylistAddDir) = LoadResPicture("playlist_adddir_d", vbResBitmap)
        Set .Pic_ButtonH(BtnPlaylistAddDir) = LoadResPicture("playlist_adddir_h", vbResBitmap)
        Set .Pic_Button(BtnPlaylistClear) = LoadResPicture("playlist_clear", vbResBitmap)
        Set .Pic_ButtonD(BtnPlaylistClear) = LoadResPicture("playlist_clear_d", vbResBitmap)
        Set .Pic_ButtonH(BtnPlaylistClear) = LoadResPicture("playlist_clear_h", vbResBitmap)
    End With
    
    With frmMain
        .vbrdTop.Picture = LoadResPicture("vtop", vbResBitmap)
        .vbrdTopRight.Picture = LoadResPicture("vtopright", vbResBitmap)
        .vbrdTopLeft.Picture = LoadResPicture("vtopleft", vbResBitmap)
        .vbrdBottom.Picture = LoadResPicture("vbottom", vbResBitmap)
        .vbrdBottomRight.Picture = LoadResPicture("vbottomright", vbResBitmap)
        .vbrdBottomLeft.Picture = LoadResPicture("vbottomleft", vbResBitmap)
        .vbrdRight.Picture = LoadResPicture("vright", vbResBitmap)
        .vbrdLeft.Picture = LoadResPicture("vleft", vbResBitmap)
        .Titlebar(1).Picture = LoadResPicture("titlebarstretch", vbResBitmap)
        .Titlebar(2).Picture = LoadResPicture("titlebarstretch", vbResBitmap)
        .Titlebar(0).Picture = LoadResPicture("titlebar", vbResBitmap)
        .PicLogo.Picture = LoadResPicture("logo", vbResBitmap)
        .brdTop.Picture = LoadResPicture("btop", vbResBitmap)
        .brdTopRight.Picture = LoadResPicture("btopright", vbResBitmap)
        .brdTopLeft.Picture = LoadResPicture("btopleft", vbResBitmap)
        .brdBottom.Picture = LoadResPicture("bbottom", vbResBitmap)
        .brdBottomRight.Picture = LoadResPicture("bbottomright", vbResBitmap)
        .brdBottomLeft.Picture = LoadResPicture("bbottomleft", vbResBitmap)
        .brdRight.Picture = LoadResPicture("bright", vbResBitmap)
        .brdLeft.Picture = LoadResPicture("bleft", vbResBitmap)
        .ProgressLeft.Picture = LoadResPicture("progress_left", vbResBitmap)
        .ProgressMiddle.Picture = LoadResPicture("progress_middle", vbResBitmap)
        .ProgressRight.Picture = LoadResPicture("progress_right", vbResBitmap)
        .VolumeLeft.Picture = LoadResPicture("volume_left", vbResBitmap)
        .VolumeRight.Picture = LoadResPicture("volume_right", vbResBitmap)
        .VolumeMiddle.Picture = LoadResPicture("volume_middle", vbResBitmap)
        .Backcolor = RGB(97, 112, 133)
        .lbl1.Font.Name = Stamp.Skin.Font
        .lbl2.Font.Name = Stamp.Skin.Font
        .lblVolume.Font.Name = Stamp.Skin.Font
        .lbl1.Font.Size = Stamp.Skin.FontSize
        .lbl2.Font.Size = Stamp.Skin.FontSize
        .lblVolume.Font.Size = Stamp.Skin.FontSize
        .lbl1.ForeColor = Stamp.Skin.FontColor
        .lbl2.ForeColor = Stamp.Skin.FontColor
        .lbl3.Font.Name = .lbl1.Font.Name
        .lbl3.Font.Size = .lbl1.Font.Size
        .lbl3.ForeColor = .lbl2.ForeColor
        .lbl4.Font.Name = .lbl2.Font.Name
        .lbl4.Font.Size = .lbl2.Font.Size
        .lbl4.ForeColor = .lbl2.ForeColor
        .lblVolume.ForeColor = vbWhite
        .lblVolume2.ForeColor = vbWhite
        .picFullscreenTitle.Font = Stamp.Skin.Font
        If Stamp.mState <> StatePlaying Then .VideoDisplay.Backcolor = Stamp.Skin.VideoDisplayColor Else .VideoDisplay.Backcolor = vbBlack
        .picPlaylistbuttons.Backcolor = RGB(97, 112, 133)
        .FullscreenControls.Backcolor = .Backcolor
        Stamp.Skin.PlaylistWidthMin = 90
        For a = 10 To 15
            .PicButton(a).Top = 45
            If a = 10 Then .PicButton(a).Left = 90 Else .PicButton(a).Left = .PicButton(a - 1).Left + .PicButton(a - 1).Width + 90
            Stamp.Skin.PlaylistWidthMin = Stamp.Skin.PlaylistWidthMin + .PicButton(a).Width + 90
        Next
    End With
    
    GUI_ResetGraphics
    frmMain.ResizeWindow
    GUI_SetInfo
    GUI_SetProgressPos
    GUI_SetVolumePos
    
End Sub

Public Sub Skin_SetUpNames()
    SkinFile(0) = "bbottom"
    SkinFile(1) = "bbottomleft"
    SkinFile(2) = "bbottomright"
    SkinFile(3) = "bleft"
    SkinFile(4) = "bright"
    SkinFile(5) = "btop"
    SkinFile(6) = "btopleft"
    SkinFile(7) = "btopright"
    SkinFile(8) = "exit"
    SkinFile(9) = "exit_d"
    SkinFile(10) = "exit_h"
    SkinFile(11) = "logo"
    SkinFile(12) = "menu"
    SkinFile(13) = "menu_d"
    SkinFile(14) = "menu_h"
    SkinFile(15) = "minimize"
    SkinFile(16) = "minimize_d"
    SkinFile(17) = "minimize_h"
    SkinFile(18) = "open"
    SkinFile(19) = "open_d"
    SkinFile(20) = "open_h"
    SkinFile(21) = "pause"
    SkinFile(22) = "pause_d"
    SkinFile(23) = "pause_h"
    SkinFile(24) = "play"
    SkinFile(25) = "play_d"
    SkinFile(26) = "play_h"
    SkinFile(27) = "playlist_add"
    SkinFile(28) = "playlist_add_d"
    SkinFile(29) = "playlist_add_h"
    SkinFile(30) = "playlist_adddir"
    SkinFile(31) = "playlist_adddir_d"
    SkinFile(32) = "playlist_adddir_h"
    SkinFile(33) = "playlist_clear"
    SkinFile(34) = "playlist_clear_d"
    SkinFile(35) = "playlist_clear_h"
    SkinFile(36) = "playlist_load"
    SkinFile(37) = "playlist_load_d"
    SkinFile(38) = "playlist_load_h"
    SkinFile(39) = "playlist_remove"
    SkinFile(40) = "playlist_remove_d"
    SkinFile(41) = "playlist_remove_h"
    SkinFile(42) = "playlist_save"
    SkinFile(43) = "playlist_save_d"
    SkinFile(44) = "playlist_save_h"
    SkinFile(45) = "progress_left"
    SkinFile(46) = "progress_middle"
    SkinFile(47) = "progress_right"
    SkinFile(48) = "progress_slider"
    SkinFile(49) = "progress_slider_d"
    SkinFile(50) = "progress_slider_h"
    SkinFile(51) = "resize"
    SkinFile(52) = "resize_d"
    SkinFile(53) = "resize_h"
    SkinFile(54) = "stop"
    SkinFile(55) = "stop_d"
    SkinFile(56) = "stop_h"
    SkinFile(57) = "titlebar"
    SkinFile(58) = "titlebarstretch"
    SkinFile(59) = "vbottom"
    SkinFile(60) = "vbottomleft"
    SkinFile(61) = "vbottomright"
    SkinFile(62) = "vleft"
    SkinFile(63) = "volume_left"
    SkinFile(64) = "volume_middle"
    SkinFile(65) = "volume_right"
    SkinFile(66) = "volume_slider"
    SkinFile(67) = "volume_slider_d"
    SkinFile(68) = "volume_slider_h"
    SkinFile(69) = "vright"
    SkinFile(70) = "vtop"
    SkinFile(71) = "vtopleft"
    SkinFile(72) = "vtopright"
End Sub
