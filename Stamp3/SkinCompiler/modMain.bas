Attribute VB_Name = "modMain"
Option Explicit

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

Dim SkinFile(72)
Dim Data As String

Sub Main()
    Dim AppPath As String
    Dim SknRAW As Stamp3SkinRAW
    Dim SknBMRAW As Stamp3SkinBitmapRAW
    Dim tmp As Variant
    Dim a As Long
    
    SetUpNames
    
    AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    If Dir(AppPath & "skin.txt") = vbNullString Then GoTo ErrOut
    
    Open AppPath & "skin.txt" For Input As #1
    
    'name
    Line Input #1, Data
    If param1 <> "name" Then GoTo ErrOut
    SknRAW.Name = param2
    
    'author
    Line Input #1, Data
    If param1 <> "author" Then GoTo ErrOut
    SknRAW.Author = param2
    
    'borderwidth
    Line Input #1, Data
    If param1 <> "borderwidth" Then GoTo ErrOut
    SknRAW.BorderWidth = param2
    
    'progressbarheight
    Line Input #1, Data
    If param1 <> "progressbarheight" Then GoTo ErrOut
    SknRAW.ProgressBarHeight = param2
    
    'volumewidth
    Line Input #1, Data
    If param1 <> "volumewidth" Then GoTo ErrOut
    SknRAW.VolumeWidth = param2
    
    'displaylineheight
    Line Input #1, Data
    If param1 <> "displaylineheight" Then GoTo ErrOut
    SknRAW.DisplayLineHeight = param2
    
    'backcolor
    Line Input #1, Data
    If param1 <> "backcolor" Then GoTo ErrOut
    tmp = Split(param2, ",")
    SknRAW.Backcolor = RGB(CByte(tmp(0)), CByte(tmp(1)), CByte(tmp(2)))
    
    'font
    Line Input #1, Data
    If param1 <> "font" Then GoTo ErrOut
    SknRAW.Font = param2
    
    'fontsize
    Line Input #1, Data
    If param1 <> "fontsize" Then GoTo ErrOut
    SknRAW.FontSize = param2
    
    'fontcolor
    Line Input #1, Data
    If param1 <> "fontcolor" Then GoTo ErrOut
    tmp = Split(param2, ",")
    SknRAW.FontColor = RGB(CByte(tmp(0)), CByte(tmp(1)), CByte(tmp(2)))
    
    'videodisplaycolor
    Line Input #1, Data
    If param1 <> "videodisplaycolor" Then GoTo ErrOut
    tmp = Split(param2, ",")
    SknRAW.VideoDisplayColor = RGB(CByte(tmp(0)), CByte(tmp(1)), CByte(tmp(2)))
        
    'commentline1
    Line Input #1, Data
    If param1 <> "commentline1" Then GoTo ErrOut
    SknRAW.CommentLine1 = param2
    
    'commentline2
    Line Input #1, Data
    If param1 <> "commentline2" Then GoTo ErrOut
    SknRAW.CommentLine2 = param2
    
    'commentline3
    Line Input #1, Data
    If param1 <> "commentline3" Then GoTo ErrOut
    SknRAW.CommentLine3 = param2
    
    SknRAW.Signature = "Stamp3Skin"
    Close #1
    
    If Dir(AppPath & SknRAW.Name & ".s3skn") <> vbNullString Then GoTo ErrOut
    
    Open AppPath & SknRAW.Name & ".s3skn" For Binary As #2
    Put #2, 1, SknRAW
    
    For a = 0 To UBound(SkinFile)
        Open AppPath & SkinFile(a) & ".bmp" For Binary As #1
        Data = String(LOF(1), 0)
        Get #1, , Data
        Data = EncodeString(Data)
        SknBMRAW.ID = a
        SknBMRAW.Size = Len(Data)
        Put #2, , SknBMRAW
        Put #2, , Data
        Close #1
    Next
    
    Close #2
    MsgBox "Skin Compiled"
    End
    
ErrOut:
    MsgBox "Invalid skin!", vbExclamation
    Reset
    End
    
End Sub

Function param1() As String
    param1 = LCase(Left(Data, InStr(1, Data, "=") - 1))
End Function

Function param2() As String
    param2 = Mid(Data, InStr(1, Data, "=") + 1)
End Function

Public Sub SetUpNames()
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
