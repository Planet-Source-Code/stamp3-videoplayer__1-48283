Attribute VB_Name = "modGUI"
Option Explicit

Public Sub GUI_SetVolumePos()
    frmMain.PicButton(BtnVolumeSlider).Left = (frmMain.conVolume.Width - frmMain.PicButton(BtnVolumeSlider).Width) / VolumeMax * Stamp.Config.Volume
End Sub

Public Function GUI_GetVolumePos() As Long
    GUI_GetVolumePos = frmMain.PicButton(BtnVolumeSlider).Left / ((frmMain.conVolume.Width - frmMain.PicButton(BtnVolumeSlider).Width) / VolumeMax)
End Function

Public Sub GUI_SetProgressPos()
    If Stamp.mLength = 0 Then
        frmMain.PicButton(BtnProgressSlider).Left = 0
    Else
        frmMain.PicButton(BtnProgressSlider).Left = (frmMain.conProgress.Width - frmMain.PicButton(BtnProgressSlider).Width) / Stamp.mLength * Stamp.mPosition
    End If
End Sub

Public Function GUI_GetProgressPos() As Long
    GUI_GetProgressPos = frmMain.PicButton(BtnProgressSlider).Left / ((frmMain.conProgress.Width - frmMain.PicButton(BtnProgressSlider).Width) / Stamp.mLength)
End Function

Public Sub GUI_ResetButtonGraphics()
    Dim a As Byte
    For a = 0 To Buttons
        frmMain.PicButton(a).Picture = Stamp.Skin.Pic_Button(a)
    Next
End Sub

Public Sub GUI_ResetGraphics()
    GUI_ResetButtonGraphics
    GUI_SetInfo
End Sub

Public Sub GUI_ScrollingVolume()
    Stamp.Config.Volume = GUI_GetVolumePos
    GUI_SetTempStatus "Set volume to " & Int(Stamp.Config.Volume / VolumeMax * 100) & "%"
    nMedia_SetVolume
End Sub

Public Sub GUI_ScrollingProgress()
    If Stamp.mState = StatePaused Or Stamp.mState = StatePlaying Then GUI_SetTempStatus "Seek to " & GUI_FormatTime(GUI_GetProgressPos)
End Sub

Public Sub GUI_ReleasedScrollProgress()
    If Stamp.mState <> StateStopped Then
        nMedia_Seek GUI_GetProgressPos
    Else
        Beep
        GUI_SetProgressPos
    End If
End Sub

Public Sub GUI_ReleasedScrollVolume()
    GUI_SetInfo
End Sub

Public Sub GUI_SetInfo()
    frmMain.lbl1 = Stamp.mTitle
    frmMain.lbl2 = Stamp.mStatus
    frmMain.lbl3 = frmMain.lbl1
    frmMain.lbl4 = frmMain.lbl2
End Sub

Public Function GUI_FormatTime(Time As Long) As String
    GUI_FormatTime = Format(Time / 1000 / 86400, "hh:mm:ss")
End Function

Public Sub GUI_SetTempStatus(Text As String)
    Stamp.cCDT = 10
    frmMain.lbl2 = Text
    frmMain.lbl4 = frmMain.lbl2
    frmMain.lbl2.Refresh
    frmMain.lbl4.Refresh
    Stamp.gLastTempState = Text
End Sub

Public Function GUI_MakeTitleString(strTitle As String) As String
    If InStr(1, strTitle, "\") <> 0 Then GUI_MakeTitleString = Right(strTitle, Len(strTitle) - InStrRev(strTitle, "\"))
    If InStr(1, GUI_MakeTitleString, ".") <> 0 Then GUI_MakeTitleString = Left(GUI_MakeTitleString, InStrRev(GUI_MakeTitleString, ".", Len(GUI_MakeTitleString)) - 1)
    GUI_MakeTitleString = Replace(GUI_MakeTitleString, "_", " ", , , vbBinaryCompare)
    GUI_MakeTitleString = StrConv(GUI_MakeTitleString, vbProperCase)
End Function

Public Sub GUI_ResizeVideodisplay()
    Dim l As Long, w As Long, h As Long, t As Long, z As Single
    If Stamp.Config.VideoZoomOn = True And frmMain.WindowState = vbMaximized Then z = Stamp.Config.Zoom Else z = 1
    If Stamp.mState = StateStopped Or Stamp.mState = StateClosed Then Exit Sub
    If frmMain.VideoDisplay.Width / Stamp.vAspectRatio > frmMain.VideoDisplay.Height Then
        If Stamp.mState = StatePlaying Or Stamp.mState = StatePaused Then
            w = (frmMain.VideoDisplay.Height * Stamp.vAspectRatio) / 15
            h = (frmMain.VideoDisplay.Height) / 15
            l = (frmMain.VideoDisplay.Width / 2 - (frmMain.VideoDisplay.Height * z * Stamp.vAspectRatio) / 2) / 15
            t = (frmMain.VideoDisplay.Height / 2 / 15) - (h * z) / 2
            Stamp.oVideoWindow.Left = l
            Stamp.oVideoWindow.Top = t
            Stamp.oVideoWindow.Width = w * z
            Stamp.oVideoWindow.Height = h * z
        End If
    Else
        If Stamp.mState = StatePlaying Or Stamp.mState = StatePaused Then
            w = frmMain.VideoDisplay.Width / 15
            h = (frmMain.VideoDisplay.Width / Stamp.vAspectRatio) / 15
            t = (frmMain.VideoDisplay.Height / 2 - (frmMain.VideoDisplay.Width * z / Stamp.vAspectRatio) / 2) / 15
            l = (frmMain.VideoDisplay.Width / 15) / 2 - (w * z) / 2
            Stamp.oVideoWindow.Left = l
            Stamp.oVideoWindow.Top = t
            Stamp.oVideoWindow.Width = w * z
            Stamp.oVideoWindow.Height = h * z
        End If
    End If
End Sub

Public Sub GUI_ShowSkinBrowser()
    frmSkinBrowser.Show
End Sub

Public Sub GUI_SetDisplayMode(Fullscreen As Boolean)
    If Fullscreen = True Then
        If frmMain.WindowState <> vbMaximized Then frmMain.WindowState = vbMaximized
        'SetWindowPos frmMain.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        frmMain.Show
        frmMain.Refresh
    Else
        If frmMain.WindowState <> vbNormal Then frmMain.WindowState = vbNormal
    End If
End Sub

Public Sub GUI_ControlsVisibleMAX()
    Dim a As Integer
    With frmMain
        For a = 0 To .Controls.Count - 1
            Select Case LCase(.Controls(a).Name)
            Case "playlist", "picplaylistbuttons", "vbrdleft"
                .Controls(a).Visible = Stamp.PlaylistVisible
            Case "videodisplay", "piclogo"
                
            Case "vbrdtop", "lbl4", "lbl3", "fullscreencontrols", "conprogress", "convolume", "lblvolume2", "progressmiddle", "progressleft", "progressright", "volumemiddle", "volumeright", "volumeleft"
                .Controls(a).Visible = Stamp.ControlsVisible
            Case "picfullscreentitle"
                
            Case "picbutton"
                Select Case .Controls(a).Index
                Case 0 To 3, 8, 9
                    .Controls(a).Visible = Stamp.ControlsVisible
                Case 10 To 15
                    .Controls(a).Visible = Stamp.PlaylistVisible
                Case Else
                    .Controls(a).Visible = False
                End Select
            Case Else
                .Controls(a).Visible = False
            End Select
        Next
    End With
End Sub

Public Sub GUI_ControlsVisibleMIN()
    Dim a As Integer
    With frmMain
        For a = 0 To .Controls.Count - 1
            Select Case LCase(.Controls(a).Name)
            Case "playlist", "picplaylistbuttons"
                .Controls(a).Visible = Stamp.PlaylistVisible
            Case "videodisplay", "piclogo"
                
            Case "lbl4", "lbl3", "fullscreencontrols"
                .Controls(a).Visible = False
            Case "picfullscreentitle"
                .Controls(a).Visible = False
            Case Else
                .Controls(a).Visible = True
            End Select
        Next
    End With
End Sub

Public Sub GUI_ToggleScreenMode()
    If frmMain.WindowState = vbNormal Then
        GUI_SetDisplayMode True
    Else
        GUI_SetDisplayMode False
    End If
End Sub

Public Function GUI_MouseMoved() As Boolean
    Static MousePositions(9) As MOUSESTAT
    Dim MousePosition As POINTAPI
    Dim a As Integer
    
    For a = 9 To 1 Step -1
        MousePositions(a).X = MousePositions(a - 1).X
        MousePositions(a).Y = MousePositions(a - 1).Y
        MousePositions(a).Button = MousePositions(a - 1).Button
    Next
    
    GetCursorPos MousePosition
    If GetKeyState(vbKeyLButton) < 0 Then MousePositions(0).Button = 1 Else MousePositions(0).Button = 0
    If GetKeyState(vbKeyRButton) < 0 Then MousePositions(0).Button = MousePositions(0).Button + 2
    MousePositions(0).X = MousePosition.X
    MousePositions(0).Y = MousePosition.Y
    
    GUI_MouseMoved = False
    For a = 0 To 8
        If MousePositions(a).Button <> MousePositions(a + 1).Button Then GUI_MouseMoved = True
        If MousePositions(a).Y <> MousePositions(a + 1).Y Then GUI_MouseMoved = True
        If MousePositions(a).X <> MousePositions(a + 1).X Then GUI_MouseMoved = True
    Next
End Function

Public Sub GUI_ShowCursor(Visible As Boolean)
    Select Case Visible
    Case False
        Do Until ShowCursor(False) < 0
        Loop
    Case True
        Do Until ShowCursor(True) > 0
        Loop
    End Select
End Sub

Public Function GUI_ToggleAlwaysOnTop()
    If Stamp.Config.AlwaysOnTop = True Then
        Stamp.Config.AlwaysOnTop = False
    Else
        Stamp.Config.AlwaysOnTop = True
        Beep
    End If
    SetAlwaysOnTop
End Function

Public Sub SetAlwaysOnTop()
    If Stamp.Config.AlwaysOnTop = True Then
        SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    frmRes.mnuAlwaysOnTop.Checked = Stamp.Config.AlwaysOnTop
End Sub

Public Sub ToggleScreenSaverActive()
    If Stamp.Config.PreventScreenSaver = True Then
        Stamp.Config.PreventScreenSaver = False
    Else
        Stamp.Config.PreventScreenSaver = True
        Beep
    End If
    SetScreenSaverPrevent
End Sub

Public Sub SetScreenSaverPrevent()
    If Stamp.Config.PreventScreenSaver = False Then
        Debug.Print "s" & SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 1, ByVal 0, ByVal 0)
        Debug.Print "p" & SystemParametersInfo(SPI_SETPOWEROFFACTIVE, 1, ByVal 0, ByVal 0)
    Else
        Debug.Print "s" & SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 0, ByVal 0, ByVal 0)
        Debug.Print "p" & SystemParametersInfo(SPI_SETPOWEROFFACTIVE, 0, ByVal 0, ByVal 0)
        Dim ret As Long
        SystemParametersInfo SPI_GETPOWEROFFACTIVE, 0, ret, 0
        Debug.Print "state: " & ret
    End If
    frmRes.mnuScreensaverToggle.Checked = Stamp.Config.PreventScreenSaver
End Sub

Public Sub TogglePriority(Priority As Long)
    Dim ret As Boolean
    
    ret = SetPriority(Priority)
    If ret = False Then
        MsgBox "Unable to set priority", vbCritical
        Exit Sub
    End If
    Select Case Priority
    Case IDLE_PRIORITY_CLASS
        frmRes.mnuPriorityHigh.Checked = False
        frmRes.mnuPriorityNormal.Checked = False
        frmRes.mnuPriorityRT.Checked = False
        frmRes.mnuPriorityIdle.Checked = True
    Case NORMAL_PRIORITY_CLASS
        frmRes.mnuPriorityHigh.Checked = False
        frmRes.mnuPriorityNormal.Checked = True
        frmRes.mnuPriorityRT.Checked = False
        frmRes.mnuPriorityIdle.Checked = False
    Case HIGH_PRIORITY_CLASS
        frmRes.mnuPriorityHigh.Checked = True
        frmRes.mnuPriorityNormal.Checked = False
        frmRes.mnuPriorityRT.Checked = False
        frmRes.mnuPriorityIdle.Checked = False
    Case REALTIME_PRIORITY_CLASS
        frmRes.mnuPriorityHigh.Checked = False
        frmRes.mnuPriorityNormal.Checked = False
        frmRes.mnuPriorityRT.Checked = True
        frmRes.mnuPriorityIdle.Checked = False
    End Select
End Sub

Public Function SetPriority(Priority As Long) As Boolean
   SetPriority = SetPriorityClass(GetCurrentProcess, Priority)
End Function

Public Function GetPriority() As Long
   GetPriority = (GetPriorityClass(GetCurrentProcess))
End Function

Public Function GetPriorityName() As String
   Dim lngPriority As Long
   lngPriority = GetPriority
   
   Select Case lngPriority
      Case 256
         GetPriorityName = "Realtime"
      Case 128
         GetPriorityName = "High"
      Case 32
         GetPriorityName = "Normal"
      Case 64
         GetPriorityName = "Idle"
   End Select
End Function

Public Function GetScreenSaverActive() As Boolean
    Dim ret As Long
    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, ret, 0
    If ret = 0 Then GetScreenSaverActive = False Else GetScreenSaverActive = True
End Function

Public Function GetPowerOffActive() As Boolean
    Dim ret As Long
    SystemParametersInfo SPI_GETPOWEROFFACTIVE, 0, ret, 0
    If ret = 0 Then GetPowerOffActive = False Else GetPowerOffActive = True
End Function

Public Sub SetScreenSaverActive(Active As Boolean)
    If Active = True Then
        SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 1, ByVal 0, ByVal 0
    Else
        SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 0, ByVal 0, ByVal 0
    End If
End Sub

Public Sub SetPowerOffActive(Active As Boolean)
    If Active = True Then
        SystemParametersInfo SPI_SETPOWEROFFACTIVE, 1, ByVal 0, ByVal 0
    Else
        SystemParametersInfo SPI_SETPOWEROFFACTIVE, 0, ByVal 0, ByVal 0
    End If
End Sub

Public Sub ResetScreenSaverActive()
    SetScreenSaverActive Stamp.ScreenSaverState
    SetPowerOffActive Stamp.PowerOffState
End Sub

Public Sub BackupScreenSaverActive()
    Stamp.ScreenSaverState = GetScreenSaverActive
    Stamp.PowerOffState = GetPowerOffActive
End Sub

Public Function GUI_MouseOverControl(objControl As Object, HasParent As Boolean, Optional objParentControl As Object) As Boolean
    Dim MousePosition As POINTAPI
    GetCursorPos MousePosition
    With objControl
        GUI_MouseOverControl = False
        If HasParent = True Then
            If MousePosition.X >= (objParentControl.Left + .Left) / Screen.TwipsPerPixelX And MousePosition.X <= (objParentControl.Left + .Left + .Width) / Screen.TwipsPerPixelX Then
                If MousePosition.Y >= (objParentControl.Top + .Top) / Screen.TwipsPerPixelY And MousePosition.Y <= (objParentControl.Top + .Top + .Height) / Screen.TwipsPerPixelY Then
                    GUI_MouseOverControl = True
                End If
            End If
        Else
            If MousePosition.X >= (.Left) / Screen.TwipsPerPixelX And MousePosition.X <= (.Left + .Width) / Screen.TwipsPerPixelX Then
                If MousePosition.Y >= (.Top) / Screen.TwipsPerPixelY And MousePosition.Y <= (.Top + .Height) / Screen.TwipsPerPixelY Then
                    GUI_MouseOverControl = True
                End If
            End If
        End If
    End With
End Function

Public Sub GUI_ResetVideoZoom()
    Stamp.Config.Zoom = 1
    GUI_ResizeVideodisplay
End Sub

Public Sub GUI_ToggleVideoZoom()
    If Stamp.Config.VideoZoomOn = True Then
        Stamp.Config.VideoZoomOn = False
    Else
        Stamp.Config.VideoZoomOn = True
        Beep
    End If
    frmRes.mnuVideoZoomOn.Checked = Stamp.Config.VideoZoomOn
    GUI_ResizeVideodisplay
End Sub
