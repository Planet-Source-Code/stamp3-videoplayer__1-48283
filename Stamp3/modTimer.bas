Attribute VB_Name = "modTimer"
Option Explicit

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Sub Timer_Proc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    
    If frmMain.WindowState = vbMaximized Then
        GUI_ShowCursor GUI_MouseMoved
    Else
        GUI_ShowCursor True
    End If
        
    If Stamp.mState <> StateStopped Then
        If frmMain.lbl2 <> Stamp.gLastTempState Then Stamp.cCDT = 0
        If Stamp.cCDT = 0 Then
            If Stamp.mState = StatePaused Then
                If Stamp.cPCDT = 0 Then
                    Stamp.mStatus = vbNullString
                    Stamp.cPCDT = 20
                ElseIf Stamp.cPCDT = 10 Then
                    Stamp.mStatus = GUI_FormatTime(Stamp.mPosition)
                End If
                Stamp.cPCDT = Stamp.cPCDT - 1
            Else
                Stamp.mPosition = nMedia_GetPosition
                Stamp.mStatus = GUI_FormatTime(Stamp.mPosition) & "/" & GUI_FormatTime(Stamp.mLength)
            End If
            GUI_SetInfo
        Else
            Stamp.cCDT = Stamp.cCDT - 1
        End If
        GUI_SetProgressPos
        If Stamp.mPosition >= Stamp.mLength And Stamp.mState = StatePlaying Then nPlaylist_Next
    End If
    
    If Stamp.cTDT > 0 Then
        Stamp.cTDT = Stamp.cTDT - 1
        If Stamp.cTDT = 0 Then
            frmMain.picFullscreenTitle.Visible = False
        End If
    End If
End Sub

Public Sub Timer_Set()
    SetTimer frmMain.hwnd, 0, 100, AddressOf Timer_Proc
    SetTimer frmMain.VideoDisplay.hwnd, 0, 10, AddressOf Timer_Proc2
End Sub

Public Sub Timer_Kill()
    KillTimer frmMain.hwnd, 0
    KillTimer frmMain.VideoDisplay.hwnd, 0
End Sub

Public Sub Timer_Proc2(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    Dim curpos As POINTAPI
    Dim Handle As Long
    Dim ret As Long
    
    GetCursorPos curpos
    Handle = WindowFromPoint(curpos.X, curpos.Y)
    ret = GetParent(Handle)
    Do Until ret = 0
        Handle = ret
        ret = GetParent(Handle)
    Loop
    If Handle <> frmMain.hwnd Then Exit Sub
    If GUI_MouseOverControl(frmMain.VideoDisplay, True, frmMain) = True Then If GetKeyState(vbKeyRButton) < 0 Then frmMain.PopupMenu frmRes.mnu
    
    If GetKeyState(40) < 0 Then
        Stamp.Config.Zoom = Stamp.Config.Zoom / 1.005
        Debug.Print Stamp.Config.Zoom
        GUI_ResizeVideodisplay
    ElseIf GetKeyState(38) < 0 Then
        Stamp.Config.Zoom = Stamp.Config.Zoom * 1.005
        Debug.Print Stamp.Config.Zoom
        GUI_ResizeVideodisplay
    End If
    
End Sub
