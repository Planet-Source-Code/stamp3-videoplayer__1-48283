Attribute VB_Name = "modComDlg32"
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_EXPLORER = &H80000
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_NODEREFERENCELINKS = &H100000

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Type BROWSEINFO
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Function ShowMultipleOpen(Filter As String, hwnd As Long, Title As String, ByRef Filelist() As String) As Long
    Dim OpenFile As OPENFILENAME
    Dim Directory As String
    Dim ret As Long
    Dim tmp As Variant
    Dim a As Long
    
    Erase Filelist
    
    With OpenFile
        .lStructSize = Len(OpenFile)
        .hWndOwner = hwnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = String(25600, 0)
        .nMaxFile = 25600
        .lpstrTitle = "Stamp3 - Add files to playlist"
        .flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER
    End With
    
    ret = GetOpenFileName(OpenFile)
    If ret = 0 Then
        ShowMultipleOpen = 0
        Exit Function
    End If
    
    Directory = Left(OpenFile.lpstrFile, OpenFile.nFileOffset - 1)
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    If Mid(OpenFile.lpstrFile, OpenFile.nFileOffset, 1) = vbNullChar Then
        tmp = Mid(OpenFile.lpstrFile, OpenFile.nFileOffset + 1, InStr(1, OpenFile.lpstrFile, vbNullChar & vbNullChar) - 1 - OpenFile.nFileOffset)
        Filelist = Split(tmp, vbNullChar)
        ShowMultipleOpen = UBound(Filelist) + 1
        For a = 0 To UBound(Filelist)
            Filelist(a) = Directory & Filelist(a)
        Next
    Else
        ReDim Filelist(0)
        ShowMultipleOpen = 1
        tmp = InStr(1, OpenFile.lpstrFile, vbNullChar)
        If tmp = 0 Then
            Filelist(0) = OpenFile.lpstrFile
        Else
            Filelist(0) = Left(OpenFile.lpstrFile, tmp - 1)
        End If
    End If
    
End Function

Public Function ShowOpen(Filter As String, hwnd As Long, Title As String) As String
    Dim OpenFile As OPENFILENAME
    Dim tmp As Variant
    Dim ret As Long
    
    With OpenFile
        .lStructSize = Len(OpenFile)
        .hWndOwner = hwnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = String(4096, 0)
        .nMaxFile = 4096
        .lpstrTitle = "Stamp3 - Open file"
        .flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER
    End With
    
    ret = GetOpenFileName(OpenFile)
    If ret = 0 Then
        ShowOpen = vbNullString
        Exit Function
    End If
    
    tmp = InStr(1, OpenFile.lpstrFile, vbNullChar)
    If tmp = 0 Then
        ShowOpen = OpenFile.lpstrFile
    Else
        ShowOpen = Left(OpenFile.lpstrFile, tmp - 1)
    End If
    
End Function

Public Function ShowSave(Filter As String, hwnd As Long, Title As String) As String
    Dim OpenFile As OPENFILENAME
    Dim tmp As Variant
    Dim ret As Long
    
    With OpenFile
        .lStructSize = Len(OpenFile)
        .hWndOwner = hwnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = String(4096, 0)
        .nMaxFile = 4096
        .lpstrTitle = Title
        .lpstrDefExt = "stn"
        .flags = OFN_EXPLORER Or OFN_OVERWRITEPROMPT Or OFN_NODEREFERENCELINKS
    End With
    
    ret = GetSaveFileName(OpenFile)
    If ret = 0 Then
        ShowSave = vbNullString
        Exit Function
    End If
    
    tmp = InStr(1, OpenFile.lpstrFile, vbNullChar)
    If tmp = 0 Then
        ShowSave = OpenFile.lpstrFile
    Else
        ShowSave = Left(OpenFile.lpstrFile, tmp - 1)
    End If
    
End Function

Public Function BrowseForFolder(hwnd As Long) As String
    Dim Browse As BROWSEINFO
    Dim ret As Long
    Dim tmp As String
    
    Browse.hWndOwner = hwnd
    ret = SHBrowseForFolder(Browse)
    If ret = 0 Then
        BrowseForFolder = vbNullString
    Else
        BrowseForFolder = String(1024, 0)
        SHGetPathFromIDList ret, BrowseForFolder
        tmp = InStr(1, BrowseForFolder, vbNullChar)
        If tmp <> 0 Then BrowseForFolder = Left(BrowseForFolder, tmp - 1)
    End If
End Function
