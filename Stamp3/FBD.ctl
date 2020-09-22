VERSION 5.00
Begin VB.UserControl FDB 
   BackColor       =   &H000000FF&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleWidth      =   240
End
Attribute VB_Name = "FDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private DirBuffer() As String, DirCount As Long
Private TreeBuffer() As String, TreeCount As Long

Public Event FileFound(FileName As String, Path As String, MaxPath As Long, HiSize As Long, LoSize As Long)
Public Event FolderFound(FolderName As String, Path As String)
Public Event CurrentDir(Path As String)

Private bCancel As Boolean

Public Sub Search(strPath As String, strFileName As String)
    Dim a As Long
    
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    ReDim DirBuffer(1 To 1)
    DirBuffer(1) = strPath
    DirCount = 1
    
    bCancel = False
    Do Until DirCount = 0
        TreeCount = DirCount
        ReDim TreeBuffer(1 To TreeCount)
        For a = 1 To TreeCount
            TreeBuffer(a) = DirBuffer(a)
        Next
        DirCount = 0
        
        For a = 1 To TreeCount
            RaiseEvent CurrentDir(TreeBuffer(a))
            Call Find(TreeBuffer(a), strFileName)
        Next
        DoEvents
        If bCancel = True Then GoTo EndSub
    Loop
    
    ReDim DirBuffer(0)
    ReDim TreeBuffer(0)

EndSub:

End Sub

Private Function Find(strPath As String, strFileName As String)
    Dim Handle As Long, ret As Long
    Dim FileData As WIN32_FIND_DATA
    Dim tmpFileName As String
    
    Handle = FindFirstFile(strPath & strFileName, FileData)
    
    If Handle = 0 Then GoTo sFinished
    Do
        tmpFileName = ClearNull(FileData.cFileName)
        If tmpFileName <> "." And tmpFileName <> ".." Then
            If Len(tmpFileName) <> 0 Then
                If FileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                    RaiseEvent FolderFound(tmpFileName, strPath)
                Else
                    RaiseEvent FileFound(tmpFileName, strPath, MAX_PATH, FileData.nFileSizeHigh, FileData.nFileSizeLow)
                End If
            End If
        End If
        ret = FindNextFile(Handle, FileData)
        If ret = 0 Then Exit Do
    Loop
    
    FindClose Handle
    Handle = FindFirstFile(strPath & "*", FileData)
    Do
        tmpFileName = ClearNull(FileData.cFileName)
        If FileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
            If tmpFileName <> "." And tmpFileName <> ".." Then
                DirCount = DirCount + 1
                ReDim Preserve DirBuffer(1 To DirCount)
                DirBuffer(DirCount) = strPath & tmpFileName & "\"
            End If
        End If
        
        ret = FindNextFile(Handle, FileData)
        If ret = 0 Then Exit Do
    Loop
    
sFinished:
    FindClose Handle

End Function

Private Function ClearNull(strString As String) As String
    ClearNull = Left(strString, InStr(1, strString, vbNullChar, vbBinaryCompare) - 1)
End Function

Private Sub UserControl_Resize()
    UserControl.Width = 15 * 16
    UserControl.Height = 15 * 16
End Sub

Public Sub Cancel()
    bCancel = True
End Sub
