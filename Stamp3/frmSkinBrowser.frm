VERSION 5.00
Begin VB.Form frmSkinBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stamp3 - Skin Browser"
   ClientHeight    =   6270
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "frmSkinBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Skins"
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         Default         =   -1  'True
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   4200
         Width           =   4695
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   2520
      Pattern         =   "*.s3skn"
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmSkinBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SknRAW As Stamp3SkinRAW

Private Sub Command1_Click()
    If List1.ListIndex = -1 Then Exit Sub
    If List1.ListIndex = 0 Then
        Stamp.Config.CurrentSkinFile = "#"
    Else
        Stamp.Config.CurrentSkinFile = SkinFiles(List1.ListIndex)
    End If
    Skin_Extract
    Unload Me
End Sub

Private Sub Form_Load()
    LoadSkinList
    SendMessage Command1.hwnd, &HF4&, &H0&, 0&
    If Stamp.Config.AlwaysOnTop = True Then SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub List1_Click()
    Dim FileHandle As Integer

    If List1.ListIndex = -1 Then Exit Sub
    
    If List1.ListIndex = 0 Then
        Text1 = "Skin name: Default Skin" & vbCrLf & "Author: Saitn" & vbCrLf & vbCrLf & "Default Stamp3 Skin" & vbCrLf & "-------------------" & vbCrLf & "Made by Saitn"
        Exit Sub
    End If
    
    FileHandle = FreeFile
    Open SkinFiles(List1.ListIndex) For Binary As FileHandle
    Get FileHandle, 1, SknRAW
    Text1 = "Skin name: " & SknRAW.Name & vbCrLf & "Author: " & SknRAW.Author & vbCrLf & vbCrLf & SknRAW.CommentLine1 & vbCrLf & SknRAW.CommentLine2 & vbCrLf & SknRAW.CommentLine3
    Close FileHandle
    
End Sub

Public Sub LoadSkinList()
    Dim FileHandle As Integer
    Dim a As Integer
    
    List1.Clear
    List1.AddItem "[Default Skin]"
    If Stamp.Config.CurrentSkinFile = "#" Then List1.ListIndex = 0
    
    If Dir(Stamp.AppPath & "Skins", vbDirectory) = vbNullString Then Exit Sub
    File1 = Stamp.AppPath & "Skins"
    
    Erase SkinFiles
    Erase SkinNames
    Skins = 0
    
    For a = 0 To File1.ListCount - 1
        FileHandle = FreeFile
        Open Stamp.AppPath & "Skins\" & File1.List(a) For Binary As FileHandle
        Get FileHandle, 1, SknRAW
        Close FileHandle
        If SknRAW.Signature = "Stamp3Skin" Then
            List1.AddItem SknRAW.Name
            Skins = Skins + 1
            ReDim Preserve SkinFiles(1 To Skins)
            ReDim Preserve SkinNames(1 To Skins)
            SkinFiles(Skins) = Stamp.AppPath & "Skins\" & File1.List(a)
            SkinNames(Skins) = SknRAW.Name
            If SkinFiles(Skins) = Stamp.Config.CurrentSkinFile Then List1.ListIndex = Skins
        End If
    Next
End Sub
