VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About ""Stamp3"""
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   855
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Homepage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Copyright Stian Østerhaug, 2001"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1 = "Stamp Version " & App.Major & "." & App.Minor
    Label2 = "Build: " & App.Revision
    Picture1.Picture = LoadResPicture("_ICON", vbResIcon)
    If Stamp.Config.AlwaysOnTop = True Then SetWindowPos frmAbout.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SendMessage Command1.hwnd, &HF4&, &H0&, 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub Label4_Click()
    ShellExecute 0, vbNullString, "http://home.online.no/~run-oes", vbNullString, vbNullString, 0
End Sub
