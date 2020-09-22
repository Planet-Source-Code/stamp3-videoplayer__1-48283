VERSION 5.00
Begin VB.Form frmFileAss 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stamp3 File Associations"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "frmFileAss.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "File types"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.ListBox List1 
         Height          =   2580
         IntegralHeight  =   0   'False
         Left            =   120
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Register"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2880
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmFileAss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a As Integer
    For a = 0 To List1.ListCount - 1
        If List1.Selected(a) = True Then
            nSetFileAssociation Stamp.AppPath & "stamp3.exe", "Stamp3.File", "Stamp3 Media File", "." & List1.List(a), Stamp.AppPath & "stamp3.exe", 2
        End If
    Next
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    Unload Me
End Sub

Private Sub Form_Load()
    SendMessage Command1.hwnd, &HF4&, &H0&, 0&
    LoadFileAssociations
End Sub

Public Sub LoadFileAssociations()
    Dim tmp() As String
    Dim a As Integer
    
    Dim strkey As String
    List1.Clear
    tmp = Split(Mid(FilterExt, 2), " ")
    For a = 0 To UBound(tmp)
        If tmp(a) <> vbNullString Then
            List1.AddItem tmp(a)
        End If
    Next
    For a = 0 To List1.ListCount - 1
        List1.Selected(a) = True
    Next
End Sub
