VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stamp3 Options"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   7011
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

