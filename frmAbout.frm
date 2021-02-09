VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CDeclFix"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3450
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF8080&
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label lblState 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   1140
      Width           =   1995
   End
   Begin VB.Label lblAbout 
      Caption         =   $"frmAbout.frx":058A
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3315
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "ver. " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Sub
