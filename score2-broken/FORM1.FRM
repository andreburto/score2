VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C0C0&
   Caption         =   "ACE Score Keeper"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton resetbtn 
      BackColor       =   &H000000FF&
      Caption         =   "Reset"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton ntrbtn 
      Caption         =   "Enter"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox school3text 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox school2text 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox school1text 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label school3lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      Caption         =   "School Three:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label school2lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      Caption         =   "School Two:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label school1lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      Caption         =   "School One:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "ACE Competiton"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public school1txt As String
Public school2txt As String
Public school3txt As String

Private Sub ntrbtn_Click()
Unload Form1
Load Form2
Form2.Show
End Sub

Private Sub resetbtn_Click()
school1txt = ""
school2txt = ""
school3txt = ""
End Sub

