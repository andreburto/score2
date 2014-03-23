VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ACE Score Keeper ver 2.0 - by: Static-Pulse"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9405
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton blue3 
      BackColor       =   &H00FF8080&
      Caption         =   "-1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton blue2 
      BackColor       =   &H00FF8080&
      Caption         =   "+2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton blue1 
      BackColor       =   &H00FF8080&
      Caption         =   "+1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton green3 
      BackColor       =   &H0080FF80&
      Caption         =   "-1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton green2 
      BackColor       =   &H0080FF80&
      Caption         =   "+2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton green1 
      BackColor       =   &H0080FF80&
      Caption         =   "+1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton entername 
      Caption         =   "E N T E R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox bluename 
      Height          =   285
      Left            =   2400
      TabIndex        =   14
      Text            =   "Blue Team Name"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox greenname 
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Text            =   "Green Team Name"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox redname 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Text            =   "Red Team Name"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton red3 
      BackColor       =   &H008080FF&
      Caption         =   "-1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton red2 
      BackColor       =   &H008080FF&
      Caption         =   "+2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton red1 
      BackColor       =   &H008080FF&
      Caption         =   "+1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1680
      Left            =   7200
      Picture         =   "score.frx":0000
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label blueteamname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blue Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   6600
      TabIndex        =   11
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label greenteamname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Green Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   3600
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label redteamname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Red Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   600
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label blueLabel1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6240
      TabIndex        =   8
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label greenLabel1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3240
      TabIndex        =   7
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label redLabel1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   240
      Shape           =   2  'Oval
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   3240
      Shape           =   2  'Oval
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label rednamelabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Red Team Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label greennamelabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Green Team Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label bluenamelabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Blue Team Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub entername_Click()

redteamname = redname
greenteamname = greenname
blueteamname = bluename

rednamelabel.Visible = False
greennamelabel.Visible = False
bluenamelabel.Visible = False
redname.Visible = False
greenname.Visible = False
bluename.Visible = False
entername.Visible = False

red1.Enabled = True
green1.Enabled = True
blue1.Enabled = True
red2.Enabled = True
green2.Enabled = True
blue2.Enabled = True
red3.Enabled = True
green3.Enabled = True
blue3.Enabled = True

End Sub

Private Sub red1_Click()
redLabel1.Caption = redLabel1.Caption + 1
End Sub

Private Sub green1_Click()
greenLabel1.Caption = greenLabel1.Caption + 1
End Sub

Private Sub blue1_Click()
blueLabel1.Caption = blueLabel1.Caption + 1
End Sub

Private Sub red2_Click()
redLabel1.Caption = redLabel1.Caption + 2
End Sub

Private Sub green2_Click()
greenLabel1.Caption = greenLabel1.Caption + 2
End Sub

Private Sub blue2_Click()
blueLabel1.Caption = blueLabel1.Caption + 2
End Sub

Private Sub red3_Click()
redLabel1.Caption = redLabel1.Caption - 1
End Sub

Private Sub green3_Click()
greenLabel1.Caption = greenLabel1.Caption - 1
End Sub

Private Sub blue3_Click()
blueLabel1.Caption = blueLabel1.Caption - 1
End Sub


