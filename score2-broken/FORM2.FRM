VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Counter"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   ScaleHeight     =   5640
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton newgme 
      Caption         =   "New Game"
      Height          =   375
      Left            =   9000
      TabIndex        =   12
      Top             =   5160
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   7680
      Max             =   100
      Min             =   -100
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   4200
      Max             =   100
      Min             =   -100
      TabIndex        =   6
      Top             =   4440
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   720
      Max             =   100
      Min             =   -100
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label zero3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label school3score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7560
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label zero2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label school2score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label school1score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label zero1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   4800
      Width           =   135
   End
   Begin VB.Shape schoolscore3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   7080
      Shape           =   5  'Rounded Square
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Shape schoolscore2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   3600
      Shape           =   5  'Rounded Square
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Shape schoolscore1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   120
      Shape           =   5  'Rounded Square
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label school3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label school2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label school1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HScroll1_Change()
school1score = CInt(HScroll1)
End Sub
Private Sub HScroll2_Change()
school2score = CInt(HScroll2)
End Sub
Private Sub HScroll3_Change()
school3score = CInt(HScroll3)
End Sub


Private Sub newgme_Click()
Unload Form2
Load Form1
Form1.Show
End Sub

Private Sub school1_Click()
school1.Caption = Form1.school1txt
End Sub
