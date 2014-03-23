VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Instructions"
   ClientHeight    =   6690
   ClientLeft      =   2850
   ClientTop       =   3495
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3952.673
   ScaleMode       =   0  'User
   ScaleWidth      =   6563.231
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin VB.TextBox fnamlabel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Text            =   "First Name"
         Top             =   5400
         Width           =   3135
      End
      Begin VB.TextBox lnamlabel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         TabIndex        =   2
         Text            =   "Last Name"
         Top             =   5400
         Width           =   3255
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   6495
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2880
      TabIndex        =   0
      Top             =   6120
      Width           =   1140
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fname As String
Public lname As String

Private Sub cmdOK_Click()
    Let fname = fnamlabel.Text
    Let lname = lnamlabel.Text
    Load Form1
    Form1.Data1.DatabaseName = CurDir & "\test.mdb" 'Added 3/23/1014
    Unload frmLogin
    Form1.Show
End Sub

Public Sub caps()
Static geoff As Boolean
    If geoff = False Then
        Let geoff = True
        Label1.Caption = "hi"
    End If
    
End Sub
