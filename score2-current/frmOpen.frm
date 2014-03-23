VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WCCS Score v.2.1"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOne 
      Caption         =   "Enter Team Names:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtPts 
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Text            =   "10"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton btnQuit 
         Caption         =   "&Quit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtTeamThree 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2760
         Width           =   4575
      End
      Begin VB.TextBox txtTeamTwo 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&ENTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         TabIndex        =   4
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtTeamOne 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Points Per:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Team One"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Team Two"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Team Three"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnQuit_Click()
    Unload frmOpen
End Sub

Private Sub Command1_Click()
    frmScore.fraOne = frmOpen.txtTeamOne
    frmScore.fraTwo = frmOpen.txtTeamTwo
    frmScore.fraThree = frmOpen.txtTeamThree
    frmScore.Show
    frmOpen.Hide
End Sub
