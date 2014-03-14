VERSION 5.00
Begin VB.Form frmScore 
   BorderStyle     =   0  'None
   Caption         =   "WCCS Score v.2"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnGO 
      Caption         =   "Game Over"
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame fraOne 
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
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton btnUpOne 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnDnOne 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblOne 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape shpOne 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         Top             =   360
         Width           =   15
      End
   End
   Begin VB.Frame fraTwo 
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
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   6135
      Begin VB.CommandButton btnUpTwo 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnDnTwo 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblTwo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape shpTwo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   735
         Left            =   120
         Top             =   360
         Width           =   15
      End
   End
   Begin VB.Frame fraThree 
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   6135
      Begin VB.CommandButton btnUpThree 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnDnThree 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblThree 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape shpThree 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         Height          =   735
         Left            =   120
         Top             =   360
         Width           =   15
      End
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public perscore As Integer

Private Sub btnGO_Click()
    Write_File

    frmScore.Hide
    frmOpen.txtTeamOne.Text = ""
    frmOpen.txtTeamTwo.Text = ""
    frmOpen.txtTeamThree.Text = ""
    
    Do While lblOne.Caption > 0
        btnDnOne_Click
    Loop
    
    Do While lblTwo.Caption > 0
        btnDnTwo_Click
    Loop
    
    Do While lblThree.Caption > 0
        btnDnThree_Click
    Loop
    
    Unload frmScore
    frmOpen.Show
End Sub

Private Sub btnUpOne_Click()
    shpOne.Width = shpOne.Width + (Screen.Width / 100)
    lblOne.Caption = lblOne.Caption + perscore
End Sub

Private Sub btnDnOne_Click()
    If lblOne.Caption > 0 Then
        shpOne.Width = shpOne.Width - (Screen.Width / 100)
        lblOne.Caption = lblOne.Caption - perscore
    End If
End Sub

Private Sub btnUpTwo_Click()
    shpTwo.Width = shpTwo.Width + (Screen.Width / 100)
    lblTwo.Caption = lblTwo.Caption + perscore
End Sub

Private Sub btnDnTwo_Click()
    If lblTwo.Caption > 0 Then
        shpTwo.Width = shpTwo.Width - (Screen.Width / 100)
        lblTwo.Caption = lblTwo.Caption - perscore
    End If
End Sub

Private Sub btnUpThree_Click()
    shpThree.Width = shpThree.Width + (Screen.Width / 100)
    lblThree.Caption = lblThree.Caption + perscore
End Sub

Private Sub btnDnThree_Click()
    If lblThree.Caption > 0 Then
        shpThree.Width = shpThree.Width - (Screen.Width / 100)
        lblThree.Caption = lblThree.Caption - perscore
    End If
End Sub

Private Sub Form_Load()
    btnUpOne.Left = Screen.Width - 750
    btnDnOne.Left = Screen.Width - 750
    btnUpTwo.Left = Screen.Width - 750
    btnDnTwo.Left = Screen.Width - 750
    btnUpThree.Left = Screen.Width - 750
    btnDnThree.Left = Screen.Width - 750
    
    fraOne.Width = Screen.Width - 250
    fraTwo.Width = Screen.Width - 250
    fraThree.Width = Screen.Width - 250
    
    fraOne.Top = 100
    fraOne.Height = (Screen.Height * 0.32) - 200
    
    fraTwo.Top = fraOne.Top + fraOne.Height + 100
    fraTwo.Height = (Screen.Height * 0.32) - 200
        
    fraThree.Top = fraTwo.Top + fraTwo.Height + 100
    fraThree.Height = (Screen.Height * 0.32) - 200

    shpOne.Height = fraOne.Height - 650
    shpTwo.Height = fraTwo.Height - 650
    shpThree.Height = fraThree.Height - 650
    
    btnGO.Left = Screen.Width - (btnGO.Width + 50)
    btnGO.Top = Screen.Height - (btnGO.Height + 50)
    
    lblOne.Top = (shpOne.Top + (shpOne.Height / 2)) - (lblOne.Height * 0.65)
    lblOne.Left = shpOne.Left + 200
    lblTwo.Top = (shpTwo.Top + (shpTwo.Height / 2)) - (lblTwo.Height * 0.65)
    lblTwo.Left = shpTwo.Left + 200
    lblThree.Top = (shpThree.Top + (shpThree.Height / 2)) - (lblThree.Height * 0.65)
    lblThree.Left = shpThree.Left + 200
      
    perscore = CInt(frmOpen.txtPts.Text)
End Sub

Sub Write_File()
    Dim fs, tfile, wfile
    Set fs = CreateObject("Scripting.FileSystemObject")
    tfile = "c:\test.txt"
    
    If fs.FileExists(tfile) <> True Then
        CreateFile
    End If
    
    Set wfile = fs.OpenTextFile(tfile, 8, 0)
    
    wfile.Write fraOne.Caption + "," + lblOne.Caption + "," _
            + fraTwo.Caption + "," + lblTwo.Caption + "," _
            + fraThree.Caption + "," + lblThree.Caption + Chr(10)
    
    wfile.Close
End Sub

Sub CreateFile()
    Dim ns, tfile, nfile
    Set ns = CreateObject("Scripting.FileSystemObject")
    tfile = "c:\test.txt"
    
    Set nfile = ns.CreateTextFile(tfile, True)
    nfile.Close
End Sub
