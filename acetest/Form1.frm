VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ACE Tester"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\JAROD\programming\vb\ACE-TEST\ACETEST\ACETEST3\test.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "thetest"
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   6480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Answer"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Answers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   8535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   11
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label answer4 
         BackStyle       =   0  'Transparent
         Caption         =   "answer4"
         DataField       =   "ans4"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   2160
         Width           =   7575
      End
      Begin VB.Label answer3 
         BackStyle       =   0  'Transparent
         Caption         =   "answer3"
         DataField       =   "ans3"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   1560
         Width           =   7575
      End
      Begin VB.Label answer2 
         BackStyle       =   0  'Transparent
         Caption         =   "answer2"
         DataField       =   "ans2"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   7575
      End
      Begin VB.Label answer1 
         BackStyle       =   0  'Transparent
         Caption         =   "answer1"
         DataField       =   "ans1"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         DataField       =   "Number"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         DataField       =   "Question"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   7575
      End
   End
   Begin VB.Label tresult 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label4 
      DataField       =   "correct"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   6015
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalSeconds As Integer
Public tally As Integer
Public george As String
Public form1titlebar As String
Public result As String

Public Sub Roxy()
Static candy As Boolean
    If candy = False Then
        Data1.Recordset.MoveFirst
        Let candy = True
    End If
End Sub

Private Sub answer1_Click()
Option1(0).Value = True
george = "1"
End Sub

Private Sub answer2_Click()
Option1(1).Value = True
george = "2"
End Sub

Private Sub answer3_Click()
Option1(2).Value = True
george = "3"
End Sub

Private Sub answer4_Click()
Option1(3).Value = True
george = "4"
End Sub

Private Sub Command1_Click()
If george = Data1.Recordset.Fields("correct").Value Then
    tally = tally + 1
End If
Label5.Caption = tally
Data1.Recordset.MoveNext
Let TotalSeconds = "0"
If Data1.Recordset.EOF = True Then
result = frmLogin.fnamlabel & " " & frmLogin.lnamlabel & " got " & tally & " out of 50 correct!"
MsgBox result, vbOK, "Game Over"
Open "c:\finalscore.txt" For Output As #1
Write #1, result
Close #1
End
End If
End Sub

Private Sub Form_Activate()
Let form1titlebar = frmLogin.fnamlabel & " " & frmLogin.lnamlabel
Let Form1.Caption = form1titlebar
End Sub


Private Sub Option1_Click(index As Integer)
If Option1(0).Value = True Then george = "1"
If Option1(1).Value = True Then george = "2"
If Option1(2).Value = True Then george = "3"
If Option1(3).Value = True Then george = "4"
End Sub

Private Sub Timer1_Timer()
TotalSeconds = TotalSeconds + 1
tresult.Caption = TotalSeconds
If TotalSeconds > 10 Then
  Let george = "0"
  Let TotalSeconds = 0
  Data1.Recordset.MoveLast
  If Data1.Recordset.EOF = True Then
    result = frmLogin.fnamlabel & " " & frmLogin.lnamlabel & " got " & tally & " out of 50 correct!"
    MsgBox result, vbOK, "Game Over"
    End
  End If
End If
End Sub
