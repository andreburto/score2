Option Strict Off
Option Explicit On
Friend Class frmOpen
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
    Public WithEvents txtPts As System.Windows.Forms.TextBox
	Public WithEvents btnQuit As System.Windows.Forms.Button
	Public WithEvents txtTeamThree As System.Windows.Forms.TextBox
	Public WithEvents txtTeamTwo As System.Windows.Forms.TextBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents txtTeamOne As System.Windows.Forms.TextBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents frmOne As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.frmOne = New System.Windows.Forms.GroupBox
        Me.txtPts = New System.Windows.Forms.TextBox
        Me.btnQuit = New System.Windows.Forms.Button
        Me.txtTeamThree = New System.Windows.Forms.TextBox
        Me.txtTeamTwo = New System.Windows.Forms.TextBox
        Me.Command1 = New System.Windows.Forms.Button
        Me.txtTeamOne = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.frmOne.SuspendLayout()
        Me.SuspendLayout()
        '
        'frmOne
        '
        Me.frmOne.BackColor = System.Drawing.SystemColors.Control
        Me.frmOne.Controls.Add(Me.txtPts)
        Me.frmOne.Controls.Add(Me.btnQuit)
        Me.frmOne.Controls.Add(Me.txtTeamThree)
        Me.frmOne.Controls.Add(Me.txtTeamTwo)
        Me.frmOne.Controls.Add(Me.Command1)
        Me.frmOne.Controls.Add(Me.txtTeamOne)
        Me.frmOne.Controls.Add(Me.Label4)
        Me.frmOne.Controls.Add(Me.Label3)
        Me.frmOne.Controls.Add(Me.Label2)
        Me.frmOne.Controls.Add(Me.Label1)
        Me.frmOne.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmOne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmOne.Location = New System.Drawing.Point(8, 8)
        Me.frmOne.Name = "frmOne"
        Me.frmOne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmOne.Size = New System.Drawing.Size(433, 225)
        Me.frmOne.TabIndex = 0
        Me.frmOne.TabStop = False
        Me.frmOne.Text = "Enter Team Names:"
        '
        'txtPts
        '
        Me.txtPts.AcceptsReturn = True
        Me.txtPts.AutoSize = False
        Me.txtPts.BackColor = System.Drawing.SystemColors.Window
        Me.txtPts.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPts.Location = New System.Drawing.Point(336, 88)
        Me.txtPts.MaxLength = 0
        Me.txtPts.Name = "txtPts"
        Me.txtPts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPts.Size = New System.Drawing.Size(81, 19)
        Me.txtPts.TabIndex = 10
        Me.txtPts.Text = "10"
        '
        'btnQuit
        '
        Me.btnQuit.BackColor = System.Drawing.SystemColors.Control
        Me.btnQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnQuit.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnQuit.Location = New System.Drawing.Point(336, 120)
        Me.btnQuit.Name = "btnQuit"
        Me.btnQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnQuit.Size = New System.Drawing.Size(81, 41)
        Me.btnQuit.TabIndex = 5
        Me.btnQuit.Text = "&Quit"
        '
        'txtTeamThree
        '
        Me.txtTeamThree.AcceptsReturn = True
        Me.txtTeamThree.AutoSize = False
        Me.txtTeamThree.BackColor = System.Drawing.SystemColors.Window
        Me.txtTeamThree.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTeamThree.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTeamThree.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTeamThree.Location = New System.Drawing.Point(16, 184)
        Me.txtTeamThree.MaxLength = 0
        Me.txtTeamThree.Name = "txtTeamThree"
        Me.txtTeamThree.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTeamThree.Size = New System.Drawing.Size(305, 25)
        Me.txtTeamThree.TabIndex = 3
        Me.txtTeamThree.Text = ""
        '
        'txtTeamTwo
        '
        Me.txtTeamTwo.AcceptsReturn = True
        Me.txtTeamTwo.AutoSize = False
        Me.txtTeamTwo.BackColor = System.Drawing.SystemColors.Window
        Me.txtTeamTwo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTeamTwo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTeamTwo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTeamTwo.Location = New System.Drawing.Point(16, 120)
        Me.txtTeamTwo.MaxLength = 0
        Me.txtTeamTwo.Name = "txtTeamTwo"
        Me.txtTeamTwo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTeamTwo.Size = New System.Drawing.Size(305, 25)
        Me.txtTeamTwo.TabIndex = 2
        Me.txtTeamTwo.Text = ""
        '
        'Command1
        '
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(336, 168)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(81, 41)
        Me.Command1.TabIndex = 4
        Me.Command1.Text = "&ENTER"
        '
        'txtTeamOne
        '
        Me.txtTeamOne.AcceptsReturn = True
        Me.txtTeamOne.AutoSize = False
        Me.txtTeamOne.BackColor = System.Drawing.SystemColors.Window
        Me.txtTeamOne.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTeamOne.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTeamOne.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTeamOne.Location = New System.Drawing.Point(16, 56)
        Me.txtTeamOne.MaxLength = 0
        Me.txtTeamOne.Name = "txtTeamOne"
        Me.txtTeamOne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTeamOne.Size = New System.Drawing.Size(305, 25)
        Me.txtTeamOne.TabIndex = 1
        Me.txtTeamOne.Text = ""
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(336, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(81, 17)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Points Per:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(137, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Team One"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(137, 17)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Team Two"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 160)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(137, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Team Three"
        '
        'frmOpen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(450, 249)
        Me.Controls.Add(Me.frmOne)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmOpen"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WCCS Score v.2.1"
        Me.frmOne.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmOpen
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmOpen
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmOpen()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Private Sub btnQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnQuit.Click
		frmOpen.DefInstance.Close()
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		frmScore.DefInstance.fraOne.Text = frmOpen.DefInstance.txtTeamOne.Text
		frmScore.DefInstance.fraTwo.Text = frmOpen.DefInstance.txtTeamTwo.Text
		frmScore.DefInstance.fraThree.Text = frmOpen.DefInstance.txtTeamThree.Text
		frmScore.DefInstance.Show()
		frmOpen.DefInstance.Hide()
	End Sub
End Class