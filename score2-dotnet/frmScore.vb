Option Strict Off
Option Explicit On
Friend Class frmScore
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
    Public WithEvents btnGO As System.Windows.Forms.Button
	Public WithEvents btnUpOne As System.Windows.Forms.Button
	Public WithEvents btnDnOne As System.Windows.Forms.Button
	Public WithEvents lblOne As System.Windows.Forms.Label
	Public WithEvents shpOne As System.Windows.Forms.Label
	Public WithEvents fraOne As System.Windows.Forms.GroupBox
	Public WithEvents btnUpTwo As System.Windows.Forms.Button
	Public WithEvents btnDnTwo As System.Windows.Forms.Button
	Public WithEvents lblTwo As System.Windows.Forms.Label
	Public WithEvents shpTwo As System.Windows.Forms.Label
	Public WithEvents fraTwo As System.Windows.Forms.GroupBox
	Public WithEvents btnUpThree As System.Windows.Forms.Button
	Public WithEvents btnDnThree As System.Windows.Forms.Button
	Public WithEvents lblThree As System.Windows.Forms.Label
	Public WithEvents shpThree As System.Windows.Forms.Label
	Public WithEvents fraThree As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnGO = New System.Windows.Forms.Button
        Me.fraOne = New System.Windows.Forms.GroupBox
        Me.btnUpOne = New System.Windows.Forms.Button
        Me.btnDnOne = New System.Windows.Forms.Button
        Me.lblOne = New System.Windows.Forms.Label
        Me.shpOne = New System.Windows.Forms.Label
        Me.fraTwo = New System.Windows.Forms.GroupBox
        Me.btnUpTwo = New System.Windows.Forms.Button
        Me.btnDnTwo = New System.Windows.Forms.Button
        Me.lblTwo = New System.Windows.Forms.Label
        Me.shpTwo = New System.Windows.Forms.Label
        Me.fraThree = New System.Windows.Forms.GroupBox
        Me.btnUpThree = New System.Windows.Forms.Button
        Me.btnDnThree = New System.Windows.Forms.Button
        Me.lblThree = New System.Windows.Forms.Label
        Me.shpThree = New System.Windows.Forms.Label
        Me.fraOne.SuspendLayout()
        Me.fraTwo.SuspendLayout()
        Me.fraThree.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnGO
        '
        Me.btnGO.BackColor = System.Drawing.SystemColors.Control
        Me.btnGO.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGO.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGO.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGO.Location = New System.Drawing.Point(352, 304)
        Me.btnGO.Name = "btnGO"
        Me.btnGO.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGO.Size = New System.Drawing.Size(65, 21)
        Me.btnGO.TabIndex = 7
        Me.btnGO.Text = "Game Over"
        '
        'fraOne
        '
        Me.fraOne.BackColor = System.Drawing.SystemColors.Control
        Me.fraOne.Controls.Add(Me.btnUpOne)
        Me.fraOne.Controls.Add(Me.btnDnOne)
        Me.fraOne.Controls.Add(Me.lblOne)
        Me.fraOne.Controls.Add(Me.shpOne)
        Me.fraOne.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraOne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraOne.Location = New System.Drawing.Point(8, 8)
        Me.fraOne.Name = "fraOne"
        Me.fraOne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraOne.Size = New System.Drawing.Size(409, 81)
        Me.fraOne.TabIndex = 9
        Me.fraOne.TabStop = False
        Me.fraOne.Text = "Team One"
        '
        'btnUpOne
        '
        Me.btnUpOne.BackColor = System.Drawing.SystemColors.Control
        Me.btnUpOne.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnUpOne.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpOne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUpOne.Location = New System.Drawing.Point(376, 16)
        Me.btnUpOne.Name = "btnUpOne"
        Me.btnUpOne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnUpOne.Size = New System.Drawing.Size(25, 25)
        Me.btnUpOne.TabIndex = 1
        Me.btnUpOne.Text = "+"
        '
        'btnDnOne
        '
        Me.btnDnOne.BackColor = System.Drawing.SystemColors.Control
        Me.btnDnOne.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDnOne.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDnOne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDnOne.Location = New System.Drawing.Point(376, 48)
        Me.btnDnOne.Name = "btnDnOne"
        Me.btnDnOne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDnOne.Size = New System.Drawing.Size(25, 25)
        Me.btnDnOne.TabIndex = 2
        Me.btnDnOne.Text = "-"
        '
        'lblOne
        '
        Me.lblOne.BackColor = System.Drawing.Color.Transparent
        Me.lblOne.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOne.Font = New System.Drawing.Font("Arial", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOne.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOne.Location = New System.Drawing.Point(0, 16)
        Me.lblOne.Name = "lblOne"
        Me.lblOne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOne.Size = New System.Drawing.Size(113, 57)
        Me.lblOne.TabIndex = 10
        Me.lblOne.Text = "0"
        Me.lblOne.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'shpOne
        '
        Me.shpOne.BackColor = System.Drawing.Color.Red
        Me.shpOne.Location = New System.Drawing.Point(8, 24)
        Me.shpOne.Name = "shpOne"
        Me.shpOne.Size = New System.Drawing.Size(1, 49)
        Me.shpOne.TabIndex = 11
        '
        'fraTwo
        '
        Me.fraTwo.BackColor = System.Drawing.SystemColors.Control
        Me.fraTwo.Controls.Add(Me.btnUpTwo)
        Me.fraTwo.Controls.Add(Me.btnDnTwo)
        Me.fraTwo.Controls.Add(Me.lblTwo)
        Me.fraTwo.Controls.Add(Me.shpTwo)
        Me.fraTwo.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraTwo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTwo.Location = New System.Drawing.Point(8, 104)
        Me.fraTwo.Name = "fraTwo"
        Me.fraTwo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTwo.Size = New System.Drawing.Size(409, 81)
        Me.fraTwo.TabIndex = 8
        Me.fraTwo.TabStop = False
        Me.fraTwo.Text = "Team Two"
        '
        'btnUpTwo
        '
        Me.btnUpTwo.BackColor = System.Drawing.SystemColors.Control
        Me.btnUpTwo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnUpTwo.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpTwo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUpTwo.Location = New System.Drawing.Point(376, 16)
        Me.btnUpTwo.Name = "btnUpTwo"
        Me.btnUpTwo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnUpTwo.Size = New System.Drawing.Size(25, 25)
        Me.btnUpTwo.TabIndex = 3
        Me.btnUpTwo.Text = "+"
        '
        'btnDnTwo
        '
        Me.btnDnTwo.BackColor = System.Drawing.SystemColors.Control
        Me.btnDnTwo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDnTwo.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDnTwo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDnTwo.Location = New System.Drawing.Point(376, 48)
        Me.btnDnTwo.Name = "btnDnTwo"
        Me.btnDnTwo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDnTwo.Size = New System.Drawing.Size(25, 25)
        Me.btnDnTwo.TabIndex = 4
        Me.btnDnTwo.Text = "-"
        '
        'lblTwo
        '
        Me.lblTwo.BackColor = System.Drawing.Color.Transparent
        Me.lblTwo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTwo.Font = New System.Drawing.Font("Arial", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTwo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTwo.Location = New System.Drawing.Point(0, 16)
        Me.lblTwo.Name = "lblTwo"
        Me.lblTwo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTwo.Size = New System.Drawing.Size(113, 57)
        Me.lblTwo.TabIndex = 11
        Me.lblTwo.Text = "0"
        Me.lblTwo.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'shpTwo
        '
        Me.shpTwo.BackColor = System.Drawing.Color.Yellow
        Me.shpTwo.Location = New System.Drawing.Point(8, 24)
        Me.shpTwo.Name = "shpTwo"
        Me.shpTwo.Size = New System.Drawing.Size(1, 49)
        Me.shpTwo.TabIndex = 12
        '
        'fraThree
        '
        Me.fraThree.BackColor = System.Drawing.SystemColors.Control
        Me.fraThree.Controls.Add(Me.btnUpThree)
        Me.fraThree.Controls.Add(Me.btnDnThree)
        Me.fraThree.Controls.Add(Me.lblThree)
        Me.fraThree.Controls.Add(Me.shpThree)
        Me.fraThree.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraThree.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraThree.Location = New System.Drawing.Point(8, 200)
        Me.fraThree.Name = "fraThree"
        Me.fraThree.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraThree.Size = New System.Drawing.Size(409, 81)
        Me.fraThree.TabIndex = 0
        Me.fraThree.TabStop = False
        Me.fraThree.Text = "Team Three"
        '
        'btnUpThree
        '
        Me.btnUpThree.BackColor = System.Drawing.SystemColors.Control
        Me.btnUpThree.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnUpThree.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpThree.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUpThree.Location = New System.Drawing.Point(376, 16)
        Me.btnUpThree.Name = "btnUpThree"
        Me.btnUpThree.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnUpThree.Size = New System.Drawing.Size(25, 25)
        Me.btnUpThree.TabIndex = 5
        Me.btnUpThree.Text = "+"
        '
        'btnDnThree
        '
        Me.btnDnThree.BackColor = System.Drawing.SystemColors.Control
        Me.btnDnThree.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDnThree.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDnThree.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDnThree.Location = New System.Drawing.Point(376, 48)
        Me.btnDnThree.Name = "btnDnThree"
        Me.btnDnThree.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDnThree.Size = New System.Drawing.Size(25, 25)
        Me.btnDnThree.TabIndex = 6
        Me.btnDnThree.Text = "-"
        '
        'lblThree
        '
        Me.lblThree.BackColor = System.Drawing.Color.Transparent
        Me.lblThree.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblThree.Font = New System.Drawing.Font("Arial", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblThree.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblThree.Location = New System.Drawing.Point(0, 16)
        Me.lblThree.Name = "lblThree"
        Me.lblThree.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblThree.Size = New System.Drawing.Size(113, 57)
        Me.lblThree.TabIndex = 12
        Me.lblThree.Text = "0"
        Me.lblThree.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'shpThree
        '
        Me.shpThree.BackColor = System.Drawing.Color.Blue
        Me.shpThree.Location = New System.Drawing.Point(8, 24)
        Me.shpThree.Name = "shpThree"
        Me.shpThree.Size = New System.Drawing.Size(1, 49)
        Me.shpThree.TabIndex = 13
        '
        'frmScore
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(421, 330)
        Me.Controls.Add(Me.btnGO)
        Me.Controls.Add(Me.fraOne)
        Me.Controls.Add(Me.fraTwo)
        Me.Controls.Add(Me.fraThree)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmScore"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WCCS Score v.2"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.fraOne.ResumeLayout(False)
        Me.fraTwo.ResumeLayout(False)
        Me.fraThree.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmScore
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmScore
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmScore()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Public perscore As Short
	
	Private Sub btnGO_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnGO.Click
		Write_File()
		
		frmScore.DefInstance.Hide()
		frmOpen.DefInstance.txtTeamOne.Text = ""
		frmOpen.DefInstance.txtTeamTwo.Text = ""
		frmOpen.DefInstance.txtTeamThree.Text = ""
		
		Do While CDbl(lblOne.Text) > 0
			btnDnOne_Click(btnDnOne, New System.EventArgs())
		Loop 
		
		Do While CDbl(lblTwo.Text) > 0
			btnDnTwo_Click(btnDnTwo, New System.EventArgs())
		Loop 
		
		Do While CDbl(lblThree.Text) > 0
			btnDnThree_Click(btnDnThree, New System.EventArgs())
		Loop 
		
		frmScore.DefInstance.Close()
		frmOpen.DefInstance.Show()
	End Sub
	
	Private Sub btnUpOne_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnUpOne.Click
		shpOne.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpOne.Width) + (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 100))
		lblOne.Text = CStr(CDbl(lblOne.Text) + perscore)
	End Sub
	
	Private Sub btnDnOne_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDnOne.Click
		If CDbl(lblOne.Text) > 0 Then
			shpOne.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpOne.Width) - (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 100))
			lblOne.Text = CStr(CDbl(lblOne.Text) - perscore)
		End If
	End Sub
	
	Private Sub btnUpTwo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnUpTwo.Click
		shpTwo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpTwo.Width) + (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 100))
		lblTwo.Text = CStr(CDbl(lblTwo.Text) + perscore)
	End Sub
	
	Private Sub btnDnTwo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDnTwo.Click
		If CDbl(lblTwo.Text) > 0 Then
			shpTwo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpTwo.Width) - (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 100))
			lblTwo.Text = CStr(CDbl(lblTwo.Text) - perscore)
		End If
	End Sub
	
	Private Sub btnUpThree_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnUpThree.Click
		shpThree.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpThree.Width) + (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 100))
		lblThree.Text = CStr(CDbl(lblThree.Text) + perscore)
	End Sub
	
	Private Sub btnDnThree_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDnThree.Click
		If CDbl(lblThree.Text) > 0 Then
			shpThree.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpThree.Width) - (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 100))
			lblThree.Text = CStr(CDbl(lblThree.Text) - perscore)
		End If
	End Sub
	
	Private Sub frmScore_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		btnUpOne.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 750)
		btnDnOne.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 750)
		btnUpTwo.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 750)
		btnDnTwo.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 750)
		btnUpThree.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 750)
		btnDnThree.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 750)
		
		fraOne.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 250)
		fraTwo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 250)
		fraThree.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 250)
		
		fraOne.Top = VB6.TwipsToPixelsY(100)
		fraOne.Height = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.32) - 200)
		
		fraTwo.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(fraOne.Top) + VB6.PixelsToTwipsY(fraOne.Height) + 100)
		fraTwo.Height = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.32) - 200)
		
		fraThree.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(fraTwo.Top) + VB6.PixelsToTwipsY(fraTwo.Height) + 100)
		fraThree.Height = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.32) - 200)
		
		shpOne.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(fraOne.Height) - 650)
		shpTwo.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(fraTwo.Height) - 650)
		shpThree.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(fraThree.Height) - 650)
		
		btnGO.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - (VB6.PixelsToTwipsX(btnGO.Width) + 50))
		btnGO.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(btnGO.Height) + 50))
		
		lblOne.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(shpOne.Top) + (VB6.PixelsToTwipsY(shpOne.Height) / 2)) - (VB6.PixelsToTwipsY(lblOne.Height) * 0.65))
		lblOne.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpOne.Left) + 200)
		lblTwo.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(shpTwo.Top) + (VB6.PixelsToTwipsY(shpTwo.Height) / 2)) - (VB6.PixelsToTwipsY(lblTwo.Height) * 0.65))
		lblTwo.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpTwo.Left) + 200)
		lblThree.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(shpThree.Top) + (VB6.PixelsToTwipsY(shpThree.Height) / 2)) - (VB6.PixelsToTwipsY(lblThree.Height) * 0.65))
		lblThree.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpThree.Left) + 200)
		
		perscore = CShort(frmOpen.DefInstance.txtPts.Text)
	End Sub
	
	Sub Write_File()
		Dim tfile, fs, wfile As Object
		fs = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object tfile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		tfile = "c:\test.txt"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.FileExists. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If fs.FileExists(tfile) <> True Then
			CreateFile()
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.OpenTextFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		wfile = fs.OpenTextFile(tfile, 8, 0)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object wfile.Write. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		wfile.Write(fraOne.Text & "," & lblOne.Text & "," & fraTwo.Text & "," & lblTwo.Text & "," & fraThree.Text & "," & lblThree.Text & Chr(10))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object wfile.Close. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		wfile.Close()
	End Sub
	
	Sub CreateFile()
		Dim tfile, ns, nfile As Object
		ns = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object tfile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		tfile = "c:\test.txt"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ns.CreateTextFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		nfile = ns.CreateTextFile(tfile, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object nfile.Close. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		nfile.Close()
	End Sub
End Class