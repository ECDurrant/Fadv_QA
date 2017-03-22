<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class QaSetup
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.lblQAauditor = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cboContactType = New System.Windows.Forms.ComboBox()
        Me.cboAgentName = New System.Windows.Forms.ComboBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtAgentTeam = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtJIRAbox = New System.Windows.Forms.TextBox()
        Me.lblUserID = New System.Windows.Forms.Label()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtCompany = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtContactName = New System.Windows.Forms.TextBox()
        Me.txtAccountNum = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtContactEmail = New System.Windows.Forms.TextBox()
        Me.txtOrderID = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtContactPhone = New System.Windows.Forms.TextBox()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.txtContactID = New System.Windows.Forms.TextBox()
        Me.txtSRNumber = New System.Windows.Forms.TextBox()
        Me.btnHide = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.lblMDrive = New System.Windows.Forms.Label()
        Me.lblSCRN = New System.Windows.Forms.Label()
        Me.GroupBox8.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.lblQAauditor)
        Me.GroupBox8.Controls.Add(Me.Label11)
        Me.GroupBox8.Controls.Add(Me.cboContactType)
        Me.GroupBox8.Controls.Add(Me.cboAgentName)
        Me.GroupBox8.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox8.Location = New System.Drawing.Point(12, 43)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(347, 167)
        Me.GroupBox8.TabIndex = 3
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Agent Information"
        '
        'lblQAauditor
        '
        Me.lblQAauditor.AutoSize = True
        Me.lblQAauditor.Font = New System.Drawing.Font("Lucida Bright", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQAauditor.ForeColor = System.Drawing.Color.Red
        Me.lblQAauditor.Location = New System.Drawing.Point(171, 24)
        Me.lblQAauditor.Name = "lblQAauditor"
        Me.lblQAauditor.Size = New System.Drawing.Size(58, 15)
        Me.lblQAauditor.TabIndex = 25
        Me.lblQAauditor.Text = "Label12"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(4, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(170, 15)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "QA Auditor/Supervisor:"
        '
        'cboContactType
        '
        Me.cboContactType.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.cboContactType.FormattingEnabled = True
        Me.cboContactType.Items.AddRange(New Object() {"Call", "Chat", "Email", "Level 2 - Call", "Level 2 - Email", "Resident - Call", "Resident - Email", "Consumer Advocacy - Call"})
        Me.cboContactType.Location = New System.Drawing.Point(12, 100)
        Me.cboContactType.Name = "cboContactType"
        Me.cboContactType.Size = New System.Drawing.Size(192, 23)
        Me.cboContactType.TabIndex = 3
        Me.cboContactType.Text = "Contact Type"
        '
        'cboAgentName
        '
        Me.cboAgentName.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.cboAgentName.FormattingEnabled = True
        Me.cboAgentName.Location = New System.Drawing.Point(12, 56)
        Me.cboAgentName.Name = "cboAgentName"
        Me.cboAgentName.Size = New System.Drawing.Size(192, 23)
        Me.cboAgentName.TabIndex = 2
        Me.cboAgentName.Text = "Agent Name"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Font = New System.Drawing.Font("Lucida Bright", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(895, 24)
        Me.MenuStrip1.TabIndex = 16
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(42, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.txtAgentTeam)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.txtJIRAbox)
        Me.GroupBox1.Controls.Add(Me.lblUserID)
        Me.GroupBox1.Controls.Add(Me.txtUserID)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.txtCompany)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtContactName)
        Me.GroupBox1.Controls.Add(Me.txtAccountNum)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtContactEmail)
        Me.GroupBox1.Controls.Add(Me.txtOrderID)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtContactPhone)
        Me.GroupBox1.Controls.Add(Me.DateTimePicker1)
        Me.GroupBox1.Controls.Add(Me.txtContactID)
        Me.GroupBox1.Controls.Add(Me.txtSRNumber)
        Me.GroupBox1.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(365, 43)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(513, 297)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Contact Information"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(7, 19)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(92, 15)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "Agent Team:"
        '
        'txtAgentTeam
        '
        Me.txtAgentTeam.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtAgentTeam.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAgentTeam.Location = New System.Drawing.Point(10, 37)
        Me.txtAgentTeam.Name = "txtAgentTeam"
        Me.txtAgentTeam.Size = New System.Drawing.Size(192, 23)
        Me.txtAgentTeam.TabIndex = 25
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(234, 194)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(48, 15)
        Me.Label12.TabIndex = 23
        Me.Label12.Text = "JIRA#:"
        '
        'txtJIRAbox
        '
        Me.txtJIRAbox.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtJIRAbox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtJIRAbox.Location = New System.Drawing.Point(237, 210)
        Me.txtJIRAbox.Name = "txtJIRAbox"
        Me.txtJIRAbox.Size = New System.Drawing.Size(192, 23)
        Me.txtJIRAbox.TabIndex = 13
        '
        'lblUserID
        '
        Me.lblUserID.AutoSize = True
        Me.lblUserID.Location = New System.Drawing.Point(234, 152)
        Me.lblUserID.Name = "lblUserID"
        Me.lblUserID.Size = New System.Drawing.Size(56, 15)
        Me.lblUserID.TabIndex = 21
        Me.lblUserID.Text = "UserID:"
        '
        'txtUserID
        '
        Me.txtUserID.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtUserID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUserID.Location = New System.Drawing.Point(237, 167)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(192, 23)
        Me.txtUserID.TabIndex = 12
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(234, 108)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(209, 15)
        Me.Label10.TabIndex = 19
        Me.Label10.Text = "Case ID/Order ID/App ID/etc:"
        '
        'txtCompany
        '
        Me.txtCompany.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtCompany.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCompany.Location = New System.Drawing.Point(237, 81)
        Me.txtCompany.Name = "txtCompany"
        Me.txtCompany.Size = New System.Drawing.Size(192, 23)
        Me.txtCompany.TabIndex = 10
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 152)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(106, 15)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Contact Name:"
        '
        'txtContactName
        '
        Me.txtContactName.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtContactName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtContactName.Location = New System.Drawing.Point(9, 169)
        Me.txtContactName.Name = "txtContactName"
        Me.txtContactName.Size = New System.Drawing.Size(192, 23)
        Me.txtContactName.TabIndex = 6
        '
        'txtAccountNum
        '
        Me.txtAccountNum.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtAccountNum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAccountNum.Location = New System.Drawing.Point(237, 38)
        Me.txtAccountNum.Name = "txtAccountNum"
        Me.txtAccountNum.Size = New System.Drawing.Size(192, 23)
        Me.txtAccountNum.TabIndex = 9
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(234, 238)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 15)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Contact Date:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(234, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(76, 15)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Company:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(234, 19)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(69, 15)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Account:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 239)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 15)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Contact Phone:"
        '
        'txtContactEmail
        '
        Me.txtContactEmail.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtContactEmail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtContactEmail.Location = New System.Drawing.Point(8, 213)
        Me.txtContactEmail.Name = "txtContactEmail"
        Me.txtContactEmail.Size = New System.Drawing.Size(192, 23)
        Me.txtContactEmail.TabIndex = 7
        '
        'txtOrderID
        '
        Me.txtOrderID.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtOrderID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOrderID.Location = New System.Drawing.Point(237, 124)
        Me.txtOrderID.Name = "txtOrderID"
        Me.txtOrderID.Size = New System.Drawing.Size(192, 23)
        Me.txtOrderID.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 195)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(105, 15)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Contact Email:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 108)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 15)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Contact ID:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 15)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "SR#:"
        '
        'txtContactPhone
        '
        Me.txtContactPhone.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtContactPhone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtContactPhone.Location = New System.Drawing.Point(8, 257)
        Me.txtContactPhone.Name = "txtContactPhone"
        Me.txtContactPhone.Size = New System.Drawing.Size(192, 23)
        Me.txtContactPhone.TabIndex = 8
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Location = New System.Drawing.Point(237, 253)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(192, 23)
        Me.DateTimePicker1.TabIndex = 14
        '
        'txtContactID
        '
        Me.txtContactID.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtContactID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtContactID.Location = New System.Drawing.Point(9, 125)
        Me.txtContactID.Name = "txtContactID"
        Me.txtContactID.Size = New System.Drawing.Size(193, 23)
        Me.txtContactID.TabIndex = 5
        '
        'txtSRNumber
        '
        Me.txtSRNumber.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtSRNumber.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSRNumber.Location = New System.Drawing.Point(9, 81)
        Me.txtSRNumber.Name = "txtSRNumber"
        Me.txtSRNumber.Size = New System.Drawing.Size(192, 23)
        Me.txtSRNumber.TabIndex = 4
        '
        'btnHide
        '
        Me.btnHide.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHide.Location = New System.Drawing.Point(779, 346)
        Me.btnHide.Name = "btnHide"
        Me.btnHide.Size = New System.Drawing.Size(97, 23)
        Me.btnHide.TabIndex = 18
        Me.btnHide.Text = "Hide Setup Box"
        Me.btnHide.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEdit.Location = New System.Drawing.Point(503, 346)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(97, 23)
        Me.btnEdit.TabIndex = 16
        Me.btnEdit.Text = "Edit info"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(365, 346)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(97, 23)
        Me.btnSave.TabIndex = 15
        Me.btnSave.Text = "Save Info"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Location = New System.Drawing.Point(641, 346)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(97, 23)
        Me.btnClear.TabIndex = 17
        Me.btnClear.Text = "Clear Fields"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(24, 267)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 23)
        Me.Button1.TabIndex = 20
        Me.Button1.Text = "Test 1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(24, 297)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(97, 23)
        Me.Button3.TabIndex = 21
        Me.Button3.Text = "Test 2"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(24, 326)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(97, 23)
        Me.Button4.TabIndex = 22
        Me.Button4.Text = "Test 3"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'lblMDrive
        '
        Me.lblMDrive.AutoSize = True
        Me.lblMDrive.Location = New System.Drawing.Point(263, 281)
        Me.lblMDrive.Name = "lblMDrive"
        Me.lblMDrive.Size = New System.Drawing.Size(40, 13)
        Me.lblMDrive.TabIndex = 23
        Me.lblMDrive.Text = "Label9"
        Me.lblMDrive.Visible = False
        '
        'lblSCRN
        '
        Me.lblSCRN.AutoSize = True
        Me.lblSCRN.Location = New System.Drawing.Point(257, 307)
        Me.lblSCRN.Name = "lblSCRN"
        Me.lblSCRN.Size = New System.Drawing.Size(46, 13)
        Me.lblSCRN.TabIndex = 24
        Me.lblSCRN.Text = "Label13"
        Me.lblSCRN.Visible = False
        '
        'QaSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.ClientSize = New System.Drawing.Size(895, 381)
        Me.Controls.Add(Me.lblSCRN)
        Me.Controls.Add(Me.lblMDrive)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnHide)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox8)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximumSize = New System.Drawing.Size(911, 420)
        Me.MinimumSize = New System.Drawing.Size(911, 420)
        Me.Name = "QaSetup"
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents cboContactType As System.Windows.Forms.ComboBox
    Friend WithEvents cboAgentName As System.Windows.Forms.ComboBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtContactPhone As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtContactID As System.Windows.Forms.TextBox
    Friend WithEvents txtSRNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtCompany As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtContactName As System.Windows.Forms.TextBox
    Friend WithEvents txtAccountNum As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtContactEmail As System.Windows.Forms.TextBox
    Friend WithEvents txtOrderID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnHide As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblQAauditor As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Label12 As Label
    Friend WithEvents txtJIRAbox As TextBox
    Friend WithEvents lblUserID As Label
    Friend WithEvents txtUserID As TextBox
    Friend WithEvents lblMDrive As Label
    Friend WithEvents lblSCRN As Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtAgentTeam As System.Windows.Forms.TextBox
End Class
