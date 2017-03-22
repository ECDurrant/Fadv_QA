<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class QADash
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
        Me.components = New System.ComponentModel.Container()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.cboTeamName = New System.Windows.Forms.ComboBox()
        Me.lblQAauditor = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cboContactType = New System.Windows.Forms.ComboBox()
        Me.cboAgentName = New System.Windows.Forms.ComboBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.QADataSet = New QAapp.QADataSet()
        Me.QAMainDBBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.QAMainDBTableAdapter = New QAapp.QADataSetTableAdapters.QAMainDBTableAdapter()
        Me.IDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SRDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewLinkColumn()
        Me.ContactIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CtypeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QAAgentDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QATeamDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QAContactDateDataGridViewTextBoxColumn = New DataGridViewAutoFilter.DataGridViewAutoFilterTextBoxColumn()
        Me.QAOrderIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QADateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QACommentsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QAOppDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CINameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CIAccountDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CICompanyDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CIPhoneDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CIEmailDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RevDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RevManagerDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RevCommentsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DisScoreDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DisNameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DisNotesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DisAppCommentsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One4DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One5DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One6DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One7DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One8DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One9DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One1NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One2NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One3NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One4NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One5NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One6NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One7NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One8NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.One9NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two4DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two5DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two6DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two7DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two8DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two9DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two1NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two2NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two3NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two4NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two5NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two6NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two7NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two8NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Two9NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three4DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three5DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three6DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three7DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three8DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three9DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three1NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three2NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three3NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three4NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three5NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three6NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three7NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three8NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Three9NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four4DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four5DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four6DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four7DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four8DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four9DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four1NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four2NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four3NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four4NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four5NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four6NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four7NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four8NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Four9NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five4DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five5DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five6DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five7DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five8DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five9DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five1NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five2NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five3NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five4NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five5NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five6NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five7NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five8NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Five9NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six4DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six5DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six6DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six7DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six8DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six9DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six1NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six2NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six3NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six4NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six5NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six6NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six7NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six8NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Six9NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven4DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven5DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven6DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven7DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven8DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven9DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven1NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven2NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven3NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven4NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven5NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven6NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven7NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven8NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Seven9NoteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QAScoreDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.JIRADataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UserIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AutoFailDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AuditorDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox8.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QADataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QAMainDBBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.cboTeamName)
        Me.GroupBox8.Controls.Add(Me.lblQAauditor)
        Me.GroupBox8.Controls.Add(Me.Label11)
        Me.GroupBox8.Controls.Add(Me.cboContactType)
        Me.GroupBox8.Controls.Add(Me.cboAgentName)
        Me.GroupBox8.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox8.Location = New System.Drawing.Point(14, 45)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(420, 254)
        Me.GroupBox8.TabIndex = 25
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Agent Information"
        '
        'cboTeamName
        '
        Me.cboTeamName.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.cboTeamName.FormattingEnabled = True
        Me.cboTeamName.Location = New System.Drawing.Point(16, 75)
        Me.cboTeamName.Name = "cboTeamName"
        Me.cboTeamName.Size = New System.Drawing.Size(259, 25)
        Me.cboTeamName.TabIndex = 26
        Me.cboTeamName.Text = "Team Name"
        '
        'lblQAauditor
        '
        Me.lblQAauditor.AutoSize = True
        Me.lblQAauditor.Font = New System.Drawing.Font("Segoe UI", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQAauditor.ForeColor = System.Drawing.Color.Red
        Me.lblQAauditor.Location = New System.Drawing.Point(171, 37)
        Me.lblQAauditor.Name = "lblQAauditor"
        Me.lblQAauditor.Size = New System.Drawing.Size(75, 15)
        Me.lblQAauditor.TabIndex = 25
        Me.lblQAauditor.Text = "Carla Hardy"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(13, 37)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(136, 15)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "QA Auditor/Supervisor:"
        '
        'cboContactType
        '
        Me.cboContactType.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.cboContactType.FormattingEnabled = True
        Me.cboContactType.Items.AddRange(New Object() {"Call", "Chat", "Email", "Level 2 - Call", "Level 2 - Email", "Resident - Call", "Resident - Email", "Consumer Advocacy - Call"})
        Me.cboContactType.Location = New System.Drawing.Point(16, 190)
        Me.cboContactType.Name = "cboContactType"
        Me.cboContactType.Size = New System.Drawing.Size(259, 25)
        Me.cboContactType.TabIndex = 3
        Me.cboContactType.Text = "Contact Type"
        '
        'cboAgentName
        '
        Me.cboAgentName.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.cboAgentName.FormattingEnabled = True
        Me.cboAgentName.Location = New System.Drawing.Point(16, 133)
        Me.cboAgentName.Name = "cboAgentName"
        Me.cboAgentName.Size = New System.Drawing.Size(259, 25)
        Me.cboAgentName.TabIndex = 2
        Me.cboAgentName.Text = "Agent Name"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(1017, 24)
        Me.MenuStrip1.TabIndex = 26
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IDDataGridViewTextBoxColumn, Me.SRDataGridViewTextBoxColumn, Me.ContactIDDataGridViewTextBoxColumn, Me.CtypeDataGridViewTextBoxColumn, Me.QAAgentDataGridViewTextBoxColumn, Me.QATeamDataGridViewTextBoxColumn, Me.QAContactDateDataGridViewTextBoxColumn, Me.QAOrderIDDataGridViewTextBoxColumn, Me.QADateDataGridViewTextBoxColumn, Me.QACommentsDataGridViewTextBoxColumn, Me.QAOppDataGridViewTextBoxColumn, Me.CINameDataGridViewTextBoxColumn, Me.CIAccountDataGridViewTextBoxColumn, Me.CICompanyDataGridViewTextBoxColumn, Me.CIPhoneDataGridViewTextBoxColumn, Me.CIEmailDataGridViewTextBoxColumn, Me.RevDateDataGridViewTextBoxColumn, Me.RevManagerDataGridViewTextBoxColumn, Me.RevCommentsDataGridViewTextBoxColumn, Me.DisScoreDataGridViewTextBoxColumn, Me.DisNameDataGridViewTextBoxColumn, Me.DisNotesDataGridViewTextBoxColumn, Me.DisAppCommentsDataGridViewTextBoxColumn, Me.One1DataGridViewTextBoxColumn, Me.One2DataGridViewTextBoxColumn, Me.One3DataGridViewTextBoxColumn, Me.One4DataGridViewTextBoxColumn, Me.One5DataGridViewTextBoxColumn, Me.One6DataGridViewTextBoxColumn, Me.One7DataGridViewTextBoxColumn, Me.One8DataGridViewTextBoxColumn, Me.One9DataGridViewTextBoxColumn, Me.One1NoteDataGridViewTextBoxColumn, Me.One2NoteDataGridViewTextBoxColumn, Me.One3NoteDataGridViewTextBoxColumn, Me.One4NoteDataGridViewTextBoxColumn, Me.One5NoteDataGridViewTextBoxColumn, Me.One6NoteDataGridViewTextBoxColumn, Me.One7NoteDataGridViewTextBoxColumn, Me.One8NoteDataGridViewTextBoxColumn, Me.One9NoteDataGridViewTextBoxColumn, Me.Two1DataGridViewTextBoxColumn, Me.Two2DataGridViewTextBoxColumn, Me.Two3DataGridViewTextBoxColumn, Me.Two4DataGridViewTextBoxColumn, Me.Two5DataGridViewTextBoxColumn, Me.Two6DataGridViewTextBoxColumn, Me.Two7DataGridViewTextBoxColumn, Me.Two8DataGridViewTextBoxColumn, Me.Two9DataGridViewTextBoxColumn, Me.Two1NoteDataGridViewTextBoxColumn, Me.Two2NoteDataGridViewTextBoxColumn, Me.Two3NoteDataGridViewTextBoxColumn, Me.Two4NoteDataGridViewTextBoxColumn, Me.Two5NoteDataGridViewTextBoxColumn, Me.Two6NoteDataGridViewTextBoxColumn, Me.Two7NoteDataGridViewTextBoxColumn, Me.Two8NoteDataGridViewTextBoxColumn, Me.Two9NoteDataGridViewTextBoxColumn, Me.Three1DataGridViewTextBoxColumn, Me.Three2DataGridViewTextBoxColumn, Me.Three3DataGridViewTextBoxColumn, Me.Three4DataGridViewTextBoxColumn, Me.Three5DataGridViewTextBoxColumn, Me.Three6DataGridViewTextBoxColumn, Me.Three7DataGridViewTextBoxColumn, Me.Three8DataGridViewTextBoxColumn, Me.Three9DataGridViewTextBoxColumn, Me.Three1NoteDataGridViewTextBoxColumn, Me.Three2NoteDataGridViewTextBoxColumn, Me.Three3NoteDataGridViewTextBoxColumn, Me.Three4NoteDataGridViewTextBoxColumn, Me.Three5NoteDataGridViewTextBoxColumn, Me.Three6NoteDataGridViewTextBoxColumn, Me.Three7NoteDataGridViewTextBoxColumn, Me.Three8NoteDataGridViewTextBoxColumn, Me.Three9NoteDataGridViewTextBoxColumn, Me.Four1DataGridViewTextBoxColumn, Me.Four2DataGridViewTextBoxColumn, Me.Four3DataGridViewTextBoxColumn, Me.Four4DataGridViewTextBoxColumn, Me.Four5DataGridViewTextBoxColumn, Me.Four6DataGridViewTextBoxColumn, Me.Four7DataGridViewTextBoxColumn, Me.Four8DataGridViewTextBoxColumn, Me.Four9DataGridViewTextBoxColumn, Me.Four1NoteDataGridViewTextBoxColumn, Me.Four2NoteDataGridViewTextBoxColumn, Me.Four3NoteDataGridViewTextBoxColumn, Me.Four4NoteDataGridViewTextBoxColumn, Me.Four5NoteDataGridViewTextBoxColumn, Me.Four6NoteDataGridViewTextBoxColumn, Me.Four7NoteDataGridViewTextBoxColumn, Me.Four8NoteDataGridViewTextBoxColumn, Me.Four9NoteDataGridViewTextBoxColumn, Me.Five1DataGridViewTextBoxColumn, Me.Five2DataGridViewTextBoxColumn, Me.Five3DataGridViewTextBoxColumn, Me.Five4DataGridViewTextBoxColumn, Me.Five5DataGridViewTextBoxColumn, Me.Five6DataGridViewTextBoxColumn, Me.Five7DataGridViewTextBoxColumn, Me.Five8DataGridViewTextBoxColumn, Me.Five9DataGridViewTextBoxColumn, Me.Five1NoteDataGridViewTextBoxColumn, Me.Five2NoteDataGridViewTextBoxColumn, Me.Five3NoteDataGridViewTextBoxColumn, Me.Five4NoteDataGridViewTextBoxColumn, Me.Five5NoteDataGridViewTextBoxColumn, Me.Five6NoteDataGridViewTextBoxColumn, Me.Five7NoteDataGridViewTextBoxColumn, Me.Five8NoteDataGridViewTextBoxColumn, Me.Five9NoteDataGridViewTextBoxColumn, Me.Six1DataGridViewTextBoxColumn, Me.Six2DataGridViewTextBoxColumn, Me.Six3DataGridViewTextBoxColumn, Me.Six4DataGridViewTextBoxColumn, Me.Six5DataGridViewTextBoxColumn, Me.Six6DataGridViewTextBoxColumn, Me.Six7DataGridViewTextBoxColumn, Me.Six8DataGridViewTextBoxColumn, Me.Six9DataGridViewTextBoxColumn, Me.Six1NoteDataGridViewTextBoxColumn, Me.Six2NoteDataGridViewTextBoxColumn, Me.Six3NoteDataGridViewTextBoxColumn, Me.Six4NoteDataGridViewTextBoxColumn, Me.Six5NoteDataGridViewTextBoxColumn, Me.Six6NoteDataGridViewTextBoxColumn, Me.Six7NoteDataGridViewTextBoxColumn, Me.Six8NoteDataGridViewTextBoxColumn, Me.Six9NoteDataGridViewTextBoxColumn, Me.Seven1DataGridViewTextBoxColumn, Me.Seven2DataGridViewTextBoxColumn, Me.Seven3DataGridViewTextBoxColumn, Me.Seven4DataGridViewTextBoxColumn, Me.Seven5DataGridViewTextBoxColumn, Me.Seven6DataGridViewTextBoxColumn, Me.Seven7DataGridViewTextBoxColumn, Me.Seven8DataGridViewTextBoxColumn, Me.Seven9DataGridViewTextBoxColumn, Me.Seven1NoteDataGridViewTextBoxColumn, Me.Seven2NoteDataGridViewTextBoxColumn, Me.Seven3NoteDataGridViewTextBoxColumn, Me.Seven4NoteDataGridViewTextBoxColumn, Me.Seven5NoteDataGridViewTextBoxColumn, Me.Seven6NoteDataGridViewTextBoxColumn, Me.Seven7NoteDataGridViewTextBoxColumn, Me.Seven8NoteDataGridViewTextBoxColumn, Me.Seven9NoteDataGridViewTextBoxColumn, Me.QAScoreDataGridViewTextBoxColumn, Me.JIRADataGridViewTextBoxColumn, Me.UserIDDataGridViewTextBoxColumn, Me.AutoFailDataGridViewTextBoxColumn, Me.AuditorDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.QAMainDBBindingSource
        Me.DataGridView1.Location = New System.Drawing.Point(440, 53)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(548, 150)
        Me.DataGridView1.TabIndex = 27
        '
        'QADataSet
        '
        Me.QADataSet.DataSetName = "QADataSet"
        Me.QADataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'QAMainDBBindingSource
        '
        Me.QAMainDBBindingSource.DataMember = "QAMainDB"
        Me.QAMainDBBindingSource.DataSource = Me.QADataSet
        '
        'QAMainDBTableAdapter
        '
        Me.QAMainDBTableAdapter.ClearBeforeFill = True
        '
        'IDDataGridViewTextBoxColumn
        '
        Me.IDDataGridViewTextBoxColumn.DataPropertyName = "ID"
        Me.IDDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IDDataGridViewTextBoxColumn.Name = "IDDataGridViewTextBoxColumn"
        Me.IDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'SRDataGridViewTextBoxColumn
        '
        Me.SRDataGridViewTextBoxColumn.DataPropertyName = "SR"
        Me.SRDataGridViewTextBoxColumn.HeaderText = "SR"
        Me.SRDataGridViewTextBoxColumn.Name = "SRDataGridViewTextBoxColumn"
        Me.SRDataGridViewTextBoxColumn.ReadOnly = True
        Me.SRDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.SRDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ContactIDDataGridViewTextBoxColumn
        '
        Me.ContactIDDataGridViewTextBoxColumn.DataPropertyName = "ContactID"
        Me.ContactIDDataGridViewTextBoxColumn.HeaderText = "ContactID"
        Me.ContactIDDataGridViewTextBoxColumn.Name = "ContactIDDataGridViewTextBoxColumn"
        Me.ContactIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CtypeDataGridViewTextBoxColumn
        '
        Me.CtypeDataGridViewTextBoxColumn.DataPropertyName = "Ctype"
        Me.CtypeDataGridViewTextBoxColumn.HeaderText = "Ctype"
        Me.CtypeDataGridViewTextBoxColumn.Name = "CtypeDataGridViewTextBoxColumn"
        Me.CtypeDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QAAgentDataGridViewTextBoxColumn
        '
        Me.QAAgentDataGridViewTextBoxColumn.DataPropertyName = "QA-Agent"
        Me.QAAgentDataGridViewTextBoxColumn.HeaderText = "QA-Agent"
        Me.QAAgentDataGridViewTextBoxColumn.Name = "QAAgentDataGridViewTextBoxColumn"
        Me.QAAgentDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QATeamDataGridViewTextBoxColumn
        '
        Me.QATeamDataGridViewTextBoxColumn.DataPropertyName = "QA-Team"
        Me.QATeamDataGridViewTextBoxColumn.HeaderText = "QA-Team"
        Me.QATeamDataGridViewTextBoxColumn.Name = "QATeamDataGridViewTextBoxColumn"
        Me.QATeamDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QAContactDateDataGridViewTextBoxColumn
        '
        Me.QAContactDateDataGridViewTextBoxColumn.DataPropertyName = "QA-ContactDate"
        Me.QAContactDateDataGridViewTextBoxColumn.HeaderText = "QA-ContactDate"
        Me.QAContactDateDataGridViewTextBoxColumn.Name = "QAContactDateDataGridViewTextBoxColumn"
        Me.QAContactDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.QAContactDateDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'QAOrderIDDataGridViewTextBoxColumn
        '
        Me.QAOrderIDDataGridViewTextBoxColumn.DataPropertyName = "QA-OrderID"
        Me.QAOrderIDDataGridViewTextBoxColumn.HeaderText = "QA-OrderID"
        Me.QAOrderIDDataGridViewTextBoxColumn.Name = "QAOrderIDDataGridViewTextBoxColumn"
        Me.QAOrderIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QADateDataGridViewTextBoxColumn
        '
        Me.QADateDataGridViewTextBoxColumn.DataPropertyName = "QA-Date"
        Me.QADateDataGridViewTextBoxColumn.HeaderText = "QA-Date"
        Me.QADateDataGridViewTextBoxColumn.Name = "QADateDataGridViewTextBoxColumn"
        Me.QADateDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QACommentsDataGridViewTextBoxColumn
        '
        Me.QACommentsDataGridViewTextBoxColumn.DataPropertyName = "QA-Comments"
        Me.QACommentsDataGridViewTextBoxColumn.HeaderText = "QA-Comments"
        Me.QACommentsDataGridViewTextBoxColumn.Name = "QACommentsDataGridViewTextBoxColumn"
        Me.QACommentsDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QAOppDataGridViewTextBoxColumn
        '
        Me.QAOppDataGridViewTextBoxColumn.DataPropertyName = "QA-Opp"
        Me.QAOppDataGridViewTextBoxColumn.HeaderText = "QA-Opp"
        Me.QAOppDataGridViewTextBoxColumn.Name = "QAOppDataGridViewTextBoxColumn"
        Me.QAOppDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CINameDataGridViewTextBoxColumn
        '
        Me.CINameDataGridViewTextBoxColumn.DataPropertyName = "CI-Name"
        Me.CINameDataGridViewTextBoxColumn.HeaderText = "CI-Name"
        Me.CINameDataGridViewTextBoxColumn.Name = "CINameDataGridViewTextBoxColumn"
        Me.CINameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CIAccountDataGridViewTextBoxColumn
        '
        Me.CIAccountDataGridViewTextBoxColumn.DataPropertyName = "CI-Account"
        Me.CIAccountDataGridViewTextBoxColumn.HeaderText = "CI-Account"
        Me.CIAccountDataGridViewTextBoxColumn.Name = "CIAccountDataGridViewTextBoxColumn"
        Me.CIAccountDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CICompanyDataGridViewTextBoxColumn
        '
        Me.CICompanyDataGridViewTextBoxColumn.DataPropertyName = "CI-Company"
        Me.CICompanyDataGridViewTextBoxColumn.HeaderText = "CI-Company"
        Me.CICompanyDataGridViewTextBoxColumn.Name = "CICompanyDataGridViewTextBoxColumn"
        Me.CICompanyDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CIPhoneDataGridViewTextBoxColumn
        '
        Me.CIPhoneDataGridViewTextBoxColumn.DataPropertyName = "CI-Phone"
        Me.CIPhoneDataGridViewTextBoxColumn.HeaderText = "CI-Phone"
        Me.CIPhoneDataGridViewTextBoxColumn.Name = "CIPhoneDataGridViewTextBoxColumn"
        Me.CIPhoneDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CIEmailDataGridViewTextBoxColumn
        '
        Me.CIEmailDataGridViewTextBoxColumn.DataPropertyName = "CI-Email"
        Me.CIEmailDataGridViewTextBoxColumn.HeaderText = "CI-Email"
        Me.CIEmailDataGridViewTextBoxColumn.Name = "CIEmailDataGridViewTextBoxColumn"
        Me.CIEmailDataGridViewTextBoxColumn.ReadOnly = True
        '
        'RevDateDataGridViewTextBoxColumn
        '
        Me.RevDateDataGridViewTextBoxColumn.DataPropertyName = "Rev-Date"
        Me.RevDateDataGridViewTextBoxColumn.HeaderText = "Rev-Date"
        Me.RevDateDataGridViewTextBoxColumn.Name = "RevDateDataGridViewTextBoxColumn"
        Me.RevDateDataGridViewTextBoxColumn.ReadOnly = True
        '
        'RevManagerDataGridViewTextBoxColumn
        '
        Me.RevManagerDataGridViewTextBoxColumn.DataPropertyName = "Rev-Manager"
        Me.RevManagerDataGridViewTextBoxColumn.HeaderText = "Rev-Manager"
        Me.RevManagerDataGridViewTextBoxColumn.Name = "RevManagerDataGridViewTextBoxColumn"
        Me.RevManagerDataGridViewTextBoxColumn.ReadOnly = True
        '
        'RevCommentsDataGridViewTextBoxColumn
        '
        Me.RevCommentsDataGridViewTextBoxColumn.DataPropertyName = "Rev-Comments"
        Me.RevCommentsDataGridViewTextBoxColumn.HeaderText = "Rev-Comments"
        Me.RevCommentsDataGridViewTextBoxColumn.Name = "RevCommentsDataGridViewTextBoxColumn"
        Me.RevCommentsDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DisScoreDataGridViewTextBoxColumn
        '
        Me.DisScoreDataGridViewTextBoxColumn.DataPropertyName = "Dis-Score"
        Me.DisScoreDataGridViewTextBoxColumn.HeaderText = "Dis-Score"
        Me.DisScoreDataGridViewTextBoxColumn.Name = "DisScoreDataGridViewTextBoxColumn"
        Me.DisScoreDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DisNameDataGridViewTextBoxColumn
        '
        Me.DisNameDataGridViewTextBoxColumn.DataPropertyName = "Dis-Name"
        Me.DisNameDataGridViewTextBoxColumn.HeaderText = "Dis-Name"
        Me.DisNameDataGridViewTextBoxColumn.Name = "DisNameDataGridViewTextBoxColumn"
        Me.DisNameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DisNotesDataGridViewTextBoxColumn
        '
        Me.DisNotesDataGridViewTextBoxColumn.DataPropertyName = "Dis-Notes"
        Me.DisNotesDataGridViewTextBoxColumn.HeaderText = "Dis-Notes"
        Me.DisNotesDataGridViewTextBoxColumn.Name = "DisNotesDataGridViewTextBoxColumn"
        Me.DisNotesDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DisAppCommentsDataGridViewTextBoxColumn
        '
        Me.DisAppCommentsDataGridViewTextBoxColumn.DataPropertyName = "Dis-AppComments"
        Me.DisAppCommentsDataGridViewTextBoxColumn.HeaderText = "Dis-AppComments"
        Me.DisAppCommentsDataGridViewTextBoxColumn.Name = "DisAppCommentsDataGridViewTextBoxColumn"
        Me.DisAppCommentsDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One1DataGridViewTextBoxColumn
        '
        Me.One1DataGridViewTextBoxColumn.DataPropertyName = "One-1"
        Me.One1DataGridViewTextBoxColumn.HeaderText = "One-1"
        Me.One1DataGridViewTextBoxColumn.Name = "One1DataGridViewTextBoxColumn"
        Me.One1DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One2DataGridViewTextBoxColumn
        '
        Me.One2DataGridViewTextBoxColumn.DataPropertyName = "One-2"
        Me.One2DataGridViewTextBoxColumn.HeaderText = "One-2"
        Me.One2DataGridViewTextBoxColumn.Name = "One2DataGridViewTextBoxColumn"
        Me.One2DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One3DataGridViewTextBoxColumn
        '
        Me.One3DataGridViewTextBoxColumn.DataPropertyName = "One-3"
        Me.One3DataGridViewTextBoxColumn.HeaderText = "One-3"
        Me.One3DataGridViewTextBoxColumn.Name = "One3DataGridViewTextBoxColumn"
        Me.One3DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One4DataGridViewTextBoxColumn
        '
        Me.One4DataGridViewTextBoxColumn.DataPropertyName = "One-4"
        Me.One4DataGridViewTextBoxColumn.HeaderText = "One-4"
        Me.One4DataGridViewTextBoxColumn.Name = "One4DataGridViewTextBoxColumn"
        Me.One4DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One5DataGridViewTextBoxColumn
        '
        Me.One5DataGridViewTextBoxColumn.DataPropertyName = "One-5"
        Me.One5DataGridViewTextBoxColumn.HeaderText = "One-5"
        Me.One5DataGridViewTextBoxColumn.Name = "One5DataGridViewTextBoxColumn"
        Me.One5DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One6DataGridViewTextBoxColumn
        '
        Me.One6DataGridViewTextBoxColumn.DataPropertyName = "One-6"
        Me.One6DataGridViewTextBoxColumn.HeaderText = "One-6"
        Me.One6DataGridViewTextBoxColumn.Name = "One6DataGridViewTextBoxColumn"
        Me.One6DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One7DataGridViewTextBoxColumn
        '
        Me.One7DataGridViewTextBoxColumn.DataPropertyName = "One-7"
        Me.One7DataGridViewTextBoxColumn.HeaderText = "One-7"
        Me.One7DataGridViewTextBoxColumn.Name = "One7DataGridViewTextBoxColumn"
        Me.One7DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One8DataGridViewTextBoxColumn
        '
        Me.One8DataGridViewTextBoxColumn.DataPropertyName = "One-8"
        Me.One8DataGridViewTextBoxColumn.HeaderText = "One-8"
        Me.One8DataGridViewTextBoxColumn.Name = "One8DataGridViewTextBoxColumn"
        Me.One8DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One9DataGridViewTextBoxColumn
        '
        Me.One9DataGridViewTextBoxColumn.DataPropertyName = "One-9"
        Me.One9DataGridViewTextBoxColumn.HeaderText = "One-9"
        Me.One9DataGridViewTextBoxColumn.Name = "One9DataGridViewTextBoxColumn"
        Me.One9DataGridViewTextBoxColumn.ReadOnly = True
        '
        'One1NoteDataGridViewTextBoxColumn
        '
        Me.One1NoteDataGridViewTextBoxColumn.DataPropertyName = "One-1Note"
        Me.One1NoteDataGridViewTextBoxColumn.HeaderText = "One-1Note"
        Me.One1NoteDataGridViewTextBoxColumn.Name = "One1NoteDataGridViewTextBoxColumn"
        Me.One1NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One2NoteDataGridViewTextBoxColumn
        '
        Me.One2NoteDataGridViewTextBoxColumn.DataPropertyName = "One-2Note"
        Me.One2NoteDataGridViewTextBoxColumn.HeaderText = "One-2Note"
        Me.One2NoteDataGridViewTextBoxColumn.Name = "One2NoteDataGridViewTextBoxColumn"
        Me.One2NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One3NoteDataGridViewTextBoxColumn
        '
        Me.One3NoteDataGridViewTextBoxColumn.DataPropertyName = "One-3Note"
        Me.One3NoteDataGridViewTextBoxColumn.HeaderText = "One-3Note"
        Me.One3NoteDataGridViewTextBoxColumn.Name = "One3NoteDataGridViewTextBoxColumn"
        Me.One3NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One4NoteDataGridViewTextBoxColumn
        '
        Me.One4NoteDataGridViewTextBoxColumn.DataPropertyName = "One-4Note"
        Me.One4NoteDataGridViewTextBoxColumn.HeaderText = "One-4Note"
        Me.One4NoteDataGridViewTextBoxColumn.Name = "One4NoteDataGridViewTextBoxColumn"
        Me.One4NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One5NoteDataGridViewTextBoxColumn
        '
        Me.One5NoteDataGridViewTextBoxColumn.DataPropertyName = "One-5Note"
        Me.One5NoteDataGridViewTextBoxColumn.HeaderText = "One-5Note"
        Me.One5NoteDataGridViewTextBoxColumn.Name = "One5NoteDataGridViewTextBoxColumn"
        Me.One5NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One6NoteDataGridViewTextBoxColumn
        '
        Me.One6NoteDataGridViewTextBoxColumn.DataPropertyName = "One-6Note"
        Me.One6NoteDataGridViewTextBoxColumn.HeaderText = "One-6Note"
        Me.One6NoteDataGridViewTextBoxColumn.Name = "One6NoteDataGridViewTextBoxColumn"
        Me.One6NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One7NoteDataGridViewTextBoxColumn
        '
        Me.One7NoteDataGridViewTextBoxColumn.DataPropertyName = "One-7Note"
        Me.One7NoteDataGridViewTextBoxColumn.HeaderText = "One-7Note"
        Me.One7NoteDataGridViewTextBoxColumn.Name = "One7NoteDataGridViewTextBoxColumn"
        Me.One7NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One8NoteDataGridViewTextBoxColumn
        '
        Me.One8NoteDataGridViewTextBoxColumn.DataPropertyName = "One-8Note"
        Me.One8NoteDataGridViewTextBoxColumn.HeaderText = "One-8Note"
        Me.One8NoteDataGridViewTextBoxColumn.Name = "One8NoteDataGridViewTextBoxColumn"
        Me.One8NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'One9NoteDataGridViewTextBoxColumn
        '
        Me.One9NoteDataGridViewTextBoxColumn.DataPropertyName = "One-9Note"
        Me.One9NoteDataGridViewTextBoxColumn.HeaderText = "One-9Note"
        Me.One9NoteDataGridViewTextBoxColumn.Name = "One9NoteDataGridViewTextBoxColumn"
        Me.One9NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two1DataGridViewTextBoxColumn
        '
        Me.Two1DataGridViewTextBoxColumn.DataPropertyName = "Two-1"
        Me.Two1DataGridViewTextBoxColumn.HeaderText = "Two-1"
        Me.Two1DataGridViewTextBoxColumn.Name = "Two1DataGridViewTextBoxColumn"
        Me.Two1DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two2DataGridViewTextBoxColumn
        '
        Me.Two2DataGridViewTextBoxColumn.DataPropertyName = "Two-2"
        Me.Two2DataGridViewTextBoxColumn.HeaderText = "Two-2"
        Me.Two2DataGridViewTextBoxColumn.Name = "Two2DataGridViewTextBoxColumn"
        Me.Two2DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two3DataGridViewTextBoxColumn
        '
        Me.Two3DataGridViewTextBoxColumn.DataPropertyName = "Two-3"
        Me.Two3DataGridViewTextBoxColumn.HeaderText = "Two-3"
        Me.Two3DataGridViewTextBoxColumn.Name = "Two3DataGridViewTextBoxColumn"
        Me.Two3DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two4DataGridViewTextBoxColumn
        '
        Me.Two4DataGridViewTextBoxColumn.DataPropertyName = "Two-4"
        Me.Two4DataGridViewTextBoxColumn.HeaderText = "Two-4"
        Me.Two4DataGridViewTextBoxColumn.Name = "Two4DataGridViewTextBoxColumn"
        Me.Two4DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two5DataGridViewTextBoxColumn
        '
        Me.Two5DataGridViewTextBoxColumn.DataPropertyName = "Two-5"
        Me.Two5DataGridViewTextBoxColumn.HeaderText = "Two-5"
        Me.Two5DataGridViewTextBoxColumn.Name = "Two5DataGridViewTextBoxColumn"
        Me.Two5DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two6DataGridViewTextBoxColumn
        '
        Me.Two6DataGridViewTextBoxColumn.DataPropertyName = "Two-6"
        Me.Two6DataGridViewTextBoxColumn.HeaderText = "Two-6"
        Me.Two6DataGridViewTextBoxColumn.Name = "Two6DataGridViewTextBoxColumn"
        Me.Two6DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two7DataGridViewTextBoxColumn
        '
        Me.Two7DataGridViewTextBoxColumn.DataPropertyName = "Two-7"
        Me.Two7DataGridViewTextBoxColumn.HeaderText = "Two-7"
        Me.Two7DataGridViewTextBoxColumn.Name = "Two7DataGridViewTextBoxColumn"
        Me.Two7DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two8DataGridViewTextBoxColumn
        '
        Me.Two8DataGridViewTextBoxColumn.DataPropertyName = "Two-8"
        Me.Two8DataGridViewTextBoxColumn.HeaderText = "Two-8"
        Me.Two8DataGridViewTextBoxColumn.Name = "Two8DataGridViewTextBoxColumn"
        Me.Two8DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two9DataGridViewTextBoxColumn
        '
        Me.Two9DataGridViewTextBoxColumn.DataPropertyName = "Two-9"
        Me.Two9DataGridViewTextBoxColumn.HeaderText = "Two-9"
        Me.Two9DataGridViewTextBoxColumn.Name = "Two9DataGridViewTextBoxColumn"
        Me.Two9DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two1NoteDataGridViewTextBoxColumn
        '
        Me.Two1NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-1Note"
        Me.Two1NoteDataGridViewTextBoxColumn.HeaderText = "Two-1Note"
        Me.Two1NoteDataGridViewTextBoxColumn.Name = "Two1NoteDataGridViewTextBoxColumn"
        Me.Two1NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two2NoteDataGridViewTextBoxColumn
        '
        Me.Two2NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-2Note"
        Me.Two2NoteDataGridViewTextBoxColumn.HeaderText = "Two-2Note"
        Me.Two2NoteDataGridViewTextBoxColumn.Name = "Two2NoteDataGridViewTextBoxColumn"
        Me.Two2NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two3NoteDataGridViewTextBoxColumn
        '
        Me.Two3NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-3Note"
        Me.Two3NoteDataGridViewTextBoxColumn.HeaderText = "Two-3Note"
        Me.Two3NoteDataGridViewTextBoxColumn.Name = "Two3NoteDataGridViewTextBoxColumn"
        Me.Two3NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two4NoteDataGridViewTextBoxColumn
        '
        Me.Two4NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-4Note"
        Me.Two4NoteDataGridViewTextBoxColumn.HeaderText = "Two-4Note"
        Me.Two4NoteDataGridViewTextBoxColumn.Name = "Two4NoteDataGridViewTextBoxColumn"
        Me.Two4NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two5NoteDataGridViewTextBoxColumn
        '
        Me.Two5NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-5Note"
        Me.Two5NoteDataGridViewTextBoxColumn.HeaderText = "Two-5Note"
        Me.Two5NoteDataGridViewTextBoxColumn.Name = "Two5NoteDataGridViewTextBoxColumn"
        Me.Two5NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two6NoteDataGridViewTextBoxColumn
        '
        Me.Two6NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-6Note"
        Me.Two6NoteDataGridViewTextBoxColumn.HeaderText = "Two-6Note"
        Me.Two6NoteDataGridViewTextBoxColumn.Name = "Two6NoteDataGridViewTextBoxColumn"
        Me.Two6NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two7NoteDataGridViewTextBoxColumn
        '
        Me.Two7NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-7Note"
        Me.Two7NoteDataGridViewTextBoxColumn.HeaderText = "Two-7Note"
        Me.Two7NoteDataGridViewTextBoxColumn.Name = "Two7NoteDataGridViewTextBoxColumn"
        Me.Two7NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two8NoteDataGridViewTextBoxColumn
        '
        Me.Two8NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-8Note"
        Me.Two8NoteDataGridViewTextBoxColumn.HeaderText = "Two-8Note"
        Me.Two8NoteDataGridViewTextBoxColumn.Name = "Two8NoteDataGridViewTextBoxColumn"
        Me.Two8NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Two9NoteDataGridViewTextBoxColumn
        '
        Me.Two9NoteDataGridViewTextBoxColumn.DataPropertyName = "Two-9Note"
        Me.Two9NoteDataGridViewTextBoxColumn.HeaderText = "Two-9Note"
        Me.Two9NoteDataGridViewTextBoxColumn.Name = "Two9NoteDataGridViewTextBoxColumn"
        Me.Two9NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three1DataGridViewTextBoxColumn
        '
        Me.Three1DataGridViewTextBoxColumn.DataPropertyName = "Three-1"
        Me.Three1DataGridViewTextBoxColumn.HeaderText = "Three-1"
        Me.Three1DataGridViewTextBoxColumn.Name = "Three1DataGridViewTextBoxColumn"
        Me.Three1DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three2DataGridViewTextBoxColumn
        '
        Me.Three2DataGridViewTextBoxColumn.DataPropertyName = "Three-2"
        Me.Three2DataGridViewTextBoxColumn.HeaderText = "Three-2"
        Me.Three2DataGridViewTextBoxColumn.Name = "Three2DataGridViewTextBoxColumn"
        Me.Three2DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three3DataGridViewTextBoxColumn
        '
        Me.Three3DataGridViewTextBoxColumn.DataPropertyName = "Three-3"
        Me.Three3DataGridViewTextBoxColumn.HeaderText = "Three-3"
        Me.Three3DataGridViewTextBoxColumn.Name = "Three3DataGridViewTextBoxColumn"
        Me.Three3DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three4DataGridViewTextBoxColumn
        '
        Me.Three4DataGridViewTextBoxColumn.DataPropertyName = "Three-4"
        Me.Three4DataGridViewTextBoxColumn.HeaderText = "Three-4"
        Me.Three4DataGridViewTextBoxColumn.Name = "Three4DataGridViewTextBoxColumn"
        Me.Three4DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three5DataGridViewTextBoxColumn
        '
        Me.Three5DataGridViewTextBoxColumn.DataPropertyName = "Three-5"
        Me.Three5DataGridViewTextBoxColumn.HeaderText = "Three-5"
        Me.Three5DataGridViewTextBoxColumn.Name = "Three5DataGridViewTextBoxColumn"
        Me.Three5DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three6DataGridViewTextBoxColumn
        '
        Me.Three6DataGridViewTextBoxColumn.DataPropertyName = "Three-6"
        Me.Three6DataGridViewTextBoxColumn.HeaderText = "Three-6"
        Me.Three6DataGridViewTextBoxColumn.Name = "Three6DataGridViewTextBoxColumn"
        Me.Three6DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three7DataGridViewTextBoxColumn
        '
        Me.Three7DataGridViewTextBoxColumn.DataPropertyName = "Three-7"
        Me.Three7DataGridViewTextBoxColumn.HeaderText = "Three-7"
        Me.Three7DataGridViewTextBoxColumn.Name = "Three7DataGridViewTextBoxColumn"
        Me.Three7DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three8DataGridViewTextBoxColumn
        '
        Me.Three8DataGridViewTextBoxColumn.DataPropertyName = "Three-8"
        Me.Three8DataGridViewTextBoxColumn.HeaderText = "Three-8"
        Me.Three8DataGridViewTextBoxColumn.Name = "Three8DataGridViewTextBoxColumn"
        Me.Three8DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three9DataGridViewTextBoxColumn
        '
        Me.Three9DataGridViewTextBoxColumn.DataPropertyName = "Three-9"
        Me.Three9DataGridViewTextBoxColumn.HeaderText = "Three-9"
        Me.Three9DataGridViewTextBoxColumn.Name = "Three9DataGridViewTextBoxColumn"
        Me.Three9DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three1NoteDataGridViewTextBoxColumn
        '
        Me.Three1NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-1Note"
        Me.Three1NoteDataGridViewTextBoxColumn.HeaderText = "Three-1Note"
        Me.Three1NoteDataGridViewTextBoxColumn.Name = "Three1NoteDataGridViewTextBoxColumn"
        Me.Three1NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three2NoteDataGridViewTextBoxColumn
        '
        Me.Three2NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-2Note"
        Me.Three2NoteDataGridViewTextBoxColumn.HeaderText = "Three-2Note"
        Me.Three2NoteDataGridViewTextBoxColumn.Name = "Three2NoteDataGridViewTextBoxColumn"
        Me.Three2NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three3NoteDataGridViewTextBoxColumn
        '
        Me.Three3NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-3Note"
        Me.Three3NoteDataGridViewTextBoxColumn.HeaderText = "Three-3Note"
        Me.Three3NoteDataGridViewTextBoxColumn.Name = "Three3NoteDataGridViewTextBoxColumn"
        Me.Three3NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three4NoteDataGridViewTextBoxColumn
        '
        Me.Three4NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-4Note"
        Me.Three4NoteDataGridViewTextBoxColumn.HeaderText = "Three-4Note"
        Me.Three4NoteDataGridViewTextBoxColumn.Name = "Three4NoteDataGridViewTextBoxColumn"
        Me.Three4NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three5NoteDataGridViewTextBoxColumn
        '
        Me.Three5NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-5Note"
        Me.Three5NoteDataGridViewTextBoxColumn.HeaderText = "Three-5Note"
        Me.Three5NoteDataGridViewTextBoxColumn.Name = "Three5NoteDataGridViewTextBoxColumn"
        Me.Three5NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three6NoteDataGridViewTextBoxColumn
        '
        Me.Three6NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-6Note"
        Me.Three6NoteDataGridViewTextBoxColumn.HeaderText = "Three-6Note"
        Me.Three6NoteDataGridViewTextBoxColumn.Name = "Three6NoteDataGridViewTextBoxColumn"
        Me.Three6NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three7NoteDataGridViewTextBoxColumn
        '
        Me.Three7NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-7Note"
        Me.Three7NoteDataGridViewTextBoxColumn.HeaderText = "Three-7Note"
        Me.Three7NoteDataGridViewTextBoxColumn.Name = "Three7NoteDataGridViewTextBoxColumn"
        Me.Three7NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three8NoteDataGridViewTextBoxColumn
        '
        Me.Three8NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-8Note"
        Me.Three8NoteDataGridViewTextBoxColumn.HeaderText = "Three-8Note"
        Me.Three8NoteDataGridViewTextBoxColumn.Name = "Three8NoteDataGridViewTextBoxColumn"
        Me.Three8NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Three9NoteDataGridViewTextBoxColumn
        '
        Me.Three9NoteDataGridViewTextBoxColumn.DataPropertyName = "Three-9Note"
        Me.Three9NoteDataGridViewTextBoxColumn.HeaderText = "Three-9Note"
        Me.Three9NoteDataGridViewTextBoxColumn.Name = "Three9NoteDataGridViewTextBoxColumn"
        Me.Three9NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four1DataGridViewTextBoxColumn
        '
        Me.Four1DataGridViewTextBoxColumn.DataPropertyName = "Four-1"
        Me.Four1DataGridViewTextBoxColumn.HeaderText = "Four-1"
        Me.Four1DataGridViewTextBoxColumn.Name = "Four1DataGridViewTextBoxColumn"
        Me.Four1DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four2DataGridViewTextBoxColumn
        '
        Me.Four2DataGridViewTextBoxColumn.DataPropertyName = "Four-2"
        Me.Four2DataGridViewTextBoxColumn.HeaderText = "Four-2"
        Me.Four2DataGridViewTextBoxColumn.Name = "Four2DataGridViewTextBoxColumn"
        Me.Four2DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four3DataGridViewTextBoxColumn
        '
        Me.Four3DataGridViewTextBoxColumn.DataPropertyName = "Four-3"
        Me.Four3DataGridViewTextBoxColumn.HeaderText = "Four-3"
        Me.Four3DataGridViewTextBoxColumn.Name = "Four3DataGridViewTextBoxColumn"
        Me.Four3DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four4DataGridViewTextBoxColumn
        '
        Me.Four4DataGridViewTextBoxColumn.DataPropertyName = "Four-4"
        Me.Four4DataGridViewTextBoxColumn.HeaderText = "Four-4"
        Me.Four4DataGridViewTextBoxColumn.Name = "Four4DataGridViewTextBoxColumn"
        Me.Four4DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four5DataGridViewTextBoxColumn
        '
        Me.Four5DataGridViewTextBoxColumn.DataPropertyName = "Four-5"
        Me.Four5DataGridViewTextBoxColumn.HeaderText = "Four-5"
        Me.Four5DataGridViewTextBoxColumn.Name = "Four5DataGridViewTextBoxColumn"
        Me.Four5DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four6DataGridViewTextBoxColumn
        '
        Me.Four6DataGridViewTextBoxColumn.DataPropertyName = "Four-6"
        Me.Four6DataGridViewTextBoxColumn.HeaderText = "Four-6"
        Me.Four6DataGridViewTextBoxColumn.Name = "Four6DataGridViewTextBoxColumn"
        Me.Four6DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four7DataGridViewTextBoxColumn
        '
        Me.Four7DataGridViewTextBoxColumn.DataPropertyName = "Four-7"
        Me.Four7DataGridViewTextBoxColumn.HeaderText = "Four-7"
        Me.Four7DataGridViewTextBoxColumn.Name = "Four7DataGridViewTextBoxColumn"
        Me.Four7DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four8DataGridViewTextBoxColumn
        '
        Me.Four8DataGridViewTextBoxColumn.DataPropertyName = "Four-8"
        Me.Four8DataGridViewTextBoxColumn.HeaderText = "Four-8"
        Me.Four8DataGridViewTextBoxColumn.Name = "Four8DataGridViewTextBoxColumn"
        Me.Four8DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four9DataGridViewTextBoxColumn
        '
        Me.Four9DataGridViewTextBoxColumn.DataPropertyName = "Four-9"
        Me.Four9DataGridViewTextBoxColumn.HeaderText = "Four-9"
        Me.Four9DataGridViewTextBoxColumn.Name = "Four9DataGridViewTextBoxColumn"
        Me.Four9DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four1NoteDataGridViewTextBoxColumn
        '
        Me.Four1NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-1Note"
        Me.Four1NoteDataGridViewTextBoxColumn.HeaderText = "Four-1Note"
        Me.Four1NoteDataGridViewTextBoxColumn.Name = "Four1NoteDataGridViewTextBoxColumn"
        Me.Four1NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four2NoteDataGridViewTextBoxColumn
        '
        Me.Four2NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-2Note"
        Me.Four2NoteDataGridViewTextBoxColumn.HeaderText = "Four-2Note"
        Me.Four2NoteDataGridViewTextBoxColumn.Name = "Four2NoteDataGridViewTextBoxColumn"
        Me.Four2NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four3NoteDataGridViewTextBoxColumn
        '
        Me.Four3NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-3Note"
        Me.Four3NoteDataGridViewTextBoxColumn.HeaderText = "Four-3Note"
        Me.Four3NoteDataGridViewTextBoxColumn.Name = "Four3NoteDataGridViewTextBoxColumn"
        Me.Four3NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four4NoteDataGridViewTextBoxColumn
        '
        Me.Four4NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-4Note"
        Me.Four4NoteDataGridViewTextBoxColumn.HeaderText = "Four-4Note"
        Me.Four4NoteDataGridViewTextBoxColumn.Name = "Four4NoteDataGridViewTextBoxColumn"
        Me.Four4NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four5NoteDataGridViewTextBoxColumn
        '
        Me.Four5NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-5Note"
        Me.Four5NoteDataGridViewTextBoxColumn.HeaderText = "Four-5Note"
        Me.Four5NoteDataGridViewTextBoxColumn.Name = "Four5NoteDataGridViewTextBoxColumn"
        Me.Four5NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four6NoteDataGridViewTextBoxColumn
        '
        Me.Four6NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-6Note"
        Me.Four6NoteDataGridViewTextBoxColumn.HeaderText = "Four-6Note"
        Me.Four6NoteDataGridViewTextBoxColumn.Name = "Four6NoteDataGridViewTextBoxColumn"
        Me.Four6NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four7NoteDataGridViewTextBoxColumn
        '
        Me.Four7NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-7Note"
        Me.Four7NoteDataGridViewTextBoxColumn.HeaderText = "Four-7Note"
        Me.Four7NoteDataGridViewTextBoxColumn.Name = "Four7NoteDataGridViewTextBoxColumn"
        Me.Four7NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four8NoteDataGridViewTextBoxColumn
        '
        Me.Four8NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-8Note"
        Me.Four8NoteDataGridViewTextBoxColumn.HeaderText = "Four-8Note"
        Me.Four8NoteDataGridViewTextBoxColumn.Name = "Four8NoteDataGridViewTextBoxColumn"
        Me.Four8NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Four9NoteDataGridViewTextBoxColumn
        '
        Me.Four9NoteDataGridViewTextBoxColumn.DataPropertyName = "Four-9Note"
        Me.Four9NoteDataGridViewTextBoxColumn.HeaderText = "Four-9Note"
        Me.Four9NoteDataGridViewTextBoxColumn.Name = "Four9NoteDataGridViewTextBoxColumn"
        Me.Four9NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five1DataGridViewTextBoxColumn
        '
        Me.Five1DataGridViewTextBoxColumn.DataPropertyName = "Five-1"
        Me.Five1DataGridViewTextBoxColumn.HeaderText = "Five-1"
        Me.Five1DataGridViewTextBoxColumn.Name = "Five1DataGridViewTextBoxColumn"
        Me.Five1DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five2DataGridViewTextBoxColumn
        '
        Me.Five2DataGridViewTextBoxColumn.DataPropertyName = "Five-2"
        Me.Five2DataGridViewTextBoxColumn.HeaderText = "Five-2"
        Me.Five2DataGridViewTextBoxColumn.Name = "Five2DataGridViewTextBoxColumn"
        Me.Five2DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five3DataGridViewTextBoxColumn
        '
        Me.Five3DataGridViewTextBoxColumn.DataPropertyName = "Five-3"
        Me.Five3DataGridViewTextBoxColumn.HeaderText = "Five-3"
        Me.Five3DataGridViewTextBoxColumn.Name = "Five3DataGridViewTextBoxColumn"
        Me.Five3DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five4DataGridViewTextBoxColumn
        '
        Me.Five4DataGridViewTextBoxColumn.DataPropertyName = "Five-4"
        Me.Five4DataGridViewTextBoxColumn.HeaderText = "Five-4"
        Me.Five4DataGridViewTextBoxColumn.Name = "Five4DataGridViewTextBoxColumn"
        Me.Five4DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five5DataGridViewTextBoxColumn
        '
        Me.Five5DataGridViewTextBoxColumn.DataPropertyName = "Five-5"
        Me.Five5DataGridViewTextBoxColumn.HeaderText = "Five-5"
        Me.Five5DataGridViewTextBoxColumn.Name = "Five5DataGridViewTextBoxColumn"
        Me.Five5DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five6DataGridViewTextBoxColumn
        '
        Me.Five6DataGridViewTextBoxColumn.DataPropertyName = "Five-6"
        Me.Five6DataGridViewTextBoxColumn.HeaderText = "Five-6"
        Me.Five6DataGridViewTextBoxColumn.Name = "Five6DataGridViewTextBoxColumn"
        Me.Five6DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five7DataGridViewTextBoxColumn
        '
        Me.Five7DataGridViewTextBoxColumn.DataPropertyName = "Five-7"
        Me.Five7DataGridViewTextBoxColumn.HeaderText = "Five-7"
        Me.Five7DataGridViewTextBoxColumn.Name = "Five7DataGridViewTextBoxColumn"
        Me.Five7DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five8DataGridViewTextBoxColumn
        '
        Me.Five8DataGridViewTextBoxColumn.DataPropertyName = "Five-8"
        Me.Five8DataGridViewTextBoxColumn.HeaderText = "Five-8"
        Me.Five8DataGridViewTextBoxColumn.Name = "Five8DataGridViewTextBoxColumn"
        Me.Five8DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five9DataGridViewTextBoxColumn
        '
        Me.Five9DataGridViewTextBoxColumn.DataPropertyName = "Five-9"
        Me.Five9DataGridViewTextBoxColumn.HeaderText = "Five-9"
        Me.Five9DataGridViewTextBoxColumn.Name = "Five9DataGridViewTextBoxColumn"
        Me.Five9DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five1NoteDataGridViewTextBoxColumn
        '
        Me.Five1NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-1Note"
        Me.Five1NoteDataGridViewTextBoxColumn.HeaderText = "Five-1Note"
        Me.Five1NoteDataGridViewTextBoxColumn.Name = "Five1NoteDataGridViewTextBoxColumn"
        Me.Five1NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five2NoteDataGridViewTextBoxColumn
        '
        Me.Five2NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-2Note"
        Me.Five2NoteDataGridViewTextBoxColumn.HeaderText = "Five-2Note"
        Me.Five2NoteDataGridViewTextBoxColumn.Name = "Five2NoteDataGridViewTextBoxColumn"
        Me.Five2NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five3NoteDataGridViewTextBoxColumn
        '
        Me.Five3NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-3Note"
        Me.Five3NoteDataGridViewTextBoxColumn.HeaderText = "Five-3Note"
        Me.Five3NoteDataGridViewTextBoxColumn.Name = "Five3NoteDataGridViewTextBoxColumn"
        Me.Five3NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five4NoteDataGridViewTextBoxColumn
        '
        Me.Five4NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-4Note"
        Me.Five4NoteDataGridViewTextBoxColumn.HeaderText = "Five-4Note"
        Me.Five4NoteDataGridViewTextBoxColumn.Name = "Five4NoteDataGridViewTextBoxColumn"
        Me.Five4NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five5NoteDataGridViewTextBoxColumn
        '
        Me.Five5NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-5Note"
        Me.Five5NoteDataGridViewTextBoxColumn.HeaderText = "Five-5Note"
        Me.Five5NoteDataGridViewTextBoxColumn.Name = "Five5NoteDataGridViewTextBoxColumn"
        Me.Five5NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five6NoteDataGridViewTextBoxColumn
        '
        Me.Five6NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-6Note"
        Me.Five6NoteDataGridViewTextBoxColumn.HeaderText = "Five-6Note"
        Me.Five6NoteDataGridViewTextBoxColumn.Name = "Five6NoteDataGridViewTextBoxColumn"
        Me.Five6NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five7NoteDataGridViewTextBoxColumn
        '
        Me.Five7NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-7Note"
        Me.Five7NoteDataGridViewTextBoxColumn.HeaderText = "Five-7Note"
        Me.Five7NoteDataGridViewTextBoxColumn.Name = "Five7NoteDataGridViewTextBoxColumn"
        Me.Five7NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five8NoteDataGridViewTextBoxColumn
        '
        Me.Five8NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-8Note"
        Me.Five8NoteDataGridViewTextBoxColumn.HeaderText = "Five-8Note"
        Me.Five8NoteDataGridViewTextBoxColumn.Name = "Five8NoteDataGridViewTextBoxColumn"
        Me.Five8NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Five9NoteDataGridViewTextBoxColumn
        '
        Me.Five9NoteDataGridViewTextBoxColumn.DataPropertyName = "Five-9Note"
        Me.Five9NoteDataGridViewTextBoxColumn.HeaderText = "Five-9Note"
        Me.Five9NoteDataGridViewTextBoxColumn.Name = "Five9NoteDataGridViewTextBoxColumn"
        Me.Five9NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six1DataGridViewTextBoxColumn
        '
        Me.Six1DataGridViewTextBoxColumn.DataPropertyName = "Six-1"
        Me.Six1DataGridViewTextBoxColumn.HeaderText = "Six-1"
        Me.Six1DataGridViewTextBoxColumn.Name = "Six1DataGridViewTextBoxColumn"
        Me.Six1DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six2DataGridViewTextBoxColumn
        '
        Me.Six2DataGridViewTextBoxColumn.DataPropertyName = "Six-2"
        Me.Six2DataGridViewTextBoxColumn.HeaderText = "Six-2"
        Me.Six2DataGridViewTextBoxColumn.Name = "Six2DataGridViewTextBoxColumn"
        Me.Six2DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six3DataGridViewTextBoxColumn
        '
        Me.Six3DataGridViewTextBoxColumn.DataPropertyName = "Six-3"
        Me.Six3DataGridViewTextBoxColumn.HeaderText = "Six-3"
        Me.Six3DataGridViewTextBoxColumn.Name = "Six3DataGridViewTextBoxColumn"
        Me.Six3DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six4DataGridViewTextBoxColumn
        '
        Me.Six4DataGridViewTextBoxColumn.DataPropertyName = "Six-4"
        Me.Six4DataGridViewTextBoxColumn.HeaderText = "Six-4"
        Me.Six4DataGridViewTextBoxColumn.Name = "Six4DataGridViewTextBoxColumn"
        Me.Six4DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six5DataGridViewTextBoxColumn
        '
        Me.Six5DataGridViewTextBoxColumn.DataPropertyName = "Six-5"
        Me.Six5DataGridViewTextBoxColumn.HeaderText = "Six-5"
        Me.Six5DataGridViewTextBoxColumn.Name = "Six5DataGridViewTextBoxColumn"
        Me.Six5DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six6DataGridViewTextBoxColumn
        '
        Me.Six6DataGridViewTextBoxColumn.DataPropertyName = "Six-6"
        Me.Six6DataGridViewTextBoxColumn.HeaderText = "Six-6"
        Me.Six6DataGridViewTextBoxColumn.Name = "Six6DataGridViewTextBoxColumn"
        Me.Six6DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six7DataGridViewTextBoxColumn
        '
        Me.Six7DataGridViewTextBoxColumn.DataPropertyName = "Six-7"
        Me.Six7DataGridViewTextBoxColumn.HeaderText = "Six-7"
        Me.Six7DataGridViewTextBoxColumn.Name = "Six7DataGridViewTextBoxColumn"
        Me.Six7DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six8DataGridViewTextBoxColumn
        '
        Me.Six8DataGridViewTextBoxColumn.DataPropertyName = "Six-8"
        Me.Six8DataGridViewTextBoxColumn.HeaderText = "Six-8"
        Me.Six8DataGridViewTextBoxColumn.Name = "Six8DataGridViewTextBoxColumn"
        Me.Six8DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six9DataGridViewTextBoxColumn
        '
        Me.Six9DataGridViewTextBoxColumn.DataPropertyName = "Six-9"
        Me.Six9DataGridViewTextBoxColumn.HeaderText = "Six-9"
        Me.Six9DataGridViewTextBoxColumn.Name = "Six9DataGridViewTextBoxColumn"
        Me.Six9DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six1NoteDataGridViewTextBoxColumn
        '
        Me.Six1NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-1Note"
        Me.Six1NoteDataGridViewTextBoxColumn.HeaderText = "Six-1Note"
        Me.Six1NoteDataGridViewTextBoxColumn.Name = "Six1NoteDataGridViewTextBoxColumn"
        Me.Six1NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six2NoteDataGridViewTextBoxColumn
        '
        Me.Six2NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-2Note"
        Me.Six2NoteDataGridViewTextBoxColumn.HeaderText = "Six-2Note"
        Me.Six2NoteDataGridViewTextBoxColumn.Name = "Six2NoteDataGridViewTextBoxColumn"
        Me.Six2NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six3NoteDataGridViewTextBoxColumn
        '
        Me.Six3NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-3Note"
        Me.Six3NoteDataGridViewTextBoxColumn.HeaderText = "Six-3Note"
        Me.Six3NoteDataGridViewTextBoxColumn.Name = "Six3NoteDataGridViewTextBoxColumn"
        Me.Six3NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six4NoteDataGridViewTextBoxColumn
        '
        Me.Six4NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-4Note"
        Me.Six4NoteDataGridViewTextBoxColumn.HeaderText = "Six-4Note"
        Me.Six4NoteDataGridViewTextBoxColumn.Name = "Six4NoteDataGridViewTextBoxColumn"
        Me.Six4NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six5NoteDataGridViewTextBoxColumn
        '
        Me.Six5NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-5Note"
        Me.Six5NoteDataGridViewTextBoxColumn.HeaderText = "Six-5Note"
        Me.Six5NoteDataGridViewTextBoxColumn.Name = "Six5NoteDataGridViewTextBoxColumn"
        Me.Six5NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six6NoteDataGridViewTextBoxColumn
        '
        Me.Six6NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-6Note"
        Me.Six6NoteDataGridViewTextBoxColumn.HeaderText = "Six-6Note"
        Me.Six6NoteDataGridViewTextBoxColumn.Name = "Six6NoteDataGridViewTextBoxColumn"
        Me.Six6NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six7NoteDataGridViewTextBoxColumn
        '
        Me.Six7NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-7Note"
        Me.Six7NoteDataGridViewTextBoxColumn.HeaderText = "Six-7Note"
        Me.Six7NoteDataGridViewTextBoxColumn.Name = "Six7NoteDataGridViewTextBoxColumn"
        Me.Six7NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six8NoteDataGridViewTextBoxColumn
        '
        Me.Six8NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-8Note"
        Me.Six8NoteDataGridViewTextBoxColumn.HeaderText = "Six-8Note"
        Me.Six8NoteDataGridViewTextBoxColumn.Name = "Six8NoteDataGridViewTextBoxColumn"
        Me.Six8NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Six9NoteDataGridViewTextBoxColumn
        '
        Me.Six9NoteDataGridViewTextBoxColumn.DataPropertyName = "Six-9Note"
        Me.Six9NoteDataGridViewTextBoxColumn.HeaderText = "Six-9Note"
        Me.Six9NoteDataGridViewTextBoxColumn.Name = "Six9NoteDataGridViewTextBoxColumn"
        Me.Six9NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven1DataGridViewTextBoxColumn
        '
        Me.Seven1DataGridViewTextBoxColumn.DataPropertyName = "Seven-1"
        Me.Seven1DataGridViewTextBoxColumn.HeaderText = "Seven-1"
        Me.Seven1DataGridViewTextBoxColumn.Name = "Seven1DataGridViewTextBoxColumn"
        Me.Seven1DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven2DataGridViewTextBoxColumn
        '
        Me.Seven2DataGridViewTextBoxColumn.DataPropertyName = "Seven-2"
        Me.Seven2DataGridViewTextBoxColumn.HeaderText = "Seven-2"
        Me.Seven2DataGridViewTextBoxColumn.Name = "Seven2DataGridViewTextBoxColumn"
        Me.Seven2DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven3DataGridViewTextBoxColumn
        '
        Me.Seven3DataGridViewTextBoxColumn.DataPropertyName = "Seven-3"
        Me.Seven3DataGridViewTextBoxColumn.HeaderText = "Seven-3"
        Me.Seven3DataGridViewTextBoxColumn.Name = "Seven3DataGridViewTextBoxColumn"
        Me.Seven3DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven4DataGridViewTextBoxColumn
        '
        Me.Seven4DataGridViewTextBoxColumn.DataPropertyName = "Seven-4"
        Me.Seven4DataGridViewTextBoxColumn.HeaderText = "Seven-4"
        Me.Seven4DataGridViewTextBoxColumn.Name = "Seven4DataGridViewTextBoxColumn"
        Me.Seven4DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven5DataGridViewTextBoxColumn
        '
        Me.Seven5DataGridViewTextBoxColumn.DataPropertyName = "Seven-5"
        Me.Seven5DataGridViewTextBoxColumn.HeaderText = "Seven-5"
        Me.Seven5DataGridViewTextBoxColumn.Name = "Seven5DataGridViewTextBoxColumn"
        Me.Seven5DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven6DataGridViewTextBoxColumn
        '
        Me.Seven6DataGridViewTextBoxColumn.DataPropertyName = "Seven-6"
        Me.Seven6DataGridViewTextBoxColumn.HeaderText = "Seven-6"
        Me.Seven6DataGridViewTextBoxColumn.Name = "Seven6DataGridViewTextBoxColumn"
        Me.Seven6DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven7DataGridViewTextBoxColumn
        '
        Me.Seven7DataGridViewTextBoxColumn.DataPropertyName = "Seven-7"
        Me.Seven7DataGridViewTextBoxColumn.HeaderText = "Seven-7"
        Me.Seven7DataGridViewTextBoxColumn.Name = "Seven7DataGridViewTextBoxColumn"
        Me.Seven7DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven8DataGridViewTextBoxColumn
        '
        Me.Seven8DataGridViewTextBoxColumn.DataPropertyName = "Seven-8"
        Me.Seven8DataGridViewTextBoxColumn.HeaderText = "Seven-8"
        Me.Seven8DataGridViewTextBoxColumn.Name = "Seven8DataGridViewTextBoxColumn"
        Me.Seven8DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven9DataGridViewTextBoxColumn
        '
        Me.Seven9DataGridViewTextBoxColumn.DataPropertyName = "Seven-9"
        Me.Seven9DataGridViewTextBoxColumn.HeaderText = "Seven-9"
        Me.Seven9DataGridViewTextBoxColumn.Name = "Seven9DataGridViewTextBoxColumn"
        Me.Seven9DataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven1NoteDataGridViewTextBoxColumn
        '
        Me.Seven1NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-1Note"
        Me.Seven1NoteDataGridViewTextBoxColumn.HeaderText = "Seven-1Note"
        Me.Seven1NoteDataGridViewTextBoxColumn.Name = "Seven1NoteDataGridViewTextBoxColumn"
        Me.Seven1NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven2NoteDataGridViewTextBoxColumn
        '
        Me.Seven2NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-2Note"
        Me.Seven2NoteDataGridViewTextBoxColumn.HeaderText = "Seven-2Note"
        Me.Seven2NoteDataGridViewTextBoxColumn.Name = "Seven2NoteDataGridViewTextBoxColumn"
        Me.Seven2NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven3NoteDataGridViewTextBoxColumn
        '
        Me.Seven3NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-3Note"
        Me.Seven3NoteDataGridViewTextBoxColumn.HeaderText = "Seven-3Note"
        Me.Seven3NoteDataGridViewTextBoxColumn.Name = "Seven3NoteDataGridViewTextBoxColumn"
        Me.Seven3NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven4NoteDataGridViewTextBoxColumn
        '
        Me.Seven4NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-4Note"
        Me.Seven4NoteDataGridViewTextBoxColumn.HeaderText = "Seven-4Note"
        Me.Seven4NoteDataGridViewTextBoxColumn.Name = "Seven4NoteDataGridViewTextBoxColumn"
        Me.Seven4NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven5NoteDataGridViewTextBoxColumn
        '
        Me.Seven5NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-5Note"
        Me.Seven5NoteDataGridViewTextBoxColumn.HeaderText = "Seven-5Note"
        Me.Seven5NoteDataGridViewTextBoxColumn.Name = "Seven5NoteDataGridViewTextBoxColumn"
        Me.Seven5NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven6NoteDataGridViewTextBoxColumn
        '
        Me.Seven6NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-6Note"
        Me.Seven6NoteDataGridViewTextBoxColumn.HeaderText = "Seven-6Note"
        Me.Seven6NoteDataGridViewTextBoxColumn.Name = "Seven6NoteDataGridViewTextBoxColumn"
        Me.Seven6NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven7NoteDataGridViewTextBoxColumn
        '
        Me.Seven7NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-7Note"
        Me.Seven7NoteDataGridViewTextBoxColumn.HeaderText = "Seven-7Note"
        Me.Seven7NoteDataGridViewTextBoxColumn.Name = "Seven7NoteDataGridViewTextBoxColumn"
        Me.Seven7NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven8NoteDataGridViewTextBoxColumn
        '
        Me.Seven8NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-8Note"
        Me.Seven8NoteDataGridViewTextBoxColumn.HeaderText = "Seven-8Note"
        Me.Seven8NoteDataGridViewTextBoxColumn.Name = "Seven8NoteDataGridViewTextBoxColumn"
        Me.Seven8NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Seven9NoteDataGridViewTextBoxColumn
        '
        Me.Seven9NoteDataGridViewTextBoxColumn.DataPropertyName = "Seven-9Note"
        Me.Seven9NoteDataGridViewTextBoxColumn.HeaderText = "Seven-9Note"
        Me.Seven9NoteDataGridViewTextBoxColumn.Name = "Seven9NoteDataGridViewTextBoxColumn"
        Me.Seven9NoteDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QAScoreDataGridViewTextBoxColumn
        '
        Me.QAScoreDataGridViewTextBoxColumn.DataPropertyName = "QAScore"
        Me.QAScoreDataGridViewTextBoxColumn.HeaderText = "QAScore"
        Me.QAScoreDataGridViewTextBoxColumn.Name = "QAScoreDataGridViewTextBoxColumn"
        Me.QAScoreDataGridViewTextBoxColumn.ReadOnly = True
        '
        'JIRADataGridViewTextBoxColumn
        '
        Me.JIRADataGridViewTextBoxColumn.DataPropertyName = "JIRA"
        Me.JIRADataGridViewTextBoxColumn.HeaderText = "JIRA"
        Me.JIRADataGridViewTextBoxColumn.Name = "JIRADataGridViewTextBoxColumn"
        Me.JIRADataGridViewTextBoxColumn.ReadOnly = True
        '
        'UserIDDataGridViewTextBoxColumn
        '
        Me.UserIDDataGridViewTextBoxColumn.DataPropertyName = "UserID"
        Me.UserIDDataGridViewTextBoxColumn.HeaderText = "UserID"
        Me.UserIDDataGridViewTextBoxColumn.Name = "UserIDDataGridViewTextBoxColumn"
        Me.UserIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'AutoFailDataGridViewTextBoxColumn
        '
        Me.AutoFailDataGridViewTextBoxColumn.DataPropertyName = "AutoFail"
        Me.AutoFailDataGridViewTextBoxColumn.HeaderText = "AutoFail"
        Me.AutoFailDataGridViewTextBoxColumn.Name = "AutoFailDataGridViewTextBoxColumn"
        Me.AutoFailDataGridViewTextBoxColumn.ReadOnly = True
        '
        'AuditorDataGridViewTextBoxColumn
        '
        Me.AuditorDataGridViewTextBoxColumn.DataPropertyName = "Auditor"
        Me.AuditorDataGridViewTextBoxColumn.HeaderText = "Auditor"
        Me.AuditorDataGridViewTextBoxColumn.Name = "AuditorDataGridViewTextBoxColumn"
        Me.AuditorDataGridViewTextBoxColumn.ReadOnly = True
        '
        'QADash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1017, 471)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.GroupBox8)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "QADash"
        Me.Text = "QADash"
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QADataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QAMainDBBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox8 As GroupBox
    Friend WithEvents cboTeamName As ComboBox
    Friend WithEvents lblQAauditor As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents cboContactType As ComboBox
    Friend WithEvents cboAgentName As ComboBox
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents QADataSet As QADataSet
    Friend WithEvents QAMainDBBindingSource As BindingSource
    Friend WithEvents QAMainDBTableAdapter As QADataSetTableAdapters.QAMainDBTableAdapter
    Friend WithEvents IDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents SRDataGridViewTextBoxColumn As DataGridViewLinkColumn
    Friend WithEvents ContactIDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CtypeDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents QAAgentDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents QATeamDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents QAContactDateDataGridViewTextBoxColumn As DataGridViewAutoFilter.DataGridViewAutoFilterTextBoxColumn
    Friend WithEvents QAOrderIDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents QADateDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents QACommentsDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents QAOppDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CINameDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CIAccountDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CICompanyDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CIPhoneDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CIEmailDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents RevDateDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents RevManagerDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents RevCommentsDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DisScoreDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DisNameDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DisNotesDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DisAppCommentsDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One4DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One5DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One6DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One7DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One8DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One9DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One1NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One2NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One3NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One4NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One5NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One6NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One7NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One8NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents One9NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two4DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two5DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two6DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two7DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two8DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two9DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two1NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two2NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two3NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two4NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two5NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two6NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two7NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two8NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Two9NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three4DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three5DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three6DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three7DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three8DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three9DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three1NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three2NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three3NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three4NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three5NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three6NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three7NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three8NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Three9NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four4DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four5DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four6DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four7DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four8DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four9DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four1NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four2NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four3NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four4NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four5NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four6NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four7NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four8NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Four9NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five4DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five5DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five6DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five7DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five8DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five9DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five1NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five2NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five3NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five4NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five5NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five6NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five7NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five8NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Five9NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six4DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six5DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six6DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six7DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six8DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six9DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six1NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six2NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six3NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six4NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six5NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six6NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six7NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six8NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Six9NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven4DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven5DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven6DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven7DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven8DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven9DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven1NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven2NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven3NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven4NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven5NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven6NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven7NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven8NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Seven9NoteDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents QAScoreDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents JIRADataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents UserIDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents AutoFailDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents AuditorDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
End Class
