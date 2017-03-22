<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form3
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.GroupBox11 = New System.Windows.Forms.GroupBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.txtReviewComments = New System.Windows.Forms.TextBox()
        Me.cboReviewManger = New System.Windows.Forms.ComboBox()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.txtDisputeScore = New System.Windows.Forms.TextBox()
        Me.txtDisputeApprovalScore = New System.Windows.Forms.TextBox()
        Me.Label50 = New System.Windows.Forms.Label()
        Me.txtDisputerName = New System.Windows.Forms.TextBox()
        Me.txtDisputeAppComments = New System.Windows.Forms.TextBox()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.txtDisputeComments = New System.Windows.Forms.TextBox()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.txtComments2 = New System.Windows.Forms.TextBox()
        Me.txtCommentes = New System.Windows.Forms.TextBox()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.GroupBox11.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.GroupBox10.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label51
        '
        Me.Label51.AutoSize = True
        Me.Label51.Font = New System.Drawing.Font("Lucida Bright", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label51.ForeColor = System.Drawing.Color.Navy
        Me.Label51.Location = New System.Drawing.Point(142, 371)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(54, 14)
        Me.Label51.TabIndex = 43
        Me.Label51.Text = "Dispute:"
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Font = New System.Drawing.Font("Lucida Bright", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label42.ForeColor = System.Drawing.Color.Navy
        Me.Label42.Location = New System.Drawing.Point(6, 0)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(49, 14)
        Me.Label42.TabIndex = 40
        Me.Label42.Text = "Review:"
        '
        'GroupBox11
        '
        Me.GroupBox11.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox11.Controls.Add(Me.Label40)
        Me.GroupBox11.Controls.Add(Me.Label46)
        Me.GroupBox11.Controls.Add(Me.Label42)
        Me.GroupBox11.Controls.Add(Me.txtReviewComments)
        Me.GroupBox11.Controls.Add(Me.cboReviewManger)
        Me.GroupBox11.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox11.Location = New System.Drawing.Point(62, 205)
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.Size = New System.Drawing.Size(639, 163)
        Me.GroupBox11.TabIndex = 41
        Me.GroupBox11.TabStop = False
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Font = New System.Drawing.Font("Lucida Bright", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(6, 62)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(113, 14)
        Me.Label40.TabIndex = 29
        Me.Label40.Text = "Review Comments:"
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(6, 24)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(123, 15)
        Me.Label46.TabIndex = 28
        Me.Label46.Text = "Review Manager:"
        '
        'txtReviewComments
        '
        Me.txtReviewComments.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReviewComments.Location = New System.Drawing.Point(9, 80)
        Me.txtReviewComments.Multiline = True
        Me.txtReviewComments.Name = "txtReviewComments"
        Me.txtReviewComments.Size = New System.Drawing.Size(622, 75)
        Me.txtReviewComments.TabIndex = 54
        '
        'cboReviewManger
        '
        Me.cboReviewManger.BackColor = System.Drawing.SystemColors.MenuBar
        Me.cboReviewManger.FormattingEnabled = True
        Me.cboReviewManger.Items.AddRange(New Object() {"Erica Anderson", "Chris Lipson", "Jared Spicer"})
        Me.cboReviewManger.Location = New System.Drawing.Point(146, 22)
        Me.cboReviewManger.Name = "cboReviewManger"
        Me.cboReviewManger.Size = New System.Drawing.Size(121, 23)
        Me.cboReviewManger.TabIndex = 53
        '
        'GroupBox9
        '
        Me.GroupBox9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox9.Controls.Add(Me.txtDisputeScore)
        Me.GroupBox9.Controls.Add(Me.txtDisputeApprovalScore)
        Me.GroupBox9.Controls.Add(Me.Label50)
        Me.GroupBox9.Controls.Add(Me.txtDisputerName)
        Me.GroupBox9.Controls.Add(Me.txtDisputeAppComments)
        Me.GroupBox9.Controls.Add(Me.Label49)
        Me.GroupBox9.Controls.Add(Me.Label48)
        Me.GroupBox9.Controls.Add(Me.Label45)
        Me.GroupBox9.Controls.Add(Me.Label44)
        Me.GroupBox9.Controls.Add(Me.txtDisputeComments)
        Me.GroupBox9.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox9.Location = New System.Drawing.Point(608, 388)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(582, 316)
        Me.GroupBox9.TabIndex = 42
        Me.GroupBox9.TabStop = False
        '
        'txtDisputeScore
        '
        Me.txtDisputeScore.Location = New System.Drawing.Point(146, 45)
        Me.txtDisputeScore.Multiline = True
        Me.txtDisputeScore.Name = "txtDisputeScore"
        Me.txtDisputeScore.Size = New System.Drawing.Size(88, 23)
        Me.txtDisputeScore.TabIndex = 56
        '
        'txtDisputeApprovalScore
        '
        Me.txtDisputeApprovalScore.Location = New System.Drawing.Point(146, 179)
        Me.txtDisputeApprovalScore.Multiline = True
        Me.txtDisputeApprovalScore.Name = "txtDisputeApprovalScore"
        Me.txtDisputeApprovalScore.Size = New System.Drawing.Size(88, 23)
        Me.txtDisputeApprovalScore.TabIndex = 58
        '
        'Label50
        '
        Me.Label50.AutoSize = True
        Me.Label50.Font = New System.Drawing.Font("Lucida Bright", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label50.Location = New System.Drawing.Point(9, 205)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(175, 14)
        Me.Label50.TabIndex = 35
        Me.Label50.Text = "Dispute Approval Comments:"
        '
        'txtDisputerName
        '
        Me.txtDisputerName.Location = New System.Drawing.Point(146, 16)
        Me.txtDisputerName.Multiline = True
        Me.txtDisputerName.Name = "txtDisputerName"
        Me.txtDisputerName.Size = New System.Drawing.Size(148, 23)
        Me.txtDisputerName.TabIndex = 55
        '
        'txtDisputeAppComments
        '
        Me.txtDisputeAppComments.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDisputeAppComments.Location = New System.Drawing.Point(9, 223)
        Me.txtDisputeAppComments.Multiline = True
        Me.txtDisputeAppComments.Name = "txtDisputeAppComments"
        Me.txtDisputeAppComments.Size = New System.Drawing.Size(565, 75)
        Me.txtDisputeAppComments.TabIndex = 59
        '
        'Label49
        '
        Me.Label49.AutoSize = True
        Me.Label49.Location = New System.Drawing.Point(9, 183)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(131, 15)
        Me.Label49.TabIndex = 31
        Me.Label49.Text = "Dispute Approval:"
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Location = New System.Drawing.Point(6, 24)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(110, 15)
        Me.Label48.TabIndex = 30
        Me.Label48.Text = "Disputer Name:"
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Font = New System.Drawing.Font("Lucida Bright", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label45.Location = New System.Drawing.Point(6, 84)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(87, 14)
        Me.Label45.TabIndex = 29
        Me.Label45.Text = "Dispute Notes:"
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(6, 53)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(107, 15)
        Me.Label44.TabIndex = 28
        Me.Label44.Text = "Dispute Score:"
        '
        'txtDisputeComments
        '
        Me.txtDisputeComments.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDisputeComments.Location = New System.Drawing.Point(9, 102)
        Me.txtDisputeComments.Multiline = True
        Me.txtDisputeComments.Name = "txtDisputeComments"
        Me.txtDisputeComments.Size = New System.Drawing.Size(565, 75)
        Me.txtDisputeComments.TabIndex = 57
        '
        'GroupBox10
        '
        Me.GroupBox10.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox10.Controls.Add(Me.txtComments2)
        Me.GroupBox10.Controls.Add(Me.txtCommentes)
        Me.GroupBox10.Controls.Add(Me.Label41)
        Me.GroupBox10.Controls.Add(Me.Label43)
        Me.GroupBox10.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox10.Location = New System.Drawing.Point(782, 12)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(350, 269)
        Me.GroupBox10.TabIndex = 44
        Me.GroupBox10.TabStop = False
        '
        'txtComments2
        '
        Me.txtComments2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtComments2.Location = New System.Drawing.Point(0, 157)
        Me.txtComments2.Multiline = True
        Me.txtComments2.Name = "txtComments2"
        Me.txtComments2.Size = New System.Drawing.Size(328, 75)
        Me.txtComments2.TabIndex = 52
        '
        'txtCommentes
        '
        Me.txtCommentes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCommentes.Location = New System.Drawing.Point(9, 37)
        Me.txtCommentes.Multiline = True
        Me.txtCommentes.Name = "txtCommentes"
        Me.txtCommentes.Size = New System.Drawing.Size(328, 75)
        Me.txtCommentes.TabIndex = 3
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(6, 136)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(158, 15)
        Me.Label41.TabIndex = 24
        Me.Label41.Text = "Areas of Opportunity:"
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(50, 12)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(114, 15)
        Me.Label43.TabIndex = 0
        Me.Label43.Text = " QA Comments:"
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1202, 716)
        Me.Controls.Add(Me.GroupBox10)
        Me.Controls.Add(Me.Label51)
        Me.Controls.Add(Me.GroupBox11)
        Me.Controls.Add(Me.GroupBox9)
        Me.Name = "Form3"
        Me.Text = "Form3"
        Me.GroupBox11.ResumeLayout(False)
        Me.GroupBox11.PerformLayout()
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox9.PerformLayout()
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox10.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents GroupBox11 As System.Windows.Forms.GroupBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents txtReviewComments As System.Windows.Forms.TextBox
    Friend WithEvents cboReviewManger As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents txtDisputeScore As System.Windows.Forms.TextBox
    Friend WithEvents txtDisputeApprovalScore As System.Windows.Forms.TextBox
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents txtDisputerName As System.Windows.Forms.TextBox
    Friend WithEvents txtDisputeAppComments As System.Windows.Forms.TextBox
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents txtDisputeComments As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents txtComments2 As System.Windows.Forms.TextBox
    Friend WithEvents txtCommentes As System.Windows.Forms.TextBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
End Class
