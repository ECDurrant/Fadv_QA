﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Splash
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Splash))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Timer333 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer222 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.QAapp.My.Resources.Resources.FA_master_v4
        Me.PictureBox1.Location = New System.Drawing.Point(2, 1)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(576, 157)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 8
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Lucida Bright", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(-1, 161)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(177, 14)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Developed By: Eric Durrant"
        Me.Label1.Visible = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(449, 161)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(129, 12)
        Me.ProgressBar1.Step = 1
        Me.ProgressBar1.TabIndex = 10
        Me.ProgressBar1.Visible = False
        '
        'Timer333
        '
        Me.Timer333.Enabled = True
        Me.Timer333.Interval = 3000
        '
        'Timer222
        '
        Me.Timer222.Interval = 600000
        '
        'Splash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(583, 179)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Splash"
        Me.Text = "Splash"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Timer333 As System.Windows.Forms.Timer
    Friend WithEvents Timer222 As System.Windows.Forms.Timer
End Class
