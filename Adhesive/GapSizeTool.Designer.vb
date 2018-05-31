<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class GapSizeTool
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GapSizeTool))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnLessGap = New System.Windows.Forms.Button()
        Me.btnMoreGap = New System.Windows.Forms.Button()
        Me.rbSetGapXY = New System.Windows.Forms.RadioButton()
        Me.rbSetGapY = New System.Windows.Forms.RadioButton()
        Me.rbSetGapX = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 135)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(187, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "VisualBounds:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnLessGap)
        Me.GroupBox1.Controls.Add(Me.btnMoreGap)
        Me.GroupBox1.Controls.Add(Me.rbSetGapXY)
        Me.GroupBox1.Controls.Add(Me.rbSetGapY)
        Me.GroupBox1.Controls.Add(Me.rbSetGapX)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(6)
        Me.GroupBox1.Size = New System.Drawing.Size(187, 129)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Change Gap"
        '
        'btnLessGap
        '
        Me.btnLessGap.Location = New System.Drawing.Point(22, 86)
        Me.btnLessGap.Name = "btnLessGap"
        Me.btnLessGap.Size = New System.Drawing.Size(75, 23)
        Me.btnLessGap.TabIndex = 4
        Me.btnLessGap.Text = "- Less"
        Me.btnLessGap.UseVisualStyleBackColor = True
        '
        'btnMoreGap
        '
        Me.btnMoreGap.Location = New System.Drawing.Point(103, 86)
        Me.btnMoreGap.Name = "btnMoreGap"
        Me.btnMoreGap.Size = New System.Drawing.Size(75, 23)
        Me.btnMoreGap.TabIndex = 3
        Me.btnMoreGap.Text = "+ More"
        Me.btnMoreGap.UseVisualStyleBackColor = True
        '
        'rbSetGapXY
        '
        Me.rbSetGapXY.AutoSize = True
        Me.rbSetGapXY.Checked = True
        Me.rbSetGapXY.Location = New System.Drawing.Point(50, 63)
        Me.rbSetGapXY.Name = "rbSetGapXY"
        Me.rbSetGapXY.Size = New System.Drawing.Size(82, 17)
        Me.rbSetGapXY.TabIndex = 2
        Me.rbSetGapXY.TabStop = True
        Me.rbSetGapXY.Text = "Both  (xx,yy)"
        Me.rbSetGapXY.UseVisualStyleBackColor = True
        '
        'rbSetGapY
        '
        Me.rbSetGapY.AutoSize = True
        Me.rbSetGapY.Location = New System.Drawing.Point(50, 44)
        Me.rbSetGapY.Name = "rbSetGapY"
        Me.rbSetGapY.Size = New System.Drawing.Size(82, 17)
        Me.rbSetGapY.TabIndex = 1
        Me.rbSetGapY.Text = "Vertical  (yy)"
        Me.rbSetGapY.UseVisualStyleBackColor = True
        '
        'rbSetGapX
        '
        Me.rbSetGapX.AutoSize = True
        Me.rbSetGapX.Location = New System.Drawing.Point(50, 25)
        Me.rbSetGapX.Name = "rbSetGapX"
        Me.rbSetGapX.Size = New System.Drawing.Size(91, 17)
        Me.rbSetGapX.TabIndex = 0
        Me.rbSetGapX.Text = "Horizontal (xx)"
        Me.rbSetGapX.UseVisualStyleBackColor = True
        '
        'GapSizeTool
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(199, 158)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(211, 191)
        Me.Name = "GapSizeTool"
        Me.Padding = New System.Windows.Forms.Padding(6)
        Me.ShowInTaskbar = False
        Me.Text = "Gap Size Tool"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnLessGap As Button
    Friend WithEvents btnMoreGap As Button
    Friend WithEvents rbSetGapXY As RadioButton
    Friend WithEvents rbSetGapY As RadioButton
    Friend WithEvents rbSetGapX As RadioButton
End Class
