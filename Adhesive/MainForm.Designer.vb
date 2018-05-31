<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.labClosedWindows = New System.Windows.Forms.Label()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.lbClosedForms = New System.Windows.Forms.CheckedListBox()
        Me.btnAddForm = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.btnConfig = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnForce = New System.Windows.Forms.Button()
        Me.btnPresets = New System.Windows.Forms.Button()
        Me.btnGapSize = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.rbSnapToAll = New System.Windows.Forms.RadioButton()
        Me.rbMoveAlong = New System.Windows.Forms.RadioButton()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel1.Padding = New System.Windows.Forms.Padding(6, 3, 3, 3)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Panel2.Padding = New System.Windows.Forms.Padding(3, 3, 6, 3)
        Me.SplitContainer1.Size = New System.Drawing.Size(408, 362)
        Me.SplitContainer1.SplitterDistance = 205
        Me.SplitContainer1.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.labClosedWindows)
        Me.GroupBox1.Controls.Add(Me.btnReset)
        Me.GroupBox1.Controls.Add(Me.lbClosedForms)
        Me.GroupBox1.Controls.Add(Me.btnAddForm)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(6, 3)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.GroupBox1.MinimumSize = New System.Drawing.Size(128, 107)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(196, 356)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Test Windows"
        '
        'labClosedWindows
        '
        Me.labClosedWindows.AutoSize = True
        Me.labClosedWindows.ForeColor = System.Drawing.SystemColors.GrayText
        Me.labClosedWindows.Location = New System.Drawing.Point(6, 49)
        Me.labClosedWindows.Name = "labClosedWindows"
        Me.labClosedWindows.Size = New System.Drawing.Size(89, 13)
        Me.labClosedWindows.TabIndex = 5
        Me.labClosedWindows.Text = "Closed Windows:"
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(45, 330)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(106, 23)
        Me.btnReset.TabIndex = 4
        Me.btnReset.Text = "Reset"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'lbClosedForms
        '
        Me.lbClosedForms.FormattingEnabled = True
        Me.lbClosedForms.HorizontalScrollbar = True
        Me.lbClosedForms.IntegralHeight = False
        Me.lbClosedForms.Location = New System.Drawing.Point(6, 65)
        Me.lbClosedForms.Name = "lbClosedForms"
        Me.lbClosedForms.Size = New System.Drawing.Size(176, 259)
        Me.lbClosedForms.TabIndex = 3
        '
        'btnAddForm
        '
        Me.btnAddForm.Location = New System.Drawing.Point(45, 19)
        Me.btnAddForm.Name = "btnAddForm"
        Me.btnAddForm.Size = New System.Drawing.Size(106, 23)
        Me.btnAddForm.TabIndex = 2
        Me.btnAddForm.Text = "Add Window"
        Me.btnAddForm.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.GroupBox5)
        Me.GroupBox2.Controls.Add(Me.GroupBox4)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(190, 356)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Adhesive Settings"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.btnConfig)
        Me.GroupBox5.Location = New System.Drawing.Point(7, 218)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(178, 58)
        Me.GroupBox5.TabIndex = 11
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "User Settings"
        '
        'btnConfig
        '
        Me.btnConfig.Location = New System.Drawing.Point(38, 19)
        Me.btnConfig.Name = "btnConfig"
        Me.btnConfig.Size = New System.Drawing.Size(104, 23)
        Me.btnConfig.TabIndex = 7
        Me.btnConfig.Text = "Edit Config"
        Me.btnConfig.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnForce)
        Me.GroupBox4.Controls.Add(Me.btnPresets)
        Me.GroupBox4.Controls.Add(Me.btnGapSize)
        Me.GroupBox4.Location = New System.Drawing.Point(7, 97)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(178, 115)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Tool Windows"
        '
        'btnForce
        '
        Me.btnForce.Location = New System.Drawing.Point(38, 77)
        Me.btnForce.Name = "btnForce"
        Me.btnForce.Size = New System.Drawing.Size(104, 23)
        Me.btnForce.TabIndex = 8
        Me.btnForce.Text = "Change Force"
        Me.btnForce.UseVisualStyleBackColor = True
        '
        'btnPresets
        '
        Me.btnPresets.Location = New System.Drawing.Point(38, 19)
        Me.btnPresets.Name = "btnPresets"
        Me.btnPresets.Size = New System.Drawing.Size(104, 23)
        Me.btnPresets.TabIndex = 7
        Me.btnPresets.Text = "Manage Presets"
        Me.btnPresets.UseVisualStyleBackColor = True
        '
        'btnGapSize
        '
        Me.btnGapSize.Location = New System.Drawing.Point(38, 48)
        Me.btnGapSize.Name = "btnGapSize"
        Me.btnGapSize.Size = New System.Drawing.Size(104, 23)
        Me.btnGapSize.TabIndex = 6
        Me.btnGapSize.Text = "Change Gap Size"
        Me.btnGapSize.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.rbSnapToAll)
        Me.GroupBox3.Controls.Add(Me.rbMoveAlong)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 20)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(179, 71)
        Me.GroupBox3.TabIndex = 9
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "On Moves / Resizes"
        '
        'rbSnapToAll
        '
        Me.rbSnapToAll.AutoSize = True
        Me.rbSnapToAll.Location = New System.Drawing.Point(39, 37)
        Me.rbSnapToAll.Name = "rbSnapToAll"
        Me.rbSnapToAll.Size = New System.Drawing.Size(100, 17)
        Me.rbSnapToAll.TabIndex = 10
        Me.rbSnapToAll.Text = "Snap To Others"
        Me.rbSnapToAll.UseVisualStyleBackColor = True
        '
        'rbMoveAlong
        '
        Me.rbMoveAlong.AutoSize = True
        Me.rbMoveAlong.Checked = True
        Me.rbMoveAlong.Location = New System.Drawing.Point(39, 19)
        Me.rbMoveAlong.Name = "rbMoveAlong"
        Me.rbMoveAlong.Size = New System.Drawing.Size(96, 17)
        Me.rbMoveAlong.TabIndex = 9
        Me.rbMoveAlong.TabStop = True
        Me.rbMoveAlong.Text = "Move All Along"
        Me.rbMoveAlong.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(414, 368)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MainForm"
        Me.Padding = New System.Windows.Forms.Padding(3)
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Adhesive Windows"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lbClosedForms As CheckedListBox
    Friend WithEvents btnAddForm As Button
    Friend WithEvents btnReset As Button
    Friend WithEvents labClosedWindows As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents btnGapSize As Button
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents rbSnapToAll As RadioButton
    Friend WithEvents rbMoveAlong As RadioButton
    Friend WithEvents btnPresets As Button
    Friend WithEvents btnForce As Button
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents btnConfig As Button
End Class
