<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ForceTool
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ForceTool))
        Me.gbFull = New System.Windows.Forms.GroupBox()
        Me.tbForce = New System.Windows.Forms.TrackBar()
        Me.gbFull.SuspendLayout()
        CType(Me.tbForce, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbFull
        '
        Me.gbFull.Controls.Add(Me.tbForce)
        Me.gbFull.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFull.Location = New System.Drawing.Point(6, 6)
        Me.gbFull.Name = "gbFull"
        Me.gbFull.Padding = New System.Windows.Forms.Padding(6)
        Me.gbFull.Size = New System.Drawing.Size(251, 80)
        Me.gbFull.TabIndex = 0
        Me.gbFull.TabStop = False
        Me.gbFull.Text = "Adhesive Force: 12"
        '
        'tbForce
        '
        Me.tbForce.Dock = System.Windows.Forms.DockStyle.Top
        Me.tbForce.Location = New System.Drawing.Point(6, 19)
        Me.tbForce.Maximum = 30
        Me.tbForce.Name = "tbForce"
        Me.tbForce.Size = New System.Drawing.Size(239, 45)
        Me.tbForce.TabIndex = 0
        Me.tbForce.Value = 12
        '
        'ForceTool
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(263, 92)
        Me.Controls.Add(Me.gbFull)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(165, 120)
        Me.Name = "ForceTool"
        Me.Padding = New System.Windows.Forms.Padding(6)
        Me.Text = "Force Tool"
        Me.gbFull.ResumeLayout(False)
        Me.gbFull.PerformLayout()
        CType(Me.tbForce, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents gbFull As GroupBox
    Friend WithEvents tbForce As TrackBar
End Class
