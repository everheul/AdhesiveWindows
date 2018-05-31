Imports AdhesiveWindows

Public Class GapSizeTool

    '----- start Adhesive code -----
    Friend adhesive As Adhesive
    Protected Overrides Sub WndProc(ByRef m As Message)
        If (adhesive Is Nothing) Then
            adhesive = New Adhesive(Me)
        End If
        If (adhesive.FilterMessage(m)) Then
            MyBase.WndProc(m)
        End If
    End Sub
    Public Overloads Sub Show()
        Adhesive.OpenToolForm(Me)
    End Sub
    '----- end Adhesive code -----

    Private Sub ToolForm_Shown() Handles Me.Shown
        SetInfoText()
        SetGapText()
    End Sub

    Private Sub ToolForm_Move() Handles Me.Move
        SetInfoText()
    End Sub

    Private Sub ToolForm_Resize() Handles Me.Resize
        SetInfoText()
    End Sub

    Private Sub SetInfoText()
        If (adhesive IsNot Nothing) Then
            If (adhesive.Loaded) Then
                Dim rect As Rectangle = adhesive.GetVisualBounds()
                Label1.Text = "size=" & rect.Width.ToString & "x" & rect.Height.ToString
            End If
        End If
    End Sub

    Private Sub GroupBox1_ClientSizeChanged() Handles GroupBox1.ClientSizeChanged
        Dim clientWidth As Integer = GroupBox1.ClientSize.Width
        Dim clientHeight As Integer = GroupBox1.ClientSize.Height
        rbSetGapX.Top = 20
        rbSetGapX.Left = CInt((clientWidth - rbSetGapX.Width) / 2)
        rbSetGapY.Top = rbSetGapX.Top + rbSetGapX.Height + 3
        rbSetGapY.Left = rbSetGapX.Left
        rbSetGapXY.Top = rbSetGapY.Top + rbSetGapY.Height + 3
        rbSetGapXY.Left = rbSetGapX.Left
        btnLessGap.Top = rbSetGapXY.Top + rbSetGapXY.Height + 12
        btnLessGap.Left = CInt((clientWidth - btnLessGap.Width - btnMoreGap.Width) / 3)
        btnMoreGap.Top = btnLessGap.Top
        btnMoreGap.Left = btnLessGap.Left * 2 + btnLessGap.Width
    End Sub

    Private Sub SetGapText()
        If (adhesive IsNot Nothing) Then
            If (adhesive.Loaded) Then
                Dim iGap As Point = adhesive.GapSize
                rbSetGapX.Text = "Horizontal ( " & iGap.X.ToString & " )"
                rbSetGapY.Text = "Vertical ( " & iGap.Y.ToString & " )"
                rbSetGapXY.Text = "Both ( " & iGap.X.ToString & ", " & iGap.Y.ToString & " )"
            End If
        End If
    End Sub

    Private Sub MoreGap_Click(sender As Object, e As EventArgs) Handles btnMoreGap.Click
        If (adhesive IsNot Nothing) Then
            If (adhesive.Loaded) Then
                Dim gap As Point = adhesive.GapSize
                If (rbSetGapX.Checked) Or (rbSetGapXY.Checked) Then
                    gap.X += 1
                End If
                If (rbSetGapY.Checked) Or (rbSetGapXY.Checked) Then
                    gap.Y += 1
                End If
                adhesive.GapSize = gap
                SetGapText()
            End If
        End If
    End Sub

    Private Sub LessGap_Click() Handles btnLessGap.Click
        If (adhesive IsNot Nothing) Then
            If (adhesive.Loaded) Then
                Dim gap As Point = adhesive.GapSize
                If (rbSetGapX.Checked) Or (rbSetGapXY.Checked) Then
                    gap.X -= 1
                End If
                If (rbSetGapY.Checked) Or (rbSetGapXY.Checked) Then
                    gap.Y -= 1
                End If
                adhesive.GapSize = gap
                SetGapText()
            End If
        End If
    End Sub

End Class