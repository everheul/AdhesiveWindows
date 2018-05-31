Imports AdhesiveWindows

Public Class TestForm

    '----- start Adhesive code -----
    Friend adhesive As Adhesive
    Protected Overrides Sub WndProc(ByRef m As Message)
        If (adhesive Is Nothing) Then adhesive = New Adhesive(Me)
        If (adhesive.FilterMessage(m)) Then MyBase.WndProc(m)
    End Sub
    Public Overloads Sub Show()
        Adhesive.OpenToolForm(Me)
    End Sub
    '----- end Adhesive code -----

    Private Sub ToolForm_Shown() Handles Me.Shown
        'remove this window from the closed-windows-list 
        MainForm.ToolFormOpened(Me.Name)
        cbFormBorderStyle.SelectedIndex = 4
        SetInfoText()
    End Sub

    Private Sub ToolForm_Move() Handles Me.Move
        SetInfoText()
    End Sub

    Private Sub ToolForm_Resize() Handles Me.Resize
        SetInfoText()
    End Sub

    Private Sub ToolForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'add this window to the closed-windows-list 
        MainForm.ToolFormClosed(Me.Name)
    End Sub

    Private Sub SetInfoText(Optional showBorder As Boolean = False)
        If (adhesive IsNot Nothing) Then
            If (adhesive.Loaded) Then
                If (showBorder) Then 'show border size:
                    Dim rect As Rectangle = adhesive.AeroBorderSize
                    If (rect <> Rectangle.Empty) Then
                        Label1.Text = "DWM Borders: left = " & rect.Left.ToString & ", top = " & rect.Top.ToString &
                                "  width = " & rect.Width.ToString & ", height = " & rect.Height.ToString
                        Return
                    End If
                Else  'show visual bounds:
                    Dim rect As Rectangle = adhesive.GetVisualBounds()
                    Label1.Text = "VisualBounds: pos=" & rect.Left.ToString & "," & rect.Top.ToString &
                                "  size=" & rect.Width.ToString & "x" & rect.Height.ToString
                End If
            End If
        End If
    End Sub

    Private Sub FormBorderStyle_SelectedIndexChanged() Handles cbFormBorderStyle.SelectedIndexChanged
        If (adhesive.Loaded) Then
            'get the choosen border type
            Dim sBorderStyle As String = cbFormBorderStyle.SelectedItem.ToString
            'convert the resulting String to a FormBorderStyle Constant
            Dim NewStyle As FormBorderStyle = CType(FormBorderStyle.Parse(GetType(FormBorderStyle), sBorderStyle), FormBorderStyle)
            'let Adhesive set the borderstyle without changing the visual bounds
            adhesive.SetBorderStyle(NewStyle)
            SetInfoText(True)
        End If
    End Sub

End Class