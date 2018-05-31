Imports AdhesiveWindows

Public Class ForceTool

    Dim adhForce As Integer

    '----- start Adhesive code -----
    Friend adhesive As Adhesive = Nothing
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

#Region "Event Handlers"
    Private Sub ForceTool_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        adhForce = adhesive.Force
        tbForce.Value = adhForce
        gbFull.Text = "Adhesive Force: " & adhForce.ToString
    End Sub

    Private Sub Force_Scroll(sender As Object, e As EventArgs) Handles tbForce.Scroll
        adhForce = tbForce.Value
        adhesive.Force = adhForce
        gbFull.Text = "Adhesive Force: " & adhForce.ToString
    End Sub
#End Region

End Class