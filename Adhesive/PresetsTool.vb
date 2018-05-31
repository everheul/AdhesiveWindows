Imports AdhesiveWindows

Public Class PresetsTool

    '----- start Adhesive code -----
    Friend adhesive As AdhesiveWindows.Adhesive
    Protected Overrides Sub WndProc(ByRef m As Message)
        If (adhesive Is Nothing) Then
            adhesive = New AdhesiveWindows.Adhesive(Me)
        End If
        If (adhesive.FilterMessage(m)) Then
            MyBase.WndProc(m)
        End If
    End Sub
    Public Overloads Sub Show()
        AdhesiveWindows.Adhesive.OpenToolForm(Me)
    End Sub
    '----- end Adhesive code -----

    Private Sub PresetsTool_Load() Handles MyBase.Load
        lstPresets.Items.Clear()
        lstPresets.Items.AddRange(adhesive.GetPresetNames())
    End Sub

    Private Sub New_Click() Handles btnNew.Click
        Dim sPresetName As String = InputBox("Save the current position and size of all windows as a new preset with name:", "Enter Preset Name")
        If (sPresetName <> String.Empty) Then
            adhesive.NewPreset(sPresetName)
            lstPresets.Items.Add(sPresetName)
            adhesive.SaveAdhesiveData(True)
        End If
    End Sub

    Private Sub Overwrite_Click() Handles btnOverwrite.Click
        If (lstPresets.SelectedItem IsNot Nothing) Then
            Dim sPresetName As String = lstPresets.SelectedItem.ToString
            adhesive.NewPreset(sPresetName)
            adhesive.SaveAdhesiveData(True)
        End If
    End Sub

    Private Sub Delete_Click() Handles btnDelete.Click
        If (lstPresets.SelectedItem IsNot Nothing) Then
            adhesive.DeletePreset(lstPresets.SelectedItem.ToString)
            lstPresets.Items.Remove(lstPresets.SelectedItem)
        End If
    End Sub

    Private Sub Apply_Click() Handles btnApply.Click
        If (lstPresets.SelectedItem IsNot Nothing) Then
            adhesive.ApplyPreset(lstPresets.SelectedItem.ToString)
        End If
        Me.Activate()
    End Sub

    Private Sub Presets_DoubleClick() Handles lstPresets.DoubleClick
        Apply_Click()
    End Sub

End Class
