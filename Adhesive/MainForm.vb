Imports AdhesiveWindows

Public Class MainForm

    '----- start Adhesive code -----
    Friend adhesive As AdhesiveWindows.Adhesive = Nothing
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

#Region "Event Handlers"

    Private Sub MainForm_Load() Handles Me.Load
        ' Seems needed in Windows 7?
        LeftGroupBox_ClientSizeChanged()
        RightGroupBox_ClientSizeChanged()
    End Sub

    Private Sub AddForm_Click() Handles btnAddForm.Click
        CreateTestTool()
    End Sub

    Private Sub Reset_Click() Handles btnReset.Click
        If (adhesive.Loaded) Then
            If (AdhesiveWindows.Adhesive.formCopyCount > 0) Then
                If (MsgBox("Delete All Test Windows?",
                           MsgBoxStyle.Exclamation + MsgBoxStyle.OkCancel, "Reset") = MsgBoxResult.Ok) Then
                    adhesive.ResetClones()
                    lbClosedForms.Items.Clear()
                End If
            End If
        End If
    End Sub

    're-open a test form that was closed before
    Private Sub ClosedForms_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles lbClosedForms.ItemCheck
        If (e.NewValue = CheckState.Checked) Then
            Dim sName = lbClosedForms.Text
            'remove it from the list (delayed)
            RemTestTool(e.Index)
            'and re-open it
            CreateTestTool(sName)
        End If
    End Sub

    Private Sub LeftGroupBox_ClientSizeChanged() Handles GroupBox1.ClientSizeChanged
        'place the buttons and listbox
        Dim clientWidth As Integer = GroupBox1.ClientSize.Width
        Dim clientHeight As Integer = GroupBox1.ClientSize.Height
        'top-down
        btnAddForm.Top = 20
        btnAddForm.Left = CInt((clientWidth - btnAddForm.Width) / 2)
        labClosedWindows.Top = btnAddForm.Top + btnAddForm.Height + 6
        'bottom-up
        btnReset.Top = clientHeight - 10 - btnReset.Height
        btnReset.Left = CInt((clientWidth - btnReset.Width) / 2)
        lbClosedForms.Top = labClosedWindows.Top + labClosedWindows.Height
        lbClosedForms.Left = 6
        lbClosedForms.Width = clientWidth - 12
        lbClosedForms.Height = btnReset.Top - lbClosedForms.Top - 6
        'Debug.WriteLine("groupbox width: {0}, height: {1}", clientWidth, clientHeight)
    End Sub

    Private Sub RightGroupBox_ClientSizeChanged() Handles GroupBox2.ClientSizeChanged
        Dim clientWidth As Integer = GroupBox2.ClientSize.Width
        Dim clientHeight As Integer = GroupBox2.ClientSize.Height
        'move along, or snap to others
        GroupBox3.Width = clientWidth - (GroupBox3.Left * 2)
        Dim clt3Width As Integer = GroupBox3.ClientSize.Width
        rbMoveAlong.Left = (clt3Width - rbMoveAlong.Width) / 2
        rbSnapToAll.Left = rbMoveAlong.Left
        rbSnapToAll.Top = rbMoveAlong.Top + rbMoveAlong.Height + 1
        GroupBox3.ClientSize = New Size(GroupBox3.ClientSize.Width, rbSnapToAll.Top + rbSnapToAll.Height + 12)
        'tool windows group
        GroupBox4.Width = GroupBox3.Width
        GroupBox4.Left = GroupBox3.Left
        GroupBox4.Top = GroupBox3.Top + GroupBox3.Height + 6
        'tool windows buttons
        btnPresets.Top = 20
        btnPresets.Left = (clt3Width - btnPresets.Width) / 2
        btnGapSize.Top = btnPresets.Top + btnPresets.Height + 6
        btnGapSize.Left = btnPresets.Left
        btnForce.Top = btnGapSize.Top + btnGapSize.Height + 6
        btnForce.Left = btnPresets.Left
        GroupBox4.ClientSize = New Size(GroupBox4.ClientSize.Width, btnForce.Top + btnForce.Height + 12)
        'debug
        GroupBox5.Width = GroupBox3.Width
        GroupBox5.Left = GroupBox3.Left
        GroupBox5.Top = GroupBox4.Top + GroupBox4.Height + 6
        'edit config
        btnConfig.Top = 20
        btnConfig.Left = btnPresets.Left
        GroupBox5.ClientSize = New Size(GroupBox5.ClientSize.Width, btnConfig.Top + btnConfig.Height + 12)
    End Sub

    Private Sub MoveAlong_CheckedChanged(sender As Object, e As EventArgs) Handles rbMoveAlong.CheckedChanged
        'called by InitializeComponent() before first event
        If (adhesive IsNot Nothing) Then
            adhesive.MainMovesAll = rbMoveAlong.Checked
        End If
    End Sub

    Private Sub Presets_Click() Handles btnPresets.Click
        PresetsTool.Show()
    End Sub

    Private Sub GapSize_Click() Handles btnGapSize.Click
        GapSizeTool.Show()
    End Sub

    Private Sub Force_Click() Handles btnForce.Click
        ForceTool.Show()
    End Sub

    Private Sub Config_Click() Handles btnConfig.Click
        Dim configFile As String = Configuration.ConfigurationManager.OpenExeConfiguration(Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath
        Process.Start("notepad", configFile)
    End Sub

#End Region

    'create a new Test tool window
    Private Sub CreateTestTool()
        AdhesiveWindows.Adhesive.formCopyCount += 1
        CreateTestTool("Test Window " & AdhesiveWindows.Adhesive.formCopyCount.ToString)
    End Sub

    'create or re-open a closed Test window
    Private Sub CreateTestTool(sName As String)
        Dim frm As New TestForm With {.Name = sName, .Text = sName}
        frm.Show()
    End Sub

    Delegate Sub RemoveFromListDelegate(lstBox As CheckedListBox, iIndex As Integer)

    Private Sub RemTestTool(iIndex As Integer)
        'remove the name from the list _after_ lbClosedForms processed its ItemCheck event
        Dim aRemToolArgs As Object() = {lbClosedForms, iIndex}
        lbClosedForms.BeginInvoke(New RemoveFromListDelegate(AddressOf RemoveFromList), aRemToolArgs)
    End Sub

    'remove an item from the re-open list 
    Public Sub RemoveFromList(lstBox As CheckedListBox, iIndex As Integer)
        Dim oItem As Object = lstBox.Items.Item(iIndex)
        If Not (oItem Is Nothing) Then
            lstBox.Items.Remove(oItem)
        End If
    End Sub

    'add a closing toolform to the re-open list
    'called from closing test forms
    Public Sub ToolFormClosed(sName As String)
        Dim index As Integer = lbClosedForms.FindStringExact(sName)
        If (index = ListBox.NoMatches) Then
            lbClosedForms.Items.Add(sName, False)
        End If
    End Sub

    'remove a closing toolform from the re-open list
    'called from loaded test forms
    Public Sub ToolFormOpened(sName As String)
        Dim index As Integer = lbClosedForms.FindStringExact(sName)
        If (index <> ListBox.NoMatches) Then
            lbClosedForms.Items.RemoveAt(index)
        End If
    End Sub

End Class
