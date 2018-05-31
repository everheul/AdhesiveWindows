' Adhesive Windows
' A Window Manager for VB.NET Windows Forms Applications.
' by Ekke Verheul, 24 may 2018

Imports System.Runtime.InteropServices

' For the Test App, some extra (normally irrelevant) code was added.
' You can remove the DEMO code by hand, or set COMPILE_MODE in your own project also (Advanced Compile Options).
#If (COMPILE_MODE = "DEMO") Then
Imports AdhesiveTest
#End If

Namespace Global.AdhesiveWindows

    Public Class Adhesive

#Region "Defaults"

        ' The default distance from where Forms snap into place.
        Public Const ADHESIVE_FORCE As Integer = 12

        ' The default gap between the Forms.
        Public Const ADHESIVE_SPACE_X As Integer = 5
        Public Const ADHESIVE_SPACE_Y As Integer = 5

#End Region

#Region "Shared data, initiated by Main"

        ' These are the Settings-stored variables and formatting routines of both FormData and PresetData.
        Friend Class SettingsData
            Friend isClosed As Boolean              ' It's a closed Form, still moving with Main.
            Friend formName As String               ' The unique name of the form, and first identification.
            Friend formBounds As Rectangle          ' Position and size of the form, normal Bounds.
            Friend className As String              ' To create a Form-inherited class with formName as Name (used by Test Windows).
            Friend Function ToSettings() As String
                Return String.Format("{0},{1},{2},{3},{4},{5},{6}", formName, isClosed, formBounds.Left, formBounds.Top,
                                 formBounds.Width, formBounds.Height, className)
            End Function
            Friend Sub FromSettings(strData As String)
                Dim data As String() = strData.Split(","c)
                If (data.Length >= 7) And (data(0).Length > 0) Then
                    Try
                        formName = data(0)
                        isClosed = Convert.ToBoolean(data(1))
                        formBounds = New Rectangle(Convert.ToInt32(data(2)), Convert.ToInt32(data(3)), Convert.ToInt32(data(4)), Convert.ToInt32(data(5)))
                        className = data(6)
                    Catch e As FormatException
                        Debug.WriteLine("Error restoring SettingsData: " & strData)
                    End Try
                Else
                    Debug.WriteLine("Error splitting SettingsData: " & strData)
                End If
            End Sub
        End Class

        ' FormData holds the positions and size of a Form during execution, exit and startup:
        Friend Class FormData
            Inherits SettingsData

            Friend formId As Integer                ' The ID, to find a Form, other then by name (not saved)
            Friend openForm As Form                 ' Pointer to the loaded Form, otherwise Nothing (not saved)
            Friend visualBounds As Rectangle        ' The last known VisualBounds, updated with every move/size event (not saved)
            Friend pMove As Size                    ' Used by AddSpace (not saved)
            ' Debug - format FormData
            Public Shadows Function ToString() As String
                Return String.Format("FormData {0}: formName= {1}, isClosed= {2}, formBounds= {3}, visualBounds= {4}, className= {5}",
                                 formId, formName, isClosed.ToString, formBounds.ToString, visualBounds.ToString, className)
            End Function
            Friend Shadows Sub FromSettings(strData As String)
                MyBase.FromSettings(strData)
                formId = nextWinId
                nextWinId += 1
            End Sub
        End Class

        ' PresetData holds the preset position and size of a Form:
        Friend Class PresetData
            Inherits SettingsData

            Friend isApplied As Boolean         ' Used by ApplyPreset() (not saved)
        End Class

        ' PresetHead holds the name of the preset and a list of all the Forms in one preset:
        Friend Class PresetHead
            Friend presetName As String
            Friend presetData As List(Of PresetData)
            Friend Function ToSettings() As String
                Dim sb As New Text.StringBuilder(String.Format("Preset={0}", presetName), 100)
                For Each pd As PresetData In presetData
                    sb.Append("|"c & pd.ToSettings)
                Next
                Return sb.ToString
            End Function
            Friend Sub FromSettings(strData As String)
                Dim data As String() = strData.Split("|"c)
                If (data.Length > 0) Then
                    presetName = data(0)
                    presetData = New List(Of PresetData)
                    For i As Integer = 1 To data.Length - 1
                        Dim pd As New PresetData()
                        pd.FromSettings(data(i))
                        presetData.Add(pd)
                    Next
                End If
            End Sub
        End Class

        ' The forms of this app are listed here.
        Private Shared allForms As List(Of FormData)

        ' All the presets are listed here.
        Private Shared allPresets As List(Of PresetHead)

        ' Pointer to the the first opening window, and the owner of all others.
        Private Shared appMainForm As Form = Nothing

        ' The ID of the next form (main = 0).
        Private Shared nextWinId As Integer

        ' The desktop scale (125% etc), needed to calc the bordersize of each window
        ' Note: this class is for PROCESS_SYSTEM_DPI_AWARE applications (dpiAware = true)
        Private Shared desktopScale As Double = 1.0

        ' Make sure these move/resize commands don't get misunderstood.
        Private Shared presetChange As Boolean = False

        ' Usersetting GapSize - the preferred x and y distances between forms
        Private Shared pGapSize As Point

        Friend Property GapSize() As Point
            Get
                Return pGapSize
            End Get
            Set(value As Point)
                ' Move the toolwindows, just to get an idea.
                SetGapSize(value)
                pGapSize = value
            End Set
        End Property

        ' Usersetting moveWithMain - Move all along, or snap main to others. 
        Private Shared moveWithMain As Boolean = True
        Friend Property MainMovesAll() As Boolean
            Get
                Return moveWithMain
            End Get
            Set(value As Boolean)
                moveWithMain = value
            End Set
        End Property

        ' Usersetting Force - How many pixels to snap. 
        Private Shared adhForce As Integer = ADHESIVE_FORCE
        Friend Property Force() As Integer
            Get
                Return adhForce
            End Get
            Set(value As Integer)
                adhForce = value
            End Set
        End Property

        ' Three different ways to calculate the borders
        Friend Enum SnapTargets As Integer
            SNAP_TO_NONE = 0    ' No snap
            SNAP_TO_MAIN = 1    ' Only snap to Mains borders
            SNAP_TO_ALL = 2     ' Snap to all other borders
            SNAP_SORTED = 3     ' Snap to all - but borders that are less then SNAP_FORCE pixels near borders of closer windows
        End Enum

        Private Shared iSnapTarget As SnapTargets = SnapTargets.SNAP_SORTED

        Friend Property SnapTarget As SnapTargets
            Get
                Return iSnapTarget
            End Get
            Set(value As SnapTargets)
                iSnapTarget = value
            End Set
        End Property

        ' 4 lists with boundaries, re-filled at every move/resize start.
        ' And shared, because there's only one Form moving at a time 
        Private leftSnapBorders As New List(Of Integer)
        Private topSnapBorders As New List(Of Integer)
        Private rightSnapBorders As New List(Of Integer)
        Private bottomSnapBorders As New List(Of Integer)

#If (COMPILE_MODE = "DEMO") Then
        ' The option to create windows as copy of a Form is used here (in MainForm) to create and name Test Windows.
        ' If there is no need for that in your solution, you could search for and remove code with a DEMO comment.
        Public Shared formCopyCount As Integer = 0
#End If

#End Region

#Region "Unshared data"

        ' Copy of FormData.formId
        Private adhesiveId As Integer = -1

        ' Copy of FormData.openForm  
        Private adhForm As Form

        ' In the process of opening a Form, the Adhesive Class is created after the first Event,
        ' when the Form is not completely ready yet. A few events later, after the first 'Active' event, 
        ' the border is painted, VisualBounds is valid and this flag is set.
        Private isLoaded As Boolean = False

        Friend ReadOnly Property Loaded() As Boolean
            Get
                Return isLoaded
            End Get
        End Property

        ' Also valid after (isLoaded = True), and empty until then:
        Private borderSize As Rectangle = New Rectangle(0, 0, 0, 0)

        Friend ReadOnly Property AeroBorderSize() As Rectangle
            Get
                Return borderSize
            End Get
        End Property

        ' Used to block not-user-related events:
        Private isActive As Boolean = True

        ' Adhesive hides and freezes all Tool Forms when Main is replaced by the system;
        ' that's on maximize, minimize and Aero auto placements. 
        ' When Main restores, the tools are restored too.

        ' Used by Tool Windows, False during Aero movements.
        Private isVisible As Boolean = True

        ' This flag is only used by Main:
        Private toolsHidden As Boolean = False

        ' True when the user is moving or sizing the Form:
        Private boundsChanging As Boolean = False

        ' Remember Mains syscommands.
        Private prevSysCommand As SysCommands = 0

#End Region

#Region "Message Loop"

        ' The Adhesive class is created on the first event of a new Form.
        Public Sub New(ByRef frm As Form)
            If (appMainForm Is Nothing) Then ' this must be Main.
                appMainForm = frm
                ' Init shared data, in this order:
                allForms = New List(Of FormData)
                allPresets = New List(Of PresetHead)
                nextWinId = 0
                desktopScale = 1.0
                pGapSize = New Point(ADHESIVE_SPACE_X, ADHESIVE_SPACE_Y)
                ' Check the desktop scale at startup (TODO)
                desktopScale = GetScaleFactor(frm)
                ' Get stored data
                LoadAdhesiveData()
            Else ' Tool Form
                frm.Owner = appMainForm
            End If
            adhForm = frm
        End Sub

        ' All WndProc messages are tested, just a few are used.
        ' And only the WM_WINDOWPOSCHANGING messages are altered sometimes, when snapping.
        ' The WMUSER messages return False, all others True (= pass on to MyBase.WndProc).
        Friend Function FilterMessage(ByRef m As Message) As Boolean
            ' Debug - show all messages in time:
            'Debug.WriteLine("{0}:{1} WndProc Event for {2}: {3} (&H{4:X})", DateTime.Now.Second, DateTime.Now.Millisecond, adhForm.Name, WndProcMsg.Format(GetType(WndProcMsg), m.Msg, "G"), m.Msg)

            Select Case m.Msg
                Case WndProcMsg.WM_WINDOWPOSCHANGING
                    WMWindowposChanging(m)

                Case WndProcMsg.WM_ENTERSIZEMOVE
                    WMEnterSizeMove()

                Case WndProcMsg.WM_EXITSIZEMOVE
                    WMExitSizeMove()

                Case WndProcMsg.WM_ACTIVATE
                    WMActivate(m)

                Case WndProcMsg.WM_SYSCOMMAND
                    WMSysCommand(m)

                Case WndProcMsg.WM_CREATE
                    WMCreate()

                Case WndProcMsg.WM_CLOSE
                    WMClose()

                Case WndProcMsg.WMUSER_SETMINIMAXI
                    WMUserSetMiniMaxi(m)
                    Return False

                Case WndProcMsg.WMUSER_MOVE_TOOL
                    WMUserMoveTool(m)
                    Return False

                Case WndProcMsg.WMUSER_SHOW_TOOL
                    WMUserShowTool(m)
                    Return False

                Case WndProcMsg.WMUSER_SHOWN
                    WMUserFormShown()
                    Return False
            End Select

            Return True
        End Function

        ' Follow the pointer in LParam and copy the RECT structure.
        Private Sub WMWindowposChanging(ByRef m As Message)
            If (isLoaded = True) Then
                ' Make a copy of the RECT structure from unmanaged memory.
                Dim evWindowPos As WINDOWPOS = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(WINDOWPOS)), WINDOWPOS)
                ' If snapped, the changes are saved to evWindowPos.
                MoveSize(evWindowPos)
                ' Copy back to unmanaged memory
                Marshal.StructureToPtr(evWindowPos, m.LParam, False)
            End If
        End Sub

        ' Start of a move/size loop; create snap borders.
        Private Sub WMEnterSizeMove()
            If (isLoaded = True) Then
                If (adhesiveId = 0) Then
                    If (moveWithMain = False) Then
                        SetSnapBorders()
                    End If
                Else
                    SetSnapBorders()
                End If
                boundsChanging = True
            End If
        End Sub

        ' End of a move/size loop.
        Private Sub WMExitSizeMove()
            If (isLoaded = True) Then
                boundsChanging = False
            End If
        End Sub

        ' A window is (de-)activated.
        Private Sub WMActivate(m As Message)
            ' Keep track of the active state of each Form.
            isActive = (LoWord(m.WParam) > 0)
            ' The first WM_ACTIVATE event is used to post the WMUSER_SHOWN message.
            If (isLoaded = False) Then
                ' GetBorderSize is done twice during the load process, because Main and the 
                '  tool windows react differently and the test-tools need the VisualBounds ASAP.
                borderSize = GetBorderSize(adhForm, desktopScale)
                isLoaded = True
                PostMessage(m.HWnd, WndProcMsg.WMUSER_SHOWN, 0, 0)
            End If
        End Sub

        ' A user clicked on one of the system icons or borders.
        Private Sub WMSysCommand(m As Message)
            If (adhesiveId = 0) Then
                ' Mains syscommands are stored and handled after the WINDOWPOSCHANGING event.
                prevSysCommand = m.WParam.ToInt32 And &HFFF0
                'Debug.WriteLine("Main - WM_SYSCOMMAND: {0} (&H{1:X})", SysCommands.Format(GetType(SysCommands), prevSysCommand, "G"), prevSysCommand)
            End If
        End Sub

        Private Sub WMCreate()
            ' Posted only once, just before anything gets painted.
            ' A perfect moment to change the bounds.
            If (isLoaded = False) Then
                ' Restore my last position and size.
                ' Note: the Aero border is not ready yet, these are no VisualBounds.
                GetLastBounds()
            End If
        End Sub

        Private Sub WMClose()
            If (isLoaded = True) Then
                If (adhesiveId = 0) Then
                    ' If Main closes, write the FormData to Settings.Adhesive
                    SaveAdhesiveData() ' (True) if save_on_close not set in IDE.
                Else
                    ' A tool was closed by the user, register that.
                    RemWindow(True)
                End If
            End If
        End Sub

        ' Prevent system amnesia after Aero Snap &| mini/maximize (see SetMiniMaximize):
        Private Sub WMUserSetMiniMaxi(m As Message)
            Select Case adhForm.WindowState
                Case FormWindowState.Maximized
                    ' Never minimize a maximized Form. Please don't.
                    If (adhForm.MinimizeBox = True) Then adhForm.MinimizeBox = False
                Case FormWindowState.Normal
                    ' If the passed wParam is True (Aero) only allow minimize. Otherwise allow both.
                    If (m.WParam) Then
                        If (adhForm.MinimizeBox = False) Then adhForm.MinimizeBox = True
                        If (adhForm.MaximizeBox = True) Then adhForm.MaximizeBox = False
                    Else
                        If (adhForm.MinimizeBox = False) Then adhForm.MinimizeBox = True
                        If (adhForm.MaximizeBox = False) Then adhForm.MaximizeBox = True
                    End If
            End Select
        End Sub

        ' Move a Tool Window.
        Private Sub WMUserMoveTool(m As Message)
            Dim x As Integer = m.WParam.ToInt32
            Dim y As Integer = m.LParam.ToInt32
            Dim pMoved As New Size(x, y)
            ' Debug.Print("Moving " & adhForm.Name & ": " & pMoved.ToString)
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formId = adhesiveId)
            If (adhData IsNot Nothing) Then
                ' Move stored bounds.
                adhData.formBounds.Location += pMoved
                adhData.visualBounds.Location += pMoved
                ' Move the form.
                adhForm.Location += pMoved
            End If
        End Sub

        Private Sub WMUserShowTool(m As Message)
            isVisible = CBool(m.WParam)
            adhForm.Visible = isVisible
        End Sub

        Private Sub WMUserFormShown()
            borderSize = GetBorderSize(adhForm, desktopScale) ' Second call, all borders ready.
            ' Save the VisualBounds for the first time.
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formId = adhesiveId)
            If (adhData IsNot Nothing) Then
                'adhData.formBounds = adhForm.Bounds
                adhData.visualBounds = GetVisualBounds(adhData.formBounds)
            End If

            If (adhesiveId = 0) Then ' Is this main?
                ReOpenToolWindows()
                adhForm.Focus()
            Else ' Tool window
                ' Disable mini/maximize and taskbar:
                adhForm.MinimizeBox = False
                adhForm.MaximizeBox = False
                adhForm.ShowInTaskbar = False
            End If
        End Sub

        ' Called by WMWindowposChanging()
        Private Sub MoveSize(ByRef wndPos As WINDOWPOS)
            If (wndPos.cx = 0) Then
                ' Many WM_WINDOWPOSCHANGING commands are sent with empty sizes,
                ' no need to check SWP_NOSIZE for that.
                Return
            End If
            If (adhesiveId = 0) Then ' Main?
                'Debug.Print("MoveSize: " & wndPos.ToString)
                ' Moved/sized by mouse?
                If (CheckWindowState(wndPos)) Then
                    ' Move or Snap?
                    If (moveWithMain = True) Then
                        MoveToolsWithMain(wndPos)
                    Else
                        SnapVisualBounds(wndPos)
                    End If
                End If
            Else
                ' A tool snaps-to-borders only if it's active, visible and not moved by any preset-change.
                If (isActive = True) And (isVisible = True) And (presetChange = False) Then
                    SnapVisualBounds(wndPos)
                End If
            End If
        End Sub

        ' Called by MoveSize()
        ' Make the tool windows visible or invisible, depending on Mains state.
        ' Return True only if Mains state is Normal.
        Private Function CheckWindowState(wndPos As WINDOWPOS) As Boolean
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formId = 0)
            If Not (adhData Is Nothing) Then
                Dim rOldPos As Rectangle = adhData.visualBounds
                Dim rNewPos As Rectangle = GetVisualBounds(wndPos)

                If (boundsChanging = True) Then
                    ' User is changing the bounds, check on Aero Auto Snap results:
                    If MovedBySystem(rOldPos, rNewPos) Then ' Mouse released in Aero Auto Snap position.
                        ' Hide the tools.
                        HideToolWindows()
                        ' Disable maximize, see SetMiniMaximize for an explanation.
                        SetMiniMaximize(True)
                    ElseIf (toolsHidden = True) Then
                        ' Moved by hand, restore.
                        ShowToolWindows()
                        SetMiniMaximize(False)
                    End If
                Else ' Not in a move/size loop.
                    If (prevSysCommand = SysCommands.SC_MAXIMIZE) Then
                        HideToolWindows()
                        ' Disable mimimize.
                        SetMiniMaximize(True)
                    ElseIf (prevSysCommand = SysCommands.SC_MINIMIZE) Then
                        HideToolWindows()
                    ElseIf (prevSysCommand = SysCommands.SC_RESTORE) Then
                        ' The system also says RESTORE when it's going from MIN/MAX back to Aero Snap position.
                        ' At this point, it's best to wait for the Size to restore.
                        If (rOldPos.Size = rNewPos.Size) Then
                            ShowToolWindows()
                            SetMiniMaximize(False)
                        Else
                            HideToolWindows()
                        End If
                    ElseIf (toolsHidden = False) And (presetChange = False) And MovedBySystem(rOldPos, rNewPos) Then
                        'Debug.Print("Moved/sized to Aero Auto position: {0}", wndPos.ToString)
                        HideToolWindows()
                        SetMiniMaximize(True)
                    ElseIf (toolsHidden = True) And (rOldPos.Size = rNewPos.Size) Then
                        'Debug.Print("Restored from Aero position.")
                        ShowToolWindows()
                        SetMiniMaximize(False)
                    End If
                    prevSysCommand = 0
                End If
            End If
            Return (toolsHidden = False)
        End Function

        ' Called by CheckWindowState()
        ' Distinct a user move/size action from Aero Auto Something.
        ' Return True only if the size changes plus two opposite bounds.
        Private Function MovedBySystem(r1 As Rectangle, r2 As Rectangle) As Boolean
            If (r1.Size = r2.Size) Then
                Return False
            ElseIf (((r1.Left <> r2.Left) And (r1.Right <> r2.Right)) Or ((r1.Top <> r2.Top) And (r1.Bottom <> r2.Bottom))) Then
                Return True
            Else
                Return False
            End If
        End Function

        ' Called by CheckWindowState()
        ' The old minimize and maximize buttons do not combine very well with the new Aero Auto Snap, or 'Multitasking' options.
        ' In some sequences, in both Windows 7 and Windows 10, the 'restore' location gets lost and, with it, the connection with the Tool Forms.
        ' Also, in Windows 7, switching between minimized and maximized alone can give results that mess up this connection.
        ' To prevent all this, Adhesive disables the maximize option when sized by Aero, and the minimized option when Main is maximized.
        Private Sub SetMiniMaximize(nomax As Boolean)
            PostMessage(adhForm.Handle, WndProcMsg.WMUSER_SETMINIMAXI, nomax, 0)
        End Sub

        ' Called by WMCreate()
        ' Get the Id and last bounds of the (loading) form, or, 
        ' if its name is unknown (first usage), add it as a new form.
        Private Sub GetLastBounds()
            Dim name As String = adhForm.Name
            Dim clsName As String = adhForm.GetType().Name
            clsName = If(clsName = name, "", clsName)
            ' Search for the name in allForms:
            Dim adhData As FormData = allForms.Find(Function(af) af.formName = name)
            If (adhData Is Nothing) Then
                ' First usage, add to allForms:
                adhesiveId = nextWinId
                nextWinId += 1
                allForms.Add(New FormData With {.formId = adhesiveId,
                                            .openForm = adhForm,
                                            .isClosed = False,
                                            .formName = name,
                                            .className = clsName,
                                            .formBounds = adhForm.Bounds})
            Else
                ' Settings loaded.
                ' Was this Tool opened before? without calling OpenToolForm()??
                If (adhData.openForm IsNot Nothing) Then
                    adhData.openForm.BringToFront()
                    ' Too bad. There is NO WAY to silently stop a new window from opening. Closing a Form in the 'CreateHandle' state is forbidden. 
                    ' It may (or may not) crash your application, and it still won't close...
                    ' The workaround was to overload the Show() routine of the forms, and if you don't:
                    Throw New System.Exception(String.Format("Error: Unhandled Form. Please use 'Adhesive.OpenToolForm({0})' in stead of '{0}.Show()'", adhForm.Name))
                Else
                    ' Normal operation; register and position the Form:
                    adhData.openForm = adhForm
                    adhData.isClosed = False
                    adhData.className = clsName
                    adhForm.Bounds = adhData.formBounds
                    adhesiveId = adhData.formId
                End If
            End If
        End Sub

#End Region

#Region "VisualBounds"

        ' Convert WINDOWPOS bounds to a VisualBounds Rectangle.
        Friend Function GetVisualBounds(ByRef wndPos As WINDOWPOS) As Rectangle
            Return New Rectangle(wndPos.x + borderSize.Left, wndPos.y + borderSize.Top,
                                wndPos.cx + borderSize.Width, wndPos.cy + borderSize.Height)
        End Function

        ' Convert Form bounds to a VisualBounds Rectangle.
        Friend Function GetVisualBounds(Optional frm As Form = Nothing) As Rectangle
            If (frm Is Nothing) Then frm = adhForm
            Return GetVisualBounds(frm.Bounds)
        End Function

        ' Convert a Form bounds Rectangle to a VisualBounds Rectangle.
        Friend Function GetVisualBounds(rect As Rectangle) As Rectangle
            Return New Rectangle(rect.Left + borderSize.Left, rect.Top + borderSize.Top,
                                rect.Width + borderSize.Width, rect.Height + borderSize.Height)
        End Function

        ' Convert and apply a VisualBounds Rectangle to a WINDOWPOS reference.
        Friend Sub SetVisualBounds(visual As Rectangle, ByRef wndPos As WINDOWPOS)
            wndPos.x = visual.X - borderSize.Left
            wndPos.y = visual.Y - borderSize.Top
            wndPos.cx = visual.Width - borderSize.Width
            wndPos.cy = visual.Height - borderSize.Height
        End Sub

        ' Convert and apply a VisualBounds Rectangle to this Form.
        Friend Sub SetVisualBounds(visual As Rectangle)
            adhForm.Bounds = ToFormBounds(visual)
        End Sub

        ' Convert a VisualBounds Rectangle to a Form bounds Rectangle.
        Friend Function ToFormBounds(visual As Rectangle) As Rectangle
            Return New Rectangle(visual.Left - borderSize.Left, visual.Top - borderSize.Top,
                                visual.Width - borderSize.Width, visual.Height - borderSize.Height)
        End Function

#End Region

#Region "Create Tool Windows"

        Private Sub ReOpenToolWindows()
            ' Re-open the tool windows that were open last session.
            ' Note: the formId's were assigned in LoadAdhesiveData()
            For Each apForm As FormData In allForms.FindAll(Function(ap) (ap.formId > 0))
                If (apForm.isClosed = False) Then
                    OpenToolForm(apForm)
#If (COMPILE_MODE = "DEMO") Then
                Else
                    ' Add it to Mains lbClosedForms CheckListBox:
                    If (apForm.className <> String.Empty) Then
                        MainForm.ToolFormClosed(apForm.formName)
                        formCopyCount += 1
                    End If
#End If
                End If
            Next
        End Sub

        ' (re-)open or activate a tool window by Form.
        ' Call this function to open your tool windows.
        Friend Shared Sub OpenToolForm(frm As Form)
            If (FormNotLoaded(frm.Name)) Then
                frm.Show()
            End If
        End Sub

        ' (re-)open or activate a tool window by FormData.
        Private Shared Sub OpenToolForm(apForm As FormData)
            OpenToolForm(apForm.formName, apForm.className)
        End Sub

        ' (re-)open or activate a tool window by name (and classname).
        Friend Shared Sub OpenToolForm(sName As String, Optional sClass As String = "")
            If (FormNotLoaded(sName)) Then
                ' Open a form by name or classname
                OpenForm(sName, sClass)
            End If
        End Sub

        ' Check whether the window is open or not, by name.
        Private Shared Function FormNotLoaded(sName As String)
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formName = sName)
            If (adhData IsNot Nothing) Then
                If (adhData.openForm IsNot Nothing) Then
                    adhData.openForm.BringToFront()
                    Return False
                End If
            End If
            Return True
        End Function

        ' Open a projects assembly window (Designer Form) by name.
        ' Uses Reflection.
        Private Shared Sub OpenForm(sName As String, sClass As String)
            ' Get to this assembly:
            Dim entryAssembly As Reflection.Assembly = Reflection.Assembly.GetEntryAssembly()
            If (entryAssembly = Nothing) Then
                Throw New System.Exception("GetEntryAssembly failed. Any unmanaged code around?")
            End If
            ' Get a list with all types in this assembly.
            Dim tTypes As Type() = Reflection.Assembly.GetEntryAssembly().GetTypes()
            ' Try to find the Form by class name:
            Dim f As Form = FindFormByName(sClass, tTypes)
            If (f IsNot Nothing) Then
                ' Open a Test window.
                f.Name = sName
                f.Text = sName
#If (COMPILE_MODE = "DEMO") Then
                formCopyCount += 1
#End If
                f.Show()
            Else
                ' Try to find the Form by name:
                f = FindFormByName(sName, tTypes)
                If (f IsNot Nothing) Then
                    f.Show()
                Else
                    ' Panic
                    Throw New System.Exception("Form not found: " & sName & ", class: " & sClass)
                End If
            End If
        End Sub

        Private Shared Function FindFormByName(sClass As String, tTypes As Type()) As Form
            If Not (sClass = String.Empty) Then
                Dim fType As Type = Array.Find(tTypes, Function(t As Type) If(t.Name = sClass, True, False))
                If (fType IsNot Nothing) Then
                    If (fType.BaseType.Name Like "*Form") Then
                        Return DirectCast(Activator.CreateInstance(fType), Form)
                    Else
                        Throw New System.Exception("Not a Form Class? Name: " & sClass)
                    End If
                End If
            End If
            Return Nothing
        End Function

#End Region

#Region "Move around Tool Windows"

        Private Sub ShowToolsAsync(bShow As Boolean)
            For Each apForm As FormData In allForms.FindAll(Function(ap) (ap.formId > 0) And (ap.openForm IsNot Nothing))
                PostMessage(apForm.openForm.Handle, WndProcMsg.WMUSER_SHOW_TOOL, bShow, 0)
            Next
        End Sub

        Private Sub HideToolWindows()
            ' Send a request to hide.
            ' The tools will unset the isVisible flag when it's done.
            If Not (toolsHidden) Then
                ShowToolsAsync(False)
                toolsHidden = True
            End If
        End Sub

        Private Sub ShowToolWindows()
            If (toolsHidden) Then
                ShowToolsAsync(True)
                toolsHidden = False
            End If
        End Sub

        ' Main has been moved and/or resized; let the tool windows move with it or get out of the way.
        Private Sub MoveToolsWithMain(ByRef wndPos As WINDOWPOS)
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formId = 0)
            If Not (adhData Is Nothing) Then
                Dim rOldPos As Rectangle = adhData.visualBounds
                Dim rNewPos As Rectangle = GetVisualBounds(wndPos)
                If (rOldPos = rNewPos) Then ' nothing changed
                    Return
                ElseIf (rNewPos.Size = rOldPos.Size) Then
                    ' If the size didn't change, it's a drag; just move the others with.
                    MoveToolWindows(New Size(rNewPos.X - rOldPos.X, rNewPos.Y - rOldPos.Y))
                Else
                    If (rNewPos.Left <> rOldPos.Left) Then
                        ' If left changed, move all forms that started LEFT of my RIGHT border.
                        MoveToolWindowsX(rNewPos.Left - rOldPos.Left, rOldPos.Right, False)
                    End If
                    If (rNewPos.Right <> rOldPos.Right) Then
                        ' If right changed, move wat started RIGHT of my LEFT border.
                        MoveToolWindowsX(rNewPos.Right - rOldPos.Right, rOldPos.Left, True)
                    End If
                    If (rNewPos.Top <> rOldPos.Top) Then ' Top changed..?
                        MoveToolWindowsY(rNewPos.Top - rOldPos.Top, rOldPos.Bottom, False)
                    End If
                    If (rNewPos.Bottom <> rOldPos.Bottom) Then ' Bottom?
                        MoveToolWindowsY(rNewPos.Bottom - rOldPos.Bottom, rOldPos.Top, True)
                    End If
                End If
                adhData.visualBounds = rNewPos
                adhData.formBounds = ToFormBounds(rNewPos) ' not yet, but soon..
            End If
        End Sub

        Private Sub MoveToolWindows(pMoved As Size)
            'Debug.Print("Moving: " & pMoved.ToString)
            For Each apForm As FormData In allForms.FindAll(Function(ap) ap.formId > 0)
                If (apForm.openForm Is Nothing) Then
                    ' Move stored bounds.
                    apForm.formBounds.Location += pMoved
                    apForm.visualBounds.Location += pMoved
                Else
                    PostMessage(apForm.openForm.Handle, WndProcMsg.WMUSER_MOVE_TOOL, pMoved.Width, pMoved.Height)
                End If
            Next
        End Sub

        ' When the Main form is resizing, the glued toolwindows glide with - but do not resize.
        ' Toolforms that started LEFT of mains RIGHT border (iFrom) must go LEFT (bDir=False),
        '  and vice verca (bDir=True).
        Private Sub MoveToolWindowsX(iMoveX As Integer, iFrom As Integer, bDir As Boolean)
            For Each apForm As FormData In allForms.FindAll(Function(ap) ap.formId > 0)
                If ((bDir = False) And (apForm.visualBounds.X < iFrom)) Or
                        ((bDir = True) And (apForm.visualBounds.X > iFrom)) Then
                    If (apForm.openForm Is Nothing) Then
                        ' Move stored bounds.
                        apForm.formBounds.X += iMoveX
                        apForm.visualBounds.X += iMoveX
                    Else
                        PostMessage(apForm.openForm.Handle, WndProcMsg.WMUSER_MOVE_TOOL, iMoveX, 0)
                    End If
                End If
            Next
        End Sub

        ' Toolforms that started ABOVE mains BOTTOM border (iFrom) must go UP (bDir=False),
        '  and vice verca (bDir=True)
        Private Sub MoveToolWindowsY(iMoveY As Integer, iFrom As Integer, bDir As Boolean)
            For Each apForm As FormData In allForms.FindAll(Function(ap) ap.formId > 0)
                If ((bDir = False) And (apForm.visualBounds.Y < iFrom)) Or
                        ((bDir = True) And (apForm.visualBounds.Y > iFrom)) Then
                    If (apForm.openForm Is Nothing) Then
                        apForm.formBounds.Y += iMoveY
                        apForm.visualBounds.Y += iMoveY
                    Else
                        PostMessage(apForm.openForm.Handle, WndProcMsg.WMUSER_MOVE_TOOL, 0, iMoveY)
                    End If
                End If
            Next
        End Sub

        ' A tool window was closed - don't stick to this Form anymore, but let formData move with main.
        Friend Sub RemWindow(closedByUser As Boolean)
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formId = adhesiveId)
            If Not (adhData Is Nothing) Then
                If (closedByUser) Then
                    adhData.isClosed = True ' Do not re-open at startup.
                End If
                adhData.openForm = Nothing
            End If
        End Sub

#End Region

#Region "Snap functions"
        ' Adhesive Forms can stick to the X and Y positions and to the right+gapX and bottom+gapY of the other forms.
        ' Before the Move and Resize events begin, the sticking borders have been calculated by SetBorders().

        ' This is called from a WM_WINDOWPOSCHANGING event - its bounds may be altered to realise a snap.
        Private Sub SnapVisualBounds(ByRef wndPos As WINDOWPOS)
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formId = adhesiveId)
            If Not (adhData Is Nothing) Then
                Dim newBounds As Rectangle = GetVisualBounds(wndPos)
                Dim oldBounds As Rectangle = adhData.visualBounds
                Dim snapBounds As Rectangle = newBounds
                If (newBounds = oldBounds) Then
                    ' Old news, ignore.
                    Return
                ElseIf (oldBounds.Size = newBounds.Size) Then
                    ' It's a move.
                    SnapHorizontalMove(oldBounds, newBounds, snapBounds)
                    SnapVerticalMove(oldBounds, newBounds, snapBounds)
                Else
                    ' It's a resize.
                    SnapLeftBoundSize(oldBounds, newBounds, snapBounds)
                    SnapRightBoundSize(oldBounds, newBounds, snapBounds)
                    SnapTopBoundSize(oldBounds, newBounds, snapBounds)
                    SnapBottomBoundSize(oldBounds, newBounds, snapBounds)
                End If
                ' If snapped, change the bounds of the message.
                If Not (snapBounds = newBounds) Then
                    SetVisualBounds(snapBounds, wndPos)
                End If
                ' Finally, save the current VisualBounds to become the oldBounds of next event.
                adhData.visualBounds = snapBounds
                adhData.formBounds = ToFormBounds(snapBounds)
            End If
        End Sub

        ' Move - A snap to a Left or Right border alters Left.
        Private Sub SnapHorizontalMove(oldBounds As Rectangle, newBounds As Rectangle, ByRef snapBounds As Rectangle)
            If Not (oldBounds.X = newBounds.X) Then
                Dim iBestLeft As Integer = newBounds.X
                Dim iSnap As Integer = adhForce
                SnapBoundMove(newBounds.X, iBestLeft, 0, iSnap, leftSnapBorders)
                SnapBoundMove(newBounds.Right, iBestLeft, oldBounds.Width, iSnap, rightSnapBorders)
                snapBounds.X = iBestLeft
            End If
        End Sub

        ' Move - A snap to a Top or Bottom border alters Top.
        Private Sub SnapVerticalMove(oldBounds As Rectangle, newBounds As Rectangle, ByRef snapBounds As Rectangle)
            If Not (oldBounds.Y = newBounds.Y) Then
                Dim iBestTop As Integer = newBounds.Y
                Dim iSnap As Integer = adhForce
                SnapBoundMove(newBounds.Y, iBestTop, 0, iSnap, topSnapBorders)
                SnapBoundMove(newBounds.Bottom, iBestTop, oldBounds.Height, iSnap, bottomSnapBorders)
                snapBounds.Y = iBestTop
            End If
        End Sub

        ' Move - Calculate the snappyest position.
        Private Sub SnapBoundMove(iNew As Integer, ByRef iBest As Integer, iBody As Integer, ByRef iSnap As Integer, borderList As List(Of Integer))
            For Each iBorder As Integer In borderList
                Dim iDiff As Integer = Math.Abs(iNew - iBorder)
                If (iDiff < iSnap) Then
                    iSnap = iDiff
                    iBest = iBorder - iBody
                End If
            Next
        End Sub

        ' Size - A snap to a Left border alters Left and Width.
        Private Sub SnapLeftBoundSize(oldBounds As Rectangle, newBounds As Rectangle, ByRef snapBounds As Rectangle)
            If Not (oldBounds.Left = newBounds.Left) Then
                Dim iBestVal As Integer = SnapBoundSize(newBounds.X, leftSnapBorders)
                snapBounds.X = iBestVal
                snapBounds.Width = oldBounds.Width + oldBounds.Left - iBestVal
            End If
        End Sub

        ' Size - A snap to a Right border alters Width.
        Private Sub SnapRightBoundSize(oldBounds As Rectangle, newBounds As Rectangle, ByRef snapBounds As Rectangle)
            If Not (oldBounds.Right = newBounds.Right) Then
                Dim iBestVal As Integer = SnapBoundSize(newBounds.Right, rightSnapBorders)
                snapBounds.Width = oldBounds.Width + iBestVal - oldBounds.Right
            End If
        End Sub

        ' Size - A snap to a Top border alters Top and Height.
        Private Sub SnapTopBoundSize(oldBounds As Rectangle, newBounds As Rectangle, ByRef snapBounds As Rectangle)
            If Not (oldBounds.Top = newBounds.Top) Then
                Dim iBestVal As Integer = SnapBoundSize(newBounds.Y, topSnapBorders)
                snapBounds.Y = iBestVal
                snapBounds.Height = oldBounds.Height + oldBounds.Y - iBestVal
            End If
        End Sub

        ' Size - A snap to a Bottom border alters Height.
        Private Sub SnapBottomBoundSize(oldBounds As Rectangle, newBounds As Rectangle, ByRef snapBounds As Rectangle)
            If Not (oldBounds.Bottom = newBounds.Bottom) Then
                Dim iBestVal As Integer = SnapBoundSize(newBounds.Bottom, bottomSnapBorders)
                snapBounds.Height = oldBounds.Height + iBestVal - oldBounds.Bottom
            End If
        End Sub

        ' Size - Calculate the snappyest position.
        Private Function SnapBoundSize(iNew As Integer, borderList As List(Of Integer)) As Integer
            Dim iSnap As Integer = adhForce
            Dim iBest As Integer = iNew
            For Each iBorder As Integer In borderList
                Dim iDiff As Integer = Math.Abs(iNew - iBorder)
                If (iDiff < iSnap) Then
                    iSnap = iDiff
                    iBest = iBorder
                End If
            Next
            Return iBest
        End Function

#End Region

#Region "Create Snap Borders"
        ' A Tool Window will, before a Resize or Move starts, first calculate all the borders it may encounter.
        ' On Move/Size events these lists can speed up calculations.

        ' Calculate the 4 shared distinct lists of boundaries for this ToolForm 
        '  before the move/size events start (on resizeStart).
        Friend Sub SetSnapBorders()
            ClearSnapBorders()
            If iSnapTarget = SnapTargets.SNAP_TO_MAIN Then
                AddFormBorders(0)
            ElseIf iSnapTarget = SnapTargets.SNAP_TO_ALL Then
                For Each apForm As FormData In allForms.FindAll(Function(ap) (ap.formId <> adhesiveId) And (ap.isClosed = False))
                    AddBorders(apForm.visualBounds, False)
                Next
            ElseIf iSnapTarget = SnapTargets.SNAP_SORTED Then
                Dim lst As List(Of KeyValuePair(Of Integer, Integer)) = GetDistances()
                ' Sort on distance.
                lst.Sort(Function(kv1, kv2)
                             Return kv1.Value.CompareTo(kv2.Value)
                         End Function)
                ' Add bounds, closest form first.
                For Each df As KeyValuePair(Of Integer, Integer) In lst
                    AddFormBorders(CInt(df.Key), True)
                Next
            End If
        End Sub

        ' Empty the border lists.
        Private Sub ClearSnapBorders()
            leftSnapBorders.Clear()
            topSnapBorders.Clear()
            rightSnapBorders.Clear()
            bottomSnapBorders.Clear()
        End Sub

        ' Add this forms borders to the border lists.
        Private Sub AddFormBorders(iID As Integer, Optional bSorted As Boolean = False)
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formId = iID)
            If Not (adhData Is Nothing) Then
                AddBorders(adhData.visualBounds, bSorted)
            End If
        End Sub

        ' Every other open form may add max 2 borders to each list.
        Private Sub AddBorders(rect As Rectangle, bSorted As Boolean)
            DistinctAdd(leftSnapBorders, rect.Left, bSorted)
            DistinctAdd(leftSnapBorders, rect.Right + pGapSize.X, bSorted)
            DistinctAdd(topSnapBorders, rect.Top, bSorted)
            DistinctAdd(topSnapBorders, rect.Bottom + pGapSize.Y, bSorted)
            DistinctAdd(rightSnapBorders, rect.Right, bSorted)
            DistinctAdd(rightSnapBorders, rect.Left - pGapSize.X, bSorted)
            DistinctAdd(bottomSnapBorders, rect.Bottom, bSorted)
            DistinctAdd(bottomSnapBorders, rect.Top - pGapSize.Y, bSorted)
        End Sub

        ' Add an integer value to a list if it's not in there yet.
        Private Sub DistinctAdd(ByRef lBounds As List(Of Integer), iNewBound As Integer, bSorted As Boolean)
            If bSorted Then
                ' Add closest forms first, and keep at least adhForce pixels between borders.
                For Each iBound As Integer In lBounds
                    If (Math.Abs(iBound - iNewBound) < adhForce) Then
                        ' A closer window listed a border too close to mine, that's that.
                        Return
                    End If
                Next
                ' Nothing comes close, so add this bound.
                lBounds.Add(iNewBound)
            Else ' Add all forms. Far away windows may add (for the user) unexpected borders.
                If (lBounds.Contains(iNewBound) = False) Then
                    lBounds.Add(iNewBound)
                End If
            End If
        End Sub

        ' Calculate an estimate of the smallest distance between this window and each of the other open windows.
        ' Returns a List Of KeyValuePairs (adhesiveId, Distance)
        Friend Function GetDistances() As List(Of KeyValuePair(Of Integer, Integer))
            Dim lst As New List(Of KeyValuePair(Of Integer, Integer))
            Dim adhData As FormData = allForms.Find(Function(ap) ap.formId = adhesiveId)
            If Not (adhData Is Nothing) Then
                Dim rThis As Rectangle = adhData.visualBounds
                Dim x1 As Double = (rThis.Left + rThis.Right) / 2
                Dim y1 As Double = (rThis.Top + rThis.Bottom) / 2
                Dim r1 As Double = If(rThis.Width < rThis.Height, rThis.Width / 2, rThis.Height / 2)
                For Each apForm As FormData In allForms.FindAll(Function(ap) (ap.formId <> adhesiveId) And (ap.isClosed = False))
                    ' Calc the distance from center to center
                    Dim rThat As Rectangle = apForm.visualBounds
                    Dim xDiff As Double = (rThat.Left + rThat.Right) / 2 - x1
                    Dim yDiff As Double = (rThat.Top + rThat.Bottom) / 2 - y1
                    Dim iDst As Integer = CInt(Math.Round(Math.Sqrt(xDiff * xDiff + yDiff * yDiff)))
                    ' Subtract (a part of) the radius
                    iDst -= r1
                    iDst -= If(rThat.Width < rThat.Height, rThat.Width / 2, rThat.Height / 2)
                    lst.Add(New KeyValuePair(Of Integer, Integer)(apForm.formId, iDst))
                Next
            End If
            Return lst
        End Function

#End Region

#Region "Gap Changing"

        ' Letting the user change the spacing for both X and Y might make up
        ' for any 'Form Bound' surprises MS can come up with in the future.
        ' But if you prefer a static gap, most of this region is otiose.
        ' This code is mainly to 'visualize' the changing gap anyway.
        ' The result may not even be perfect, since the sizes do not change, 
        ' but of course that can easily be corrected by the user.

        Friend Sub SetGapSize(pNew As Point)
            ' Any change?
            If (pGapSize = pNew) Then
                Return
            End If
            Dim sAdd As New Size(pNew - pGapSize)
            Dim adhData As FormData
            ' Clear all move points (.pMove)
            For Each adhData In allForms
                adhData.pMove = New Size(0, 0)
            Next
            ' Calculate the nescessary moves of the forms around Main, 
            ' on Mains (and their, recursive) borders and store them in .pMove
            adhData = allForms.Find(Function(ap) ap.formId = 0)
            If (adhData IsNot Nothing) Then
                Dim rect As Rectangle = adhData.visualBounds
                If (sAdd.Height <> 0) Then
                    MoveFromTopBound(rect.Top - pGapSize.Y, sAdd.Height, sAdd.Height)
                    MoveFromBottomBound(rect.Bottom + pGapSize.Y, sAdd.Height, sAdd.Height)
                End If
                If (sAdd.Width <> 0) Then
                    MoveFromRightBound(rect.Right + pGapSize.X, sAdd.Width, sAdd.Width)
                    MoveFromLeftBound(rect.Left - pGapSize.X, sAdd.Width, sAdd.Width)
                End If
            End If
            ' Save the new positions and move all open SnapToolForms:
            For Each adhData In allForms.FindAll(Function(ap) (ap.formId > 0))
                If (adhData.isClosed = False) Then
                    adhData.openForm.Location += adhData.pMove
                End If
                ' Move closed windows too.
                adhData.formBounds.Location += adhData.pMove
                adhData.visualBounds.Location += adhData.pMove
            Next
        End Sub

        ' MoveFrom*Bound: Calculate the new locations of the toolforms. 
        ' These functions are recursive and change the gap between all forms in this direction
        Friend Sub MoveFromTopBound(iBorder As Integer, iMove As Integer, iAdd As Integer)
            For Each adhData As FormData In allForms.FindAll(Function(ap) (ap.formId > 0) And (ap.visualBounds.Bottom = iBorder))
                MoveFromTopBound(adhData.visualBounds.Top - pGapSize.Y, iMove + iAdd, iAdd)
                adhData.pMove.Height = -iMove
            Next
        End Sub

        Friend Sub MoveFromBottomBound(iBorder As Integer, iMove As Integer, iAdd As Integer)
            For Each adhData As FormData In allForms.FindAll(Function(ap) (ap.formId > 0) And (ap.visualBounds.Top = iBorder))
                MoveFromBottomBound(adhData.visualBounds.Bottom + pGapSize.Y, iMove + iAdd, iAdd)
                adhData.pMove.Height = iMove
            Next
        End Sub

        Friend Sub MoveFromRightBound(iBorder As Integer, iMove As Integer, iAdd As Integer)
            For Each adhData As FormData In allForms.FindAll(Function(ap) (ap.formId > 0) And (ap.visualBounds.Left = iBorder))
                MoveFromRightBound(adhData.visualBounds.Right + pGapSize.X, iMove + iAdd, iAdd)
                adhData.pMove.Width = iMove
            Next
        End Sub

        Friend Sub MoveFromLeftBound(iBorder As Integer, iMove As Integer, iAdd As Integer)
            For Each adhData As FormData In allForms.FindAll(Function(ap) (ap.formId > 0) And (ap.visualBounds.Right = iBorder))
                MoveFromLeftBound(adhData.visualBounds.Left - pGapSize.X, iMove + iAdd, iAdd)
                adhData.pMove.Width = -iMove
            Next
        End Sub

#End Region

#Region "Settings & Presets"

        ' All stored settings (form names, booleans and integers) are placed in one user-settings string, called "Adhesive".
        ' The used conversion method is based on a few special characters that may not be used in form names: #=,|

        Friend Sub LoadAdhesiveData()
            ' Debug - If you need to get to the file where your settings are stored, run this:
            'Debug.WriteLine(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath, "User Settings")

            Dim mySettings As String = My.Settings.Adhesive.Trim()
            Dim data As String() = mySettings.Split("#"c)
            For Each varData As String In data
                varData.Trim()
                Dim p As Integer = varData.IndexOf("="c)
                If (p > 0) Then
                    Dim sData As String = varData.Substring(p + 1)

                    Select Case varData.Substring(0, p)

                        Case "gapSize"        ' gapSize=x,y#
                            Try
                                Dim spaceStr As String() = sData.Split(","c)
                                pGapSize = New Point(Convert.ToInt32(spaceStr(0)), Convert.ToInt32(spaceStr(1)))
                            Catch ex As Exception
                                Debug.WriteLine("Error restoring SnapData: " & varData)
                            End Try

                        Case "desktopScale"     ' desktopScale=1.0# 
                            Try
                                ' If the stored locations and sizes were made on a different scale, 
                                '  they are deleted. They could be kept though, for later use (TODO?)
                                Dim lastDesktopScale As Double = Convert.ToDouble(sData)
                                If (lastDesktopScale <> desktopScale) Then
                                    ' A scale change since last run! Now what?
                                    ' One could give the windows a scaled size/location, but that 
                                    '  could interfere with the "AutoScaleMode" settings.
                                    ' So here, all the stored locations & sizes are thrown away.
                                    ' Luckyly, not many people regularly change the scale...
                                    Exit For
                                End If
                            Catch ex As Exception
                                Debug.WriteLine("Error restoring SnapData: " & varData)
                            End Try

                        Case "adhesiveForce"     ' adhesiveForce=12#
                            Try
                                adhForce = Convert.ToDecimal(sData)
                            Catch ex As Exception
                                Debug.WriteLine("Error restoring SnapData: " & varData)
                            End Try

                        Case "allForms"
                            ' allForms=name,closed,x,y,w,h,classname|name,...|...,classname#
                            Dim setFrms As String() = sData.Split("|"c)
                            For Each setForm As String In setFrms
                                Dim fd As New FormData
                                fd.FromSettings(setForm)
                                allForms.Add(fd)
                            Next

                        Case "Preset"
                            ' Like allForms, but starting with presetname|
                            Dim ph As New PresetHead
                            ph.FromSettings(sData)
                            allPresets.Add(ph)

                    End Select
                End If
            Next
        End Sub

        Friend Sub SaveAdhesiveData(Optional pleaseSave As Boolean = False)
            Dim serialData As New Text.StringBuilder()
            serialData.AppendFormat("gapSize={0},{1}#", pGapSize.X, pGapSize.Y)
            serialData.AppendFormat("desktopScale={0}#", desktopScale.ToString)
            serialData.AppendFormat("adhesiveForce={0}#", adhForce.ToString)
            ' Save current positions.
            serialData.Append("allForms=")
            For Each apForm As FormData In allForms
                serialData.Append(apForm.ToSettings)
                If (apForm IsNot allForms.Last) Then
                    serialData.Append("|"c)
                End If
            Next
            ' Save presets.
            For Each ph As PresetHead In allPresets
                serialData.Append("#" & ph.ToSettings)
            Next
            ' Apply.
            My.Settings.Adhesive = serialData.ToString
            If (pleaseSave) Then My.Settings.Save()
        End Sub

        ' Returns a array with the names of all currently saved presets.
        Friend Function GetPresetNames() As String()
            Dim names As New List(Of String)
            For Each ph As PresetHead In allPresets
                names.Add(ph.presetName)
            Next
            Return names.ToArray
        End Function

        Private Function AddPreset(name As String) As PresetHead
            Dim ph As PresetHead = allPresets.Find(Function(pd) pd.presetName = name)
            If (ph IsNot Nothing) Then
                ph.presetData.Clear()
            Else
                ph = New PresetHead With {.presetName = name, .presetData = New List(Of PresetData)}
                allPresets.Add(ph)
            End If
            Return ph
        End Function

        ' Adds, or overwrites, a new layout preset.
        Friend Sub NewPreset(name As String)
            ' Replace any invalid characters in the name:
            Dim _invalidChars As Char() = {"#"c, "="c, "|"c, ","c}
            For Each c As Char In _invalidChars
                name = name.Replace(c, "_")
            Next
            ' Set the name and add all windows: 
            Dim pl As PresetHead = AddPreset(name)
            For Each fd As FormData In allForms
                pl.presetData.Add(New PresetData With {
                              .formName = fd.formName,
                              .formBounds = fd.formBounds,
                              .className = fd.className,
                              .isClosed = fd.isClosed})
            Next
        End Sub

        Friend Sub DeletePreset(name As String)
            Dim pl As PresetHead = allPresets.Find(Function(ph) ph.presetName = name)
            If (pl IsNot Nothing) Then
                pl.presetData.Clear()
                allPresets.Remove(pl)
            End If
        End Sub

        ' Change the position, size and closure of all application forms to make them match a preset.
        Friend Sub ApplyPreset(name As String)

            ' Make sure mains next rebound is not seen as a system action.
            presetChange = True

            Dim pHead As PresetHead = allPresets.Find(Function(ph) ph.presetName = name)
            If (pHead IsNot Nothing) Then
                ' Turn off the applied flag.
                pHead.presetData.ForEach(Function(pdi) pdi.isApplied = False)

                ' Try to find each of the Forms in this preset:
                For Each adhData As FormData In allForms
                    Dim pd As PresetData = pHead.presetData.Find(Function(pdi) pdi.formName = adhData.formName)
                    If (pd IsNot Nothing) Then  ' We have a match.
                        pd.isApplied = True
                        If (pd.isClosed = True) Then
                            If (adhData.openForm IsNot Nothing) Then
                                adhData.openForm.Close()
                            End If
                            adhData.openForm = Nothing
                        End If
                        adhData.isClosed = pd.isClosed
                        adhData.formBounds = pd.formBounds
                        adhData.visualBounds = GetVisualBounds(pd.formBounds)
                        If (pd.isClosed = False) Then
                            If (adhData.openForm Is Nothing) Then
                                OpenToolForm(pd.formName, pd.className)
                            Else
                                ' Move the form..
                                adhData.openForm.Bounds = pd.formBounds
                            End If
                        End If
                    Else
                        'Form not in preset, close form.
                        If (adhData.openForm IsNot Nothing) Then
                            adhData.openForm.Close()
                        End If
                        adhData.isClosed = True
                        adhData.openForm = Nothing
                    End If
                Next

                ' Then, for those who were still unknown (first usage?), create new PresetData:
                For Each pd As PresetData In pHead.presetData.FindAll(Function(pdi) pdi.isApplied = False)
                    allForms.Add(New FormData With {.formName = pd.formName,
                                                .className = pd.className,
                                                .formBounds = pd.formBounds,
                                                .isClosed = pd.isClosed})
                    If (pd.isClosed = False) Then
                        OpenForm(pd.formName, pd.className)
                    End If
                Next
            End If

            Application.DoEvents()
            presetChange = False

        End Sub

#End Region

#If (COMPILE_MODE = "DEMO") Then
        ' Not needed for projects without Form clones.
        Public Sub ResetClones()
            For Each apForm As FormData In allForms.FindAll(Function(ap) (ap.className <> String.Empty))
                If (apForm.openForm IsNot Nothing) Then
                    apForm.openForm.Close()
                End If
            Next
            allForms.RemoveAll(Function(ap) (ap.className <> String.Empty))
            formCopyCount = 0
        End Sub
        ' Only used to change a Forms FormBorderStyle without changing the visualbounds.
        Friend Sub SetBorderStyle(NewStyle As FormBorderStyle)
            If Not (adhForm.FormBorderStyle = NewStyle) Then
                isActive = False
                Dim rOldBounds As Rectangle = GetVisualBounds()
                adhForm.Opacity = 0
                adhForm.FormBorderStyle = NewStyle
                borderSize = GetBorderSize(adhForm, desktopScale)
                SetVisualBounds(rOldBounds)
                adhForm.Opacity = 1.0
                isActive = True
            End If
        End Sub
#End If

    End Class

End Namespace
