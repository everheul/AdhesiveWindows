Imports System.Runtime.InteropServices

Namespace Global.AdhesiveWindows

    Friend Module NativeMethods

        <DllImport("dwmapi.dll")>
        Friend Function DwmGetWindowAttribute(ByVal hwnd As IntPtr,
                                          ByVal dwAttribute As Integer,
                                          ByRef pvAttribute As RECT,
                                          ByVal cbAttribute As Integer) As <MarshalAs(UnmanagedType.I4)> Integer
        End Function

        <DllImport("gdi32.dll")>
        Friend Function GetDeviceCaps(ByVal hdc As IntPtr,
                                  ByVal nIndex As Integer) As <MarshalAs(UnmanagedType.I4)> Integer
        End Function

        <DllImport("user32.dll")>
        Friend Function PostMessage(ByVal hWnd As IntPtr,
                                ByVal Msg As UInteger,
                                wParam As IntPtr,
                                lParam As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("dwmapi.dll")>
        Friend Sub DwmIsCompositionEnabled(ByRef uType As Integer)
        End Sub

        ' Used GetDeviceCaps constants:
        Friend Const DEVCAPS_VERTRES As Integer = 10
        Friend Const DEVCAPS_DESKTOPVERTRES As Integer = 117
        Friend Const DEVCAPS_LOGPIXELSY As Integer = 90

        ' Used DwmGetWindowAttribute constants:
        Friend Const DWMWA_EXTENDED_FRAME_BOUNDS As Long = 9

#Region "WndProcMsg Constants"
        Public Enum WndProcMsg
            WM_USER = &H400
            WMUSER_SHOWN = WM_USER + 1       ' SHOWN event simulation
            WMUSER_SETMINIMAXI = WM_USER + 2 ' en/disable mains mini/maximize buttons
            WMUSER_MOVE_TOOL = WM_USER + 3  ' move all tool windows the same way
            WMUSER_SHOW_TOOL = WM_USER + 4  ' hide or show the tool windows

            ' Most std WndProc Messages, with comments from https://www.pinvoke.net/default.aspx/Constants.WM
            WM_ACTIVATE = &H6               'The WM_ACTIVATE message is sent when a window is being activated or deactivated. This message is sent first to the window procedure of the top-level window being deactivated; it is then sent to the window procedure of the top-level window being activated.
            WM_ACTIVATEAPP = &H1C           'The WM_ACTIVATEAPP message is sent when a window belonging to a different application than the active window is about to be activated. The message is sent to the application whose window is being activated and to the application whose window is being deactivated.
            WM_AFXFIRST = &H360             'The WM_AFXFIRST specifies the first afx message.
            WM_AFXLAST = &H37F              'The WM_AFXFIRST specifies the last afx message.
            WM_APP = &H8000                 'The WM_APP constant is used by applications to help define private messages, usually of the form WM_APP+X, where X is an integer value.
            WM_ASKCBFORMATNAME = &H30C      'The WM_ASKCBFORMATNAME message is sent to the clipboard owner by a clipboard viewer window to request the name of a CF_OWNERDISPLAY clipboard format.
            WM_CANCELJOURNAL = &H4B         'The WM_CANCELJOURNAL message is posted to an application when a user cancels the application's journaling activities. The message is posted with a NULL window handle.
            WM_CANCELMODE = &H1F            'The WM_CANCELMODE message is sent to cancel certain modes, such as mouse capture. For example, the system sends this message to the active window when a dialog box or message box is displayed. Certain functions also send this message explicitly to the specified window regardless of whether it is the active window. For example, the EnableWindow function sends this message when disabling the specified window.
            WM_CAPTURECHANGED = &H215       'The WM_CAPTURECHANGED message is sent to the window that is losing the mouse capture.
            WM_CHANGECBCHAIN = &H30D        'The WM_CHANGECBCHAIN message is sent to the first window in the clipboard viewer chain when a window is being removed from the chain.
            WM_CHANGEUISTATE = &H127        'An application sends the WM_CHANGEUISTATE message to indicate that the user interface (UI) state should be changed.
            WM_CHAR = &H102                 'The WM_CHAR message is posted to the window with the keyboard focus when a WM_KEYDOWN message is translated by the TranslateMessage function. The WM_CHAR message contains the character code of the key that was pressed.
            WM_CHARTOITEM = &H2F            'Sent by a list box with the LBS_WANTKEYBOARDINPUT style to its owner in response to a WM_CHAR message.
            WM_CHILDACTIVATE = &H22         'The WM_CHILDACTIVATE message is sent to a child window when the user clicks the window's title bar or when the window is activated, moved, or sized.
            WM_CLEAR = &H303                'An application sends a WM_CLEAR message to an edit control or combo box to delete (clear) the current selection, if any, from the edit control.
            WM_CLOSE = &H10                 'The WM_CLOSE message is sent as a signal that a window or an application should terminate.
            WM_COMMAND = &H111              'The WM_COMMAND message is sent when the user selects a command item from a menu, when a control sends a notification message to its parent window, or when an accelerator keystroke is translated.
            WM_COMPACTING = &H41            'The WM_COMPACTING message is sent to all top-level windows when the system detects more than 12.5 percent of system time over a 30- to 60-second interval is being spent compacting memory. This indicates that system memory is low.
            WM_COMPAREITEM = &H39           'The system sends the WM_COMPAREITEM message to determine the relative position of a new item in the sorted list of an owner-drawn combo box or list box. Whenever the application adds a new item, the system sends this message to the owner of a combo box or list box created with the CBS_SORT or LBS_SORT style.
            WM_CONTEXTMENU = &H7B           'The WM_CONTEXTMENU message notifies a window that the user clicked the right mouse button (right-clicked) in the window.
            WM_COPY = &H301                 'An application sends the WM_COPY message to an edit control or combo box to copy the current selection to the clipboard in CF_TEXT format.
            WM_COPYDATA = &H4A              'An application sends the WM_COPYDATA message to pass data to another application.
            WM_CREATE = &H1                 'The WM_CREATE message is sent when an application requests that a window be created by calling the CreateWindowEx or CreateWindow function. (The message is sent before the function returns.) The window procedure of the new window receives this message after the window is created, but before the window becomes visible.
            WM_CTLCOLORBTN = &H135          'The WM_CTLCOLORBTN message is sent to the parent window of a button before drawing the button. The parent window can change the button's text and background colors. However, only owner-drawn buttons respond to the parent window processing this message.
            WM_CTLCOLORDLG = &H136          'The WM_CTLCOLORDLG message is sent to a dialog box before the system draws the dialog box. By responding to this message, the dialog box can set its text and background colors using the specified display device context handle.
            WM_CTLCOLOREDIT = &H133         'An edit control that is not read-only or disabled sends the WM_CTLCOLOREDIT message to its parent window when the control is about to be drawn. By responding to this message, the parent window can use the specified device context handle to set the text and background colors of the edit control.
            WM_CTLCOLORLISTBOX = &H134      'Sent to the parent window of a list box before the system draws the list box. By responding to this message, the parent window can set the text and background colors of the list box by using the specified display device context handle.
            WM_CTLCOLORMSGBOX = &H132       'The WM_CTLCOLORMSGBOX message is sent to the owner window of a message box before Windows draws the message box. By responding to this message, the owner window can set the text and background colors of the message box by using the given display device context handle.
            WM_CTLCOLORSCROLLBAR = &H137    'The WM_CTLCOLORSCROLLBAR message is sent to the parent window of a scroll bar control when the control is about to be drawn. By responding to this message, the parent window can use the display context handle to set the background color of the scroll bar control.
            WM_CTLCOLORSTATIC = &H138       'A static control, or an edit control that is read-only or disabled, sends the WM_CTLCOLORSTATIC message to its parent window when the control is about to be drawn. By responding to this message, the parent window can use the specified device context handle to set the text and background colors of the static control.
            WM_CUT = &H300                  'An application sends a WM_CUT message to an edit control or combo box to delete (cut) the current selection, if any, in the edit control and copy the deleted text to the clipboard in CF_TEXT format.
            WM_DEADCHAR = &H103             'The WM_DEADCHAR message is posted to the window with the keyboard focus when a WM_KEYUP message is translated by the TranslateMessage function. WM_DEADCHAR specifies a character code generated by a dead key. A dead key is a key that generates a character, such as the umlaut (double-dot), that is combined with another character to form a composite character. For example, the umlaut-O character (Ö) is generated by typing the dead key for the umlaut character, and then typing the O key.
            WM_DELETEITEM = &H2D            'Sent to the owner of a list box or combo box when the list box or combo box is destroyed or when items are removed by the LB_DELETESTRING, LB_RESETCONTENT, CB_DELETESTRING, or CB_RESETCONTENT message. The system sends a WM_DELETEITEM message for each deleted item. The system sends the WM_DELETEITEM message for any deleted list box or combo box item with nonzero item data.
            WM_DESTROY = &H2                'The WM_DESTROY message is sent when a window is being destroyed. It is sent to the window procedure of the window being destroyed after the window is removed from the screen. This message is sent first to the window being destroyed and then to the child windows (if any) as they are destroyed. During the processing of the message, it can be assumed that all child windows still exist.
            WM_DESTROYCLIPBOARD = &H307     'The WM_DESTROYCLIPBOARD message is sent to the clipboard owner when a call to the EmptyClipboard function empties the clipboard.
            WM_DEVICECHANGE = &H219         'Notifies an application of a change to the hardware configuration of a device or the computer.
            WM_DEVMODECHANGE = &H1B         'The WM_DEVMODECHANGE message is sent to all top-level windows whenever the user changes device-mode settings.
            WM_DISPLAYCHANGE = &H7E         'The WM_DISPLAYCHANGE message is sent to all windows when the display resolution has changed.
            WM_DRAWCLIPBOARD = &H308        'The WM_DRAWCLIPBOARD message is sent to the first window in the clipboard viewer chain when the content of the clipboard changes. This enables a clipboard viewer window to display the new content of the clipboard.
            WM_DRAWITEM = &H2B              'The WM_DRAWITEM message is sent to the parent window of an owner-drawn button, combo box, list box, or menu when a visual aspect of the button, combo box, list box, or menu has changed.
            WM_DROPFILES = &H233            'Sent when the user drops a file on the window of an application that has registered itself as a recipient of dropped files.
            WM_ENABLE = &HA                 'The WM_ENABLE message is sent when an application changes the enabled state of a window. It is sent to the window whose enabled state is changing. This message is sent before the EnableWindow function returns, but after the enabled state (WS_DISABLED style bit) of the window has changed.
            WM_ENDSESSION = &H16            'The WM_ENDSESSION message is sent to an application after the system processes the results of the WM_QUERYENDSESSION message. The WM_ENDSESSION message informs the application whether the session is ending.
            WM_ENTERIDLE = &H121            'The WM_ENTERIDLE message is sent to the owner window of a modal dialog box or menu that is entering an idle state. A modal dialog box or menu enters an idle state when no messages are waiting in its queue after it has processed one or more previous messages.
            WM_ENTERMENULOOP = &H211        'The WM_ENTERMENULOOP message informs an application's main window procedure that a menu modal loop has been entered.
            WM_ENTERSIZEMOVE = &H231        'The WM_ENTERSIZEMOVE message is sent one time to a window after it enters the moving or sizing modal loop. The window enters the moving or sizing modal loop when the user clicks the window's title bar or sizing border, or when the window passes the WM_SYSCOMMAND message to the DefWindowProc function and the wParam parameter of the message specifies the SC_MOVE or SC_SIZE value. The operation is complete when DefWindowProc returns.
            WM_ERASEBKGND = &H14            'The WM_ERASEBKGND message is sent when the window background must be erased (for example, when a window is resized). The message is sent to prepare an invalidated portion of a window for painting.
            WM_EXITMENULOOP = &H212         'The WM_EXITMENULOOP message informs an application's main window procedure that a menu modal loop has been exited.
            WM_EXITSIZEMOVE = &H232         'The WM_EXITSIZEMOVE message is sent one time to a window, after it has exited the moving or sizing modal loop. The window enters the moving or sizing modal loop when the user clicks the window's title bar or sizing border, or when the window passes the WM_SYSCOMMAND message to the DefWindowProc function and the wParam parameter of the message specifies the SC_MOVE or SC_SIZE value. The operation is complete when DefWindowProc returns.
            WM_FONTCHANGE = &H1D            'An application sends the WM_FONTCHANGE message to all top-level windows in the system after changing the pool of font resources.
            WM_GETDLGCODE = &H87            'The WM_GETDLGCODE message is sent to the window procedure associated with a control. By default, the system handles all keyboard input to the control; the system interprets certain types of keyboard input as dialog box navigation keys. To override this default behavior, the control can respond to the WM_GETDLGCODE message to indicate the types of input it wants to process itself.
            WM_GETFONT = &H31               'An application sends a WM_GETFONT message to a control to retrieve the font with which the control is currently drawing its text.
            WM_GETHOTKEY = &H33             'An application sends a WM_GETHOTKEY message to determine the hot key associated with a window.
            WM_GETICON = &H7F               'The WM_GETICON message is sent to a window to retrieve a handle to the large or small icon associated with a window. The system displays the large icon in the ALT+TAB dialog, and the small icon in the window caption.
            WM_GETMINMAXINFO = &H24         'The WM_GETMINMAXINFO message is sent to a window when the size or position of the window is about to change. An application can use this message to override the window's default maximized size and position, or its default minimum or maximum tracking size.
            WM_GETOBJECT = &H3D             'Active Accessibility sends the WM_GETOBJECT message to obtain information about an accessible object contained in a server application. Applications never send this message directly. It is sent only by Active Accessibility in response to calls to AccessibleObjectFromPoint, AccessibleObjectFromEvent, or AccessibleObjectFromWindow. However, server applications handle this message.
            WM_GETTEXT = &HD                'An application sends a WM_GETTEXT message to copy the text that corresponds to a window into a buffer provided by the caller.
            WM_GETTEXTLENGTH = &HE          'An application sends a WM_GETTEXTLENGTH message to determine the length, in characters, of the text associated with a window.
            WM_HANDHELDFIRST = &H358        'Definition Needed
            WM_HANDHELDLAST = &H35F         'Definition Needed
            WM_HELP = &H53                  'Indicates that the user pressed the F1 key. If a menu is active when F1 is pressed, WM_HELP is sent to the window associated with the menu; otherwise, WM_HELP is sent to the window that has the keyboard focus. If no window has the keyboard focus, WM_HELP is sent to the currently active window.
            WM_HOTKEY = &H312               'The WM_HOTKEY message is posted when the user presses a hot key registered by the RegisterHotKey function. The message is placed at the top of the message queue associated with the thread that registered the hot key.
            WM_HSCROLL = &H114              'This message is sent to a window when a scroll event occurs in the window's standard horizontal scroll bar. This message is also sent to the owner of a horizontal scroll bar control when a scroll event occurs in the control.
            WM_HSCROLLCLIPBOARD = &H30E     'The WM_HSCROLLCLIPBOARD message is sent to the clipboard owner by a clipboard viewer window. This occurs when the clipboard contains data in the CF_OWNERDISPLAY format and an event occurs in the clipboard viewer's horizontal scroll bar. The owner should scroll the clipboard image and update the scroll bar values.
            WM_ICONERASEBKGND = &H27        'Windows NT 3.51 and earlier: The WM_ICONERASEBKGND message is sent to a minimized window when the background of the icon must be filled before painting the icon. A window receives this message only if a class icon is defined for the window; otherwise, WM_ERASEBKGND is sent. This message is not sent by newer versions of Windows.
            WM_IME_CHAR = &H286             'Sent to an application when the IME gets a character of the conversion result. A window receives this message through its WindowProc function.
            WM_IME_COMPOSITION = &H10F      'Sent to an application when the IME changes composition status as a result of a keystroke. A window receives this message through its WindowProc function.
            WM_IME_COMPOSITIONFULL = &H284  'Sent to an application when the IME window finds no space to extend the area for the composition window. A window receives this message through its WindowProc function.
            WM_IME_CONTROL = &H283          'Sent by an application to direct the IME window to carry out the requested command. The application uses this message to control the IME window that it has created. To send this message, the application calls the SendMessage function with the following parameters.
            WM_IME_ENDCOMPOSITION = &H10E   'Sent to an application when the IME ends composition. A window receives this message through its WindowProc function.
            WM_IME_KEYDOWN = &H290          'Sent to an application by the IME to notify the application of a key press and to keep message order. A window receives this message through its WindowProc function.
            WM_IME_KEYLAST = &H10F          'Definition Needed
            WM_IME_KEYUP = &H291            'Sent to an application by the IME to notify the application of a key release and to keep message order. A window receives this message through its WindowProc function.
            WM_IME_NOTIFY = &H282           'Sent to an application to notify it of changes to the IME window. A window receives this message through its WindowProc function.
            WM_IME_REQUEST = &H288          'Sent to an application to provide commands and request information. A window receives this message through its WindowProc function.
            WM_IME_SELECT = &H285           'Sent to an application when the operating system is about to change the current IME. A window receives this message through its WindowProc function.
            WM_IME_SETCONTEXT = &H281       'Sent to an application when a window is activated. A window receives this message through its WindowProc function.
            WM_IME_STARTCOMPOSITION = &H10D 'Sent immediately before the IME generates the composition string as a result of a keystroke. A window receives this message through its WindowProc function.
            WM_INITDIALOG = &H110           'The WM_INITDIALOG message is sent to the dialog box procedure immediately before a dialog box is displayed. Dialog box procedures typically use this message to initialize controls and carry out any other initialization tasks that affect the appearance of the dialog box.
            WM_INITMENU = &H116             'The WM_INITMENU message is sent when a menu is about to become active. It occurs when the user clicks an item on the menu bar or presses a menu key. This allows the application to modify the menu before it is displayed.
            WM_INITMENUPOPUP = &H117        'The WM_INITMENUPOPUP message is sent when a drop-down menu or submenu is about to become active. This allows an application to modify the menu before it is displayed, without changing the entire menu.
            WM_INPUTLANGCHANGE = &H51       'The WM_INPUTLANGCHANGE message is sent to the topmost affected window after an application's input language has been changed. You should make any application-specific settings and pass the message to the DefWindowProc function, which passes the message to all first-level child windows. These child windows can pass the message to DefWindowProc to have it pass the message to their child windows, and so on.
            WM_KEYDOWN = &H100              'The WM_KEYDOWN message is posted to the window with the keyboard focus when a nonsystem key is pressed. A nonsystem key is a key that is pressed when the ALT key is not pressed.
            WM_KEYFIRST = &H100             'This message filters for keyboard messages.
            WM_KEYLAST = &H108              'This message filters for keyboard messages.
            WM_KEYUP = &H101                'The WM_KEYUP message is posted to the window with the keyboard focus when a nonsystem key is released. A nonsystem key is a key that is pressed when the ALT key is not pressed, or a keyboard key that is pressed when a window has the keyboard focus.
            WM_KILLFOCUS = &H8              'The WM_KILLFOCUS message is sent to a window immediately before it loses the keyboard focus.
            WM_LBUTTONDBLCLK = &H203        'The WM_LBUTTONDBLCLK message is posted when the user double-clicks the left mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_LBUTTONDOWN = &H201          'The WM_LBUTTONDOWN message is posted when the user presses the left mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_LBUTTONUP = &H202            'The WM_LBUTTONUP message is posted when the user releases the left mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_MBUTTONDBLCLK = &H209        'The WM_MBUTTONDBLCLK message is posted when the user double-clicks the middle mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_MBUTTONDOWN = &H207          'The WM_MBUTTONDOWN message is posted when the user presses the middle mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_MBUTTONUP = &H208            'The WM_MBUTTONUP message is posted when the user releases the middle mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_MDIACTIVATE = &H222          'An application sends the WM_MDIACTIVATE message to a multiple-document interface (MDI) client window to instruct the client window to activate a different MDI child window.
            WM_MDICASCADE = &H227           'An application sends the WM_MDICASCADE message to a multiple-document interface (MDI) client window to arrange all its child windows in a cascade format.
            WM_MDICREATE = &H220            'An application sends the WM_MDICREATE message to a multiple-document interface (MDI) client window to create an MDI child window.
            WM_MDIDESTROY = &H221           'An application sends the WM_MDIDESTROY message to a multiple-document interface (MDI) client window to close an MDI child window.
            WM_MDIGETACTIVE = &H229         'An application sends the WM_MDIGETACTIVE message to a multiple-document interface (MDI) client window to retrieve the handle to the active MDI child window.
            WM_MDIICONARRANGE = &H228       'An application sends the WM_MDIICONARRANGE message to a multiple-document interface (MDI) client window to arrange all minimized MDI child windows. It does not affect child windows that are not minimized.
            WM_MDIMAXIMIZE = &H225          'An application sends the WM_MDIMAXIMIZE message to a multiple-document interface (MDI) client window to maximize an MDI child window. The system resizes the child window to make its client area fill the client window. The system places the child window's window menu icon in the rightmost position of the frame window's menu bar, and places the child window's restore icon in the leftmost position. The system also appends the title bar text of the child window to that of the frame window.
            WM_MDINEXT = &H224              'An application sends the WM_MDINEXT message to a multiple-document interface (MDI) client window to activate the next or previous child window.
            WM_MDIREFRESHMENU = &H234       'An application sends the WM_MDIREFRESHMENU message to a multiple-document interface (MDI) client window to refresh the window menu of the MDI frame window.
            WM_MDIRESTORE = &H223           'An application sends the WM_MDIRESTORE message to a multiple-document interface (MDI) client window to restore an MDI child window from maximized or minimized size.
            WM_MDISETMENU = &H230           'An application sends the WM_MDISETMENU message to a multiple-document interface (MDI) client window to replace the entire menu of an MDI frame window, to replace the window menu of the frame window, or both.
            WM_MDITILE = &H226              'An application sends the WM_MDITILE message to a multiple-document interface (MDI) client window to arrange all of its MDI child windows in a tile format.
            WM_MEASUREITEM = &H2C           'The WM_MEASUREITEM message is sent to the owner window of a combo box, list box, list view control, or menu item when the control or menu is created.
            WM_MENUCHAR = &H120             'The WM_MENUCHAR message is sent when a menu is active and the user presses a key that does not correspond to any mnemonic or accelerator key. This message is sent to the window that owns the menu.
            WM_MENUCOMMAND = &H126          'The WM_MENUCOMMAND message is sent when the user makes a selection from a menu.
            WM_MENUDRAG = &H123             'The WM_MENUDRAG message is sent to the owner of a drag-and-drop menu when the user drags a menu item.
            WM_MENUGETOBJECT = &H124        'The WM_MENUGETOBJECT message is sent to the owner of a drag-and-drop menu when the mouse cursor enters a menu item or moves from the center of the item to the top or bottom of the item.
            WM_MENURBUTTONUP = &H122        'The WM_MENURBUTTONUP message is sent when the user releases the right mouse button while the cursor is on a menu item.
            WM_MENUSELECT = &H11F           'The WM_MENUSELECT message is sent to a menu's owner window when the user selects a menu item.
            WM_MOUSEACTIVATE = &H21         'The WM_MOUSEACTIVATE message is sent when the cursor is in an inactive window and the user presses a mouse button. The parent window receives this message only if the child window passes it to the DefWindowProc function.
            WM_MOUSEFIRST = &H200           'Use WM_MOUSEFIRST to specify the first mouse message. Use the PeekMessage() Function.
            WM_MOUSEHOVER = &H2A1           'The WM_MOUSEHOVER message is posted to a window when the cursor hovers over the client area of the window for the period of time specified in a prior call to TrackMouseEvent.
            WM_MOUSELAST = &H20D            'Definition Needed
            WM_MOUSELEAVE = &H2A3           'The WM_MOUSELEAVE message is posted to a window when the cursor leaves the client area of the window specified in a prior call to TrackMouseEvent.
            WM_MOUSEMOVE = &H200            'The WM_MOUSEMOVE message is posted to a window when the cursor moves. If the mouse is not captured, the message is posted to the window that contains the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_MOUSEWHEEL = &H20A           'The WM_MOUSEWHEEL message is sent to the focus window when the mouse wheel is rotated. The DefWindowProc function propagates the message to the window's parent. There should be no internal forwarding of the message, since DefWindowProc propagates it up the parent chain until it finds a window that processes it.
            WM_MOUSEHWHEEL = &H20E          'The WM_MOUSEHWHEEL message is sent to the focus window when the mouse's horizontal scroll wheel is tilted or rotated. The DefWindowProc function propagates the message to the window's parent. There should be no internal forwarding of the message, since DefWindowProc propagates it up the parent chain until it finds a window that processes it.
            WM_MOVE = &H3                   'The WM_MOVE message is sent after a window has been moved.
            WM_MOVING = &H216               'The WM_MOVING message is sent to a window that the user is moving. By processing this message, an application can monitor the position of the drag rectangle and, if needed, change its position.
            WM_NCACTIVATE = &H86            'Non Client Area Activated Caption(Title) of the Form
            WM_NCCALCSIZE = &H83            'The WM_NCCALCSIZE message is sent when the size and position of a window's client area must be calculated. By processing this message, an application can control the content of the window's client area when the size or position of the window changes.
            WM_NCCREATE = &H81              'The WM_NCCREATE message is sent prior to the WM_CREATE message when a window is first created.
            WM_NCDESTROY = &H82             'The WM_NCDESTROY message informs a window that its nonclient area is being destroyed. The DestroyWindow function sends the WM_NCDESTROY message to the window following the WM_DESTROY message. WM_DESTROY is used to free the allocated memory object associated with the window.
            WM_NCHITTEST = &H84             'The WM_NCHITTEST message is sent to a window when the cursor moves, or when a mouse button is pressed or released. If the mouse is not captured, the message is sent to the window beneath the cursor. Otherwise, the message is sent to the window that has captured the mouse.
            WM_NCLBUTTONDBLCLK = &HA3       'The WM_NCLBUTTONDBLCLK message is posted when the user double-clicks the left mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCLBUTTONDOWN = &HA1         'The WM_NCLBUTTONDOWN message is posted when the user presses the left mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCLBUTTONUP = &HA2           'The WM_NCLBUTTONUP message is posted when the user releases the left mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCMBUTTONDBLCLK = &HA9       'The WM_NCMBUTTONDBLCLK message is posted when the user double-clicks the middle mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCMBUTTONDOWN = &HA7         'The WM_NCMBUTTONDOWN message is posted when the user presses the middle mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCMBUTTONUP = &HA8           'The WM_NCMBUTTONUP message is posted when the user releases the middle mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCMOUSEHOVER = &H2A0         'The WM_NCMOUSEHOVER message is posted to a window when the cursor hovers over the nonclient area of the window for the period of time specified in a prior call to TrackMouseEvent.
            WM_NCMOUSELEAVE = &H2A2         'The WM_NCMOUSELEAVE message is posted to a window when the cursor leaves the nonclient area of the window specified in a prior call to TrackMouseEvent.
            WM_NCMOUSEMOVE = &HA0           'The WM_NCMOUSEMOVE message is posted to a window when the cursor is moved within the nonclient area of the window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCPAINT = &H85               'The WM_NCPAINT message is sent to a window when its frame must be painted.
            WM_NCRBUTTONDBLCLK = &HA6       'The WM_NCRBUTTONDBLCLK message is posted when the user double-clicks the right mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCRBUTTONDOWN = &HA4         'The WM_NCRBUTTONDOWN message is posted when the user presses the right mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCRBUTTONUP = &HA5           'The WM_NCRBUTTONUP message is posted when the user releases the right mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCXBUTTONDBLCLK = &HAD       'The WM_NCXBUTTONDBLCLK message is posted when the user double-clicks the first or second X button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCXBUTTONDOWN = &HAB         'The WM_NCXBUTTONDOWN message is posted when the user presses the first or second X button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCXBUTTONUP = &HAC           'The WM_NCXBUTTONUP message is posted when the user releases the first or second X button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse, this message is not posted.
            WM_NCUAHDRAWCAPTION = &HAE      'The WM_NCUAHDRAWCAPTION message is an undocumented message related to themes. When handling WM_NCPAINT, this message should also be handled.
            WM_NCUAHDRAWFRAME = &HAF        'The WM_NCUAHDRAWFRAME message is an undocumented message related to themes. When handling WM_NCPAINT, this message should also be handled.
            WM_NEXTDLGCTL = &H28            'The WM_NEXTDLGCTL message is sent to a dialog box procedure to set the keyboard focus to a different control in the dialog box
            WM_NEXTMENU = &H213             'The WM_NEXTMENU message is sent to an application when the right or left arrow key is used to switch between the menu bar and the system menu.
            WM_NOTIFY = &H4E                'Sent by a common control to its parent window when an event has occurred or the control requires some information.
            WM_NOTIFYFORMAT = &H55          'Determines if a window accepts ANSI or Unicode structures in the WM_NOTIFY notification message. WM_NOTIFYFORMAT messages are sent from a common control to its parent window and from the parent window to the common control.
            WM_NULL = &H0                   'The WM_NULL message performs no operation. An application sends the WM_NULL message if it wants to post a message that the recipient window will ignore.
            WM_PAINT = &HF                  'Occurs when the control needs repainting
            WM_PAINTCLIPBOARD = &H309       'The WM_PAINTCLIPBOARD message is sent to the clipboard owner by a clipboard viewer window when the clipboard contains data in the CF_OWNERDISPLAY format and the clipboard viewer's client area needs repainting.
            WM_PAINTICON = &H26             'Windows NT 3.51 and earlier: The WM_PAINTICON message is sent to a minimized window when the icon is to be painted. This message is not sent by newer versions of Microsoft Windows, except in unusual circumstances explained in the Remarks.
            WM_PALETTECHANGED = &H311       'This message is sent by the OS to all top-level and overlapped windows after the window with the keyboard focus realizes its logical palette. This message enables windows that do not have the keyboard focus to realize their logical palettes and update their client areas.
            WM_PALETTEISCHANGING = &H310    'The WM_PALETTEISCHANGING message informs applications that an application is going to realize its logical palette.
            WM_PARENTNOTIFY = &H210         'The WM_PARENTNOTIFY message is sent to the parent of a child window when the child window is created or destroyed, or when the user clicks a mouse button while the cursor is over the child window. When the child window is being created, the system sends WM_PARENTNOTIFY just before the CreateWindow or CreateWindowEx function that creates the window returns. When the child window is being destroyed, the system sends the message before any processing to destroy the window takes place.
            WM_PASTE = &H302                'An application sends a WM_PASTE message to an edit control or combo box to copy the current content of the clipboard to the edit control at the current caret position. Data is inserted only if the clipboard contains data in CF_TEXT format.
            WM_PENWINFIRST = &H380          'Definition Needed
            WM_PENWINLAST = &H38F           'Definition Needed
            WM_POWER = &H48                 'Notifies applications that the system, typically a battery-powered personal computer, is about to enter a suspended mode. Obsolete : use POWERBROADCAST instead
            WM_POWERBROADCAST = &H218       'Notifies applications that a power-management event has occurred.
            WM_PRINT = &H317                'The WM_PRINT message is sent to a window to request that it draw itself in the specified device context, most commonly in a printer device context.
            WM_PRINTCLIENT = &H318          'The WM_PRINTCLIENT message is sent to a window to request that it draw its client area in the specified device context, most commonly in a printer device context.
            WM_QUERYDRAGICON = &H37         'The WM_QUERYDRAGICON message is sent to a minimized (iconic) window. The window is about to be dragged by the user but does not have an icon defined for its class. An application can return a handle to an icon or cursor. The system displays this cursor or icon while the user drags the icon.
            WM_QUERYENDSESSION = &H11       'The WM_QUERYENDSESSION message is sent when the user chooses to end the session or when an application calls one of the system shutdown functions. If any application returns zero, the session is not ended. The system stops sending WM_QUERYENDSESSION messages as soon as one application returns zero. After processing this message, the system sends the WM_ENDSESSION message with the wParam parameter set to the results of the WM_QUERYENDSESSION message.
            WM_QUERYNEWPALETTE = &H30F      'This message informs a window that it is about to receive the keyboard focus, giving the window the opportunity to realize its logical palette when it receives the focus.
            WM_QUERYOPEN = &H13             'The WM_QUERYOPEN message is sent to an icon when the user requests that the window be restored to its previous size and position.
            WM_QUEUESYNC = &H23             'The WM_QUEUESYNC message is sent by a computer-based training (CBT) application to separate user-input messages from other messages sent through the WH_JOURNALPLAYBACK Hook procedure.
            WM_QUIT = &H12                  'Once received, it ends the application's Message Loop, signaling the application to end. It can be sent by pressing Alt+F4, Clicking the X in the upper right-hand of the program, or going to File->Exit.
            WM_RBUTTONDBLCLK = &H206        'The WM_RBUTTONDBLCLK message is posted when the user double-clicks the right mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_RBUTTONDOWN = &H204          'The WM_RBUTTONDOWN message is posted when the user presses the right mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_RBUTTONUP = &H205            'The WM_RBUTTONUP message is posted when the user releases the right mouse button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_RENDERALLFORMATS = &H306     'The WM_RENDERALLFORMATS message is sent to the clipboard owner before it is destroyed, if the clipboard owner has delayed rendering one or more clipboard formats. For the content of the clipboard to remain available to other applications, the clipboard owner must render data in all the formats it is capable of generating, and place the data on the clipboard by calling the SetClipboardData function.
            WM_RENDERFORMAT = &H305         'The WM_RENDERFORMAT message is sent to the clipboard owner if it has delayed rendering a specific clipboard format and if an application has requested data in that format. The clipboard owner must render data in the specified format and place it on the clipboard by calling the SetClipboardData function.
            WM_SETCURSOR = &H20             'The WM_SETCURSOR message is sent to a window if the mouse causes the cursor to move within a window and mouse input is not captured.
            WM_SETFOCUS = &H7               'When the controll got the focus
            WM_SETFONT = &H30               'An application sends a WM_SETFONT message to specify the font that a control is to use when drawing text.
            WM_SETHOTKEY = &H32             'An application sends a WM_SETHOTKEY message to a window to associate a hot key with the window. When the user presses the hot key, the system activates the window.
            WM_SETICON = &H80               'An application sends the WM_SETICON message to associate a new large or small icon with a window. The system displays the large icon in the ALT+TAB dialog box, and the small icon in the window caption.
            WM_SETREDRAW = &HB              'An application sends the WM_SETREDRAW message to a window to allow changes in that window to be redrawn or to prevent changes in that window from being redrawn.
            WM_SETTEXT = &HC                'Text / Caption changed on the control. An application sends a WM_SETTEXT message to set the text of a window.
            WM_SETTINGCHANGE = &H1A         'An application sends the WM_SETTINGCHANGE message to all top-level windows after making a change to the WIN.INI file. The SystemParametersInfo function sends this message after an application uses the function to change a setting in WIN.INI.
            WM_SHOWWINDOW = &H18            'The WM_SHOWWINDOW message is sent to a window when the window is about to be hidden or shown
            WM_SIZE = &H5                   'The WM_SIZE message is sent to a window after its size has changed.
            WM_SIZECLIPBOARD = &H30B        'The WM_SIZECLIPBOARD message is sent to the clipboard owner by a clipboard viewer window when the clipboard contains data in the CF_OWNERDISPLAY format and the clipboard viewer's client area has changed size.
            WM_SIZING = &H214               'The WM_SIZING message is sent to a window that the user is resizing. By processing this message, an application can monitor the size and position of the drag rectangle and, if needed, change its size or position.
            WM_SPOOLERSTATUS = &H2A         'The WM_SPOOLERSTATUS message is sent from Print Manager whenever a job is added to or removed from the Print Manager queue.
            WM_STYLECHANGED = &H7D          'The WM_STYLECHANGED message is sent to a window after the SetWindowLong function has changed one or more of the window's styles.
            WM_STYLECHANGING = &H7C         'The WM_STYLECHANGING message is sent to a window when the SetWindowLong function is about to change one or more of the window's styles.
            WM_SYNCPAINT = &H88             'The WM_SYNCPAINT message is used to synchronize painting while avoiding linking independent GUI threads.
            WM_SYSCHAR = &H106              'The WM_SYSCHAR message is posted to the window with the keyboard focus when a WM_SYSKEYDOWN message is translated by the TranslateMessage function. It specifies the character code of a system character key — that is, a character key that is pressed while the ALT key is down.
            WM_SYSCOLORCHANGE = &H15        'This message is sent to all top-level windows when a change is made to a system color setting.
            WM_SYSCOMMAND = &H112           'A window receives this message when the user chooses a command from the Window menu (formerly known as the system or control menu) or when the user chooses the maximize button, minimize button, restore button, or close button.
            WM_SYSDEADCHAR = &H107          'The WM_SYSDEADCHAR message is sent to the window with the keyboard focus when a WM_SYSKEYDOWN message is translated by the TranslateMessage function. WM_SYSDEADCHAR specifies the character code of a system dead key — that is, a dead key that is pressed while holding down the ALT key.
            WM_SYSKEYDOWN = &H104           'The WM_SYSKEYDOWN message is posted to the window with the keyboard focus when the user presses the F10 key (which activates the menu bar) or holds down the ALT key and then presses another key. It also occurs when no window currently has the keyboard focus; in this case, the WM_SYSKEYDOWN message is sent to the active window. The window that receives the message can distinguish between these two contexts by checking the context code in the lParam parameter.
            WM_SYSKEYUP = &H105             'The WM_SYSKEYUP message is posted to the window with the keyboard focus when the user releases a key that was pressed while the ALT key was held down. It also occurs when no window currently has the keyboard focus; in this case, the WM_SYSKEYUP message is sent to the active window. The window that receives the message can distinguish between these two contexts by checking the context code in the lParam parameter.
            WM_TCARD = &H52                 'Sent to an application that has initiated a training card with Microsoft Windows Help. The message informs the application when the user clicks an authorable button. An application initiates a training card by specifying the HELP_TCARD command in a call to the WinHelp function.
            WM_TIMECHANGE = &H1E            'A message that is sent whenever there is a change in the system time.
            WM_TIMER = &H113                'The WM_TIMER message is posted to the installing thread's message queue when a timer expires. The message is posted by the GetMessage or PeekMessage function.
            WM_UNDO = &H304                 'An application sends a WM_UNDO message to an edit control to undo the last operation. When this message is sent to an edit control, the previously deleted text is restored or the previously added text is deleted.
            WM_UNINITMENUPOPUP = &H125      'The WM_UNINITMENUPOPUP message is sent when a drop-down menu or submenu has been destroyed.
            WM_USERCHANGED = &H54           'The WM_USERCHANGED message is sent to all windows after the user has logged on or off. When the user logs on or off, the system updates the user-specific settings. The system sends this message immediately after updating the settings.
            WM_VKEYTOITEM = &H2E            'Sent by a list box with the LBS_WANTKEYBOARDINPUT style to its owner in response to a WM_KEYDOWN message.
            WM_VSCROLL = &H115              'The WM_VSCROLL message is sent to a window when a scroll event occurs in the window's standard vertical scroll bar. This message is also sent to the owner of a vertical scroll bar control when a scroll event occurs in the control.
            WM_VSCROLLCLIPBOARD = &H30A     'The WM_VSCROLLCLIPBOARD message is sent to the clipboard owner by a clipboard viewer window when the clipboard contains data in the CF_OWNERDISPLAY format and an event occurs in the clipboard viewer's vertical scroll bar. The owner should scroll the clipboard image and update the scroll bar values.
            WM_WINDOWPOSCHANGED = &H47      'The WM_WINDOWPOSCHANGED message is sent to a window whose size, position, or place in the Z order has changed as a result of a call to the SetWindowPos function or another window-management function.
            WM_WINDOWPOSCHANGING = &H46     'The WM_WINDOWPOSCHANGING message is sent to a window whose size, position, or place in the Z order is about to change as a result of a call to the SetWindowPos function or another window-management function.
            WM_WININICHANGE = &H1A          'An application sends the WM_WININICHANGE message to all top-level windows after making a change to the WIN.INI file. The SystemParametersInfo function sends this message after an application uses the function to change a setting in WIN.INI. Note The WM_WININICHANGE message is provided only for compatibility with earlier versions of the system. Applications should use the WM_SETTINGCHANGE message.
            WM_XBUTTONDBLCLK = &H20D        'The WM_XBUTTONDBLCLK message is posted when the user double-clicks the first or second X button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_XBUTTONDOWN = &H20B          'The WM_XBUTTONDOWN message is posted when the user presses the first or second X button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
            WM_XBUTTONUP = &H20C            'The WM_XBUTTONUP message is posted when the user releases the first or second X button while the cursor is in the client area of a window. If the mouse is not captured, the message is posted to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
        End Enum
#End Region

        ' Commands sent with the WndProc WM_SYSCOMMAND events:
        Public Enum SysCommands As Integer
            SC_SIZE = &HF000
            SC_MOVE = &HF010
            SC_MINIMIZE = &HF020
            SC_MAXIMIZE = &HF030
            SC_NEXTWINDOW = &HF040
            SC_PREVWINDOW = &HF050
            SC_CLOSE = &HF060
            SC_VSCROLL = &HF070
            SC_HSCROLL = &HF080
            SC_MOUSEMENU = &HF090
            SC_KEYMENU = &HF100
            SC_ARRANGE = &HF110
            SC_RESTORE = &HF120
            SC_TASKLIST = &HF130
            SC_SCREENSAVE = &HF140
            SC_HOTKEY = &HF150
            SC_DEFAULT = &HF160
            SC_MONITORPOWER = &HF170
            SC_CONTEXTHELP = &HF180
        End Enum

        ' Used with the WndProc WM_WINDOWPOSCHANGING events
        <StructLayout(LayoutKind.Sequential)>
        Public Structure WINDOWPOS
            Friend hwnd As IntPtr
            Friend hwndInsertAfter As IntPtr
            Public x As Integer
            Public y As Integer
            Public cx As Integer
            Public cy As Integer
            Public flags As UInteger
            ' Debug - format WINDOWPOS struct:
            Public Overrides Function ToString() As String
                Return String.Format("pos: ({0},{1})  size: ({2},{3})  flags: {4}", x, y, cx, cy, WindowPosFlags.Format(GetType(WindowPosFlags), flags, "G"))
            End Function
        End Structure

        ' Commands (flags) sent with the WndProc WM_WINDOWPOSCHANGING events:
        <FlagsAttribute()>
        Public Enum WindowPosFlags As UInteger
            SWP_NOSIZE = &H1
            SWP_NOMOVE = &H2
            SWP_NOZORDER = &H4
            SWP_NOREDRAW = &H8
            SWP_NOACTIVATE = &H10
            SWP_NOACTION = SWP_NOMOVE And SWP_NOSIZE And SWP_NOZORDER And SWP_NOACTIVATE
            SWP_FRAMECHANGED = &H20
            SWP_SHOWWINDOW = &H40
            SWP_HIDEWINDOW = &H80
            SWP_NOCOPYBITS = &H100
            SWP_NOOWNERZORDER = &H200
            SWP_NOSENDCHANGING = &H400
            SWP_DEFERERASE = &H2000
            SWP_ASYNCWINDOWPOS = &H4000
            SWP_FULLSCREENHINT = &H8000         ' Undocumented - seen with Aero Snap Auto Maximize movements, and with mini/maximize/restore (!)
            SWP_SPLITSCREENHINT = &H100000      ' Undocumented - seen with Aero Snap Auto Arrange movements in Windows 10
            SWP_SPLITSCREENHINT7 = &H1000000    ' Undocumented - seen with Aero Snap Auto Arrange movements in Windows 7
        End Enum

        ' Used with DwmGetWindowAttribute() to get the Aero bordersize
        <StructLayout(LayoutKind.Sequential)>
        Friend Structure RECT
            Dim Left As Integer
            Dim Top As Integer
            Dim Right As Integer
            Dim Bottom As Integer
            ' Used to scale the Aero bounds to scaled desktops (dpi)
            Friend Sub Scale(scl As Double)
                Left = CInt(Convert.ToDouble(Left) * scl)
                Top = CInt(Convert.ToDouble(Top) * scl)
                Right = CInt(Convert.ToDouble(Right) * scl)
                Bottom = CInt(Convert.ToDouble(Bottom) * scl)
            End Sub
        End Structure

        ' VB.net version of the HIWORD macro
        Friend Function HiWord(lValue As Long) As Integer
            Return If(lValue And &H80000000, (lValue \ 65535) - 1, HiWord = lValue \ 65535)
        End Function

        ' VB.net version of the LOWORD macro
        Friend Function LoWord(lValue As Long) As Integer
            Return If(lValue And &H8000&, &H8000 Or (lValue And &H7FFF&), lValue And &HFFFF&)
        End Function

        Friend Function DwmEnabled() As Boolean
            ' Don't test anything older then Vista.
            If (Environment.OSVersion.Version.Major >= 6) Then
                Dim res As Integer
                DwmIsCompositionEnabled(res)
                Return CBool(res)
            Else
                Return False
            End If
        End Function

        ' Find and return the desktop scale user setting (125% etc).
        ' This function is called at startup, so this should work for PROCESS_SYSTEM_DPI_AWARE applications (dpiAware = true).
        ' Changing the DPI during execution may have a disturbing effect on Adhesive Windows.
        Friend Function GetScaleFactor(frm As Form) As Double
            Dim scale As Double = 1.0
            If (DwmEnabled()) Then
                ' Get the vertical dpi-related values only (horizontal and vertical should be identical)
                Using grfx As Graphics = frm.CreateGraphics()
                    Dim desktop As IntPtr = grfx.GetHdc()
                    Dim physicalHeight As Integer = GetDeviceCaps(desktop, DEVCAPS_DESKTOPVERTRES)
                    Dim scaledHeight As Integer = GetDeviceCaps(desktop, DEVCAPS_VERTRES)
                    Dim logPixY As Integer = GetDeviceCaps(desktop, DEVCAPS_LOGPIXELSY)
                    If (scaledHeight <> physicalHeight) Then
                        ' This works well with Windows10:
                        scale = scaledHeight / physicalHeight
                    ElseIf (logPixY <> 96) Then
                        ' This with Windows 7 (Vista?):
                        scale = 96.0 / logPixY
                    End If
                    grfx.ReleaseHdc(desktop)
                End Using
            End If
            Return scale
        End Function

        ' Get the size of any (in)visible Aero borders.
        Friend Function GetBorderSize(frm As Form, scale As Double) As Rectangle

            ' Is DWM (Desktop Window Manager) active?
            If (DwmEnabled()) Then
                Try
                    ' Get the Visual Borders from DwmGetWindowAttribute and calc the
                    '  difference between that and the Form Bounds.
                    Dim rVisual As New RECT
                    Dim iRes As Long = DwmGetWindowAttribute(frm.Handle, DWMWA_EXTENDED_FRAME_BOUNDS, rVisual, Marshal.SizeOf(rVisual))
                    If (iRes = 0) Then
                        Dim rBnds As Rectangle = frm.Bounds

                        ' rVisual holds the bounds excluding the invisible borders, but unscaled (see GetDPI).
                        ' For instance, if scaling is set to 125%, these bounds must first be multiplied by 0.8
                        rVisual.Scale(scale)

                        ' Save the difference between the two bounds as Rectangle.
                        ' That will be (7, 0), (-14, -7) for scalable windows in Windows 10, but who knows..
                        Dim iVisualWidth As Integer = rVisual.Right - rVisual.Left
                        Dim iVisualHeight As Integer = rVisual.Bottom - rVisual.Top
                        Return New Rectangle(rVisual.Left - rBnds.Left, rVisual.Top - rBnds.Top,
                                                   iVisualWidth - rBnds.Width, iVisualHeight - rBnds.Height)
                    Else
                        Debug.WriteLine("DwmGetWindowAttribute returned: " & iRes.ToString)
                    End If
                Catch ex As Exception
                    Debug.WriteLine("DwmGetWindowAttribute error: " & ex.Message)
                End Try
            End If
            Return New Rectangle(0, 0, 0, 0)
        End Function

    End Module

End Namespace
