VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolbarMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements WinSubHookGX.GXISubclass

Private Const TPM_NOANIMATION       As Long = &H4000&

Private Enum TRACKINGSTATE
    TRACK_NONE = 0
    TRACK_BUTTON
    TRACK_POPUP
End Enum

Private Type TPMPARAMS
    cbSize                          As Long
    rcExclude                       As RECT
End Type


Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
                                                  ByVal nPos As Long) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long

Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, _
                                                        ByVal un As Long, _
                                                        ByVal n1 As Long, _
                                                        ByVal n2 As Long, _
                                                        ByVal hwnd As Long, _
                                                        lpTPMParams As TPMPARAMS) As Long

Private m_bProcessRightArrow    As Boolean
Private m_bProcessLeftArrow     As Boolean
Private m_bIn                   As Boolean
Private m_bEscapeWasPressed     As Boolean
Private m_hMenuTracking         As Long
Private m_iPopupTracking        As Long
Private m_iNewPopup             As Long
Private m_hWnd                  As Long
Private m_lPtr                  As Long
Private m_iExit                 As Long
Private m_tPMouse               As POINTAPI
Private m_iTrackingState        As TRACKINGSTATE
Private m_TbarSubclass          As GXMSubclass

Private Sub Class_Initialize()
    Set m_TbarSubclass = New GXMSubclass
End Sub

Private Sub CancelMenuAndTrackNewOne(ByVal iNewPopup As Long)

Dim cBar       As clsRCMMenu
Dim hMenuPopup As Long

    If iNewPopup > 0 Then
        If iNewPopup <> m_iPopupTracking Then
            If GetBar(cBar) Then
                hMenuPopup = cBar.GetMenuHandle(iNewPopup)
                If hMenuPopup <> 0 Then
                    PostMessage m_hWnd, &H1F, 0&, 0&
                    m_iNewPopup = iNewPopup
                End If
            End If
        End If
    End If
    
End Sub

Private Sub ComputeMenuTrackPoint(ByRef rc As RECT, _
                                  tPM As TPMPARAMS, _
                                  tP As POINTAPI)
    
    tP.x = rc.left
    tP.y = rc.Bottom
    LSet tPM.rcExclude = rc
    
End Sub

Private Function GetBar(ByRef cBar As clsRCMMenu) As Boolean

    If Not m_lPtr = 0 Then
        Set cBar = ObjectFromPtr(m_lPtr)
        GetBar = True
    End If
    
End Function

Private Function GetButtonCount() As Long

Dim cBar As clsRCMMenu

    If GetBar(cBar) Then
        GetButtonCount = cBar.p_MenuCount
    End If
    
End Function

Private Function GetButtonVisible(ByVal iButton As Long) As Boolean
    GetButtonVisible = True
End Function

Private Function GetHotItem() As Long

Dim cBar As clsRCMMenu

    'If GetBar(cBar) Then
   '     GetHotItem = cBar.HotItem
   ' End If
    
End Function

Private Function GetNextOrPrevButton(ByVal iButton As Long, _
                                     ByVal bPrev As Boolean) As Long

Dim iSB    As Long
Dim bfound As Boolean

    If bPrev Then
        iSB = iButton
        Do While Not bfound
            iButton = iButton - 1
            If iButton < 1 Then
                iButton = GetButtonCount()
            End If
            If Not (GetButtonVisible(iButton)) Then
                If iButton = iSB Then
                    iButton = -1
                    Exit Do
                End If
            Else
                bfound = True
            End If
        Loop
    Else
        iSB = iButton
        Do While Not bfound
            iButton = iButton + 1
            If iButton > GetButtonCount() Then
                iButton = 1
            End If
            If Not GetButtonVisible(iButton) Then
                If iButton = iSB Then
                    iButton = -1
                    Exit Do
                End If
            Else
                bfound = True
            End If
        Loop
    End If
    GetNextOrPrevButton = iButton
    
End Function

Private Sub GetRect(ByVal iButton As Long, _
                    ByRef tR As RECT)

Dim cBar As clsRCMMenu

    With tR
        .left = 0
        .top = 0
        .Bottom = 0
        .Right = 0
    End With
    If GetBar(cBar) Then
        If iButton > 0 Then
            If iButton <= cBar.p_MenuCount Then
                cBar.GetRect iButton, tR
            End If
        End If
    End If
    
End Sub

Private Function HitTest(pt As POINTAPI) As Long

Dim cBar As clsRCMMenu
    If GetBar(cBar) Then
        HitTest = cBar.HitTest(pt)
    End If
    
End Function

Private Property Get IsValidButton(ByVal iButton As Long) As Boolean

    If iButton > 0 Then
        IsValidButton = True
    End If
    
End Property

Friend Function MenuInput(m As Msg) As Boolean

Dim iMsg    As Long
Dim vKey    As Long
Dim tP      As POINTAPI
Dim iButton As Long

    Debug.Assert m_iTrackingState = TRACK_POPUP  '; // sanity check
    iMsg = m.message
    If iMsg = &H100 Then
        '// handle left/right-arrow.
        vKey = m.wParam
        If ((vKey = vbKeyLeft And m_bProcessLeftArrow) Or (vKey = vbKeyRight And m_bProcessRightArrow)) Then
            CancelMenuAndTrackNewOne GetNextOrPrevButton(m_iPopupTracking, vKey = vbKeyLeft)
            MenuInput = True ' // eat it
            ' // escape:
        ElseIf (vKey = vbKeyEscape) Then
            m_bEscapeWasPressed = True ';    // (menu will abort itself)
        End If
    ElseIf (iMsg = &H200 Or iMsg = &H201) Then
        '// handle mouse move or click
        LSet tP = m.pt
        'ScreenToClient m_hWndBand, tP
        If iMsg = &H200 Then
            iButton = HitTest(tP)
            If IsValidButton(iButton) Then
                If iButton <> m_iPopupTracking Then
                    '// user moved mouse over a different button: track its popup
                    CancelMenuAndTrackNewOne iButton
                End If
            End If
            LSet m_tPMouse = tP
        ElseIf iMsg = &H201 Then
            If HitTest(tP) = m_iPopupTracking Then
                '// user clicked on same button I am tracking: cancel menu
                CancelMenuAndTrackNewOne -1
                MenuInput = True ' // eat it
            End If
        End If
    ElseIf iMsg = &H202 Or iMsg = &H205 Then
    End If

End Function

Private Sub MenuSelect(ByVal hMenu As Long, _
                       ByVal iItem As Long)
    
    If m_iTrackingState > 0 Then
        '// process right-arrow if item is NOT a submenu
        m_bProcessRightArrow = (GetSubMenu(hMenu, iItem) = 0)
        '// process left-arrow if curent menu is one I'm tracking
        m_bProcessLeftArrow = (hMenu = m_hMenuTracking)
    End If
    
End Sub

Private Sub PressButton(ByVal iButton As Long, _
                        ByVal bState As Boolean)

Dim cBar   As clsRCMMenu

    If GetBar(cBar) Then
        If iButton > 0 Then
            If iButton <= cBar.p_MenuCount Then
                cBar.PressButton iButton, bState
            End If
        End If
    End If
    
End Sub

Private Function SetHotItem(ByVal iButton As Long) As Long

Dim cBar As clsRCMMenu

    If GetBar(cBar) Then
        cBar.HotItem = iButton
    End If
    
End Function

Private Sub SetTrackingState(ByVal iState As TRACKINGSTATE, _
                             ByVal iButton As Long)
    
    If iState <> m_iTrackingState Then
        If iState = TRACK_NONE Then
            iButton = -1
        End If
        SetHotItem iButton
        '// could be none (-1)
        If iState = TRACK_POPUP Then
            '// set related state stuff
            m_bEscapeWasPressed = False 'FALSE;   // assume Esc key not pressed
            m_bProcessRightArrow = True        '// assume left/right arrow..
            m_bProcessLeftArrow = True         '; // ..will move to prev/next popup
            m_iPopupTracking = iButton          '// which popup I'm tracking
        End If
        m_iTrackingState = iState
    End If
    
End Sub

Private Sub ToggleTrackButtonMode()

    If (m_iTrackingState = TRACK_NONE Or m_iTrackingState = TRACK_BUTTON) Then
        If m_iTrackingState = TRACK_NONE Then
            SetTrackingState TRACK_BUTTON, 1
        Else
            SetTrackingState TRACK_NONE, 1
        End If
    End If
    
End Sub

Friend Function TrackPopup(ByVal iButton As Long, _
                           Optional ByVal lMenuOffsetY As Long, _
                           Optional bTransparent As Boolean) As Long

Dim hMenuPopup  As Long
Dim lR          As Long
Dim lRtnID      As Long
Dim tPM         As TPMPARAMS
Dim rcButton    As RECT
Dim pt          As POINTAPI
Dim cBar        As clsRCMMenu
Dim lFlags      As Long

    If Not m_bIn Then
        m_bIn = True
        m_iNewPopup = iButton
        If GetBar(cBar) Then
            Do While (m_iNewPopup > -1)
                '// while user selects another menu
                lRtnID = 0
                m_iNewPopup = -1
                '// assume quit after this
                PressButton iButton, True
                '// press the button
                'UpdateWindow ToolbarhWnd(m_hWnd)
                '// and force repaint now
                SetTrackingState TRACK_POPUP, iButton
                '// enter tracking state
                '// Need to install a hook to trap menu input in order to make
                '// left/right-arrow keys and "hot" mouse tracking work.
                AttachMsgHook Me
                '// get submenu and display it beneath button
                GetRect iButton, rcButton
                'ClientRectToScreen m_hWndBand, rcButton
                tPM.cbSize = Len(tPM)
                ComputeMenuTrackPoint rcButton, tPM, pt
                'hMenuPopup = GetSubMenu(m_hMenu, iButton)
                hMenuPopup = cBar.GetMenuHandle(iButton)
                If Not hMenuPopup = 0 Then
                    ' Show the menu:
                    m_hMenuTracking = hMenuPopup
                    '/*** offset secondary menubarbg here with image height
                    '/*** turn off animation flags for transparency
                    
                    If bTransparent Then
                        lFlags = lFlags Or TPM_NOANIMATION
                    Else
                        lFlags = TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_VERTICAL
                    End If
                    lR = TrackPopupMenuEx(hMenuPopup, lFlags, pt.x, pt.y + lMenuOffsetY, m_hWnd, tPM)
                    'lR is the ID of the menu
                    lRtnID = lR
                End If
                '// uninstall hook.
                DetachMsgHook
                PressButton iButton, False
                '// un-press button
                'UpdateWindow ToolbarhWNd(m_hWnd)
                '// and force repaint now
                '// If the user exited the menu loop by pressing Escape,
                '// return to track-button state; otherwise normal non-tracking state.
                If m_bEscapeWasPressed Then
                    SetTrackingState TRACK_NONE, iButton
                Else
                    SetTrackingState TRACK_NONE, iButton
                End If
                '// If the user moved mouse to a new top-level popup (eg from File to
                '// Edit button), I will have posted a &H1F to quit
                '// the first popup, and set m_iNewPopup to the new menu to show.
                '// Otherwise, m_iNewPopup will be -1 as set above.
                '// So just set iButton to the next popup menu and keep looping...
                iButton = m_iNewPopup
            Loop
            'Set hot button if mouse is over, otherwise not:
            'The ID of the selected menu
            TrackPopup = lRtnID
        End If
        m_bIn = False
    End If
    
End Function

Friend Sub CoolMenuAttach(ByRef hWndA As Long, _
                          ByVal cBar As clsRCMMenu)

    m_iExit = 0
    CoolMenuDetach
    If hWndA = 0 Then Exit Sub
    m_hWnd = hWndA
    SendMessageLong m_hWnd, &H211, 0&, 0&
    With m_TbarSubclass
        .Subclass m_hWnd, Me
        .AddMessage m_hWnd, &H11F, MSG_BEFORE
    End With
    m_lPtr = ObjPtr(cBar)

End Sub

Friend Sub CoolMenuDetach()

    If Not m_hWnd = 0 Then
        SendMessageLong m_hWnd, &H212, 0&, 0&
        With m_TbarSubclass
            .DeleteMessage m_hWnd, &H11F, MSG_BEFORE
            .UnSubclass m_hWnd
        End With
        m_hWnd = 0
    End If
    m_hWnd = 0
    m_lPtr = 0
    
End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As WinSubHookGX.eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)
    
    Select Case uMsg
    Case &H11F
        MenuSelect lParam, (wParam And &HFFFF&)
    End Select

End Sub

Private Sub Class_Terminate()
    CoolMenuDetach
    Set m_TbarSubclass = Nothing
End Sub
