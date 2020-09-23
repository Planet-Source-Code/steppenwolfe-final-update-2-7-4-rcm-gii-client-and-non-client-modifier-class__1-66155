Attribute VB_Name = "mTimer"
Option Explicit

Public Const TPM_LEFTALIGN                  As Long = &H0
Public Const TPM_LEFTBUTTON                 As Long = &H0
Private Const TPM_NONOTIFY                  As Long = &H80
Private Const TPM_HORIZONTAL                As Long = &H0
Public Const TPM_VERTICAL                   As Long = &H40
Private Const WH_MSGFILTER                  As Long = (-1)
Private Const cTimerMax                     As Long = 100

Public Type RECT
    left                                    As Long
    top                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type
Public Type POINTAPI
    x                                       As Long
    y                                       As Long
End Type

Public Type Msg
    hwnd                                    As Long
    message                                 As Long
    wParam                                  As Long
    lParam                                  As Long
    time                                    As Long
    pt                                      As POINTAPI
End Type

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal nIDEvent As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
                                                      ByVal nCode As Long, _
                                                      ByVal wParam As Long, _
                                                      ByVal lParam As Long) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                                                                  ByVal lpFn As Long, _
                                                                                  ByVal hMod As Long, _
                                                                                  ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private m_hMsgHook                          As Long
Private m_lMsgHookPtr                       As Long
Private aTimers(1 To cTimerMax)             As clsTimer
Private m_cTimerCount                       As Long

Public Sub AttachMsgHook(cThis As clsToolbarMenu)

Dim lpFn As Long

    DetachMsgHook
    m_lMsgHookPtr = ObjPtr(cThis)
    lpFn = HookAddress(AddressOf MenuInputFilter)
    m_hMsgHook = SetWindowsHookEx(WH_MSGFILTER, lpFn, 0&, GetCurrentThreadId())
    Debug.Assert (m_hMsgHook <> 0)

End Sub

Public Sub DetachMsgHook()

    If m_hMsgHook <> 0 Then
        UnhookWindowsHookEx m_hMsgHook
        m_hMsgHook = 0
    End If
    
End Sub

Private Function HookAddress(ByVal lPtr As Long) As Long
    HookAddress = lPtr
End Function

Private Function MenuInputFilter(ByVal nCode As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long

Dim cM    As clsToolbarMenu
Dim lpMsg As Msg

    If nCode = 2 Then
        If Not m_lMsgHookPtr = 0 Then
            Set cM = ObjectFromPtr(m_lMsgHookPtr)
            CopyMemory lpMsg, ByVal lParam, Len(lpMsg)
            If cM.MenuInput(lpMsg) Then
                MenuInputFilter = 1
                Exit Function
            End If
        End If
    End If
    
    MenuInputFilter = CallNextHookEx(m_hMsgHook, nCode, wParam, lParam)
    
End Function

Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object

Dim objT As Object

    If Not (lPtr = 0) Then
        '/* Turn the pointer into an illegal, uncounted interface
        CopyMemory objT, lPtr, 4
        '/* Do NOT hit the End button here! You will crash!
        '/* Assign to legal reference
        Set ObjectFromPtr = objT
        ' Still do NOT hit the End button here! You will still crash!
        '/* Destroy the illegal reference
        CopyMemory objT, 0&, 4
    End If

End Property

Public Function TimerCreate(myTimer As clsTimer) As Boolean

'/* Create the timer
Dim i As Long

    myTimer.TimerID = SetTimer(0&, 0&, myTimer.Interval, AddressOf TimerProc)
    If myTimer.TimerID Then
        TimerCreate = True
        For i = 1 To cTimerMax
            If aTimers(i) Is Nothing Then
                Set aTimers(i) = myTimer
                If i > m_cTimerCount Then
                    m_cTimerCount = i
                End If
                TimerCreate = True
                Exit Function
            End If
        Next i
        myTimer.ErrRaise eeTooManyTimers
    Else
        '/* TimerCreate = False
        myTimer.TimerID = 0
        myTimer.Interval = 0
    End If

End Function

Public Function TimerDestroy(myTimer As clsTimer) As Long
'/* TimerDestroy = False
'/* Find and remove this timer

Dim i As Long

    '/* SPM - no need to count past the last timer set up in the
    '/* aTimer array:
    For i = 1 To m_cTimerCount
        ' Find timer in array
        If Not aTimers(i) Is Nothing Then
            If myTimer.TimerID = aTimers(i).TimerID Then
                KillTimer 0&, myTimer.TimerID
                ' Remove timer and set reference to nothing
                Set aTimers(i) = Nothing
                TimerDestroy = True
                Exit For
            End If
        End If
    Next i

End Function

Private Sub TimerProc(ByVal lngHwnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal idEvent As Long, _
                      ByVal dwTime As Long)

Dim i As Long

    '/* Find the timer with this ID
    For i = 1 To m_cTimerCount
        If Not (aTimers(i) Is Nothing) Then
            If idEvent = aTimers(i).TimerID Then
                '/* Generate the event
                aTimers(i).PulseTimer
                Exit Sub
            End If
        End If
    Next i
    '/* If the timer count is zero, it is extremelly likely that the IDE
    '/* has stopped. Kill the timer:
    If m_cTimerCount = 0 Then
        KillTimer 0&, idEvent
    End If

End Sub

Private Function StoreTimer(myTimer As clsTimer) As Boolean

Dim i As Long

    For i = 1 To m_cTimerCount
        If aTimers(i) Is Nothing Then
            Set aTimers(i) = myTimer
            StoreTimer = True
            Exit For
        End If
    Next

End Function
