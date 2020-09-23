Attribute VB_Name = "mKeyboardHook"
Option Explicit


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                                                                  ByVal lpFn As Long, _
                                                                                  ByVal hMod As Long, _
                                                                                  ByVal dwThreadId As Long) As Long


Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
                                                      ByVal nCode As Long, _
                                                      ByVal wParam As Long, _
                                                      lParam As Any) As Long


Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Const HC_ACTION = 0
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const VK_TAB = &H9
Private Const VK_CONTROL = &H11
Private Const VK_ESCAPE = &H1B
Private Const WH_KEYBOARD_LL = 13
Private Const LLKHF_ALTDOWN = &H20


Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private m_lHookPtr  As Long
Private m_lObjPtr   As Long
Private m_tHook     As KBDLLHOOKSTRUCT


Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object

Dim objT As Object

       CopyMemory objT, lPtr, 4
       Set ObjectFromPtr = objT
       CopyMemory objT, 0&, 4
    
End Property

Private Property Get PtrFromObject(ByRef oObj As Object) As Long

    PtrFromObject = ObjPtr(oObj)

  End Property

Private Function KeyboardProc(ByVal nCode As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long

Dim bDiscard    As Boolean
Dim cRc         As clsRCM

'/* I'll add in alt menu hooks lator(maybe)

    If (nCode = HC_ACTION) Then
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            CopyMemory m_tHook, ByVal lParam, Len(m_tHook)
            bDiscard = (Not (m_tHook.flags And LLKHF_ALTDOWN) = 0)
        End If
    End If

    If bDiscard Then
        KeyboardProc = -1
        On Error Resume Next
        If Not m_lObjPtr = 0 Then
            Set cRc = ObjectFromPtr(m_lObjPtr)
            cRc.Draw_Rollover
        End If
    Else
        KeyboardProc = CallNextHookEx(0&, nCode, wParam, ByVal lParam)
    End If

End Function

Public Function InstallKeyboardHook(cRc As clsRCM) As Boolean

    If m_lHookPtr = 0 Then
        m_lHookPtr = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardProc, App.hInstance, 0)
        m_lObjPtr = PtrFromObject(cRc)
        InstallKeyboardHook = True
    Else
        InstallKeyboardHook = (UnhookWindowsHookEx(m_lHookPtr) = 0)
        m_lHookPtr = 0
        m_lObjPtr = 0
    End If

End Function

Public Function RemoveKeyboardHook() As Boolean

    If Not m_lHookPtr = 0 Then
        UnhookWindowsHookEx m_lHookPtr
        m_lHookPtr = 0
        m_lObjPtr = 0
    End If
    
End Function
