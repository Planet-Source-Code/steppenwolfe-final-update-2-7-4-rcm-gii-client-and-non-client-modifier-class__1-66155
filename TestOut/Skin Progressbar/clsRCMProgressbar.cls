VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMProgressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements WinSubHookGX.GXISubclass

Private Const CLR_DEFAULT                               As Long = &HFFFFFF
Private Const BAR_DEFAULT                               As Long = &HCC0011
Private Const CCM_FIRST                                 As Long = &H2000
Private Const WM_USER                                   As Long = &H400
Private Const CCM_LAST                                  As Long = (CCM_FIRST + &H200)
Private Const CCM_SETBKCOLOR                            As Long = (CCM_FIRST + 1)
Private Const PBM_SETPOS                                As Long = (WM_USER + 2)
Private Const PBM_SETSTEP                               As Long = (WM_USER + 4)
Private Const PBM_STEPIT                                As Long = (WM_USER + 5)
Private Const PBM_SETRANGE32                            As Long = (WM_USER + 6)
Private Const PBM_GETRANGE                              As Long = (WM_USER + 7)
Private Const PBM_GETPOS                                As Long = (WM_USER + 8)
Private Const PBM_SETBARCOLOR                           As Long = (WM_USER + 9)
Private Const PBM_SETBKCOLOR                            As Long = CCM_SETBKCOLOR
'/* frame styles
Private Const WS_EX_CLIENTEDGE                          As Long = &H200
Private Const WS_EX_STATICEDGE                          As Long = &H20000
Private Const GWL_STYLE                                 As Long = (-16)
Private Const GWL_EXSTYLE                               As Long = (-20)
Private Const WS_BORDER                                 As Long = &H800000
Private Const SWP_NOMOVE                                As Long = &H2
Private Const SWP_NOSIZE                                As Long = &H1
Private Const SWP_FRAMECHANGED                          As Long = &H20
Private Const SWP_NOACTIVATE                            As Long = &H10
Private Const SWP_NOZORDER                              As Long = &H4
Private Const SWP_CHANGE                                As Long = SWP_NOACTIVATE Or _
    SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE


Public Enum PBSTYLE
    GRAPHICAL = 1
    FLATBRUSH = 2
    GRADBRUSH = 3
    XPSTYLE = 4
End Enum

Private Type RECT
    left                                                As Long
    top                                                 As Long
    Right                                               As Long
    Bottom                                              As Long
End Type

Private Type BITMAP
    bmType                                              As Long
    bmWidth                                             As Long
    bmHeight                                            As Long
    bmWidthBytes                                        As Long
    bmPlanes                                            As Integer
    bmBitsPixel                                         As Integer
    bmBits                                              As Long
End Type

Private Type PBRANGE
    lLow                                                As Long
    lHigh                                               As Long
End Type

Private Type PAINTSTRUCT
    hdc                                     As Long
    fErase                                  As Boolean
    rcPaint                                 As RECT
    fRestore                                As Boolean
    fIncUpdate                              As Boolean
    rgbReserved(32)                         As Byte
End Type


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bHasCalc                      As Boolean
Private m_bVertical                     As Boolean
Private m_bPainting                     As Boolean
Private m_lProgressBarObj               As Long
Private m_lDestWidth                    As Long
Private m_lDestHeight                   As Long
Private m_lStepSize                     As Long
Private m_lMaxStep                      As Long
Private m_lBarGradColor                 As Long
Private m_lCtrlHwnd                     As Long
Private m_lBackColor                    As Long
Private m_lForeColor                    As Long
Private m_lBarColor                     As Long
Private m_tRect                         As RECT
Private m_eBarStyle                     As PBSTYLE
Private m_pPBar                         As StdPicture
Private m_cRender                       As clsRender
Private m_cProgBar                      As clsStoreDc
Private m_GSubclass                     As GXMSubclass


Private Sub Class_Initialize()
    InitCommonControls
    m_lBackColor = -1
    m_lBarColor = -1
    m_lBarGradColor = -1
    m_eBarStyle = 2
End Sub

Public Property Get p_OProgressBarObj() As Object

Dim oTemp   As Object

    If m_lProgressBarObj = 0 Then
        Set p_OProgressBarObj = Nothing
    Else
        CopyMemory oTemp, m_lProgressBarObj, 4
        Set p_OProgressBarObj = oTemp
        CopyMemory oTemp, 0&, 4
    End If

End Property

Public Property Set p_OProgressBarObj(ByVal PropVal As Object)
    m_lProgressBarObj = ObjPtr(PropVal)
End Property

Public Property Get p_IPBar() As StdPicture
'/* bar image
    Set p_IPBar = m_pPBar
End Property

Public Property Set p_IPBar(PropVal As StdPicture)
    Set m_pPBar = PropVal
End Property

Public Property Get p_ForeColor() As Long
'/* text forecolor
    p_ForeColor = m_lForeColor
End Property

Public Property Let p_ForeColor(ByVal PropVal As Long)
    m_lForeColor = TranslateColor(PropVal)
End Property

Public Property Get p_BackColor() As Long
'/* control backcolor
    p_BackColor = m_lBackColor
End Property

Public Property Let p_BackColor(ByVal PropVal As Long)
    m_lBackColor = TranslateColor(PropVal)
End Property

Public Property Get p_BarColor() As Long
'/* control barcolor
    p_BarColor = m_lBarColor
End Property

Public Property Let p_BarColor(ByVal PropVal As Long)
    m_lBarColor = TranslateColor(PropVal)
End Property

Public Property Get p_BarGradColor() As Long
'/* gradient offset color
    p_BarGradColor = m_lBarGradColor
End Property

Public Property Let p_BarGradColor(ByVal PropVal As Long)
    m_lBarGradColor = TranslateColor(PropVal)
End Property

Public Property Get p_BarStyle() As PBSTYLE
'/* control barcolor
    p_BarStyle = m_eBarStyle
End Property

Public Property Let p_BarStyle(PropVal As PBSTYLE)
    m_eBarStyle = PropVal
End Property

Public Sub Progressbar_Attach()

    If p_OProgressBarObj Is Nothing Then Exit Sub
    m_lCtrlHwnd = p_OProgressBarObj.hwnd

    '/* change frame style
    Flatten_Container m_lCtrlHwnd

    '/* orientation
    m_bVertical = (p_OProgressBarObj.Orientation = 1)

    '/* set ctrl bg and bar colors
    SetBackColor m_lCtrlHwnd, m_lBackColor

    '/* use control -color change only
    If m_eBarStyle = FLATBRUSH Then
        SetBarColor m_lCtrlHwnd, m_lBarColor
        m_lCtrlHwnd = 0
        Exit Sub
    ElseIf m_eBarStyle = GRAPHICAL Then
        If p_IPBar Is Nothing Then
            m_eBarStyle = FLATBRUSH
            m_lCtrlHwnd = 0
            Exit Sub
        Else
            Image_Store m_lCtrlHwnd
        End If
    End If

    '/* create the mask window
    SetBarColor m_lCtrlHwnd, m_lBackColor

    '/* subclass control
    Set m_cRender = New clsRender
    Set m_GSubclass = New GXMSubclass
    Control_Attach m_lCtrlHwnd

End Sub

Private Function Get_Minimum() As Long
    Get_Minimum = p_OProgressBarObj.Min
End Function

Private Function Get_Maximum() As Long
    Get_Maximum = p_OProgressBarObj.Max
End Function

Private Sub Flatten_Container(ByVal lHwnd As Long, _
                              Optional ByVal bRestore As Boolean)

'/* flatten control container

Dim lStyle      As Long

    lStyle = GetWindowLong(lHwnd, GWL_STYLE)
    If bRestore Then
        lStyle = lStyle And WS_BORDER
    Else
        lStyle = lStyle And Not WS_BORDER
    End If
    SetWindowLong lHwnd, GWL_STYLE, lStyle

    lStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
    If bRestore Then
        lStyle = lStyle And WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    Else
        lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    End If

    SetWindowLong lHwnd, GWL_EXSTYLE, lStyle
    SetWindowPos lHwnd, 0&, 0&, 0&, 0&, 0&, SWP_CHANGE

End Sub

Private Sub SetBackColor(ByVal lHwnd As Long, _
                         ByVal lColor As Long)
'/* set back color

    If lHwnd = 0 Then Exit Sub
    If lColor = -1 Then lColor = CLR_DEFAULT
    SendMessage lHwnd, PBM_SETBKCOLOR, 0&, ByVal lColor

End Sub

Private Sub SetBarColor(ByVal lHwnd As Long, _
                        ByVal lColor As Long)
'/* set bar color

    If lHwnd = 0 Then Exit Sub
    If lColor = -1 Then lColor = BAR_DEFAULT
    SendMessage lHwnd, PBM_SETBARCOLOR, 0&, ByVal lColor

End Sub

Private Function GetHighRange(ByVal lHwnd As Long) As Long
'/* get max - doesn't work as described on msdn, always 30,000

    If lHwnd = 0 Then Exit Function
    GetHighRange = SendMessage(lHwnd, PBM_GETRANGE, 0&, 0&)

End Function

Private Function GetRange(ByVal lHwnd As Long) As PBRANGE
'/* get min/max struct - doesn't work, always returns 0 and 30,000
'/* regardless of pbar min/max settings, just more m$ rubbish..
'/* quote msdn: 'Retrieves information about the current high and low limits of a given progress bar control'
'/* umm, no.. some of the constants listed in m$ ref do not even exist, others do not work with vb control.

Dim tRange      As PBRANGE

    If lHwnd = 0 Then Exit Function
    SendMessage lHwnd, PBM_GETRANGE, 0&, tRange
    LSet GetRange = tRange

End Function

Private Sub SetRange32(ByVal lHwnd As Long, _
                       ByRef tRange As PBRANGE)

'/* set min/max 32

    If lHwnd = 0 Then Exit Sub
    SendMessage lHwnd, PBM_SETRANGE32, tRange.lLow, tRange.lHigh

End Sub

Private Sub SetStep(ByVal lHwnd As Long, _
                    ByVal lStep As Long)

'/* set step increment

    If lHwnd = 0 Then Exit Sub
    SendMessage lHwnd, PBM_SETSTEP, lStep, 0&

End Sub

Private Sub StepPBar(ByVal lHwnd As Long)
'/* advance bar 1 step

    If lHwnd = 0 Then Exit Sub
    SendMessage lHwnd, PBM_STEPIT, 0&, 0&

End Sub

Private Sub SetPosition(ByVal lHwnd As Long, _
                        ByVal lPosition As Long)

'/* set bar to position

    If lHwnd = 0 Then Exit Sub
    SendMessage lHwnd, PBM_SETPOS, lPosition, 0&

End Sub

Private Function GetPosition(ByVal lHwnd As Long) As Long
'/* get current step pos

    If lHwnd = 0 Then Exit Function
    GetPosition = SendMessage(lHwnd, PBM_GETPOS, 0&, 0&)

End Function

Private Sub Store_Size()

    GetClientRect m_lCtrlHwnd, m_tRect

End Sub

Private Function Calc_Step(ByVal lHwnd As Long) As Long

Dim lStep       As Long

On Error Resume Next

    lStep = (GetHighRange(lHwnd) / (Get_Maximum - Get_Minimum))
    m_lMaxStep = GetHighRange(lHwnd) / lStep
    Calc_Step = lStep

On Error GoTo 0

End Function

Private Sub PBar_Increment(ByVal lHwnd As Long)

Dim lIncr       As Long
Dim lPos        As Long
Dim tRect       As RECT

    If Not m_bHasCalc Then
        m_lStepSize = Calc_Step(lHwnd)
        Store_Size
        m_bHasCalc = True
    End If
    
    lIncr = GetPosition(lHwnd) / m_lStepSize
    LSet tRect = m_tRect
    If m_bVertical Then
        lPos = (m_tRect.Bottom / m_lMaxStep) * lIncr
        tRect.top = m_tRect.Bottom - lPos
        tRect.Bottom = lPos
        If lIncr = m_lMaxStep Then
            Paint_Background m_lCtrlHwnd
            tRect.top = 0
        End If
    Else
        lPos = (m_tRect.Right / m_lMaxStep) * lIncr
        tRect.Right = lPos
        If lIncr = m_lMaxStep Then
            Paint_Background m_lCtrlHwnd
            tRect.Right = 0
        End If
    End If

    Select Case m_eBarStyle
    '/* graphical
    Case 1
        PB_Graphical m_lCtrlHwnd, tRect
    '/* gradient
    Case 3
        PB_Gradient m_lCtrlHwnd, tRect
    '/* xp
    Case 4
        PB_XPStyle m_lCtrlHwnd, tRect
    End Select

End Sub

Private Sub Paint_Background(ByVal lHwnd As Long)

Dim lHdc            As Long
Dim lhBrush         As Long

    lHdc = GetDC(lHwnd)
    lhBrush = CreateSolidBrush(m_lBackColor)
    FillRect lHdc, m_tRect, lhBrush
    ReleaseDC lHwnd, lHdc
    DeleteObject lhBrush

End Sub

Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = &H0
    End If

End Function

Private Sub PB_Gradient(ByVal lHwnd As Long, _
                        ByRef tRect As RECT)

Dim lHdc        As Long

    lHdc = GetDC(lHwnd)
    With tRect
        If m_bVertical Then
            m_cRender.Gradient lHdc, .left, .Right, .top, .Bottom, m_lBarColor, m_lBarGradColor, Fill_Vertical, True
        Else
            m_cRender.Gradient lHdc, .left, .Right, .top, .Bottom, m_lBarColor, m_lBarGradColor, Fill_Horizontal
        End If
    End With
    ReleaseDC lHwnd, lHdc

End Sub

Private Sub PB_XPStyle(ByVal lHwnd As Long, _
                       ByRef tRect As RECT)

Dim lHdc            As Long
Dim lColor          As Long
Dim lClrCrt         As Long

    lHdc = GetDC(lHwnd)
    With m_cRender
        lColor = .XPShift(m_lBarColor, 50)
        lClrCrt = .XPShift(m_lBarColor, 150)
    End With
    
    With tRect
        If m_bVertical Then
            m_cRender.Gradient lHdc, .left, .Right, .top, .Bottom, m_lBarColor, m_lBarGradColor, Fill_Horizontal, True
        Else
            m_cRender.Gradient lHdc, .left, .Right, .top, .Bottom, lClrCrt, lColor, Fill_Vertical
        End If
    End With
    ReleaseDC lHwnd, lHdc

End Sub

Private Sub Image_Store(ByVal lHwnd As Long)

Dim tRect       As RECT

    Set m_cProgBar = New clsStoreDc
    m_cProgBar.CreateFromPicture m_pPBar
    GetClientRect lHwnd, tRect
    With tRect
        If m_bVertical Then
            m_lDestWidth = .Right
            m_lDestHeight = .Bottom / 10
        Else
            m_lDestWidth = .Right / 10
            m_lDestHeight = .Bottom
        End If
    End With

End Sub

Private Sub PB_Graphical(ByVal lHwnd As Long, _
                         ByRef tRect As RECT)

Dim lHdc            As Long
Dim lTmpDc          As Long
Dim lWndHnd         As Long
Dim lDrawDc         As Long
Dim lBmp            As Long
Dim lBmpOld         As Long
Dim lQuot           As Long
Dim lDstWdh         As Long

    lWndHnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lWndHnd)
    lDrawDc = CreateCompatibleDC(lTmpDc)
    lBmp = CreateCompatibleBitmap(lTmpDc, m_lDestWidth, m_lDestHeight)
    lBmpOld = SelectObject(lDrawDc, lBmp)
    lHdc = GetDC(lHwnd)

    If Not m_bVertical Then
        lQuot = Abs(tRect.Right) / m_lDestWidth
        lDstWdh = lQuot * m_lDestWidth
    End If

    With tRect
        m_cRender.Stretch lDrawDc, 0, 0, m_lDestWidth, m_lDestHeight, m_cProgBar.hdc, 0, 0, m_cProgBar.Width, m_cProgBar.Height, SRCCOPY
        m_cRender.Tile lHdc, .left, .top, lDstWdh, .Bottom, lDrawDc, m_lDestWidth, m_lDestHeight
    End With

    SelectObject lDrawDc, lBmpOld
    DeleteObject lBmp
    DeleteDC lDrawDc
    ReleaseDC lWndHnd, lTmpDc
    ReleaseDC lHwnd, lHdc

End Sub

Private Sub Control_Attach(ByVal lHwnd As Long)

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .Subclass lHwnd, Me
        .AddMessage lHwnd, WM_PAINT, MSG_BEFORE
        .AddMessage lHwnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
        .AddMessage lHwnd, WM_STYLECHANGED, MSG_BEFORE
        .AddMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
    End With

End Sub

Private Sub Control_Detach(ByVal lHwnd As Long)

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .DeleteMessage lHwnd, WM_PAINT, MSG_BEFORE
        .DeleteMessage lHwnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
        .DeleteMessage lHwnd, WM_STYLECHANGED, MSG_BEFORE
        .DeleteMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
        .UnSubclass lHwnd
    End With

End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As WinSubHookGX.eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)

Dim tPAINT  As PAINTSTRUCT

    If IsWindowVisible(lHwnd) = 0 Then
        lReturn = 1
        bHandled = True
        Exit Sub
    End If
    
    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            BeginPaint lHwnd, tPAINT
            PBar_Increment m_lCtrlHwnd
            EndPaint lHwnd, tPAINT
            m_bPainting = False
            bHandled = True
        Else
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If
    
    Case WM_ERASEBKGND
        lReturn = 1
    
    Case WM_STYLECHANGED, WM_WINDOWPOSCHANGED
        m_bHasCalc = False
    End Select

End Sub

Private Sub Cleanup()

    If Not m_lCtrlHwnd = 0 Then
        Control_Detach m_lCtrlHwnd
        If Not m_cProgBar Is Nothing Then Set m_cProgBar = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
    End If
    m_lCtrlHwnd = 0

End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub
