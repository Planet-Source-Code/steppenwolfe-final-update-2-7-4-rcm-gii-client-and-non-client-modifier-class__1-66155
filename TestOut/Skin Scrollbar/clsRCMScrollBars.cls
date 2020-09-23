VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMScrollBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/~ Hats off to Carles.. I tried to hijack hwnd of scrollbar in control, but I think it may be
'/~ impossible. So I opted to use windows on windows technique. All written from scratch using
'/~ my own methods and Carles great coolscrollbar project as a guide..
'/~ http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63527&lngWId=1

Implements WinSubHookGX.GXISubclass

Private Const SIF_RANGE                                 As Long = &H1
Private Const SIF_PAGE                                  As Long = &H2
Private Const SIF_POS                                   As Long = &H4
Private Const SIF_ALL                                   As Double = (SIF_RANGE Or SIF_PAGE Or SIF_POS)
Private Const OBJID_HSCROLL                             As Long = &HFFFFFFFA
Private Const OBJID_VSCROLL                             As Long = &HFFFFFFFB
Private Const SS_CUSTOMDRAW                             As Long = &HD
Private Const GWL_STYLE                                 As Long = -16
Private Const GWL_EXSTYLE                               As Long = -20
Private Const WS_EX_TOOLWINDOW                          As Long = &H80&
Private Const WS_EX_CLIENTEDGE                          As Long = &H200
Private Const WS_EX_STATICEDGE                          As Long = &H20000
Private Const ESB_DISABLE_LEFT                          As Long = &H1
Private Const ESB_DISABLE_RIGHT                         As Long = &H2
Private Const WS_BORDER                                 As Long = &H800000
Private Const WS_CHILD                                  As Long = &H40000000
Private Const SWP_SHOWWINDOW                            As Long = &H40
Private Const SWP_NOMOVE                                As Long = &H2
Private Const SWP_NOSIZE                                As Long = &H1
Private Const SWP_FRAMECHANGED                          As Long = &H20
Private Const SWP_NOACTIVATE                            As Long = &H10
Private Const SWP_NOZORDER                              As Long = &H4
Private Const SW_HIDE                                   As Long = &H0
Private Const SW_NORMAL                                 As Long = &H1
Private Const SWP_CHANGE                                As Long = SWP_NOACTIVATE Or _
    SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE

Private Enum SCR_STYLE
    SB_HORZ = 0
    SB_VERT = 1
    SB_BOTH = 3
    SB_SZR = 4
End Enum

Private Enum SC_MESSAGE
    WM_NULL = &H0
    WM_CREATE = &H1
    WM_DESTROY = &H2
    WM_MOVE = &H3
    WM_SIZE = &H5
    WM_PAINT = &HF
    WM_ERASEBKGND = &H14
    WM_MEASUREITEM = &H2C
    WM_NOTIFY = &H4E
    WM_TIMER = &H113
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_MOUSELEAVE = &H2A3
    WM_DRAWITEM = &H2B
    WM_CTLCOLORSCROLLBAR = &H137
    WM_WINDOWPOSCHANGED = &H47
    WM_STYLECHANGED = &H7D
    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCCALCSIZE = &H83
    WM_NCHITTEST = &H84
    WM_NCPAINT = &H85
    WM_NCACTIVATE = &H86
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCMOUSEMOVE = &HA0
End Enum

'/* system metrics
Private Enum SYSTEM_METRICS
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CYSMCAPTION = 51
    SM_CXMINIMIZED = 57
    SM_CYMINIMIZED = 58
    SM_CXMAXTRACK = 59
    SM_CYMAXTRACK = 60
    SM_CXMAXIMIZED = 61
    SM_CYMAXIMIZED = 62
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum


Private Type RECT
    Left                                                As Long
    Top                                                 As Long
    Right                                               As Long
    Bottom                                              As Long
End Type

Private Type POINTAPI
    x                                                   As Long
    y                                                   As Long
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

Private Type SCROLLINFO
    cbSize                                              As Long
    fMask                                               As Long
    nMin                                                As Long
    nMax                                                As Long
    nPage                                               As Long
    nPos                                                As Long
    nTrackPos                                           As Long
End Type

Private Type SCROLLBARINFO
    cbSize                                              As Long
    rcScrollBar                                         As RECT
    dxyLineButton                                       As Long
    xyThumbTop                                          As Long
    xyThumbBottom                                       As Long
    reserved                                            As Long
    rgstate(0 To 5)                                     As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                              As Long
    dwFlags                                             As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                           As Long
    dwHoverTime                                         As Long
End Type


Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetScrollBarInfo Lib "user32" (ByVal hwnd As Long, _
                                                        ByVal idObject As Long, _
                                                        psbi As SCROLLBARINFO) As Long

Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, _
                                                     ByVal n As Long, _
                                                     lpScrollInfo As SCROLLINFO) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                                                              ByVal lpClassName As String, _
                                                                              ByVal lpWindowName As String, _
                                                                              ByVal dwStyle As Long, _
                                                                              ByVal x As Long, _
                                                                              ByVal y As Long, _
                                                                              ByVal nWidth As Long, _
                                                                              ByVal nHeight As Long, _
                                                                              ByVal hWndParent As Long, _
                                                                              ByVal hMenu As Long, _
                                                                              ByVal hInstance As Long, _
                                                                              lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal Length As Long)

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                                   ByVal lpRect As Long, _
                                                                                   ByVal bErase As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bHasVertical                                  As Boolean
Private m_bHasHorizontal                                As Boolean
Private m_bHasSizer                                     As Boolean
Private m_bTrackUser32                                  As Boolean
Private m_bPainting                                     As Boolean
Private m_lScrollBarObj                                 As Long
Private m_lHBtnWidth                                    As Long
Private m_lVBtnWidth                                    As Long
Private m_hWndParent                                    As Long
Private m_lEdgeDepth                                    As Long
Private m_lVXScroll                                     As Long
Private m_lVYScroll                                     As Long
Private m_lHXScroll                                     As Long
Private m_lHYScroll                                     As Long
Private m_lVYThumb                                      As Long
Private m_lHXThumb                                      As Long
Private m_lCtrlHnd                                      As Long
Private m_lVScrollHnd                                   As Long
Private m_lHScrollHnd                                   As Long
Private m_lSizerHnd                                     As Long
Private m_tVertRect                                     As RECT
Private m_tVScrInfo                                     As SCROLLINFO
Private m_tVScrBar                                      As SCROLLBARINFO
Private m_pVTrack                                       As StdPicture
Private m_pVThumb                                       As StdPicture
Private m_pVBtDwn                                       As StdPicture
Private m_pVBtUp                                        As StdPicture
Private m_cVTrackDc                                     As clsStoreDc
Private m_cVThumbDc                                     As clsStoreDc
Private m_cBtDwnDc                                      As clsStoreDc
Private m_cBtUpDc                                       As clsStoreDc
Private m_tHorzRect                                     As RECT
Private m_tHScrInfo                                     As SCROLLINFO
Private m_tHScrBar                                      As SCROLLBARINFO
Private m_pHTrack                                       As StdPicture
Private m_pHThumb                                       As StdPicture
Private m_pHBtLft                                       As StdPicture
Private m_pHBtRgt                                       As StdPicture
Private m_pSizer                                        As StdPicture
Private m_cHTrackDc                                     As clsStoreDc
Private m_cHThumbDc                                     As clsStoreDc
Private m_cHBtLftDc                                     As clsStoreDc
Private m_cHBtRgtDc                                     As clsStoreDc
Private m_cSzrDc                                        As clsStoreDc
Private m_cRender                                       As clsRender
Private m_GSubclass                                     As GXMSubclass


Private Sub Class_Initialize()
    InitCommonControls
End Sub


Public Property Get p_OScrollBarObj() As Object

Dim oTemp   As Object

    If m_lScrollBarObj = 0 Then
        Set p_OScrollBarObj = Nothing
    Else
        CopyMemory oTemp, m_lScrollBarObj, 4
        Set p_OScrollBarObj = oTemp
        CopyMemory oTemp, 0&, 4
    End If

End Property

Public Property Set p_OScrollBarObj(ByVal PropVal As Object)
    m_lScrollBarObj = ObjPtr(PropVal)
End Property

Public Property Get p_IHBtLft() As StdPicture
'/* scrollbar horizontal button left
    Set p_IHBtLft = m_pHBtLft
End Property

Public Property Set p_IHBtLft(PropVal As StdPicture)
    Set m_pHBtLft = PropVal
End Property

Public Property Get p_IHBtRgt() As StdPicture
'/* scrollbar horizontal button right
    Set p_IHBtRgt = m_pHBtRgt
End Property

Public Property Set p_IHBtRgt(PropVal As StdPicture)
    Set m_pHBtRgt = PropVal
End Property

Public Property Get p_IHThumb() As StdPicture
'/* scrollbar horizontal thumb
    Set p_IHThumb = m_pHThumb
End Property

Public Property Set p_IHThumb(PropVal As StdPicture)
    Set m_pHThumb = PropVal
End Property

Public Property Get p_IHTrack() As StdPicture
'/* scrollbar horizontal track
    Set p_IHTrack = m_pHTrack
End Property

Public Property Set p_IHTrack(PropVal As StdPicture)
    Set m_pHTrack = PropVal
End Property

Public Property Get p_ISizer() As StdPicture
'/* scrollbar sizer
    Set p_ISizer = m_pSizer
End Property

Public Property Set p_ISizer(PropVal As StdPicture)
    Set m_pSizer = PropVal
End Property

Public Property Get p_IVBtDwn() As StdPicture
'/* scrollbar vertical button down
    Set p_IVBtDwn = m_pVBtDwn
End Property

Public Property Set p_IVBtDwn(PropVal As StdPicture)
    Set m_pVBtDwn = PropVal
End Property

Public Property Get p_IVBtUp() As StdPicture
'/* scrollbar vertical button up
    Set p_IVBtUp = m_pVBtUp
End Property

Public Property Set p_IVBtUp(PropVal As StdPicture)
    Set m_pVBtUp = PropVal
End Property

Public Property Get p_IVThumb() As StdPicture
'/* scrollbar vertical thumb
    Set p_IVThumb = m_pVThumb
End Property

Public Property Set p_IVThumb(PropVal As StdPicture)
    Set m_pVThumb = PropVal
End Property

Public Property Get p_IVTrack() As StdPicture
'/* scrollbar vertical track
    Set p_IVTrack = m_pVTrack
End Property

Public Property Set p_IVTrack(PropVal As StdPicture)
    Set m_pVTrack = PropVal
End Property


Public Sub Scrollbar_Attach()

Dim sClass As String

    m_lCtrlHnd = p_OScrollBarObj.hwnd
    If m_lCtrlHnd = 0 Then Exit Sub

    '/* reset border style
    Flatten_Container m_lCtrlHnd

    '/* store system sizes
    Scrollbar_Metrics

    '/* test for scrollbars
    Build_PreTest

    '/* load image classes
    If m_bHasVertical Then
        Create_Scrollbar SB_VERT
        Init_VImages
    End If

    If m_bHasHorizontal Then
        Create_Scrollbar SB_HORZ
        Init_HImages
    End If

    If m_bHasSizer Then
        Create_Scrollbar SB_SZR
        Init_SZImage
    End If

    '/* does not work this way with lvw
    If Not InStrB(1, IsClassName(m_lCtrlHnd), "ListView") > 0 Then
        Align_Scrollbars
    End If
    
    Reset_SCPos
    Set m_GSubclass = New GXMSubclass
    Set m_cRender = New clsRender
    Control_Attatch m_lCtrlHnd

End Sub

Private Sub Scrollbar_Metrics()
'/* store scrollbar metrics

    m_lVXScroll = GetSystemMetrics(SM_CXVSCROLL)
    m_lVYScroll = GetSystemMetrics(SM_CYVSCROLL)
    m_lHXScroll = GetSystemMetrics(SM_CXHSCROLL)
    m_lHYScroll = GetSystemMetrics(SM_CYHSCROLL)
    m_lVYThumb = GetSystemMetrics(SM_CYVTHUMB)
    m_lHXThumb = GetSystemMetrics(SM_CXHTHUMB)

End Sub

Private Sub Align_Scrollbars()

Dim lEdge               As Long
Dim lVertOffset         As Long
Dim lHorzOffset         As Long

    lEdge = Edge_Depth(m_lCtrlHnd)
    lVertOffset = m_lVYScroll * 2
    lHorzOffset = m_lHYScroll * 2

    If m_bHasHorizontal Then
        If Store_Horizontal(m_lCtrlHnd) Then
            OffsetRect m_tHorzRect, -(lEdge * 2), (lEdge * 2)
            With m_tHorzRect
                SetWindowPos m_lHScrollHnd, 0, .Left, .Top - lHorzOffset, (.Right - .Left), (.Bottom - .Top), SWP_SHOWWINDOW
            End With
        End If
    End If

    If m_bHasVertical Then
        If Store_Vertical(m_lCtrlHnd) Then
            OffsetRect m_tVertRect, -(lEdge * 2), (lEdge * 2)
            With m_tVertRect
                SetWindowPos m_lVScrollHnd, 0, .Left, .Top - lVertOffset, (.Right - .Left), (.Bottom - .Top), SWP_SHOWWINDOW
            End With
        End If
    End If

    If m_bHasSizer Then
        If m_bHasHorizontal And m_bHasVertical Then
            With m_tVertRect
                SetWindowPos m_lSizerHnd, 0, .Left, .Bottom - lVertOffset, m_lVXScroll, m_lVYScroll, SWP_SHOWWINDOW
            End With
        End If
    End If

End Sub

Private Sub Build_PreTest()

    If Store_Vertical(m_lCtrlHnd) Then
        m_bHasVertical = True
    End If
    If Store_Horizontal(m_lCtrlHnd) Then
        m_bHasHorizontal = True
    End If
    If m_bHasHorizontal And m_bHasVertical Then
        m_bHasSizer = True
    End If

End Sub

Private Sub Build_PostTest()

Dim sClass      As String

    sClass = IsClassName(m_lCtrlHnd)
    If InStrB(1, sClass, "ListView") > 0 Then
        Reset_SCPos
    Else
        If Not m_bHasHorizontal Then
            If Store_Horizontal(m_lCtrlHnd) Then
                Create_Scrollbar SB_HORZ
                Init_HImages
                m_bHasHorizontal = True
            End If
        End If

        If Not m_bHasVertical Then
            If Store_Vertical(m_lCtrlHnd) Then
                Create_Scrollbar SB_VERT
                Init_VImages
                m_bHasVertical = True
            End If
        End If

        If Not m_bHasSizer Then
            If m_bHasHorizontal And m_bHasVertical Then
                Create_Scrollbar SB_SZR
                Init_SZImage
                m_bHasSizer = True
            End If
        End If
    End If

End Sub

Private Function Create_Scrollbar(eStyle As SCR_STYLE) As Long

Dim bOnDesktop      As Boolean
Dim lEdge           As Long

    m_hWndParent = GetParent(m_lCtrlHnd)
    bOnDesktop = (m_hWndParent = GetDesktopWindow())
    lEdge = Edge_Depth(m_lCtrlHnd)

    Select Case eStyle
    Case SB_HORZ
        With m_tHorzRect
            m_lHScrollHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
        End With

    Case SB_VERT
        With m_tVertRect
            m_lVScrollHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
        End With

    Case SB_SZR
        m_lSizerHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
    End Select

End Function

Private Function Destroy_Scrollbar(eStyle As SCR_STYLE) As Long

    Select Case eStyle
    Case SB_HORZ
        If Not m_lHScrollHnd = 0 Then
            DestroyWindow m_lHScrollHnd
            m_lHScrollHnd = 0
        End If
    Case SB_VERT
        If Not m_lVScrollHnd = 0 Then
            DestroyWindow m_lVScrollHnd
            m_lVScrollHnd = 0
        End If
    Case SB_BOTH
        If Not m_lHScrollHnd = 0 Then
            DestroyWindow m_lHScrollHnd
            m_lHScrollHnd = 0
        End If
        If Not m_lVScrollHnd = 0 Then
            DestroyWindow m_lVScrollHnd
            m_lHScrollHnd = 0
        End If
        If Not m_lSizerHnd = 0 Then
            DestroyWindow m_lSizerHnd
            m_lSizerHnd = 0
        End If
    Case SB_SZR
        If Not m_lSizerHnd = 0 Then
            DestroyWindow m_lSizerHnd
            m_lSizerHnd = 0
        End If
    End Select

End Function

Private Sub Destroy_VImages()

    If Not m_cVTrackDc Is Nothing Then Set m_cVTrackDc = Nothing
    If Not m_cVThumbDc Is Nothing Then Set m_cVThumbDc = Nothing
    If Not m_cBtDwnDc Is Nothing Then Set m_cBtDwnDc = Nothing
    If Not m_cBtUpDc Is Nothing Then Set m_cBtUpDc = Nothing

End Sub

Private Function Edge_Depth(ByVal lHwnd As Long) As Long
'/* frame style depth

Dim lStyle          As Long
Dim lExStyle        As Long
Dim lEdge           As Long

    lStyle = GetWindowLong(lHwnd, GWL_STYLE)
    lExStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
    lEdge = -1 * ((lStyle And WS_BORDER) = WS_BORDER)
    lEdge = lEdge + -2 * ((lExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE)

    If InStrB(1, IsClassName(lHwnd), "ListView") > 0 Then
        If lEdge = 0 Then
            lEdge = 1
        End If
    Else
        If lEdge = 0 Then
            lEdge = 2
        End If
    End If
    Edge_Depth = lEdge

End Function

Private Sub Flatten_Container(ByVal lHwnd As Long, _
                              Optional ByVal bRestore As Boolean)

'/* toggle border styles
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
    SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_CHANGE

End Sub

Private Sub HBar_Paint(ByVal lOffset As Long)
'/* paint horz scrollbar

On Error GoTo Handler

Dim lHdc        As Long
Dim lDiff       As Long
Dim lHOffset    As Long
Dim lLOffset    As Long
Dim lSclHdc     As Long
Dim lSzrDc      As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long
Dim tRect       As RECT
Dim tPos        As POINTAPI

    '/* thumb size
    With m_tHScrBar
        lDiff = .xyThumbBottom - .xyThumbTop
    End With
    
    '/* backbuffer the image
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)
    lHdc = CreateCompatibleDC(lTmpDc)
    With m_tHorzRect
        lBmp = CreateCompatibleBitmap(lTmpDc, .Right, .Bottom)
        OffsetRect m_tHorzRect, -.Left, -.Top
    End With
    lBmpOld = SelectObject(lHdc, lBmp)
    lSclHdc = GetDC(m_lHScrollHnd)
    
    '/* apply offset
    If lOffset = 0 Then
        lLOffset = 1
    ElseIf lOffset = 1 Then
        lHOffset = 1
    End If
    
    With m_tHorzRect
        '/* build to temp dc
        m_cRender.Stretch lHdc, 0, 0, .Right, m_lHXScroll, m_cHTrackDc.hdc, 0, 0, m_cHTrackDc.Width, m_cHTrackDc.Height, SRCCOPY
        m_cRender.Blit lHdc, 0, 0, m_lHXScroll, .Bottom, m_cHBtRgtDc.hdc, m_lHBtnWidth * lLOffset, 0, SRCCOPY
        m_cRender.Blit lHdc, .Right - m_lHBtnWidth, .Top, .Right, .Bottom, m_cHBtLftDc.hdc, m_lHBtnWidth * lHOffset, 0, SRCCOPY
        If lDiff > m_cHThumbDc.Width Then
            m_cRender.Stretch lHdc, m_tHScrBar.xyThumbTop, 0, 3, m_lHYScroll + 1, m_cHThumbDc.hdc, 0, 0, 3, m_cHThumbDc.Height, SRCCOPY
            m_cRender.Stretch lHdc, m_tHScrBar.xyThumbTop + 3, 0, lDiff - 6, m_lHYScroll + 1, m_cHThumbDc.hdc, 3, 0, m_cHThumbDc.Width - 6, m_cHThumbDc.Height, SRCCOPY
            m_cRender.Stretch lHdc, m_tHScrBar.xyThumbBottom - 3, 0, 3, m_lHYScroll + 1, m_cHThumbDc.hdc, m_cHThumbDc.Width - 3, 0, 3, m_cHThumbDc.Height, SRCCOPY
        Else
            m_cRender.Stretch lHdc, m_tHScrBar.xyThumbTop, 0, lDiff, m_lHYScroll, m_cHThumbDc.hdc, 0, 0, m_cHThumbDc.Width, m_cHThumbDc.Height, SRCCOPY
        End If
        '/* copy to dest
        m_cRender.Blit lSclHdc, 0, 0, .Right, .Bottom, lHdc, 0, 0, SRCCOPY
    End With
    
    If m_bHasSizer Then
        lSzrDc = GetDC(m_lSizerHnd)
        With m_tHorzRect
            m_cRender.Stretch lSzrDc, 0, 0, m_lVXScroll, .Bottom, m_cSzrDc.hdc, 0, 0, m_cSzrDc.Width, m_cSzrDc.Height, SRCCOPY
        End With
        ReleaseDC m_lSizerHnd, lSzrDc
    End If
    '/* cleanup
    SelectObject lHdc, lBmpOld
    DeleteObject lBmp
    DeleteDC lHdc
    ReleaseDC lHwnd, lTmpDc
    ReleaseDC m_lVScrollHnd, lSclHdc

Handler:
    On Error GoTo 0

End Sub

Private Sub VBar_Paint(ByVal lOffset As Long)
'/* paint vert scrollbar

On Error GoTo Handler

Dim lHdc        As Long
Dim lDiff       As Long
Dim lHOffset    As Long
Dim lLOffset    As Long
Dim lSclHdc     As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long
Dim tRect       As RECT
Dim tPos        As POINTAPI

    '/* thumb size
    With m_tVScrBar
        lDiff = .xyThumbBottom - .xyThumbTop
    End With
    
    '/* backbuffer the image
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)
    lHdc = CreateCompatibleDC(lTmpDc)
    With m_tVertRect
        lBmp = CreateCompatibleBitmap(lTmpDc, .Right, .Bottom)
        OffsetRect m_tVertRect, -.Left, -.Top
    End With
    lBmpOld = SelectObject(lHdc, lBmp)
    lSclHdc = GetDC(m_lVScrollHnd)
    
    '/* apply offset
    If lOffset = 0 Then
        lLOffset = 1
    ElseIf lOffset = 1 Then
        lHOffset = 1
    End If
    
    With m_tVertRect
        '/* build to temp dc
        m_cRender.Stretch lHdc, 0, 0, m_lVXScroll, .Bottom, m_cVTrackDc.hdc, 0, 0, m_cVTrackDc.Width, m_cVTrackDc.Height, SRCCOPY
        m_cRender.Blit lHdc, 0, 0, m_lVXScroll, .Bottom, m_cBtUpDc.hdc, m_lVBtnWidth * lLOffset, 0, SRCCOPY
        m_cRender.Blit lHdc, 0, .Bottom - m_lVYScroll, m_lVXScroll, .Bottom, m_cBtDwnDc.hdc, m_lVBtnWidth * lHOffset, 0, SRCCOPY
        If lDiff > m_cVThumbDc.Height Then
            '/* top
            m_cRender.Stretch lHdc, 0, m_tVScrBar.xyThumbBottom - lDiff, m_lVXScroll + 1, 3, m_cVThumbDc.hdc, 0, 0, m_cVThumbDc.Width, 3, SRCCOPY
            '/* center
            m_cRender.Stretch lHdc, 0, (m_tVScrBar.xyThumbBottom - (lDiff - 3)), m_lVXScroll + 1, lDiff - 6, m_cVThumbDc.hdc, 0, 3, m_cVThumbDc.Width, m_cVThumbDc.Height - 6, SRCCOPY
            '/* bottom
            m_cRender.Stretch lHdc, 0, m_tVScrBar.xyThumbBottom - 3, m_lVXScroll + 1, 3, m_cVThumbDc.hdc, 0, m_cVThumbDc.Height - 3, m_cVThumbDc.Width, 3, SRCCOPY
        Else
            m_cRender.Stretch lHdc, 0, m_tVScrBar.xyThumbBottom - lDiff, m_lVXScroll, lDiff, m_cVThumbDc.hdc, 0, 0, m_cVThumbDc.Width, m_cVThumbDc.Height, SRCCOPY
        End If
        '/* copy to dest
        m_cRender.Blit lSclHdc, 0, 0, .Right, .Bottom, lHdc, 0, 0, SRCCOPY
    End With
    
    '/* cleanup
    SelectObject lHdc, lBmpOld
    DeleteObject lBmp
    DeleteDC lHdc
    ReleaseDC lHwnd, lTmpDc
    ReleaseDC m_lVScrollHnd, lSclHdc
    
Handler:
    On Error GoTo 0

End Sub

Private Sub Init_HImages()
'/* init hz image classes

Dim tBmp        As BITMAP

    '/* image size
    GetObject p_IHBtLft.handle, Len(tBmp), tBmp
    With tBmp
        m_lHBtnWidth = (.bmWidth / 2)
    End With
    
    '/* create the image classes
    Set m_cHTrackDc = New clsStoreDc
    m_cHTrackDc.CreateFromPicture p_IHTrack

    Set m_cHThumbDc = New clsStoreDc
    m_cHThumbDc.CreateFromPicture p_IHThumb

    Set m_cHBtLftDc = New clsStoreDc
    m_cHBtLftDc.CreateFromPicture p_IHBtLft

    Set m_cHBtRgtDc = New clsStoreDc
    m_cHBtRgtDc.CreateFromPicture p_IHBtRgt

End Sub

Private Sub Init_SZImage()

    '/* sizer
    Set m_cSzrDc = New clsStoreDc
    m_cSzrDc.CreateFromPicture p_ISizer

End Sub

Private Sub Init_VImages()
'/* init vt image classes

Dim tBmp        As BITMAP

    '/* cleanup
    Destroy_VImages
    GetObject p_IVBtDwn.handle, Len(tBmp), tBmp
    With tBmp
        m_lVBtnWidth = (.bmWidth / 2)
    End With
    
    '/* create the image classes
    Set m_cVTrackDc = New clsStoreDc
    m_cVTrackDc.CreateFromPicture p_IVTrack

    Set m_cVThumbDc = New clsStoreDc
    m_cVThumbDc.CreateFromPicture p_IVThumb

    Set m_cBtDwnDc = New clsStoreDc
    m_cBtDwnDc.CreateFromPicture p_IVBtDwn

    Set m_cBtUpDc = New clsStoreDc
    m_cBtUpDc.CreateFromPicture p_IVBtUp

End Sub

Private Function IsClassName(ByVal lHwnd As Long) As String

Dim lRet   As Long
Dim sClass As String

    sClass = Space$(255)
    lRet = GetClassName(lHwnd, sClass, 255)
    IsClassName = Left$(sClass, lRet)

End Function

Private Sub Control_Attatch(ByVal lHwnd As Long)

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .Subclass lHwnd, Me
        .AddMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
        .AddMessage lHwnd, WM_PAINT, MSG_BEFORE
        .AddMessage lHwnd, WM_VSCROLL, MSG_BEFORE
        .AddMessage lHwnd, WM_HSCROLL, MSG_BEFORE
        .AddMessage lHwnd, WM_STYLECHANGED, MSG_BEFORE
        .AddMessage lHwnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
    End With

End Sub

Private Sub Control_Detatch(ByVal lHwnd As Long)

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .DeleteMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
        .DeleteMessage lHwnd, WM_PAINT, MSG_BEFORE
        .DeleteMessage lHwnd, WM_VSCROLL, MSG_BEFORE
        .DeleteMessage lHwnd, WM_HSCROLL, MSG_BEFORE
        .DeleteMessage lHwnd, WM_STYLECHANGED, MSG_BEFORE
        .DeleteMessage lHwnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
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

Static bSelect  As Boolean
Static bFirst As Boolean

    If IsWindowVisible(m_lCtrlHnd) = 0 Then
        Mask_Show m_lHScrollHnd, False
        Mask_Show m_lVScrollHnd, False
        Mask_Show m_lSizerHnd, False
        Exit Sub
    Else
        If IsWindowVisible(m_lHScrollHnd) = 0 Then
            Mask_Show m_lHScrollHnd, True
        End If
        If IsWindowVisible(m_lVScrollHnd) = 0 Then
            Mask_Show m_lVScrollHnd, True
        End If
        If IsWindowVisible(m_lSizerHnd) = 0 Then
            Mask_Show m_lSizerHnd, True
        End If
    End If

    Select Case uMsg
    '/* client paint
    Case WM_PAINT
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        If Not bSelect Then
            If m_bHasHorizontal Then
                Store_Horizontal lHwnd
                HBar_Paint -1
            End If
            If m_bHasVertical Then
                Store_Vertical lHwnd
                VBar_Paint -1
            End If
        End If
        bHandled = True
        
    '/* horizontal scroll
    Case WM_HSCROLL
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Store_Horizontal lHwnd
        HBar_Paint wParam
        bSelect = True
        bHandled = True
    
    '/* vertical Scroll
    Case WM_VSCROLL
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Store_Vertical lHwnd
        VBar_Paint wParam
        bSelect = True
        bHandled = True

    Case WM_WINDOWPOSCHANGED, WM_STYLECHANGED
        Build_PostTest
        Reset_SCPos
    
    Case WM_ERASEBKGND
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        If m_bHasHorizontal Then
            Store_Horizontal lHwnd
            HBar_Paint -1
        End If
        If m_bHasVertical Then
            Store_Vertical lHwnd
            VBar_Paint -1
        End If
        bHandled = True
    End Select
    
End Sub

Private Sub Control_Redraw(ByVal lHwnd As Long)

    InvalidateRectAsNull lHwnd, 0&, 0&
    UpdateWindow lHwnd

End Sub

Private Sub Mask_Show(ByVal lHwnd As Long, _
                      ByVal bVisible As Boolean)

    If bVisible Then
        ShowWindow lHwnd, SW_NORMAL
    Else
        ShowWindow lHwnd, SW_HIDE
    End If
    
End Sub

Private Sub Reset_SCPos()
'/* sneaky ms.. hides scrollbars to left/bottom of lv control..

Dim lEdge       As Long
Dim uPt         As POINTAPI
Dim tClRt       As RECT
Dim tOnRt       As RECT

    GetWindowRect m_lCtrlHnd, tOnRt
    GetClientRect m_lCtrlHnd, tClRt
    CopyMemory uPt, tOnRt, LenB(uPt)
    ScreenToClient m_hWndParent, uPt
    With tOnRt
        OffsetRect tOnRt, -.Left, -.Top
    End With

    lEdge = Edge_Depth(m_lCtrlHnd)
    With tOnRt
        SetWindowPos m_lVScrollHnd, 0, uPt.x + (.Right - (m_lVXScroll + (1))), uPt.y + lEdge, m_lVXScroll, .Bottom - .Top, SWP_SHOWWINDOW
        SetWindowPos m_lHScrollHnd, 0, uPt.x + lEdge, (uPt.y + .Bottom) - m_lHYScroll, .Right - m_lVXScroll, m_lVYScroll, SWP_SHOWWINDOW
        SetWindowPos m_lSizerHnd, 0, uPt.x + (.Right - (m_lVXScroll + lEdge)), (uPt.y + .Bottom) - (m_lHYScroll + lEdge), m_lHXScroll, m_lHYScroll, SWP_SHOWWINDOW
    End With

End Sub

Private Function Store_Horizontal(ByVal lHwnd As Long) As Boolean

Dim lState      As Long
Dim tRect       As RECT
Dim tRect2      As RECT
Dim tPoint      As POINTAPI

    On Error GoTo Handler

    With m_tHScrInfo
        .cbSize = Len(m_tHScrInfo)
        .fMask = SIF_ALL
    End With

    '/* copy vertical scrollbar structure
    lState = GetScrollInfo(lHwnd, SB_HORZ, m_tHScrInfo)
    '/* no scrollbar
    If lState = 0 Then Exit Function
    '/* get bar params
    m_tHScrBar.cbSize = Len(m_tHScrBar)
    GetScrollBarInfo lHwnd, OBJID_HSCROLL, m_tHScrBar
    
    '/* copy size to working rect
    CopyRect m_tHorzRect, m_tHScrBar.rcScrollBar
    Store_Horizontal = True

Handler:
    On Error GoTo 0

End Function

Private Function Store_Vertical(ByVal lHwnd As Long) As Boolean

Dim lState      As Long
Dim tRect       As RECT
Dim tRect2      As RECT
Dim tPoint      As POINTAPI

On Error GoTo Handler

    With m_tVScrInfo
        .cbSize = Len(m_tVScrInfo)
        .fMask = SIF_ALL
    End With

    '/* copy vertical scrollbar structure
    lState = GetScrollInfo(lHwnd, SB_VERT, m_tVScrInfo)
    '/* no scrollbar
    If lState = 0 Then Exit Function
    '/* get bar params
    m_tVScrBar.cbSize = Len(m_tVScrBar)
    GetScrollBarInfo lHwnd, OBJID_VSCROLL, m_tVScrBar
    '/* copy size to working rect
    CopyRect m_tVertRect, m_tVScrBar.rcScrollBar
    '/* apply client offsets
    Store_Vertical = True

Handler:
    On Error GoTo 0

End Function

Private Sub Cleanup()

    If Not m_lCtrlHnd = 0 Then
        Control_Detatch m_lCtrlHnd
        Destroy_Scrollbar SB_BOTH
        If Not m_cVTrackDc Is Nothing Then Set m_cVTrackDc = Nothing
        If Not m_cVThumbDc Is Nothing Then Set m_cVThumbDc = Nothing
        If Not m_cBtDwnDc Is Nothing Then Set m_cBtDwnDc = Nothing
        If Not m_cBtUpDc Is Nothing Then Set m_cBtUpDc = Nothing
        If Not m_cHTrackDc Is Nothing Then Set m_cHTrackDc = Nothing
        If Not m_cHThumbDc Is Nothing Then Set m_cHThumbDc = Nothing
        If Not m_cHBtLftDc Is Nothing Then Set m_cHBtLftDc = Nothing
        If Not m_cHBtRgtDc Is Nothing Then Set m_cHBtRgtDc = Nothing
        If Not m_cSzrDc Is Nothing Then Set m_cSzrDc = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_pVTrack Is Nothing Then Set m_pVTrack = Nothing
        If Not m_pVThumb Is Nothing Then Set m_pVThumb = Nothing
        If Not m_pVBtDwn Is Nothing Then Set m_pVBtDwn = Nothing
        If Not m_pVBtUp Is Nothing Then Set m_pVBtUp = Nothing
        If Not m_pHTrack Is Nothing Then Set m_pHTrack = Nothing
        If Not m_pHThumb Is Nothing Then Set m_pHThumb = Nothing
        If Not m_pHBtLft Is Nothing Then Set m_pHBtLft = Nothing
        If Not m_pHBtRgt Is Nothing Then Set m_pHBtRgt = Nothing
        If Not m_pSizer Is Nothing Then Set m_pSizer = Nothing
        If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
        If Not p_OScrollBarObj Is Nothing Then Set p_OScrollBarObj = Nothing
        m_lCtrlHnd = 0
    End If

End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

