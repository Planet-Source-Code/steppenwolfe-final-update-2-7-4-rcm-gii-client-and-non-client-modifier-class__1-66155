VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements WinSubHookGX.GXISubclass

'/* border styles
Private Const GWL_WNDPROC                       As Long = (-4)
Private Const GW_CHILD                          As Long = &H5
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_DRAWFRAME                     As Long = SWP_FRAMECHANGED
Private Const CB_GETDROPPEDSTATE                As Long = &H157
Private Const CB_GETDROPPEDCONTROLRECT          As Long = &H152
Private Const CB_GETCOMBOBOXINFO                As Long = 164

#If (WINVER >= &H400) Then
Private Const CB_MSGMAX                         As Long = &H162
#Else
Private Const CB_MSGMAX                         As Long = &H15B
#End If

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Type COMBOBOXINFO
   cbSize                                       As Long
   rcItem                                       As RECT
   rcButton                                     As RECT
   stateButton                                  As Long
   hwndCombo                                    As Long
   hwndEdit                                     As Long
   hwndList                                     As Long
End Type


Private Type POINTAPI
    x                                           As Long
    y                                           As Long
End Type

Private Type BITMAP
    bmType                                      As Long
    bmWidth                                     As Long
    bmHeight                                    As Long
    bmWidthBytes                                As Long
    bmPlanes                                    As Integer
    bmBitsPixel                                 As Integer
    bmBits                                      As Long
End Type

Private Type PAINTSTRUCT
    hdc                                     As Long
    fErase                                  As Boolean
    rcPaint                                 As RECT
    fRestore                                As Boolean
    fIncUpdate                              As Boolean
    rgbReserved(32)                         As Byte
End Type


Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                      As Long
    dwFlags                                     As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                   As Long
    dwHoverTime                                 As Long
End Type

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Enum SC_MESSAGE
    WM_NULL = &H0
    WM_SIZE = &H5
    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_ENABLE = &HA
    WM_PAINT = &HF
    WM_ERASEBKGND = &H14
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_MOUSELEAVE = &H2A3
    WM_WINDOWPOSCHANGED = &H47
    WM_STYLECHANGED = &H7D
End Enum

Public Enum CB_STYLE
    FrameFlat = &H1
    'FrameRaised = &H2
    'FrameInset = &H3
End Enum


Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                                                          ByVal hWnd2 As Long, _
                                                                          ByVal lpsz1 As String, _
                                                                          lpsz2 As Any) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal hPalette As Long, _
                                                               pccolorref As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, _
                                                       CBInfo As COMBOBOXINFO) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long


Private m_bTrackUser32                      As Boolean
Private m_bEnabled                          As Boolean
Private m_bPreRelease                       As Boolean
Private m_bExtendedStyle                    As Boolean
Private m_bPainting                         As Boolean
Private m_lParentHwnd                       As Long
Private m_lComboWidth                       As Long
Private m_lComboHeight                      As Long
Private m_lTransparentColor                 As Long
Private m_lFrameHighLite                    As Long
Private m_lFrameColor                       As Long
Private m_lAccentColor                      As Long
Private m_lLtShadowColor                    As Long
Private m_lDrkShadowColor                   As Long
Private m_lHdc()                            As Long
Private m_lBmp()                            As Long
Private m_lBmpOld()                         As Long
Private m_eFrameStyle                       As CB_STYLE
Private c_CtrlBg                            As Collection
Private c_InstTrack                         As Collection
Private c_CtrlType                          As Collection
Private m_pComboImg                         As StdPicture
Private m_cRender                           As clsRender
Private m_cComboDc                          As clsStoreDc
Private m_GSubclass                         As GXMSubclass
Private m_lParentObj                        As Long


Private Sub Class_Initialize()
    InitCommonControls
    m_bTrackUser32 = Function_Exported("TrackMouseEvent", "User32")
    m_lTransparentColor = -1
    m_lFrameColor = -1
    m_lFrameHighLite = -1
    m_eFrameStyle = -1
End Sub

Public Property Get p_OParentObj() As Object

Dim oTemp   As Object

    If m_lParentObj = 0 Then
        Set p_OParentObj = Nothing
    Else
        CopyMemory oTemp, m_lParentObj, 4
        Set p_OParentObj = oTemp
        CopyMemory oTemp, 0&, 4
    End If

End Property

Public Property Set p_OParentObj(ByVal PropVal As Object)
    m_lParentObj = ObjPtr(PropVal)
End Property

Public Property Get p_FrameColor() As Long
'/* frame color
    p_FrameColor = m_lFrameColor
End Property

Public Property Let p_FrameColor(ByVal PropVal As Long)
    m_lFrameColor = TranslateColor(PropVal)
End Property

Public Property Get p_FrameHighLite() As Long
'/* frame rollover color
    p_FrameHighLite = m_lFrameHighLite
End Property

Public Property Let p_FrameHighLite(ByVal PropVal As Long)
    m_lFrameHighLite = TranslateColor(PropVal)
End Property

Public Property Get p_FrameStyle() As CB_STYLE
'/* frame style
    p_FrameStyle = m_eFrameStyle
End Property

Public Property Let p_FrameStyle(PropVal As CB_STYLE)
    m_eFrameStyle = PropVal
End Property

Public Property Get p_IComboImg() As StdPicture
'/* command button image
    Set p_IComboImg = m_pComboImg
End Property

Public Property Set p_IComboImg(PropVal As StdPicture)
    Set m_pComboImg = PropVal
End Property

Public Property Get p_TransparentColor() As Long
'/* transparency color
    p_TransparentColor = m_lTransparentColor
End Property

Public Property Let p_TransparentColor(ByVal PropVal As Long)
    m_lTransparentColor = PropVal
End Property

Public Sub Attatch_ComboBox()

Dim tBmp        As BITMAP

On Error GoTo Handler

    '/* test for defaults
    If (p_IComboImg Is Nothing) Or _
        (p_OParentObj Is Nothing) Then Exit Sub

    '/* store parent hwnd
    m_lParentHwnd = p_OParentObj.hwnd

    '/* default sizes
    GetObject p_IComboImg.handle, Len(tBmp), tBmp
    With tBmp
        m_lComboWidth = (.bmWidth / 4)
        m_lComboHeight = .bmHeight
    End With
    
    '/* start subclasser
    Set m_GSubclass = New GXMSubclass
    
    '/* transparency color default
    If m_lTransparentColor = -1 Then
        m_lTransparentColor = &HFF00FF
    End If
    '/* control information
    Set c_InstTrack = New Collection
    Set c_CtrlType = New Collection
    
    '/* init image classes
    Set m_cComboDc = New clsStoreDc
    m_cComboDc.CreateFromPicture p_IComboImg
    Set m_cRender = New clsRender

    '/* store images
    Backbuffer_Images
    '/* add controls
    Control_Allocation
    
    '/* custom border styles
    m_eFrameStyle = FrameFlat
    m_bExtendedStyle = Not m_eFrameStyle = -1
    If m_bExtendedStyle Then
        Set c_CtrlBg = New Collection
        If m_lFrameColor = -1 Then
            m_lFrameColor = &H333333
        End If
        If m_lFrameHighLite = -1 Then
            m_lFrameHighLite = &H777777
        End If
        BuildPalette
    End If
    
Handler:
    On Error GoTo 0

End Sub

Private Sub Control_Allocation()

Dim lVcnt           As Long
Dim lHwndStore      As Long
Dim lCtrHnd         As Long
Dim oCtrl           As Control
Dim sCtrl           As String

On Error Resume Next

    For Each oCtrl In p_OParentObj
        '/* command button
        sCtrl = IsClassName(oCtrl.hwnd)
        If InStrB(1, sCtrl, "ComboBox") > 0 Or _
            InStrB(1, sCtrl, "DriveListBox") > 0 Or _
            InStrB(1, sCtrl, "ImageCombo") > 0 Then
            With oCtrl
                '/* filter invalids or user exempted via tag
                If .hwnd = m_lParentHwnd Then GoTo Skip
                If (.hwnd = 0) Or (.Tag = "NO") Then GoTo Skip
                lCtrHnd = .hwnd
                '/* object type ID
                Select Case True
                Case InStrB(1, sCtrl, "ComboBox") > 0
                    '/* add hwnd to tracker
                    c_CtrlType.Add 1, CStr(lCtrHnd)
                Case InStrB(1, sCtrl, "DriveListBox") > 0
                    c_CtrlType.Add 2, CStr(lCtrHnd)
                Case InStrB(1, sCtrl, "ImageCombo") > 0
                    '/* get combo hnd in control
                    lCtrHnd = FindWindowEx(lCtrHnd, 0&, "ComboBox", ByVal 0&)
                    If lCtrHnd = 0 Then
                        GoTo Skip
                    End If
                    c_CtrlType.Add 4, CStr(lCtrHnd)
                Case Else
                    GoTo Skip
                End Select
                '/* filter duplications
                lHwndStore = c_InstTrack.Item(CStr(lCtrHnd))
                If Not lHwndStore = 0 Then
                    lHwndStore = 0
                    GoTo Skip
                End If
                '/* add to handle tracker
                c_InstTrack.Add lCtrHnd, CStr(lCtrHnd)
                '/* start subclassing
                Control_Attatch lCtrHnd
                '/* test initial enabled state
                m_bEnabled = IsWindowEnabled(lCtrHnd)
                c_CtrlBg.Add .BackColor, CStr(lCtrHnd)
                If Not m_bEnabled Then
                    SendMessage lCtrHnd, WM_ENABLE, 0&, 0&
                End If
Skip:
            End With
        End If
    Next oCtrl

On Error GoTo 0

End Sub

Private Function IsClassName(ByVal lHwnd As Long) As String

Dim lRet        As Long
Dim sClass      As String

    sClass = Space$(255)
    lRet = GetClassName(lHwnd, sClass, 255)
    IsClassName = left(sClass, lRet)

End Function

Private Sub Backbuffer_Images()

Dim lCt         As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long

    ReDim m_lHdc(3)
    ReDim m_lBmp(3)
    ReDim m_lBmpOld(3)

    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)

    For lCt = 0 To 3
        m_lHdc(lCt) = CreateCompatibleDC(lTmpDc)
        m_lBmp(lCt) = CreateCompatibleBitmap(lTmpDc, m_lComboWidth, m_lComboHeight)
        m_lBmpOld(lCt) = SelectObject(m_lHdc(lCt), m_lBmp(lCt))
        m_cRender.Blit m_lHdc(lCt), 0, 0, m_lComboWidth, m_lComboHeight, m_cComboDc.hdc, (m_lComboWidth * lCt), 0, SRCCOPY
    Next lCt

    ReleaseDC lHwnd, lTmpDc

End Sub

Private Sub Backbuffer_Destroy()

Dim lCt     As Long

    For lCt = 0 To 3
        SelectObject m_lHdc(lCt), m_lBmpOld(lCt)
        DeleteObject m_lBmp(lCt)
        DeleteDC m_lHdc(lCt)
        m_lBmpOld(lCt) = 0
        m_lHdc(lCt) = 0
        m_lBmp(lCt) = 0
    Next lCt

End Sub

Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, _
                            ByVal oColorTo As OLE_COLOR, _
                            Optional ByVal alpha As Long = 128) As Long

Dim lCFrom      As Long
Dim lCTo        As Long
Dim lSrcR       As Long
Dim lSrcG       As Long
Dim lSrcB       As Long
Dim lDstR       As Long
Dim lDstG       As Long
Dim lDstB       As Long

    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000

    BlendColor = RGB(((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255))

End Function

Private Sub BuildPalette()

    m_lAccentColor = BlendColor(vbHighlight, m_lFrameColor, 80)
    m_lLtShadowColor = BlendColor(vbButtonShadow, m_lFrameHighLite, 18)
    m_lDrkShadowColor = BlendColor(&H333333, m_lFrameColor, 120)

End Sub

Private Sub Control_Release()

Dim vInst       As Variant

    On Error Resume Next

    m_bPreRelease = True
    For Each vInst In c_InstTrack
        If Not CLng(vInst) = 0 Then
            Control_Detatch CLng(vInst)
        End If
    Next vInst

    On Error GoTo 0

End Sub

Private Function Control_State(ByVal lHwnd As Long, _
                               ByVal lState As Long) As Long

'/* 4 states
'/* 0-normal, 1-focus, 2-down, 3-disabled

Dim bEnabled        As Boolean
Static bDropped     As Boolean

    If lHwnd = 0 Then Exit Function
    If IsWindowVisible(lHwnd) = 0 Then Exit Function
    
    If c_CtrlType.Item(CStr(lHwnd)) = 4 Then
        If IsWindowEnabled(GetParent(lHwnd)) = 0 Then
            Image_Render lHwnd, 3
            Exit Function
        End If
    Else
        If IsWindowEnabled(lHwnd) = 0 Then
            Image_Render lHwnd, 3
            Exit Function
        End If
    End If

    If Not SendMessageLong(lHwnd, CB_GETDROPPEDSTATE, 0&, 0&) = 0 Then
        bDropped = True
        Image_Render lHwnd, 2
        Exit Function
    Else
        If bDropped Then
            bDropped = False
        End If
    End If

    Image_Render lHwnd, lState

End Function

Public Function GetLngColor(Color As Long) As Long

    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If

End Function

Private Sub Button_Size(ByVal lHwnd As Long, tCbInfo As COMBOBOXINFO)

    GetComboBoxInfo lHwnd, tCbInfo
    
End Sub

Private Sub Image_Render(ByVal lHwnd As Long, _
                         ByVal lOffset As Long)

Dim lHdc        As Long
Dim lEditHnd    As Long
Dim tRect       As RECT
Dim tCbInfo     As COMBOBOXINFO

    lHdc = GetDC(lHwnd)
    If lHdc = 0 Then Exit Sub
    GetClientRect lHwnd, tRect
    tCbInfo.cbSize = Len(tCbInfo)
    Button_Size lHwnd, tCbInfo

    With tCbInfo.rcButton
        If Not m_lTransparentColor = -1 Then
            m_cRender.Stretch lHdc, .left, .top, (.Right - .left) - 1, (.Bottom - .top), m_lHdc(lOffset), 0, 0, m_lComboWidth, m_lComboHeight, SRCCOPY
        Else
            m_cRender.Transparent lHdc, .left, .top, (.Right - .left) - 1, (.Bottom - .top), m_lHdc(lOffset), 0, 0, m_lComboWidth, m_lComboHeight, m_lTransparentColor
        End If
        If m_bExtendedStyle Then
            Render_Frame tRect, lHwnd, lHdc, lOffset
        End If
    End With
    
    ReleaseDC lHwnd, lHdc

On Error GoTo 0

End Sub

Private Sub Render_Frame(ByRef tRect As RECT, _
                         ByVal lHwnd As Long, _
                         ByVal lHdc As Long, _
                         ByVal lOffset As Long)

Dim lColor      As Long
Dim lHiLite     As Long
Dim lHBrush     As Long
Dim tTemp       As RECT
    
    Select Case m_bExtendedStyle
    '/* flat style
    Case 1
        Select Case lOffset
        '/* normal
        Case 0
            With tRect
                InflateRect tTemp, .Right, .Bottom
            End With
            lHBrush = CreateSolidBrush(m_lFrameHighLite)
            
        '/* hover
        Case 1
            With tRect
                InflateRect tTemp, .Right, .Bottom
            End With
            lHBrush = CreateSolidBrush(m_lFrameColor)
        '/* down
        Case 2
            With tRect
                InflateRect tTemp, .Right - 1, .Bottom - 1
            End With
            lHBrush = CreateSolidBrush(m_lLtShadowColor)
        '/* disabled
        Case 3
            With tRect
                InflateRect tTemp, .Right - 1, .Bottom - 1
            End With
            lHBrush = CreateSolidBrush(m_lDrkShadowColor)
        End Select
    End Select

    FrameRect lHdc, tRect, lHBrush
    DeleteObject lHBrush
            
End Sub

Private Sub Track_Mouse(ByVal lHwnd As Long)

Dim tme As TRACKMOUSEEVENT_STRUCT

    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_LEAVE
        .hwndTrack = lHwnd
    End With

    If m_bTrackUser32 Then
        TrackMouseEvent tme
    Else
        TrackMouseEventComCtl tme
    End If

End Sub

Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = &H0
    End If

End Function

Private Function Function_Exported(ByVal sFunction As String, _
                                   ByVal sModule As String) As Boolean

'/* test for library support

Dim lModule     As Long
Dim bLoaded     As Boolean

    lModule = GetModuleHandleA(sModule)
    If lModule = 0 Then
        lModule = LoadLibraryA(sModule)
    End If
    If lModule Then bLoaded = True

    If lModule Then
        If GetProcAddress(lModule, sFunction) Then
            Function_Exported = True
        End If
    End If

    If bLoaded Then
        FreeLibrary lModule
    End If

End Function

Private Sub Control_Attatch(ByVal lHwnd As Long)

On Error GoTo Handler

    With m_GSubclass
        .Subclass lHwnd, Me
        .AddMessage lHwnd, WM_LBUTTONDOWN, MSG_BEFORE
        .AddMessage lHwnd, WM_LBUTTONUP, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .AddMessage lHwnd, WM_PAINT, MSG_BEFORE
        .AddMessage lHwnd, WM_SETFOCUS, MSG_BEFORE
        .AddMessage lHwnd, WM_KILLFOCUS, MSG_BEFORE
        .AddMessage lHwnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
        .AddMessage lHwnd, WM_STYLECHANGED, MSG_BEFORE
        .AddMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
    End With

Handler:
    On Error GoTo 0

End Sub

Private Sub Control_Detatch(ByVal lHwnd As Long)

On Error GoTo Handler

    With m_GSubclass
        .DeleteMessage lHwnd, WM_LBUTTONDOWN, MSG_BEFORE
        .DeleteMessage lHwnd, WM_LBUTTONUP, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_PAINT, MSG_BEFORE
        .DeleteMessage lHwnd, WM_SETFOCUS, MSG_BEFORE
        .DeleteMessage lHwnd, WM_KILLFOCUS, MSG_BEFORE
        .DeleteMessage lHwnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
        .DeleteMessage lHwnd, WM_STYLECHANGED, MSG_BEFORE
        .DeleteMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
        .UnSubclass lHwnd
    End With

Handler:
    On Error GoTo 0

End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As WinSubHookGX.eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)

Dim lHdc        As Long
Dim tPAINT      As PAINTSTRUCT
Dim tCbInfo     As COMBOBOXINFO
Dim tRect       As RECT

    If IsWindowVisible(lHwnd) = 0 Then
        bHandled = True
        Exit Sub
    End If
    
    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            lHdc = BeginPaint(lHwnd, tPAINT)
            Image_Render lHwnd, 1
            EndPaint lHwnd, tPAINT
            m_bPainting = False
            bHandled = True
        Else
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If

    Case WM_MOUSELEAVE
        Control_State lHwnd, 1

    Case WM_MOUSEMOVE
        Control_State lHwnd, 2
        Track_Mouse lHwnd
        
    Case WM_SETFOCUS
        Control_State lHwnd, 2

   Case WM_KILLFOCUS
        Control_State lHwnd, 1
        
    Case WM_WINDOWPOSCHANGED, WM_STYLECHANGED
        Control_State lHwnd, 1
        
    Case WM_ERASEBKGND
        tCbInfo.cbSize = Len(tCbInfo)
        Button_Size lHwnd, tCbInfo
        If Not tCbInfo.hwndEdit = 0 Then
            Dim lHBrush As Long
            lHdc = GetDC(tCbInfo.hwndCombo)
            GetClientRect lHwnd, tRect
            With tRect
                OffsetRect tRect, 1, 1
                .Right = .Right - 2
                .Bottom = .Bottom - 2
                lHBrush = CreateSolidBrush(&HFFFFFF)
                FrameRect lHdc, tRect, lHBrush
                OffsetRect tRect, 1, 1
                .Right = .Right - 1
                .Bottom = .Bottom - 2
            End With
            FrameRect lHdc, tRect, lHBrush
            ReleaseDC tCbInfo.hwndCombo, lHdc
            DeleteObject lHBrush
            bHandled = True
        End If
        lReturn = 1
    End Select

End Sub

Private Sub Cleanup()

On Error GoTo Handler

    If Not m_lParentHwnd = 0 Then
        Control_Release
        Backbuffer_Destroy
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cComboDc Is Nothing Then Set m_cComboDc = Nothing
        If Not p_IComboImg Is Nothing Then Set p_IComboImg = Nothing
        If Not c_InstTrack Is Nothing Then Set c_InstTrack = Nothing
        If Not c_CtrlType Is Nothing Then Set c_CtrlType = Nothing
        If Not c_CtrlBg Is Nothing Then Set c_CtrlBg = Nothing
        Erase m_lHdc
        Erase m_lBmp
        Erase m_lBmpOld
    End If
    m_lParentHwnd = 0

Exit Sub

Handler:
    On Error GoTo 0

End Sub

Private Sub Class_Terminate()
    Cleanup
    If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
    If Not p_OParentObj Is Nothing Then Set p_OParentObj = Nothing
End Sub

