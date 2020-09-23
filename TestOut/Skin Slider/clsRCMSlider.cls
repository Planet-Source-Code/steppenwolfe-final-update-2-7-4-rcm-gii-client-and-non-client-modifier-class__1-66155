VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements WinSubHookGX.GXISubclass

Private Const WM_USER                               As Long = &H400
Private Const TTM_ACTIVATE                          As Long = (WM_USER + 1)
Private Const TBM_GETTHUMBRECT                      As Long = (WM_USER + 25)
Private Const TBM_GETCHANNELRECT                    As Long = (WM_USER + 26)

Private Type RECT
    left                                            As Long
    top                                             As Long
    Right                                           As Long
    Bottom                                          As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type BITMAP
    bmType                                          As Long
    bmWidth                                         As Long
    bmHeight                                        As Long
    bmWidthBytes                                    As Long
    bmPlanes                                        As Integer
    bmBitsPixel                                     As Integer
    bmBits                                          As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                          As Long
    dwFlags                                         As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                       As Long
    dwHoverTime                                     As Long
End Type


Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Enum SC_MESSAGE
    WM_NULL = &H0
    WM_PAINT = &HF
    WM_STYLECHANGED = &H7D
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_MOUSELEAVE = &H2A3
    WM_CTLCOLORSTATIC = &H138
End Enum


Private Declare Function SetRect Lib "user32" (lpRect As RECT, _
                                               ByVal x1 As Long, _
                                               ByVal y1 As Long, _
                                               ByVal x2 As Long, _
                                               ByVal y2 As Long) As Long

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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String, _
                                                                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, _
                                                                      ByVal lpString As String) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                                   ByVal lpRect As Long, _
                                                                                   ByVal bErase As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bTrackUser32                  As Boolean
Private m_bUseToolTips                  As Boolean
Private m_bPainting                     As Boolean
Private m_lParentObj                    As Long
Private m_lParentHwnd                   As Long
Private m_lBackColor                    As Long
Private m_lHBgBrush                     As Long
Private m_lVBgBrush                     As Long
Private m_lHorzWidth                    As Long
Private m_lVertWidth                    As Long
Private c_InstTrack                     As Collection
Private c_ParStore                      As Collection
Private m_pSldVTrack                    As StdPicture
Private m_pSldVThumb                    As StdPicture
Private m_pSldHTrack                    As StdPicture
Private m_pSldHThumb                    As StdPicture
Private m_cVTrackDc                     As clsStoreDc
Private m_cVThumbDc                     As clsStoreDc
Private m_cHTrackDc                     As clsStoreDc
Private m_cHThumbDc                     As clsStoreDc
Private m_cRender                       As clsRender
Private m_GSubclass                     As GXMSubclass


Private Sub Class_Initialize()
    InitCommonControls
    m_bTrackUser32 = Function_Exported("TrackMouseEvent", "User32")
    m_lBackColor = -1
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

Public Property Get p_ISldVTrack() As StdPicture
'/* verical track
    Set p_ISldVTrack = m_pSldVTrack
End Property

Public Property Set p_ISldVTrack(PropVal As StdPicture)
    Set m_pSldVTrack = PropVal
End Property

Public Property Get p_ISldVThumb() As StdPicture
'/* vertical thumb
    Set p_ISldVThumb = m_pSldVThumb
End Property

Public Property Set p_ISldVThumb(PropVal As StdPicture)
    Set m_pSldVThumb = PropVal
End Property

Public Property Get p_ISldHTrack() As StdPicture
'/* horizontal tarck
    Set p_ISldHTrack = m_pSldHTrack
End Property

Public Property Set p_ISldHTrack(PropVal As StdPicture)
    Set m_pSldHTrack = PropVal
End Property

Public Property Get p_ISldHThumb() As StdPicture
'/* horizontal thumb
    Set p_ISldHThumb = m_pSldHThumb
End Property

Public Property Set p_ISldHThumb(PropVal As StdPicture)
    Set m_pSldHThumb = PropVal
End Property

Public Property Get p_UseToolTips() As Boolean
'/* toggle tooltips
    p_UseToolTips = m_bUseToolTips
End Property

Public Property Let p_UseToolTips(ByVal PropVal As Boolean)
    m_bUseToolTips = PropVal
End Property

Public Property Get p_BackColor() As Long
'/* toggle tooltips
    p_BackColor = m_lBackColor
End Property

Public Property Let p_BackColor(ByVal PropVal As Long)
    m_lBackColor = TranslateColor(PropVal)
End Property

Public Sub Attatch_Slider()

Dim tBmp        As BITMAP

On Error GoTo Handler

    '/* test for defaults
    If (p_OParentObj Is Nothing) Or _
        (p_ISldVTrack Is Nothing) Or _
        (p_ISldVThumb Is Nothing) Or _
        (p_ISldHTrack Is Nothing) Or _
        (p_ISldHThumb Is Nothing) Then Exit Sub

    '/* store parent hwnd
    m_lParentHwnd = p_OParentObj.hwnd
    Set c_ParStore = New Collection
    
    '/* default sizes
    GetObject p_ISldVThumb.handle, Len(tBmp), tBmp
    m_lVertWidth = (tBmp.bmWidth / 2)
    GetObject p_ISldHThumb.handle, Len(tBmp), tBmp
    m_lHorzWidth = (tBmp.bmWidth / 2)

    '/* start subclasser
    Set m_GSubclass = New GXMSubclass
    '/* control information
    Set c_InstTrack = New Collection

    '/* init image classes
    Set m_cVTrackDc = New clsStoreDc
    m_cVTrackDc.CreateFromPicture p_ISldVTrack
    Set m_cVThumbDc = New clsStoreDc
    m_cVThumbDc.CreateFromPicture p_ISldVThumb
    Set m_cHTrackDc = New clsStoreDc
    m_cHTrackDc.CreateFromPicture p_ISldHTrack
    Set m_cHThumbDc = New clsStoreDc
    m_cHThumbDc.CreateFromPicture p_ISldHThumb

    Set m_cRender = New clsRender

    '/* add controls
    Control_Allocation

Handler:
    On Error GoTo 0

End Sub

Private Sub Control_Allocation()

Dim oCtrl           As Control
Dim lHwndStore      As Long
Dim lCtrHnd         As Long
Dim sCtrl           As String

On Error Resume Next

    For Each oCtrl In p_OParentObj
        sCtrl = IsClassName(oCtrl.hwnd)
        If InStrB(1, sCtrl, "Slider") > 0 Then
            With oCtrl
                If .hwnd = m_lParentHwnd Then GoTo Skip
                '/* filter invalids or user exempted via tag
                If (.hwnd = 0) Or (.Tag = "NO") Then GoTo Skip
                lCtrHnd = .hwnd
                '/* filter duplications
                lHwndStore = c_InstTrack.Item(CStr(lCtrHnd))
                If Not lHwndStore = 0 Then
                    lHwndStore = 0
                    GoTo Skip
                End If
                '/* add to handle tracker
                c_InstTrack.Add lCtrHnd, CStr(lCtrHnd)
                SetBackColor lCtrHnd
                'Slider_ToolTips lCtrHnd
                '/* start subclassing
                Control_Attach lCtrHnd
                SetFocus lCtrHnd
Skip:
            End With
        End If
    Next oCtrl

On Error GoTo 0

End Sub

Private Sub SetBackColor(ByVal lHwnd As Long)
'/* set back color

Dim lColor      As Long

On Error Resume Next

    If lHwnd = 0 Then Exit Sub
    If m_lBackColor = -1 Then
        If Slider_Orientation(lHwnd) = 1 Then
            lColor = GetPixel(m_cVThumbDc.hdc, 0&, 0&)
            m_lVBgBrush = CreateSolidBrush(lColor)
        Else
            lColor = GetPixel(m_cHThumbDc.hdc, 0&, 0&)
            m_lHBgBrush = CreateSolidBrush(lColor)
        End If
    Else
        m_lVBgBrush = CreateSolidBrush(m_lBackColor)
        m_lHBgBrush = CreateSolidBrush(m_lBackColor)
    End If

On Error GoTo 0

End Sub

Private Function TranslateColor(ByVal Clr As OLE_COLOR)

   If (Clr And &H80000000) = &H80000000 Then
      TranslateColor = GetSysColor(Clr And &HFF)
   Else
      TranslateColor = Clr
   End If

End Function

Private Sub DestroyBgBrush()

    If Not m_lVBgBrush = 0 Then
        DeleteObject m_lVBgBrush
        m_lVBgBrush = 0
    End If
    If Not m_lHBgBrush = 0 Then
        DeleteObject m_lHBgBrush
        m_lHBgBrush = 0
    End If

End Sub

Private Function Slider_Orientation(ByVal lHwnd As Long) As Long

Dim tRect       As RECT

    If lHwnd = 0 Then Exit Function
    SendMessage lHwnd, TBM_GETTHUMBRECT, 0&, tRect
    With tRect
        If .Right - .left > .Bottom - .top Then
            Slider_Orientation = 1
        Else
            Slider_Orientation = 0
        End If
    End With

End Function

Private Sub Slider_ToolTips(ByVal lHwnd As Long)

    If lHwnd = 0 Then Exit Sub
    SendMessage lHwnd, TTM_ACTIVATE, ByVal Abs(m_bUseToolTips), ByVal 0&

End Sub

Private Sub Control_Release()

Dim vInst       As Variant

On Error Resume Next

    For Each vInst In c_InstTrack
        If Not CLng(vInst) = 0 Then
            Control_Detach CLng(vInst)
            RemoveProp CLng(vInst), "msg"
        End If
    Next vInst
    
    Parent_Detach

On Error GoTo 0

End Sub

Private Function Function_Exported(ByVal sFunction As String, _
                                   ByVal sModule As String) As Boolean

'/* test for library support

Dim lModule As Long
Dim bLoaded As Boolean

    lModule = GetModuleHandleA(sModule)
    If lModule = 0 Then
        lModule = LoadLibraryA(sModule)
    End If
    If lModule Then
        bLoaded = True
    End If

    If lModule Then
        If GetProcAddress(lModule, sFunction) Then
            Function_Exported = True
        End If
    End If

    If bLoaded Then
        FreeLibrary lModule
    End If

End Function

Private Function IsClassName(ByVal lHwnd As Long) As String

Dim lRet        As Long
Dim sClass      As String

    sClass = Space$(255)
    lRet = GetClassName(lHwnd, sClass, 255)
    IsClassName = left(sClass, lRet)

End Function

Private Sub Control_Redraw(ByVal lHwnd As Long)

    InvalidateRectAsNull lHwnd, 0&, 0&
    UpdateWindow lHwnd

End Sub

Private Sub Image_Render(ByVal lHwnd As Long, ByVal lOffset As Long)

Dim lHdc        As Long
Dim tRTrck      As RECT
Dim tRThmb      As RECT
Dim tRect       As RECT

    lHdc = GetDC(lHwnd)
    If lHdc = 0 Then Exit Sub
    SendMessage lHwnd, TBM_GETCHANNELRECT, 0&, tRTrck
    SendMessage lHwnd, TBM_GETTHUMBRECT, 0&, tRThmb

    '/* vertical
    If Slider_Orientation(lHwnd) = 1 Then
        '/ invert rect
        With tRTrck
            SetRect tRect, .top, .left, .Bottom, .Right
        End With
        With tRect
            m_cRender.Stretch lHdc, .left, .top, (.Right - .left), 3, m_cVTrackDc.hdc, 0, 0, m_cVTrackDc.Width, 3, SRCCOPY
            m_cRender.Stretch lHdc, .left, .top + 3, (.Right - .left), (.Bottom - .top) - 6, m_cVTrackDc.hdc, 0, 3, m_cVTrackDc.Width, (m_cVTrackDc.Height - 6), SRCCOPY
            m_cRender.Stretch lHdc, .left, (.Bottom - 3), (.Right - .left), 3, m_cVTrackDc.hdc, 0, (m_cVTrackDc.Height - 3), m_cVTrackDc.Width, 3, SRCCOPY
        End With
        With tRThmb
            m_cRender.Stretch lHdc, .left, .top, (.Right - .left), (.Bottom - .top), m_cVThumbDc.hdc, m_lVertWidth * lOffset, 0, m_lVertWidth, m_cVThumbDc.Height, SRCCOPY
        End With
    '/* horizontal
    Else
        With tRTrck
            m_cRender.Stretch lHdc, .left, .top, 3, (.Bottom - .top), m_cHTrackDc.hdc, 0, 0, 3, m_cHTrackDc.Height, SRCCOPY
            m_cRender.Stretch lHdc, (.left + 3), .top, (.Right - .left) - 6, (.Bottom - .top), m_cHTrackDc.hdc, 3, 0, m_cHTrackDc.Width - 6, m_cHTrackDc.Height, SRCCOPY
            m_cRender.Stretch lHdc, (.Right - 3), .top, 3, (.Bottom - .top), m_cHTrackDc.hdc, (m_cHTrackDc.Width - 3), 0, 3, m_cHTrackDc.Height, SRCCOPY
        End With
        With tRThmb
            m_cRender.Stretch lHdc, .left, .top, (.Right - .left), (.Bottom - .top), m_cHThumbDc.hdc, m_lHorzWidth * lOffset, 0, m_lHorzWidth, m_cHThumbDc.Height, SRCCOPY
        End With
    End If

    ReleaseDC lHwnd, lHdc

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

Private Sub Control_Attach(ByVal lHwnd As Long)

Dim lParHnd         As Long
Dim lParChk         As Long

    If lHwnd = 0 Then Exit Sub
    lParHnd = GetParent(lHwnd)
    If lParHnd = 0 Then Exit Sub
    On Error Resume Next
    lParChk = c_ParStore.Item(CStr(lParHnd))
    On Error GoTo 0

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .Subclass lHwnd, Me
        .AddMessage lHwnd, WM_PAINT, MSG_BEFORE
        .AddMessage lHwnd, WM_LBUTTONDOWN, MSG_BEFORE
        .AddMessage lHwnd, WM_LBUTTONUP, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .AddMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
        If lParChk = 0 Then
            .Subclass lParHnd, Me
            .AddMessage lParHnd, WM_CTLCOLORSTATIC, MSG_BEFORE
            c_ParStore.Add lParHnd, CStr(lParHnd)
        End If
    End With

End Sub

Private Sub Control_Detach(ByVal lHwnd As Long)

Dim lParent As Long

On Error GoTo Handler

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .DeleteMessage lHwnd, WM_PAINT, MSG_BEFORE
        .DeleteMessage lHwnd, WM_LBUTTONDOWN, MSG_BEFORE
        .DeleteMessage lHwnd, WM_LBUTTONUP, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
        .UnSubclass lHwnd
    End With

Handler:
    On Error GoTo 0

End Sub

Private Sub Parent_Detach()

On Error Resume Next

Dim vI As Variant

    With m_GSubclass
        For Each vI In c_ParStore
            If Not CLng(vI) = 0 Then
                .DeleteMessage CLng(vI), WM_CTLCOLORSTATIC, MSG_BEFORE
                .UnSubclass CLng(vI)
            End If
        Next vI
    End With

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

Static bSelect  As Boolean

    If IsWindowVisible(lHwnd) = 0 Then
        lReturn = 1
     '  bHandled = True
        Exit Sub
    End If
    
    '/* change slider backcolor
    If uMsg = WM_CTLCOLORSTATIC Then
        On Error Resume Next
        If lParam = 0 Then Exit Sub
        If Not c_InstTrack.Item(CStr(lParam)) > 0 Then Exit Sub
        On Error GoTo 0
        If Slider_Orientation(lParam) = 0 Then
            If Not m_lHBgBrush = 0 Then
                lReturn = m_lHBgBrush
                bHandled = True
            End If
        Else
            If Not m_lVBgBrush = 0 Then
                lReturn = m_lVBgBrush
                bHandled = True
            End If
        End If
        Exit Sub
    End If

    Select Case uMsg
    Case WM_PAINT
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Image_Render lHwnd, 0
        bHandled = True

    Case WM_LBUTTONDOWN
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Image_Render lHwnd, 0
        bHandled = True
        bSelect = True
        
    Case WM_LBUTTONUP
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Image_Render lHwnd, 0
        bHandled = True
        bSelect = False
        
    Case WM_MOUSELEAVE
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Image_Render lHwnd, 0
        bHandled = True
        
    Case WM_MOUSEMOVE
        If Not bSelect Then
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
            Image_Render lHwnd, 1
            bHandled = True
        End If
        Track_Mouse lHwnd
        
    Case WM_ERASEBKGND
        Control_Redraw lHwnd
        lReturn = 1
        bHandled = True
    End Select

End Sub

Private Sub Cleanup()

On Error Resume Next

    If Not m_lParentHwnd = 0 Then
        Control_Release
        DestroyBgBrush
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cVTrackDc Is Nothing Then Set m_cVTrackDc = Nothing
        If Not p_ISldVTrack Is Nothing Then Set p_ISldVTrack = Nothing
        If Not m_cVThumbDc Is Nothing Then Set m_cVThumbDc = Nothing
        If Not p_ISldVThumb Is Nothing Then Set p_ISldVThumb = Nothing
        If Not m_cHTrackDc Is Nothing Then Set m_cHTrackDc = Nothing
        If Not p_ISldHTrack Is Nothing Then Set p_ISldHTrack = Nothing
        If Not m_cHThumbDc Is Nothing Then Set m_cHThumbDc = Nothing
        If Not p_ISldHThumb Is Nothing Then Set p_ISldHThumb = Nothing
        If Not c_InstTrack Is Nothing Then Set c_InstTrack = Nothing
        If Not c_ParStore Is Nothing Then Set c_ParStore = Nothing
    End If
    m_lParentHwnd = 0

On Error GoTo 0

End Sub

Private Sub Class_Terminate()
    Cleanup
    If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
    If Not p_OParentObj Is Nothing Then Set p_OParentObj = Nothing
End Sub

