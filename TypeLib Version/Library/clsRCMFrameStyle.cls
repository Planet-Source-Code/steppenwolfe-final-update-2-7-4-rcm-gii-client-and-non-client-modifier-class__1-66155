VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMFrameStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements WinSubHookGX.GXISubclass

'/* border styles
Private Const GWL_STYLE                         As Long = (-16)
Private Const GWL_EXSTYLE                       As Long = (-20)
Private Const WS_BORDER                         As Long = &H800000
Private Const WS_EX_CLIENTEDGE                  As Long = &H200
Private Const WS_EX_STATICEDGE                  As Long = &H20000
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_DRAWFRAME                     As Long = SWP_FRAMECHANGED
Private Const SWP_CHANGE                        As Long = SWP_NOACTIVATE Or _
    SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE


Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum


Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                      As Long
    dwFlags                                     As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                   As Long
    dwHoverTime                                 As Long
End Type


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

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bTrackUser32                          As Boolean
Private m_bEnabled                              As Boolean
Private m_bPreRelease                           As Boolean
Private m_bExtendedStyle                        As Boolean
Private m_bFramePictureBox                      As Boolean
Private m_bFrameListBox                         As Boolean
Private m_bFrameDirListBox                      As Boolean
Private m_bFrameFileListBox                     As Boolean
Private m_bFrameTreeView                        As Boolean
Private m_bFrameListView                        As Boolean
Private m_lParentObj                            As Long
Private m_lParentHwnd                           As Long
Private m_lColorNormal                          As Long
Private m_lColorHighLite                        As Long
Private m_lColorFocused                         As Long
Private m_lLtShadowColor                        As Long
Private m_lDkShadowColor                        As Long
Private m_lColorHover                           As Long
Private c_InstTrack                             As Collection
Private m_GSubclass                             As GXMSubclass


Private Sub Class_Initialize()
    InitCommonControls
    m_lColorNormal = -1
    m_lColorHighLite = -1
    m_bTrackUser32 = Function_Exported("TrackMouseEvent", "User32")
End Sub


Public Property Get p_OParentObj() As Object

Dim oTemp As Object

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

Public Property Get p_ColorFocused() As Long
'/* frame focused color
    p_ColorFocused = m_lColorFocused
End Property

Public Property Let p_ColorFocused(ByVal PropVal As Long)
    m_lColorFocused = TranslateColor(PropVal)
End Property

Public Property Get p_ColorHover() As Long
'/* frame rollover color
    p_ColorHover = m_lColorHover
End Property

Public Property Let p_ColorHover(ByVal PropVal As Long)
    m_lColorHover = TranslateColor(PropVal)
End Property

Public Property Get p_ColorNormal() As Long
'/* normal frame color
    p_ColorNormal = m_lColorNormal
End Property

Public Property Let p_ColorNormal(ByVal PropVal As Long)
    m_lColorNormal = TranslateColor(PropVal)
End Property

Public Property Get p_FrameDirListBox() As Boolean
'/* frame DirListBox
    p_FrameDirListBox = m_bFrameDirListBox
End Property

Public Property Let p_FrameDirListBox(ByVal PropVal As Boolean)
    m_bFrameDirListBox = PropVal
End Property

Public Property Get p_FrameFileListBox() As Boolean
'/* frame FileListBox
    p_FrameFileListBox = m_bFrameFileListBox
End Property

Public Property Let p_FrameFileListBox(ByVal PropVal As Boolean)
    m_bFrameFileListBox = PropVal
End Property

Public Property Get p_FrameListBox() As Boolean
'/* frame ListBox
    p_FrameListBox = m_bFrameListBox
End Property

Public Property Let p_FrameListBox(ByVal PropVal As Boolean)
    m_bFrameListBox = PropVal
End Property

Public Property Get p_FrameListView() As Boolean
'/* frame ListView
    p_FrameListView = m_bFrameListView
End Property

Public Property Let p_FrameListView(ByVal PropVal As Boolean)
    m_bFrameListView = PropVal
End Property

Public Property Get p_FramePictureBox() As Boolean
'/* frame PictureBox
    p_FramePictureBox = m_bFramePictureBox
End Property

Public Property Let p_FramePictureBox(ByVal PropVal As Boolean)
    m_bFramePictureBox = PropVal
End Property

Public Property Get p_FrameTreeView() As Boolean
'/* frame TreeView
    p_FrameTreeView = m_bFrameTreeView
End Property

Public Property Let p_FrameTreeView(ByVal PropVal As Boolean)
    m_bFrameTreeView = PropVal
End Property


Public Sub Attatch_Frame()

On Error GoTo Handler

    '/* test for defaults
    If p_OParentObj Is Nothing Then Exit Sub
    '/* store parent hwnd
    m_lParentHwnd = p_OParentObj.hwnd
    '/* start subclasser
    Set m_GSubclass = New GXMSubclass
    '/* control information
    Set c_InstTrack = New Collection
    '/* store colors
    BuildPalette
    '/* add controls
    Control_Allocation

Handler:
    On Error GoTo 0

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

    m_lLtShadowColor = BlendColor(vbButtonShadow, m_lColorHighLite, 18)
    m_lDkShadowColor = BlendColor(&H333333, m_lColorNormal, 120)

End Sub

Private Sub Control_Allocation()
'/* assign subclass members
'/* capture the control tag, this
'/* will allow us to make unique control
'/* style decisions via tag data..
'/* I removed textbox because w/ watermark
'/* caused double-subclassing crashes

Dim oCtrl       As Control
Dim sCtClass    As String

On Error Resume Next

    For Each oCtrl In p_OParentObj
        With oCtrl
            If .hwnd = m_lParentHwnd Then GoTo Skip
            If (.hwnd = 0) Or (.Tag = "NO") Then GoTo Skip
            sCtClass = IsClassName(.hwnd)
            Select Case True
            '/* treeview
            Case InStr(1, sCtClass, "TreeView") > 0
                If m_bFrameTreeView Then
                    c_InstTrack.Add .hwnd, CStr(.hwnd)
                    Flatten_Container .hwnd
                    Control_Attach .hwnd
                End If
            '/* filelistbox
            Case InStr(1, sCtClass, "FileListBox") > 0
                If m_bFrameFileListBox Then
                    c_InstTrack.Add .hwnd, CStr(.hwnd)
                    Flatten_Container .hwnd
                    Control_Attach .hwnd
                End If
            '/* dirlistbox
            Case InStr(1, sCtClass, "DirListBox") > 0
                If m_bFrameDirListBox Then
                    c_InstTrack.Add .hwnd, CStr(.hwnd)
                    Flatten_Container .hwnd
                    Control_Attach .hwnd
                End If
            '/* listbox
            Case InStr(1, sCtClass, "ListBox") > 0
                If m_bFrameListBox Then
                    c_InstTrack.Add .hwnd, CStr(.hwnd)
                    Flatten_Container .hwnd
                    Control_Attach .hwnd
                End If
            '/* picturebox
            Case InStr(1, sCtClass, "PictureBox") > 0
                If m_bFramePictureBox Then
                    c_InstTrack.Add .hwnd, CStr(.hwnd)
                    Flatten_Container .hwnd
                    Control_Attach .hwnd
                End If
            End Select
Skip:
        End With
    Next oCtrl

On Error GoTo 0

End Sub

Private Sub Control_Release()

Dim vInst       As Variant

On Error Resume Next

    m_bPreRelease = True
    For Each vInst In c_InstTrack
        If Not CLng(vInst) = 0 Then
            Control_Detach CLng(vInst)
        End If
    Next vInst

On Error GoTo 0

End Sub

Private Sub Draw_Frame(ByRef tRect As RECT, _
                       ByVal lHwnd As Long, _
                       ByVal lHdc As Long, _
                       ByVal lOffset As Long)

    Select Case lOffset
    '/* normal
    Case 0
        Render_Frame tRect, lHdc, m_lColorNormal
    '/* hover
    Case 1
        Render_Frame tRect, lHdc, m_lColorHover
    '/* down
    Case 2
        Render_Frame tRect, lHdc, m_lColorFocused
    '/* disabled
    Case 3
        Render_Frame tRect, lHdc, m_lDkShadowColor
    '/* none
    Case 4
        Render_Frame tRect, lHdc, -1
    End Select

End Sub

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
    SetWindowPos lHwnd, 0&, 0&, 0&, 0&, 0&, SWP_CHANGE

End Sub

Private Function Function_Exported(ByVal sFunction As String, _
                                   ByVal sModule As String) As Boolean

'/* test for library support

Dim lModule     As Long
Dim bLoaded     As Boolean

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

Private Function GetLngColor(Color As Long) As Long

    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If

End Function

Private Function IsClassName(ByVal lHwnd As Long) As String

Dim lRet        As Long
Dim sClass      As String

    sClass = Space$(255)
    lRet = GetClassName(lHwnd, sClass, 255)
    IsClassName = left(sClass, lRet)

End Function

Private Sub Control_Attach(ByVal lHwnd As Long)

On Error GoTo Handler

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .Subclass lHwnd, Me
        .AddMessage lHwnd, WM_LBUTTONDOWN, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .AddMessage lHwnd, WM_SETFOCUS, MSG_BEFORE
        .AddMessage lHwnd, WM_KILLFOCUS, MSG_BEFORE
        .AddMessage lHwnd, WM_SIZE, MSG_BEFORE
    End With

Handler:
    On Error GoTo 0

End Sub

Private Sub Control_Detach(ByVal lHwnd As Long)

On Error GoTo Handler

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .DeleteMessage lHwnd, WM_LBUTTONDOWN, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_SETFOCUS, MSG_BEFORE
        .DeleteMessage lHwnd, WM_KILLFOCUS, MSG_BEFORE
        .DeleteMessage lHwnd, WM_SIZE, MSG_BEFORE
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

Static bFocused         As Boolean

    If IsWindowVisible(lHwnd) = 0 Then
        lReturn = 0
        bHandled = True
        Exit Sub
    End If

    Select Case uMsg
    Case WM_LBUTTONDOWN
        Control_State lHwnd, 2
        bFocused = True

    Case WM_MOUSELEAVE
        Control_State lHwnd, 0
        bFocused = False

    Case WM_MOUSEMOVE
        If Not bFocused Then
            Control_State lHwnd, 1
        End If
        Track_Mouse lHwnd

    Case WM_SIZE
        Control_State lHwnd, 4

    Case WM_SETFOCUS
        Control_State lHwnd, 2

    Case WM_KILLFOCUS
        Control_State lHwnd, 0
        bFocused = False
    End Select

End Sub

Private Function Control_State(ByVal lHwnd As Long, _
                               ByVal lState As Long) As Long

Dim lHdc        As Long
Dim tRect       As RECT

    If IsWindowVisible(lHwnd) = 0 Then Exit Function
    lHdc = GetDC(lHwnd)
    If Not lHdc = 0 Then
        GetClientRect lHwnd, tRect
        If Not IsWindowEnabled(lHwnd) = 1 Then
            lState = 3
        End If
        Draw_Frame tRect, lHwnd, lHdc, lState
        ReleaseDC lHwnd, lHdc
    End If
    
End Function

Private Sub Render_Frame(ByRef tRect As RECT, _
                         ByVal lHdc As Long, _
                         ByVal lColor As Long)

Dim lhBrush     As Long

    '/* paint and draw frame
    If Not lColor = -1 Then
        lhBrush = CreateSolidBrush(lColor)
        With tRect
            OffsetRect tRect, -.left, -.top
        End With
        FrameRect lHdc, tRect, lhBrush
        DeleteObject lhBrush
    End If

End Sub

Private Sub Track_Mouse(ByVal lHwnd As Long)

Dim tme         As TRACKMOUSEEVENT_STRUCT

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

Private Function TranslateColor(ByVal Clr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(Clr, hPal, TranslateColor) Then
        TranslateColor = &H0
    End If

End Function

Private Sub Cleanup()

On Error Resume Next

    If Not m_lParentHwnd = 0 Then
        Control_Release
        If Not c_InstTrack Is Nothing Then Set c_InstTrack = Nothing
        If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
        If Not p_OParentObj Is Nothing Then Set p_OParentObj = Nothing
    End If
    m_lParentHwnd = 0

On Error GoTo 0

End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

