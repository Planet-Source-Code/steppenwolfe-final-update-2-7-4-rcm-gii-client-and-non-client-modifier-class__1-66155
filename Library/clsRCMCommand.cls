VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements GXISubclass

Private Const BN_HILITE                                 As Long = &H2&
Private Const BN_UNHILITE                               As Long = &H3&
Private Const BN_DOUBLECLICKED                          As Long = &H5&
Private Const DT_CENTER                                 As Long = &H1
Private Const DT_VCENTER                                As Long = &H4
Private Const DT_SINGLELINE                             As Long = &H20
Private Const NONANTIALIASED_QUALITY                    As Long = &H3
Private Const ANTIALIASED_QUALITY                       As Long = &H4
Private Const DSS_MONO                                  As Long = &H80
Private Const DSS_NORMAL                                As Long = &H0
Private Const DST_ICON                                  As Long = &H3


Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Enum ICON_POSITION
    ICN_LEFT = &H1
    ICN_RIGHT = &H2
    ICN_TOP = &H3
End Enum


Private Type POINTAPI
    x                                                   As Long
    y                                                   As Long
End Type

Private Type RECT
    left                                                As Long
    top                                                 As Long
    Right                                               As Long
    Bottom                                              As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType                                             As Long
    CtlID                                               As Long
    itemID                                              As Long
    itemAction                                          As Long
    itemState                                           As Long
    hwndItem                                            As Long
    hdc                                                 As Long
    rcItem                                              As RECT
    itemData                                            As Long
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

Private Type LOGBRUSH
    lbStyle                                             As Long
    lbColor                                             As Long
    lbHatch                                             As Long
End Type

Private Type LOGFONT
    lfHeight                                            As Long
    lfWidth                                             As Long
    lfEscapement                                        As Long
    lfOrientation                                       As Long
    lfWeight                                            As Long
    lfItalic                                            As Byte
    lfUnderline                                         As Byte
    lfStrikeOut                                         As Byte
    lfCharSet                                           As Byte
    lfOutPrecision                                      As Byte
    lfClipPrecision                                     As Byte
    lfQuality                                           As Byte
    lfPitchAndFamily                                    As Byte
    lfFaceName                                          As String * 32
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                              As Long
    dwFlags                                             As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                           As Long
    dwHoverTime                                         As Long
End Type


Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, _
                                                    ByVal nBkMode As Long) As Long

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                      ByVal lpStr As String, _
                                                                      ByVal nCount As Long, _
                                                                      ByRef lpRect As RECT, _
                                                                      ByVal wFormat As Long) As Long

Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, _
                                                       ByVal crColor As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                    ByVal hBrush As Long, _
                                                                    ByVal lpDrawStateProc As Long, _
                                                                    ByVal lParam As Long, _
                                                                    ByVal wParam As Long, _
                                                                    ByVal x As Long, _
                                                                    ByVal y As Long, _
                                                                    ByVal cx As Long, _
                                                                    ByVal cy As Long, _
                                                                    ByVal fuFlags As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lDest As Any, _
                                                                     lSource As Any, _
                                                                     ByVal cbCopy As Long)

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bTrackUser32                  As Boolean
Private m_bEnabled                      As Boolean
Private m_bDrawFocused                  As Boolean
Private m_bTextAntiAliased              As Boolean
Private m_lParentHwnd                   As Long
Private m_lCmdWidth                     As Long
Private m_lCmdHeight                    As Long
Private m_lRenderOffsetX                As Long
Private m_lRenderOffsetY                As Long
Private m_lForeColor                    As Long
Private m_lColorHiLite                  As Long
Private m_lTransparentColor             As Long
Private m_lHdc()                        As Long
Private m_lBmp()                        As Long
Private m_lBmpOld()                     As Long
Private m_sCmdCaption                   As String
Private m_eIconPosition                 As ICON_POSITION
Private c_InstTrack                     As Collection
Private c_CaptionText                   As Collection
Private c_CaptionFont                   As Collection
Private c_CaptionSize                   As Collection
Private c_CaptionOptions                As Collection
Private c_ButtonIcon                    As Collection
Private c_ParStore                      As Collection
Private m_pCmdImg                       As StdPicture
Private m_cRender                       As clsRender
Private m_cCmdDc                        As clsStoreDc
Private m_GSubclass                     As GXMSubclass
Private m_lParentObj                    As Long


Private Sub Class_Initialize()
    InitCommonControls
    m_bTrackUser32 = Function_Exported("TrackMouseEvent", "User32")
    m_lTransparentColor = -1
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

Public Property Get p_ICmdImg() As StdPicture
'/* command button image
    Set p_ICmdImg = m_pCmdImg
End Property

Public Property Set p_ICmdImg(PropVal As StdPicture)
    Set m_pCmdImg = PropVal
End Property

Public Property Get p_CmdCaption() As String
'/* control caption
    p_CmdCaption = m_sCmdCaption
End Property

Public Property Let p_CmdCaption(ByVal PropVal As String)
    m_sCmdCaption = PropVal
End Property

Public Property Get p_ColorHiLite() As Long
'/* button over text color
    p_ColorHiLite = m_lColorHiLite
End Property

Public Property Let p_ColorHiLite(ByVal PropVal As Long)
    m_lColorHiLite = TranslateColor(PropVal)
End Property

Public Property Get p_DrawFocused() As Boolean
'/* control hwnd
    p_DrawFocused = m_bDrawFocused
End Property

Public Property Let p_DrawFocused(ByVal PropVal As Boolean)
    m_bDrawFocused = PropVal
End Property

Public Property Get p_ForeColor() As Long
'/* button text color
    p_ForeColor = m_lForeColor
End Property

Public Property Let p_ForeColor(ByVal PropVal As Long)
    m_lForeColor = TranslateColor(PropVal)
End Property

Public Property Get p_IconPosition() As ICON_POSITION
'/* icon offset
    p_IconPosition = m_eIconPosition
End Property

Public Property Let p_IconPosition(PropVal As ICON_POSITION)
    m_eIconPosition = PropVal
End Property

Public Property Get p_RenderOffsetX() As Long
'/* stretchblt offset X
    p_RenderOffsetX = m_lRenderOffsetX
End Property

Public Property Let p_RenderOffsetX(ByVal PropVal As Long)
    m_lRenderOffsetX = PropVal
End Property

Public Property Get p_RenderOffsetY() As Long
''/* stretchblt offset Y
    p_RenderOffsetY = m_lRenderOffsetY
End Property

Public Property Let p_RenderOffsetY(ByVal PropVal As Long)
    m_lRenderOffsetY = PropVal
End Property

Public Property Get p_TextAntiAliased() As Boolean
'/* control hwnd
    p_TextAntiAliased = m_bTextAntiAliased
End Property

Public Property Let p_TextAntiAliased(ByVal PropVal As Boolean)
    m_bTextAntiAliased = PropVal
End Property

Public Property Get p_TransparentColor() As Long
'/* transparency color
    p_TransparentColor = m_lTransparentColor
End Property

Public Property Let p_TransparentColor(ByVal PropVal As Long)
    m_lTransparentColor = PropVal
End Property


Public Sub Attatch_Command()
'up/pressed/disabled/over
Dim tBmp        As BITMAP

On Error GoTo Handler

    '/* test for defaults
    If (p_ICmdImg Is Nothing) Or _
        (p_OParentObj Is Nothing) Then Exit Sub

    '/* store parent hwnd
    m_lParentHwnd = p_OParentObj.hwnd

    '/* default sizes
    GetObject p_ICmdImg.handle, Len(tBmp), tBmp
    With tBmp
        m_lCmdWidth = (.bmWidth / 5)
        m_lCmdHeight = .bmHeight
    End With

    '/* default image offsets
    If m_lRenderOffsetX = 0 Then
        m_lRenderOffsetX = 3
    End If

    '/* color defaults
    If m_lForeColor = 0 Then
        m_lForeColor = &H222222
    End If
    If m_lColorHiLite = 0 Then
        m_lColorHiLite = m_lForeColor
    End If

    Set c_ParStore = New Collection

    '/* init image classes
    Set m_GSubclass = New GXMSubclass
    Set m_cCmdDc = New clsStoreDc
    m_cCmdDc.UseAlpha = True
    m_cCmdDc.CreateFromPicture p_ICmdImg
    
    '/* drawing class
    Set m_cRender = New clsRender
    
    '/* icon pos
    If m_eIconPosition = 0 Then
        m_eIconPosition = ICN_RIGHT
    End If
    
    '/* create image buffer
    Backbuffer_Images
    
    '/* add controls
    Control_Allocation

Handler:
    On Error GoTo 0

End Sub

Private Sub Control_Allocation()
'/* assign object params to variant array

Dim oCtrl           As Control
Dim lVcnt           As Long
Dim sOptions        As String
Dim lHwndStore      As Long
Dim sCtrl           As String

On Error Resume Next

    Set c_InstTrack = New Collection
    Set c_CaptionText = New Collection
    Set c_CaptionFont = New Collection
    Set c_CaptionSize = New Collection
    Set c_CaptionOptions = New Collection
    Set c_ButtonIcon = New Collection

    For Each oCtrl In p_OParentObj
        sCtrl = IsClassName(oCtrl.hwnd)
        If InStrB(1, sCtrl, "CommandButton") > 0 Then
            sOptions = vbNullString
            With oCtrl
                '/* filter invalids or user exempted via tag
                If (.hwnd = 0) Or (.Tag = "NO") Then GoTo Skip
                '/* filter duplications
                lHwndStore = c_InstTrack.Item(CStr(.hwnd))
                If Not lHwndStore = 0 Then
                    lHwndStore = 0
                    GoTo Skip
                End If
                c_InstTrack.Add .hwnd, CStr(.hwnd)
                c_CaptionText.Add .Caption, CStr(.hwnd)
                .Style = 1
                '/* font
                If Len(.Font) < 2 Then
                    c_CaptionFont.Add "MS Sans Serif" & vbNullChar, CStr(.hwnd)
                Else
                    c_CaptionFont.Add .Font, CStr(.hwnd)
                End If
                '/* font size
                If Len(.Font.Size) = 0 Then
                    c_CaptionSize.Add 8, CStr(.hwnd)
                Else
                    c_CaptionSize.Add .Font.Size, CStr(.hwnd)
                End If
                '/* font attributes
                If .Font.Bold Then sOptions = "B"
                If .Font.Underline Then sOptions = sOptions & "U"
                If .Font.Strikethrough Then sOptions = sOptions & "S"
                If .Font.Italic Then sOptions = sOptions & "I"
                c_CaptionOptions.Add sOptions, CStr(.hwnd)
                If Not .Picture.handle = 0 Then
                    c_ButtonIcon.Add .Picture.handle, CStr(.hwnd)
                End If
                '/* start subclassing
                Control_Attatch .hwnd
                '/* test initial enabled state
                m_bEnabled = IsWindowEnabled(.hwnd)
                '/* paint disabled
                If Not m_bEnabled Then
                    SendMessage .hwnd, WM_ENABLE, 0&, 0&
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

Private Sub Render_Image(ByVal lHwnd As Long, _
                         ByVal lOffset As Long, _
                         ByRef tRect As RECT)

Dim lHdc            As Long
Dim lOldBmp         As Long
Dim lNewBmp         As Long
Dim lDrawDc         As Long
Dim lActive         As Long
Dim lIconHnd        As Long

On Error Resume Next

    lHdc = GetDC(lHwnd)
    If lHdc = 0 Then Exit Sub
    lDrawDc = CreateCompatibleDC(lHdc)
    lNewBmp = CreateCompatibleBitmap(lHdc, tRect.Right, tRect.Bottom)
    lOldBmp = SelectObject(lDrawDc, lNewBmp)

    With tRect
        '/* left side
        m_cRender.Stretch lDrawDc, .left, .top, m_lRenderOffsetX, .Bottom, m_lHdc(lOffset), 0, 0, m_lRenderOffsetX, m_lCmdHeight, SRCCOPY, True
        '/* middle
        m_cRender.Stretch lDrawDc, .left + m_lRenderOffsetX, 0, (.Right - (m_lRenderOffsetX * 2)), .Bottom, m_lHdc(lOffset), m_lRenderOffsetX, 0, (m_lCmdWidth - (m_lRenderOffsetX * 2)), m_lCmdHeight, SRCCOPY, True
        '/* right side
        m_cRender.Stretch lDrawDc, .Right - m_lRenderOffsetX, 0, m_lRenderOffsetX, .Bottom, m_lHdc(lOffset), m_lCmdWidth - m_lRenderOffsetX, 0, m_lRenderOffsetX, m_lCmdHeight, SRCCOPY, True
    End With

    lIconHnd = c_ButtonIcon.Item(CStr(lHwnd))
    If Not lIconHnd = 0 Then
        If lOffset = 1 Then
            Render_Icon lHwnd, lIconHnd, lDrawDc, True
        Else
            Render_Icon lHwnd, lIconHnd, lDrawDc
        End If
        '/* draw the caption
        Render_Caption lOffset, lHwnd, lDrawDc, tRect, True
    Else
        Render_Caption lOffset, lHwnd, lDrawDc, tRect, False
    End If
    
    '/* copy to button
    With tRect
        If m_lTransparentColor = -1 Then
            '/* copy to button
            m_cRender.Blit lHdc, 0, 0, .Right, .Bottom, lDrawDc, 0, 0, SRCCOPY
        Else
            m_cRender.Transparent lHdc, 0, 0, .Right, .Bottom, lDrawDc, 0, 0, .Right, .Bottom, m_lTransparentColor
        End If
    End With
    
    '/* cleanup
    SelectObject lDrawDc, lOldBmp
    DeleteObject lNewBmp
    DeleteDC lDrawDc
    ReleaseDC lHwnd, lHdc

On Error GoTo 0

End Sub

Private Sub Render_Caption(ByVal lOffset As Long, _
                           ByVal lHwnd As Long, _
                           ByVal lHdc As Long, _
                           ByRef tRect As RECT, _
                           Optional ByVal bIcon As Boolean)

Dim lHFont          As Long
Dim lFontOld        As Long
Dim sCaption        As String
Dim sOptions        As String
Dim tLogFont        As LOGFONT

On Error Resume Next

    sCaption = c_CaptionText(CStr(lHwnd))
    sOptions = c_CaptionOptions(CStr(lHwnd))
    With tLogFont
        .lfFaceName = c_CaptionFont(CStr(lHwnd)) & vbNullChar
        .lfHeight = (c_CaptionSize(CStr(lHwnd)) * -20) / Screen.TwipsPerPixelY
        If InStrB(1, sOptions, "B") Then
            .lfWeight = 900
        Else
            .lfWeight = 500
        End If
        If LenB(sOptions) > 0 Then
            .lfItalic = IIf(InStrB(1, sOptions, "I"), 1, 0)
            .lfUnderline = IIf(InStrB(1, sOptions, "U"), 1, 0)
            .lfStrikeOut = IIf(InStrB(1, sOptions, "S"), 1, 0)
        End If
        If m_bTextAntiAliased Then
            .lfQuality = ANTIALIASED_QUALITY
        Else
            .lfQuality = NONANTIALIASED_QUALITY
        End If
    End With

    '/* create font
    lHFont = CreateFontIndirect(tLogFont)
    '/* place in dc
    lFontOld = SelectObject(lHdc, lHFont)
    '/* bg transparent
    SetBkMode lHdc, &H1
    '/* uses an icon
    If bIcon Then
        '/* shift text position
        Select Case m_eIconPosition
        '/* left
        Case 1
            With tRect
                .left = .left + 16
            End With
        '/* right
        Case 2
            With tRect
                .left = .left - 16
            End With
        '/* top
        Case 3
            With tRect
                .top = .top + 16
            End With
        End Select
    End If

    Select Case lOffset
    '/* down
    Case 1
        With tRect
            .top = .top + 1
        End With
        SetTextColor lHdc, m_lColorHiLite
        DrawText lHdc, sCaption, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    '/* disabled
    Case 2
        SetTextColor lHdc, GetSysColor(vb3DHighlight And &H1F)
        OffsetRect tRect, 1, 1
        DrawText lHdc, sCaption, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        SetTextColor lHdc, GetSysColor(vbButtonShadow And &H1F)
        OffsetRect tRect, -1, -1
        DrawText lHdc, sCaption, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    '/* over
    Case 3
        SetTextColor lHdc, m_lColorHiLite
        DrawText lHdc, sCaption, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        '/* focused
    Case 4
        If m_bDrawFocused Then
            DrawText lHdc, sCaption, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        Else
            SetTextColor lHdc, m_lForeColor
            DrawText lHdc, sCaption, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        End If
    '/* normal
    Case Else
        SetTextColor lHdc, m_lForeColor
        DrawText lHdc, sCaption, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    End Select

    '/* delete fonts
    SelectObject lHdc, lFontOld
    DeleteObject lHFont

End Sub

Private Sub Render_Icon(ByVal lHwnd As Long, _
                        ByVal lIcnHnd As Long, _
                        ByVal lHdc As Long, _
                        Optional ByVal bDown As Boolean)

Dim lSDc            As Long
Dim lBmp            As Long
Dim lBmpOld         As Long
Dim lTop            As Long
Dim lLeft           As Long
Dim lTextLen        As Long
Dim hBr             As Long
Dim hBe             As Long
Dim lLBrsh          As LOGBRUSH
Dim tRect           As RECT
Dim tCmd            As RECT

    '/* create a temp image dc
    lSDc = CreateCompatibleDC(lHdc)
    lBmp = CreateCompatibleBitmap(lHdc, 16, 16)
    lBmpOld = SelectObject(lSDc, lBmp)

    With tRect
        .Right = 16
        .Bottom = 16
    End With
    
    '/* backpaint
    lLBrsh.lbColor = vbMagenta
    hBr = CreateBrushIndirect(lLBrsh)
    FillRect lSDc, tRect, hBr
    DeleteObject hBr
    hBr = 0
    
    '/* draw icon
    If IsWindowEnabled(lHwnd) = 0 Then
        hBr = CreateSolidBrush(&H808080)
        hBe = CreateSolidBrush(m_lColorHiLite)
        DrawState lSDc, hBe, 0&, lIcnHnd, 0&, 1&, 1&, 17&, 17&, DST_ICON Or DSS_MONO
        DrawState lSDc, hBr, 0&, lIcnHnd, 0&, 0&, 0&, 16&, 16&, DST_ICON Or DSS_MONO
        DeleteObject hBr
        DeleteObject hBe
    ElseIf bDown Then
        hBr = CreateSolidBrush(&H808080)
        hBe = CreateSolidBrush(m_lColorHiLite)
        DrawState lSDc, hBe, 0&, lIcnHnd, 0&, 1&, 1&, 17&, 17&, DST_ICON Or DSS_MONO
        DrawState lSDc, hBr, 0&, lIcnHnd, 0&, 0&, 0&, 16&, 16&, DST_ICON Or DSS_NORMAL
        DeleteObject hBr
        DeleteObject hBe
    Else
        DrawState lSDc, 0&, 0&, lIcnHnd, 0&, 0&, 0&, 16&, 16&, DST_ICON Or DSS_NORMAL
    End If

    GetWindowRect lHwnd, tCmd
    OffsetRect tCmd, -tCmd.left, -tCmd.top

    Select Case m_eIconPosition
    '/* left
    Case 1
        With tCmd
            lLeft = 8
            lTop = ((.Bottom - tRect.Bottom) / 2) + 1
        End With
    '/* right
    Case 2
        With tCmd
            lLeft = (.Right - tRect.Right) - 8
            lTop = ((.Bottom - tRect.Bottom) / 2) + 1
        End With
    '/* top
    Case 3
        With tCmd
            lLeft = (.Right - tRect.Right) / 2
            lTop = 2
        End With
    End Select
    If bDown Then
        lTop = lTop + 1
    End If

    '/* paint to dc
    m_cRender.Transparent lHdc, lLeft, lTop, 16&, 16&, lSDc, 0&, 0&, 16&, 16&, vbMagenta
    
    SelectObject lSDc, lBmpOld
    DeleteObject lBmp
    DeleteDC lSDc

End Sub

Private Function TranslateColor(ByVal Clr As OLE_COLOR)

   If (Clr And &H80000000) = &H80000000 Then
      TranslateColor = GetSysColor(Clr And &HFF)
   Else
      TranslateColor = Clr
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

Private Sub Control_Attatch(ByVal lHwnd As Long)

Dim lParHnd As Long
Dim lParChk As Long

    If lHwnd = 0 Then Exit Sub
    lParHnd = GetParent(lHwnd)
    If lParHnd = 0 Then Exit Sub
    On Error Resume Next
    lParChk = c_ParStore.Item(CStr(lParHnd))
    On Error GoTo 0

    With m_GSubclass
        .Subclass lHwnd, Me
        .AddMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        If lParChk = 0 Then
            .Subclass lParHnd, Me
            .AddMessage lParHnd, WM_DRAWITEM, MSG_BEFORE
            c_ParStore.Add lParHnd, CStr(lParHnd)
        End If
    End With

End Sub

Private Sub Control_Detach(ByVal lHwnd As Long)

Dim lParHnd   As Long

    If lHwnd = 0 Then Exit Sub
    lParHnd = GetParent(lHwnd)
    If lParHnd = 0 Then Exit Sub
    With m_GSubclass
        .DeleteMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
    End With

End Sub

Private Sub Parent_Detach()

On Error Resume Next

Dim vI As Variant

    With m_GSubclass
        For Each vI In c_ParStore
            If Not CLng(vI) = 0 Then
                .DeleteMessage CLng(vI), WM_DRAWITEM, MSG_BEFORE
            End If
        Next vI
    End With

On Error GoTo 0

End Sub

Private Sub Control_Release()

Dim vInst       As Variant

On Error Resume Next

    For Each vInst In c_InstTrack
        If Not CLng(vInst) = 0 Then
            Control_Detach CLng(vInst)
        End If
    Next vInst
    
    Parent_Detach

On Error GoTo 0

End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)

Dim tDIstc      As DRAWITEMSTRUCT
Dim tRect       As RECT

    Select Case uMsg
    '/* mouse over
    Case WM_MOUSEMOVE
        GetWindowRect lHwnd, tRect
        With tRect
            OffsetRect tRect, -.left, -.top
        End With
        Control_State lHwnd, 3, tRect
        Track_Mouse lHwnd
        lReturn = 1

    Case WM_MOUSELEAVE
        GetWindowRect lHwnd, tRect
        With tRect
            OffsetRect tRect, -.left, -.top
        End With
        Control_State lHwnd, 0&, tRect
        lReturn = 1

    Case WM_DRAWITEM
        CopyMemory tDIstc, ByVal lParam, Len(tDIstc)
        With tDIstc
            OffsetRect .rcItem, -.rcItem.left, -.rcItem.top
            LSet tRect = .rcItem
            Select Case .itemState
            Case 0
                Control_State .hwndItem, 0, tRect
            Case 4
                Control_State .hwndItem, 2, tRect
            Case 16
                Control_State .hwndItem, 4, tRect
            Case 17
                Control_State .hwndItem, 1, tRect
            End Select
        End With
        bHandled = True
        lReturn = 1
    End Select

End Sub

Private Function Control_State(ByVal lHwnd As Long, _
                               ByVal lState As Long, _
                               ByRef tRect As RECT) As Long

'/* 0 normal, 1, pressed, 2 disabled, 3 over, 4 focused

    If IsWindowVisible(lHwnd) = 0 Then Exit Function
    If IsWindowEnabled(lHwnd) = 0 Then
        Render_Image lHwnd, 2, tRect
    Else
        Render_Image lHwnd, lState, tRect
    End If

End Function

Private Sub Backbuffer_Images()

Dim lCt         As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long

    ReDim m_lHdc(4)
    ReDim m_lBmp(4)
    ReDim m_lBmpOld(4)

    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)

    For lCt = 0 To 4
        m_lHdc(lCt) = CreateCompatibleDC(lTmpDc)
        m_lBmp(lCt) = CreateCompatibleBitmap(lTmpDc, m_lCmdWidth, m_lCmdHeight)
        m_lBmpOld(lCt) = SelectObject(m_lHdc(lCt), m_lBmp(lCt))
        m_cRender.Blit m_lHdc(lCt), 0, 0, m_lCmdWidth, m_lCmdHeight, m_cCmdDc.hdc, (m_lCmdWidth * lCt), 0, SRCCOPY
    Next lCt

    ReleaseDC lHwnd, lTmpDc

End Sub

Private Sub Backbuffer_Destroy()

Dim lCt     As Long

    For lCt = 0 To 4
        SelectObject m_lHdc(lCt), m_lBmpOld(lCt)
        DeleteObject m_lBmp(lCt)
        DeleteDC m_lHdc(lCt)
        m_lHdc(lCt) = 0
        m_lBmp(lCt) = 0
        m_lBmpOld(lCt) = 0
    Next lCt
        
End Sub

Public Sub Cleanup()

On Error Resume Next

    If Not m_lParentHwnd = 0 Then
        Control_Release
        Backbuffer_Destroy
        If Not m_cCmdDc Is Nothing Then Set m_cCmdDc = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_pCmdImg Is Nothing Then Set m_pCmdImg = Nothing
        If Not c_InstTrack Is Nothing Then Set c_InstTrack = Nothing
        If Not c_CaptionText Is Nothing Then Set c_CaptionText = Nothing
        If Not c_CaptionFont Is Nothing Then Set c_CaptionFont = Nothing
        If Not c_CaptionSize Is Nothing Then Set c_CaptionSize = Nothing
        If Not c_ParStore Is Nothing Then Set c_ParStore = Nothing
        If Not c_CaptionOptions Is Nothing Then Set c_CaptionOptions = Nothing
        If Not c_ButtonIcon Is Nothing Then Set c_ButtonIcon = Nothing
        If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
        If Not p_OParentObj Is Nothing Then Set p_OParentObj = Nothing
        Erase m_lHdc
        Erase m_lBmp
        Erase m_lBmpOld
    End If
    m_lParentHwnd = 0

On Error GoTo 0

End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub
