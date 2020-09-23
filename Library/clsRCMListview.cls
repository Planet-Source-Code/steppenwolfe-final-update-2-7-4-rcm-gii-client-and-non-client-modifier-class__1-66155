VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMListview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements GXISubclass

'/* frame styles
Private Const SWP_NOSIZE                                As Long = &H1
Private Const SWP_NOMOVE                                As Long = &H2
Private Const SWP_NOZORDER                              As Long = &H4
Private Const SWP_NOACTIVATE                            As Long = &H10
Private Const SWP_FRAMECHANGED                          As Long = &H20
Private Const WS_EX_CLIENTEDGE                          As Long = &H200
Private Const WS_EX_STATICEDGE                          As Long = &H20000
Private Const GWL_STYLE                                 As Long = (-16)
Private Const GWL_EXSTYLE                               As Long = (-20)
Private Const WS_BORDER                                 As Long = &H800000
Private Const SWP_CHANGE                                As Long = SWP_NOACTIVATE Or _
    SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
'/* lvdata
Private Const LVM_FIRST                                 As Long = &H1000
Private Const LVM_GETCOLUMN                             As Long = (LVM_FIRST + 25)
Private Const LVM_GETHEADER                             As Long = (LVM_FIRST + 31)
Private Const LVCF_TEXT                                 As Long = &H4
Private Const LVCF_IMAGE                                As Long = &H10
Private Const HDM_FIRST                                 As Long = &H1200
Private Const HDM_GETITEMCOUNT                          As Long = (HDM_FIRST + 0)
'/* header data
Private Const HDM_HITTEST                               As Double = (&H1200 + 6)
Private Const HDM_GETITEMRECT                           As Double = (&H1200 + 7)
Private Const HDM_GETIMAGELIST                          As Double = (&H1200 + 9)
Private Const H_MAX                                     As Long = &HFFFF + 1
Private Const HDN_FIRST                                 As Long = H_MAX - 300
Private Const HDN_LAST                                  As Long = H_MAX - 399
Private Const NM_FIRST                                  As Long = H_MAX
Private Const NONANTIALIASED_QUALITY                    As Long = &H3
Private Const ANTIALIASED_QUALITY                       As Long = &H4
'/* font align
Private Const DT_CENTER                                 As Long = &H1
Private Const DT_VCENTER                                As Long = &H4
Private Const DT_SINGLELINE                             As Long = &H20
Private Const DT_CALCRECT                               As Long = &H400


Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
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

Private Type POINTAPI
    x                                                   As Long
    y                                                   As Long
End Type

Private Type HDHITTESTINFO
    pt                                                  As POINTAPI
    flags                                               As Long
    iItem                                               As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                              As Long
    dwFlags                                             As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                           As Long
    dwHoverTime                                         As Long
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

Private Type LVCOLUMN
    Mask                                                As Long
    fmt                                                 As Long
    cx                                                  As Long
    pszText                                             As Long
    cchTextMax                                          As Long
    iSubItem                                            As Long
    iImage                                              As Long
    iOrder                                              As Long
End Type

Private Type NMHDR
    hwndFrom                                            As Long
    idfrom                                              As Long
    code                                                As Long
End Type

Private Type TCRD
    Lo                                                  As Integer
    Hi                                                  As Integer
End Type

Private Type TCLC
    lBase                                               As Long
End Type

Private Type PAINTSTRUCT
    hdc                                                 As Long
    fErase                                              As Boolean
    rcPaint                                             As RECT
    fRestore                                            As Boolean
    fIncUpdate                                          As Boolean
    rgbReserved(32)                                     As Byte
End Type


Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long


Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

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

Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, _
                                                    ByVal nBkMode As Long) As Long

Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                                   ByVal lpRect As Long, _
                                                                                   ByVal bErase As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

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

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Function ImageList_Draw Lib "Comctl32" (ByVal hImageList As Long, _
                                                        ByVal lIndex As Long, _
                                                        ByVal hdc As Long, _
                                                        ByVal x As Long, _
                                                        ByVal y As Long, _
                                                        ByVal fStyle As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function PathCompactPath Lib "shlwapi.dll" Alias "PathCompactPathA" (ByVal hdc As Long, _
                                                                                     ByVal pszPath As String, _
                                                                                     ByVal dX As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bTrackUser32                      As Boolean
Private m_bTextAntiAliased                  As Boolean
Private m_bPainting                         As Boolean
Private m_lCurrState                        As Long
Private m_lListViewObj                      As Long
Private m_lHdrHandle                        As Long
Private m_lTextForeColor                    As Long
Private m_lTextHighLite                     As Long
Private m_lLastMsg                          As Long
Private m_lParentHwnd                       As Long
Private m_lHeaderWidth                      As Long
Private m_lHeaderHeight                     As Long
Private m_lImlHandle                        As Long
Private m_lHdc()                            As Long
Private m_lBmp()                            As Long
Private m_lBmpOld()                         As Long
Private m_tFont                             As LOGFONT
Private c_ColumnText                        As Collection
Private c_ColumnIcon                        As Collection
Private m_pHeader                           As StdPicture
Private m_cRender                           As clsRender
Private m_cHeaderDc                         As clsStoreDc
Private m_GSubclass                         As GXMSubclass


Private Sub Class_Initialize()
    InitCommonControls
    m_bTrackUser32 = Function_Exported("TrackMouseEvent", "User32")
    m_lTextForeColor = &H111111
    m_lTextHighLite = &H676767
End Sub

Public Property Get p_OListViewObj() As Object

Dim oTemp   As Object

    If m_lListViewObj = 0 Then
        Set p_OListViewObj = Nothing
    Else
        CopyMemory oTemp, m_lListViewObj, 4
        Set p_OListViewObj = oTemp
        CopyMemory oTemp, 0&, 4
    End If

End Property

Public Property Set p_OListViewObj(ByVal PropVal As Object)
    m_lListViewObj = ObjPtr(PropVal)
End Property

Public Property Get p_IHeader() As StdPicture
'/* header image
    Set p_IHeader = m_pHeader
End Property

Public Property Set p_IHeader(PropVal As StdPicture)
    Set m_pHeader = PropVal
End Property

Public Property Get p_TextAntiAliased() As Boolean
'/* anti-aliased text
    p_TextAntiAliased = m_bTextAntiAliased
End Property

Public Property Let p_TextAntiAliased(ByVal PropVal As Boolean)
    m_bTextAntiAliased = PropVal
End Property

Public Property Get p_TextForeColor() As Long
'/* column text forecolor
    p_TextForeColor = m_lTextForeColor
End Property

Public Property Let p_TextForeColor(ByVal PropVal As Long)
    m_lTextForeColor = TranslateColor(PropVal)
End Property

Public Property Get p_TextHighLite() As Long
'/* column text highlite color
    p_TextHighLite = m_lTextHighLite
End Property

Public Property Let p_TextHighLite(ByVal PropVal As Long)
    m_lTextHighLite = TranslateColor(PropVal)
End Property

Public Property Get p_ColumnIcon(ByVal lHwnd As Long, _
                                 ByVal lColumn As Long) As Long

Dim uLVC    As LVCOLUMN

    If lHwnd = 0 Then Exit Property
    uLVC.Mask = LVCF_IMAGE
    SendMessage lHwnd, LVM_GETCOLUMN, lColumn, uLVC
    p_ColumnIcon = uLVC.iImage

End Property

Private Property Get p_ColumnCount(ByVal lHeaderHnd As Long) As Long

    If lHeaderHnd = 0 Then Exit Property
    p_ColumnCount = SendMessageLong(lHeaderHnd, HDM_GETITEMCOUNT, 0&, 0&)

End Property

Private Property Get p_HeaderHwnd(ByVal lHwnd As Long) As Long

    If lHwnd = 0 Then Exit Property
    p_HeaderHwnd = SendMessageLong(lHwnd, LVM_GETHEADER, 0&, 0&)

End Property

Private Property Get p_ColumnText(ByVal lHeaderHnd As Long, _
                                  ByVal lColumn As Long) As String

Dim aText(261)      As Byte
Dim uLVC            As LVCOLUMN

On Error GoTo Handler

    If lHeaderHnd = 0 Then Exit Property
    With uLVC
        .pszText = VarPtr(aText(0))
        .cchTextMax = UBound(aText)
        .Mask = LVCF_TEXT
    End With

    SendMessage lHeaderHnd, LVM_GETCOLUMN, lColumn, uLVC
    p_ColumnText = left(StrConv(aText(), vbUnicode), uLVC.cchTextMax)
    p_ColumnText = left(p_ColumnText, InStr(1, p_ColumnText, vbNullChar) - 1)

Handler:
    On Error GoTo 0

End Property

Public Sub Attatch_Listview()

Dim tBmp        As BITMAP

On Error GoTo Handler

    '/* test for defaults
    If (p_IHeader Is Nothing) Or _
        (p_OListViewObj Is Nothing) Then Exit Sub
    
    '/* store parent hwnd
    m_lParentHwnd = p_OListViewObj.hwnd

    '/* default sizes
    GetObject p_IHeader.handle, Len(tBmp), tBmp
    With tBmp
        m_lHeaderWidth = (.bmWidth / 4)
        m_lHeaderHeight = .bmHeight
    End With
    
    Set m_GSubclass = New GXMSubclass
    '/* init image classes
    Set m_cHeaderDc = New clsStoreDc
    m_cHeaderDc.CreateFromPicture p_IHeader
    Set m_cRender = New clsRender
    
    Backbuffer_Images
    '/* add controls
    Control_Allocation

Handler:
    On Error GoTo 0

End Sub

Private Sub Control_Allocation()
'/* assign object params to variant array

Dim sCtrl       As String

On Error Resume Next

    sCtrl = IsClassName(p_OListViewObj.hwnd)
    If InStrB(1, sCtrl, "ListView") > 0 Then
        With p_OListViewObj
            '/* filter invalids or user exempted via tag
            If (.hwnd = 0) Or (.Tag = "NO") Then GoTo Skip
            m_lHdrHandle = p_HeaderHwnd(.hwnd)
            If m_lHdrHandle = 0 Then GoTo Skip
            '/* get font settings
            With m_tFont
                .lfHeight = (p_OListViewObj.Font.Size * -20) / Screen.TwipsPerPixelY
                .lfCharSet = 1
                .lfItalic = p_OListViewObj.Font.Italic
                .lfUnderline = p_OListViewObj.Font.Underline
                .lfFaceName = p_OListViewObj.Font.Name & vbNullChar
                If m_bTextAntiAliased Then
                    .lfQuality = ANTIALIASED_QUALITY
                Else
                    .lfQuality = NONANTIALIASED_QUALITY
                End If
                .lfWeight = 500
            End With
            '/* change style
            Flatten_Container .hwnd
            '/* get column text
            Store_ColumnText
            '/* store column icons
            Store_ColumnIcon
            '/* start subclassing
            Control_Attatch
Skip:
        End With
    End If
    
On Error GoTo 0

End Sub

Private Function IsClassName(ByVal lHwnd As Long) As String

Dim lRet   As Long
Dim sClass As String

    sClass = Space$(255)
    lRet = GetClassName(lHwnd, sClass, 255)
    IsClassName = left(sClass, lRet)

End Function

Private Function Function_Exported(ByVal sFunction As String, _
                                   ByVal sModule As String) As Boolean
'/* test for library support

Dim lModule         As Long
Dim bLoaded         As Boolean

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

Private Function Column_Focused(ByVal lHwnd As Long, _
                                ByVal lParam As Long) As Long

'/* draw focused header

Dim tHdHit      As HDHITTESTINFO
Dim tCoord      As TCRD
Dim tLng        As TCLC
Dim tRect       As RECT

    tLng.lBase = lParam
    LSet tCoord = tLng
    With tHdHit.pt
        .x = tCoord.Lo
        .y = tCoord.Hi
    End With
    
    SendMessage lHwnd, HDM_HITTEST, 0&, tHdHit
    SendMessage lHwnd, HDM_GETITEMRECT, tHdHit.iItem, tRect
    If Not tHdHit.iItem = -1 Then
        Column_Focused = tHdHit.iItem
    End If

End Function

Private Function Column_LostFocus(ByVal lHwnd As Long, _
                                  ByVal lCol As Long) As Boolean

'/* reset panel after lost focus

Dim tRect       As RECT
Dim lHdc        As Long

    SendMessage lHwnd, HDM_GETITEMRECT, lCol, tRect
    lHdc = GetDC(lHwnd)
    Image_Render 0, lCol
    ReleaseDC lHwnd, lHdc

End Function

Private Sub Image_Render(ByVal lOffset As Long, _
                         ByVal lColNum As Long)

'/* render column panel

Dim lCt             As Long
Dim lTmpDc          As Long
Dim lHwnd           As Long
Dim lDrawDc         As Long
Dim lBmp            As Long
Dim lBmpOld         As Long
Dim lX              As Long
Dim lHdc            As Long
Dim tHdr            As RECT
Dim tTemp           As RECT
Dim tRect           As RECT
Dim tTR             As RECT

On Error GoTo Handler

    GetClientRect m_lHdrHandle, tHdr
    GetClientRect m_lParentHwnd, tTR
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)
    lDrawDc = CreateCompatibleDC(lTmpDc)
    lHdc = GetDC(m_lHdrHandle)
    LSet tRect = tHdr

    With tHdr
        OffsetRect tHdr, -.left, -.top
        lBmp = CreateCompatibleBitmap(lTmpDc, tTR.Right, .Bottom)
    End With
    lBmpOld = SelectObject(lDrawDc, lBmp)

    lX = 0
    For lCt = 0 To p_ColumnCount(m_lHdrHandle) - 1
        SendMessage m_lHdrHandle, HDM_GETITEMRECT, lCt, tTemp
        With tTemp
            If lCt = lColNum Then: lX = lOffset: Else: lX = 0
            '/* left side
            m_cRender.Blit lDrawDc, .left, 0, 3, .Bottom - 2, m_lHdc(lX), 0, 0, SRCCOPY
            '/* center
            m_cRender.Stretch lDrawDc, .left + 3, 0, .Right - 3, .Bottom, m_lHdc(lX), 3, 0, m_lHeaderWidth - 3, m_lHeaderHeight, SRCCOPY
        End With
        If lX = 2 Then
            Column_Text lDrawDc, lCt, True, tTemp
        Else
            Column_Text lDrawDc, lCt, False, tTemp
        End If
        Column_Icon lCt, lDrawDc, tTemp
    Next lCt

    With tRect
        m_cRender.Blit lHdc, .left, .top, .Right, .Bottom, lDrawDc, 0, 0, SRCCOPY
    End With

    SelectObject lDrawDc, lBmpOld
    DeleteObject lBmp
    DeleteDC lDrawDc
    ReleaseDC lHwnd, lTmpDc
    ReleaseDC m_lHdrHandle, lHdc

Handler:
    On Error GoTo 0

End Sub

Private Sub Column_Text(ByVal lHdc As Long, _
                        ByVal lColNum As Long, _
                        ByVal bHighLite As Boolean, _
                        ByRef tRect As RECT)

'/* draw panel caption
Dim sText       As String
Dim tPnt        As POINTAPI
Dim lHFont      As Long
Dim lFontOld    As Long

On Error Resume Next

    sText = c_ColumnText.Item("h:" & lColNum)
    If LenB(sText) = 0 Then Exit Sub
    lHFont = CreateFontIndirect(m_tFont)
    lFontOld = SelectObject(lHdc, lHFont)
    SetBkMode lHdc, 1

    If bHighLite Then
        SetTextColor lHdc, m_lTextHighLite
    Else
        SetTextColor lHdc, m_lTextForeColor
    End If
    
    If Not m_lImlHandle = 0 Then
        Dim tShft As RECT
        CopyRect tShft, tRect
        '/ calculate text size
        GetTextExtentPoint32 lHdc, sText, lstrlen(sText) + 2, tPnt
        With tRect
            tShft.left = .left + 20
            '/* test min size
            If (.Right - .left) < (tPnt.x + 40) Then
                Dim sTmp As String
                sTmp = String(255, 0)
                sTmp = sText & vbNullChar
                '/* compact text
                PathCompactPath lHdc, sTmp, (.Right - .left) - 40
            Else
                sTmp = sText
            End If
        End With
        DrawText lHdc, sTmp, -1, tShft, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Else
        DrawText lHdc, sText, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    End If
    
    SelectObject lHdc, lFontOld
    DeleteObject lHFont

On Error GoTo 0

End Sub

Public Sub Store_ColumnText()
'/* store column captions

Dim lCt         As Long
Dim sText       As String

On Error Resume Next

    Set c_ColumnText = New Collection
    If m_lParentHwnd = 0 Then Exit Sub
    If m_lHdrHandle = 0 Then Exit Sub
    For lCt = 0 To p_ColumnCount(m_lHdrHandle) - 1
        sText = p_ColumnText(m_lParentHwnd, lCt)
        If Not LenB(sText) = 0 Then
            c_ColumnText.Add sText, "h:" & lCt
        End If
    Next lCt

On Error GoTo 0

End Sub

Private Sub Store_ColumnIcon()
'/* store icon indece and header iml handle

Dim lCt             As Long
Dim lIcnIdx         As Long

On Error Resume Next

    Set c_ColumnIcon = New Collection
    If m_lParentHwnd = 0 Then Exit Sub
    If m_lHdrHandle = 0 Then Exit Sub
    m_lImlHandle = SendMessageLong(m_lHdrHandle, HDM_GETIMAGELIST, 0&, 0&)
    If m_lImlHandle = 0 Then Exit Sub
    For lCt = 0 To p_ColumnCount(m_lHdrHandle) - 1
        lIcnIdx = p_ColumnIcon(m_lParentHwnd, lCt)
        If Not lIcnIdx = -1 Then
            c_ColumnIcon.Add lIcnIdx, "h:" & lCt
        End If
    Next lCt

On Error GoTo 0

End Sub

Private Sub Column_Icon(ByVal lColNum As Long, _
                        ByVal lHdc As Long, _
                        ByRef tRect As RECT)
'/* draw header icon

Dim lIcnIdx         As Long

On Error Resume Next

    If c_ColumnIcon Is Nothing Then Exit Sub
    If lColNum = -1 Then Exit Sub
    If m_lImlHandle = 0 Then Exit Sub
    
    lIcnIdx = c_ColumnIcon.Item("h:" & lColNum)
    With tRect
        ImageList_Draw m_lImlHandle, lIcnIdx, lHdc, .left + 10&, .top + 2, &H1
    End With

On Error GoTo 0

End Sub

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

Private Function TranslateColor(ByVal Clr As OLE_COLOR)

   If (Clr And &H80000000) = &H80000000 Then
      TranslateColor = GetSysColor(Clr And &HFF)
   Else
      TranslateColor = Clr
   End If

End Function

Private Sub Backbuffer_Images()
'/* store state images

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
        m_lBmp(lCt) = CreateCompatibleBitmap(lTmpDc, m_lHeaderWidth, m_lHeaderHeight)
        m_lBmpOld(lCt) = SelectObject(m_lHdc(lCt), m_lBmp(lCt))
        m_cRender.Blit m_lHdc(lCt), 0, 0, m_lHeaderWidth, m_lHeaderHeight, m_cHeaderDc.hdc, (m_lHeaderWidth * lCt), 0, SRCCOPY
    Next lCt

    ReleaseDC lHwnd, lTmpDc

End Sub

Private Sub Backbuffer_Destroy()
'/* destroy backbuffer

Dim lCt As Long

    For lCt = 0 To 3
        SelectObject m_lHdc(lCt), m_lBmpOld(lCt)
        DeleteObject m_lBmp(lCt)
        DeleteDC m_lHdc(lCt)
        m_lBmpOld(lCt) = 0
        m_lHdc(lCt) = 0
        m_lBmp(lCt) = 0
    Next lCt

End Sub

Private Sub Control_Attatch()

    With m_GSubclass
        If Not m_lHdrHandle = 0 Then
            .Subclass m_lHdrHandle, Me
            .AddMessage m_lHdrHandle, WM_MOUSEMOVE, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_MOUSELEAVE, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_PAINT, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_LBUTTONDOWN, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_LBUTTONUP, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_ERASEBKGND, MSG_BEFORE
        Else
            Exit Sub
        End If
        If Not m_lParentHwnd = 0 Then
            .Subclass m_lParentHwnd, Me
            .AddMessage m_lParentHwnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
            .AddMessage m_lParentHwnd, WM_STYLECHANGED, MSG_BEFORE
        End If
    End With

End Sub

Private Sub Control_Detatch()

    With m_GSubclass
        If Not m_lHdrHandle = 0 Then
            .DeleteMessage m_lHdrHandle, WM_MOUSEMOVE, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_MOUSELEAVE, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_PAINT, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_LBUTTONDOWN, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_LBUTTONUP, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_ERASEBKGND, MSG_BEFORE
            .UnSubclass m_lHdrHandle
        Else
            Exit Sub
        End If
        If Not m_lParentHwnd = 0 Then
            .DeleteMessage m_lParentHwnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
            .DeleteMessage m_lParentHwnd, WM_STYLECHANGED, MSG_BEFORE
            .UnSubclass m_lParentHwnd
        End If
    End With

End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)

Static lSelected        As Long
Static bSelect          As Boolean
Dim tPAINT              As PAINTSTRUCT

    If IsWindowVisible(m_lParentHwnd) = 0 Then
        lReturn = 0
        bHandled = True
        Exit Sub
    End If
    
    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            BeginPaint lHwnd, tPAINT
            Select Case m_lCurrState
            Case 1
                Image_Render 1, lSelected
            Case 2
                Image_Render 2, lSelected
            Case Else
                Image_Render 0, 0
            End Select
            EndPaint lHwnd, tPAINT
            m_bPainting = False
            bHandled = True
        Else
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If
        
    Case WM_LBUTTONDOWN
        m_lCurrState = 1
        bSelect = True
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&

    Case WM_LBUTTONUP
        m_lCurrState = 0
        bSelect = False
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&

    Case WM_ERASEBKGND
        lReturn = 1
        bHandled = True

    Case WM_MOUSEMOVE
        If GetAsyncKeyState(vbLeftButton) = 0 Then
            lSelected = Column_Focused(lHwnd, lParam)
            m_lCurrState = 2
            If Not bSelect Then
                SendMessageLong lHwnd, WM_PAINT, 0&, 0&
            End If
            Track_Mouse lHwnd
        End If
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)

    Case WM_MOUSELEAVE
        m_lCurrState = 0
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)

    Case WM_WINDOWPOSCHANGED, WM_STYLECHANGED
        m_lCurrState = 0
        SendMessageLong m_lHdrHandle, WM_PAINT, 0&, 0&
    End Select

    m_lLastMsg = uMsg

End Sub

Private Sub Cleanup()

On Error Resume Next

    If Not m_lParentHwnd = 0 Then
        Control_Detatch
        Backbuffer_Destroy
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cHeaderDc Is Nothing Then Set m_cHeaderDc = Nothing
        If Not p_IHeader Is Nothing Then Set p_IHeader = Nothing
        If Not c_ColumnIcon Is Nothing Then Set c_ColumnIcon = Nothing
        If Not c_ColumnText Is Nothing Then Set c_ColumnText = Nothing
        If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
        If Not p_OListViewObj Is Nothing Then Set p_OListViewObj = Nothing
        Erase m_lHdc
        Erase m_lBmp
        Erase m_lBmpOld
    End If
    m_lHdrHandle = 0
    m_lImlHandle = 0
    m_lParentHwnd = 0

On Error GoTo 0

End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

