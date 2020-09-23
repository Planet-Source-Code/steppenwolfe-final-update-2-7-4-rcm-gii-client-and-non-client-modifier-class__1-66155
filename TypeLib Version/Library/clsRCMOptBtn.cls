VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMOptBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'/* thanks to Mario Flores Gonzalez and his great winxpc for inspiring some of the message hacks..

Implements WinSubHookGX.GXISubclass

'/* Button Control Messages
Private Const BM_GETCHECK                   As Long = &HF0
Private Const BM_GETSTATE                   As Long = &HF2
Private Const DT_LEFT                       As Long = &H0
Private Const DT_VCENTER                    As Long = &H4
Private Const DT_SINGLELINE                 As Long = &H20
Private Const DEFAULT_QUALITY               As Long = &H0
Private Const DRAFT_QUALITY                 As Long = &H1
Private Const PROOF_QUALITY                 As Long = &H2
Private Const NONANTIALIASED_QUALITY        As Long = &H3
Private Const ANTIALIASED_QUALITY           As Long = &H4


Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum


Private Type RECT
    left                                    As Long
    top                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type

Private Type BITMAP
    bmType                                  As Long
    bmWidth                                 As Long
    bmHeight                                As Long
    bmWidthBytes                            As Long
    bmPlanes                                As Integer
    bmBitsPixel                             As Integer
    bmBits                                  As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                  As Long
    dwFlags                                 As TRACKMOUSEEVENT_FLAGS
    hwndTrack                               As Long
    dwHoverTime                             As Long
End Type

Private Type PAINTSTRUCT
    hdc                                     As Long
    fErase                                  As Boolean
    rcPaint                                 As RECT
    fRestore                                As Boolean
    fIncUpdate                              As Boolean
    rgbReserved(32)                         As Byte
End Type

Private Type LOGFONT
    lfHeight                                As Long
    lfWidth                                 As Long
    lfEscapement                            As Long
    lfOrientation                           As Long
    lfWeight                                As Long
    lfItalic                                As Byte
    lfUnderline                             As Byte
    lfStrikeOut                             As Byte
    lfCharSet                               As Byte
    lfOutPrecision                          As Byte
    lfClipPrecision                         As Byte
    lfQuality                               As Byte
    lfPitchAndFamily                        As Byte
    lfFaceName                              As String * 32
End Type


Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
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

Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                                   ByVal lpRect As Long, _
                                                                                   ByVal bErase As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                                ByVal wMsg As Long, _
                                                                                ByVal wParam As Long, _
                                                                                ByVal lParam As String) As Long

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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, _
                                                       ByVal crColor As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bTrackUser32                  As Boolean
Private m_bEnabled                      As Boolean
Private m_bPainting                     As Boolean
Private m_bPreUnload                    As Boolean
Private m_lParentObj                    As Long
Private m_lLastSel                      As Long
Private m_lParentHwnd                   As Long
Private m_lOptWidth                     As Long
Private m_lOptHeight                    As Long
Private m_lTransparentColor             As Long
Private m_lHdc()                        As Long
Private m_lBmp()                        As Long
Private m_lBmpOld()                     As Long
Private c_InstTrack                     As Collection
Private m_cCtrlPtr                      As Collection
Private m_pOptImg                       As StdPicture
Private m_cRender                       As clsRender
Private m_cOptDc                        As clsStoreDc
Private m_GSubclass                     As GXMSubclass

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

Public Property Get p_IOptImg() As StdPicture
'/* command button image
    Set p_IOptImg = m_pOptImg
End Property

Public Property Set p_IOptImg(PropVal As StdPicture)
    Set m_pOptImg = PropVal
End Property

Public Property Get p_TransparentColor() As Long
'/* transparency color
    p_TransparentColor = m_lTransparentColor
End Property

Public Property Let p_TransparentColor(PropVal As Long)
    m_lTransparentColor = PropVal
End Property

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object

Dim oThis As Object

    CopyMemory oThis, lPtr, 4
    Set ObjectFromPtr = oThis
    CopyMemory oThis, 0&, 4

End Property

Private Property Get PtrFromObject(ByRef oThis As Object) As Long
    PtrFromObject = ObjPtr(oThis)
End Property

Private Sub Class_Initialize()
    InitCommonControls
    m_bTrackUser32 = Function_Exported("TrackMouseEvent", "User32")
End Sub
  
Public Sub Attatch_OptBtn()

Dim tBmp    As BITMAP

On Error GoTo Handler
    
    '/* test for defaults
    If (p_IOptImg Is Nothing) Or _
       (p_OParentObj Is Nothing) Then
        Exit Sub
    End If

    '/* store parent hwnd
    m_lParentHwnd = p_OParentObj.hwnd

    '/* default sizes
    GetObject p_IOptImg.handle, Len(tBmp), tBmp
    With tBmp
        m_lOptWidth = (.bmWidth / 8)
        m_lOptHeight = .bmHeight
    End With
    
    '/* transparency color default
    If m_lTransparentColor = 0 Then
        m_lTransparentColor = &HFF00FF
    End If
    
    Set m_GSubclass = New GXMSubclass
    '/* init image classes
    Set m_cOptDc = New clsStoreDc
    m_cOptDc.CreateFromPicture p_IOptImg
    Set m_cRender = New clsRender
    
    Backbuffer_Images
    '/* add controls
    Control_Allocation

Handler:
    On Error GoTo 0

End Sub


Private Sub Control_Allocation()
'/* assign object params to variant array

Dim oCtrl       As Control
Dim lVcnt       As Long
Dim sOptions    As String
Dim lHwndStore  As Long
Dim sCtrl       As String

On Error Resume Next

    Set c_InstTrack = New Collection
    Set m_cCtrlPtr = New Collection
    
    For Each oCtrl In p_OParentObj
        sCtrl = IsClassName(oCtrl.hwnd)
        If InStrB(1, sCtrl, "OptionButton") > 0 Then
            If oCtrl.hwnd = m_lParentHwnd Then GoTo Skip
            sOptions = ""
            With oCtrl
                '/* filter invalids or user exempted via tag
                If (.hwnd = 0) Or (.Tag = "NO") Then
                    GoTo Skip
                End If
                '/* filter duplications
                lHwndStore = c_InstTrack.Item(CStr(.hwnd))
                If Not lHwndStore = 0 Then
                    lHwndStore = 0
                    GoTo Skip
                End If
                c_InstTrack.Add .hwnd, CStr(.hwnd)
                m_cCtrlPtr.Add PtrFromObject(oCtrl), CStr(.hwnd)
                '/* start subclassing
                Control_Attach .hwnd
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
    
    '/* refresh controls
    m_bPreUnload = False
    Redraw_All
    
On Error GoTo 0

End Sub

Private Function IsClassName(ByVal lHwnd As Long) As String

Dim lRet        As Long
Dim sClass      As String

    sClass = Space(255)
    lRet = GetClassName(lHwnd, sClass, 255)
    IsClassName = left$(sClass, lRet)

End Function

Private Sub Control_Release()

Dim vInst   As Variant
Dim lCt     As Long

On Error Resume Next

    If c_InstTrack Is Nothing Then Exit Sub
    For Each vInst In c_InstTrack
        If Not CLng(vInst) = 0 Then
            Control_Detach CLng(vInst)
        End If
    Next vInst
    
On Error GoTo 0

End Sub

Private Sub Image_Render(ByVal lHwnd As Long, _
                         ByVal lOffset As Long)

Dim lHdc        As Long
Dim tRect       As RECT


On Error Resume Next

    lHdc = GetDC(lHwnd)
    If lHdc = 0 Then Exit Sub
    GetWindowRect lHwnd, tRect
    OffsetRect tRect, -tRect.left, -tRect.top
    
    With tRect
        '/* copy from stored dc
        m_cRender.Transparent lHdc, .left, (.Bottom - 13) / 2, m_lOptWidth, m_lOptHeight, m_lHdc(lOffset), 0, 0, m_lOptWidth, m_lOptHeight, m_lTransparentColor
    End With

    ReleaseDC lHwnd, lHdc

On Error GoTo 0

End Sub

Private Sub Draw_Caption(ByVal lHwnd As Long, _
                         ByVal lHdc As Long, _
                         ByRef tSize As RECT, _
                         Optional ByVal bDisabled As Boolean)

Dim hBrush      As Long
Dim lClr        As Long
Dim lTmp        As Long
Dim lHFont      As Long
Dim lFontOld    As Long
Dim sCap        As String
Dim oTmp        As Object
Dim tRTmp       As RECT
Dim tLogFont    As LOGFONT

On Error Resume Next
    
    lTmp = m_cCtrlPtr.Item(CStr(lHwnd))
    If lTmp = 0 Then Exit Sub
    Set oTmp = ObjectFromPtr(lTmp)
    lClr = oTmp.BackColor
    hBrush = CreateSolidBrush(TranslateColor(lClr))
    FillRect lHdc, tSize, hBrush
    DeleteObject hBrush
    
    With tLogFont
        .lfFaceName = oTmp.Font + vbNullChar
        .lfHeight = (oTmp.Font.Size * -20) / Screen.TwipsPerPixelY
        .lfWeight = 500
        .lfItalic = IIf(oTmp.FontItalic, 1, 0)
        .lfUnderline = IIf(oTmp.FontUnderline, 1, 0)
        .lfStrikeOut = IIf(oTmp.FontStrikethru, 1, 0)
        .lfQuality = ANTIALIASED_QUALITY
    End With
    
    sCap = oTmp.Caption
    lHFont = CreateFontIndirect(tLogFont)
    lFontOld = SelectObject(lHdc, lHFont)
    SetBkMode lHdc, 1
    SetTextColor lHdc, oTmp.ForeColor
    CopyRect tRTmp, tSize
    tRTmp.left = m_lOptWidth + 4
    DrawText lHdc, sCap, -1, tRTmp, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    If bDisabled Then
        OffsetRect tRTmp, -1, -1
        SetTextColor lHdc, &HCCCCCC
        DrawText lHdc, sCap, -1, tRTmp, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    End If
    SelectObject lHdc, lFontOld
    DeleteObject lHFont
    Set oTmp = Nothing

On Error GoTo 0

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

Private Sub Backbuffer_Images()

Dim lCt     As Long
Dim lTmpDc  As Long
Dim lHwnd   As Long

    ReDim m_lHdc(7)
    ReDim m_lBmp(7)
    ReDim m_lBmpOld(7)
    
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)

    For lCt = 0 To 7
        m_lHdc(lCt) = CreateCompatibleDC(lTmpDc)
        m_lBmp(lCt) = CreateCompatibleBitmap(lTmpDc, m_lOptWidth, m_lOptHeight)
        m_lBmpOld(lCt) = SelectObject(m_lHdc(lCt), m_lBmp(lCt))
        m_cRender.Blit m_lHdc(lCt), 0, 0, m_lOptWidth, m_lOptHeight, m_cOptDc.hdc, (m_lOptWidth * lCt), 0, SRCCOPY
    Next lCt

    ReleaseDC lHwnd, lTmpDc
    
End Sub

Private Sub Backbuffer_Destoy()

Dim lCt As Long

    For lCt = 0 To 7
        SelectObject m_lHdc(lCt), m_lBmpOld(lCt)
        DeleteObject m_lBmp(lCt)
        DeleteDC m_lHdc(lCt)
        m_lHdc(lCt) = 0
        m_lBmp(lCt) = 0
        m_lBmpOld(lCt) = 0
    Next lCt
    
End Sub

Public Sub Redraw_All()

Dim vI As Variant
Dim lHnd As Long

    If m_bPreUnload Then Exit Sub
    'If c_InstTrack Is Nothing Then Exit Sub
    For Each vI In c_InstTrack
        If Not CLng(vI) = 0 Then
            lHnd = CLng(vI)
            Image_Render lHnd, Control_State(lHnd, True)
        End If
    Next vI

End Sub

Public Sub Control_Redraw(ByVal lHwnd As Long)

    InvalidateRectAsNull lHwnd, 0&, 0&
    UpdateWindow lHwnd

End Sub

Private Sub Control_Attach(ByVal lHwnd As Long)

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .Subclass lHwnd, Me
        .AddMessage lHwnd, WM_ENABLE, MSG_BEFORE
        .AddMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
        .AddMessage lHwnd, WM_KEYDOWN, MSG_BEFORE
        .AddMessage lHwnd, WM_KEYUP, MSG_BEFORE
        .AddMessage lHwnd, WM_KILLFOCUS, MSG_BEFORE
        .AddMessage lHwnd, WM_LBUTTONDOWN, MSG_BEFORE
        .AddMessage lHwnd, WM_LBUTTONUP, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .AddMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        .AddMessage lHwnd, WM_PAINT, MSG_BEFORE
        .AddMessage lHwnd, WM_SETFOCUS, MSG_BEFORE
        .AddMessage lHwnd, WM_SIZE, MSG_BEFORE
        .AddMessage lHwnd, WM_STYLECHANGED, MSG_BEFORE
    End With
    
End Sub

Private Sub Control_Detach(ByVal lHwnd As Long)

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .DeleteMessage lHwnd, WM_ENABLE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_ERASEBKGND, MSG_BEFORE
        .DeleteMessage lHwnd, WM_KEYDOWN, MSG_BEFORE
        .DeleteMessage lHwnd, WM_KEYUP, MSG_BEFORE
        .DeleteMessage lHwnd, WM_KILLFOCUS, MSG_BEFORE
        .DeleteMessage lHwnd, WM_LBUTTONDOWN, MSG_BEFORE
        .DeleteMessage lHwnd, WM_LBUTTONUP, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSELEAVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_PAINT, MSG_BEFORE
        .DeleteMessage lHwnd, WM_SETFOCUS, MSG_BEFORE
        .DeleteMessage lHwnd, WM_SIZE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_STYLECHANGED, MSG_BEFORE
        .UnSubclass lHwnd
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

    
Dim lHdc        As Long
Dim tPAINT      As PAINTSTRUCT
Static lHndMem  As Long

    If IsWindowVisible(lHwnd) = 0 Then
        bHandled = True
        Exit Sub
    End If
    
    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            lHdc = BeginPaint(lHwnd, tPAINT)
            If IsWindowEnabled(lHwnd) = 0 Then
                Draw_Caption lHwnd, lHdc, tPAINT.rcPaint, True
            Else
                Draw_Caption lHwnd, lHdc, tPAINT.rcPaint
            End If
            If Not m_lLastSel = 0 Then
                Image_Render lHwnd, m_lLastSel
                m_lLastSel = 0
            Else
                Image_Render lHwnd, Control_State(lHwnd, True)
            End If
            EndPaint lHwnd, tPAINT
            lReturn = 1
            m_bPainting = False
            bHandled = True
        Else
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
            bHandled = True
        End If
    
    Case WM_SETFOCUS
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        If Not lHndMem = 0 Then
            SendMessageLong lHwnd, WM_PAINT, 0&, 0&
            SendMessageLong lHndMem, WM_PAINT, 0&, 0&
        Else
            Redraw_All
        End If
        bHandled = True
    
    Case WM_KILLFOCUS
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&
        bHandled = True
        
    Case WM_ERASEBKGND
        Control_Redraw lHwnd
        lReturn = 1
        bHandled = True
    
    Case WM_LBUTTONUP
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        If Not lHndMem = 0 Then
            SendMessageLong lHwnd, WM_PAINT, 0&, 0&
            SendMessageLong lHndMem, WM_PAINT, 0&, 0&
        Else
            Redraw_All
        End If
        bHandled = True
    
    Case WM_LBUTTONDOWN
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&
        bHandled = True
        
    Case WM_MOUSELEAVE
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&
        bHandled = True
        
    Case WM_MOUSEMOVE
        If SendMessageLong(lHwnd, BM_GETSTATE, 0&, 0&) = 0 Then
            m_lLastSel = 2
        Else
            m_lLastSel = 6
        End If
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&
        Track_Mouse lHwnd
        bHandled = True
    
    Case WM_STYLECHANGED, WM_SIZE
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&
        lReturn = 1
        bHandled = True
    
    Case WM_ENABLE
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&
        lReturn = 1
        bHandled = True
    
    Case Else
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        SendMessageLong lHwnd, WM_PAINT, 0&, 0&
        bHandled = True
    End Select
    
End Sub

Private Function Control_State(ByVal lHwnd As Long, _
                               Optional ByVal bNoFocus As Boolean) As Long

'/* 0-[empty] normal, 1-focused, 2-pressed, 3-disabled, 4-[value] normal, 5-focused, 6-pressed, 7-disabled
'SendMessageLong(m_hWnd, BM_GETCHECK, 0&, 0&)

Dim lState      As Long
Dim lValue      As Long
Dim bEnabled    As Boolean

    lState = SendMessageLong(lHwnd, BM_GETSTATE, 0&, 0&)
    lValue = SendMessageLong(lHwnd, BM_GETCHECK, 0&, 0&)
    bEnabled = IsWindowEnabled(lHwnd)
    
    If bNoFocus Then
        If lValue = 0 Then
            If bEnabled Then
                Control_State = 0
            Else
                Control_State = 3
            End If
            Exit Function
        ElseIf lValue = 1 Then
            If bEnabled Then
                Control_State = 4
            Else
                Control_State = 7
            End If
            Exit Function
        End If
    End If
    
    Select Case lState
    '/* empty hover
    Case 0
        Control_State = 1
    '/* lost focus
    Case 1
        Control_State = 0
    '/* got focus
    Case 8
        Control_State = 1
    '/* value hover
    Case 9
        Control_State = 5
    '/* empty press abort
    Case 104
        Control_State = 1
    '/* value press abort
    Case 105
        Control_State = 5
    '/* empty press
    Case 108
        Control_State = 2
    '/* value press
    Case 109
        Control_State = 6
    '/* got focus
    Case 120
        Control_State = 1
    Case Else
        Control_State = 4
    End Select

End Function

Public Sub Cleanup()

On Error Resume Next

    m_bPreUnload = True
    If Not m_lParentHwnd = 0 Then
        Control_Release
        Backbuffer_Destoy
        If Not m_cOptDc Is Nothing Then Set m_cOptDc = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not p_IOptImg Is Nothing Then Set p_IOptImg = Nothing
        If Not c_InstTrack Is Nothing Then Set c_InstTrack = Nothing
        If Not m_cCtrlPtr Is Nothing Then Set m_cCtrlPtr = Nothing
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

