VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMWaterMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/* adapted from vbaccelerators textbox image demo

Implements WinSubHookGX.GXISubclass

Private Const WM_PRINT                          As Long = &H317
Private Const EN_HSCROLL                        As Long = &H601
Private Const EN_VSCROLL                        As Long = &H602
Private Const PRF_CHECKVISIBLE                  As Long = &H1
Private Const PRF_CLIENT                        As Long = &H4


Public Enum WMK_POSITION
    WMK_TOPLEFT = &H1
    WMK_TOPRIGHT = &H2
    WMK_BOTTOMLEFT = &H3
    WMK_BOTTOMRIGHT = &H4
    WMK_CENTER = &H5
    WMK_TILE = &H6
End Enum


Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type


Private Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT, _
                                                     ByVal bErase As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bCE                               As Boolean
Private m_bFixBackground                    As Boolean
Private m_lTextBoxObj                       As Long
Private m_hWndParent                        As Long
Private m_iScroll                           As Long
Private m_xTileOffset                       As Long
Private m_yTileOffset                       As Long
Private m_lCtrlHnd                          As Long
Private m_eWtrMrkPos                        As WMK_POSITION
Private m_tRect                             As RECT
Private m_pWaterMark                        As StdPicture
Private m_cTempDc                           As clsStoreDc
Private m_cBgDC                             As clsStoreDc
Private m_cRender                           As clsRender
Private m_GSubclass                         As GXMSubclass


Private Sub Class_Initialize()
    InitCommonControls
    m_bFixBackground = True
    m_eWtrMrkPos = WMK_CENTER
End Sub

Public Property Get p_OTextBoxObj() As Object

Dim oTemp As Object

    If m_lTextBoxObj = 0 Then
        Set p_OTextBoxObj = Nothing
    Else
        CopyMemory oTemp, m_lTextBoxObj, 4
        Set p_OTextBoxObj = oTemp
        CopyMemory oTemp, 0&, 4
    End If

End Property

Public Property Set p_OTextBoxObj(ByVal PropVal As Object)
    m_lTextBoxObj = ObjPtr(PropVal)
End Property

Public Property Get TileOffsetX() As Long
    TileOffsetX = m_xTileOffset
End Property

Public Property Let TileOffsetX(ByVal xOffset As Long)
    m_xTileOffset = xOffset
End Property

Public Property Get TileOffsetY() As Long
    TileOffsetY = m_yTileOffset
End Property

Public Property Let TileOffsetY(ByVal yOffset As Long)
    m_yTileOffset = yOffset
End Property

Public Property Get p_IWaterMark() As StdPicture
'/* watermark image
    Set p_IWaterMark = m_pWaterMark
End Property

Public Property Set p_IWaterMark(PropVal As StdPicture)
    Set m_pWaterMark = PropVal
End Property

Public Property Get p_WaterMarkPosition() As WMK_POSITION
'/* frame style
    p_WaterMarkPosition = m_eWtrMrkPos
End Property

Public Property Let p_WaterMarkPosition(PropVal As WMK_POSITION)
    m_eWtrMrkPos = PropVal
End Property


Public Sub Attach()

    If p_IWaterMark Is Nothing Then Exit Sub
    m_lCtrlHnd = p_OTextBoxObj.hwnd
    If m_lCtrlHnd = 0 Then Exit Sub
    
    Set m_GSubclass = New GXMSubclass
    Set m_cTempDc = New clsStoreDc
    m_cTempDc.UseAlpha = True
    Set m_cBgDC = New clsStoreDc
    m_cBgDC.UseAlpha = True
    
    Set m_cRender = New clsRender
    If Not SetBackdrop Then
        Cleanup
    End If
    Attach_Control m_lCtrlHnd

End Sub

Private Sub PostPaint(Optional ByVal bUseArea As Boolean = False)

Dim lHdc            As Long
Dim lhBrush         As Long
Dim lOffsetX        As Long
Dim lOffsetY        As Long
Dim tRect           As RECT
Dim tBg             As RECT

    If m_bFixBackground Then
        bUseArea = False
    End If

    '/* control size
    GetClientRect m_lCtrlHnd, tRect

    '/* size image
    With tRect
        m_cTempDc.Width = .Right - .left + 1
        m_cTempDc.Height = .Bottom - .top + 1
    End With

    ' /* Copy background
    Select Case m_eWtrMrkPos
        '/* use tiling
    Case WMK_TILE
        '/* render image
        If Not m_bFixBackground Then
            With m_tRect
                m_cRender.Tile m_cTempDc.hdc, .left, .top, .Right - .left + 1, .Bottom - .top + 1, m_cBgDC.hdc, m_cBgDC.Width, m_cBgDC.Height
            End With
        Else
            With m_cTempDc
                m_cRender.Tile .hdc, 0, 0, .Width, .Height, m_cBgDC.hdc, m_cBgDC.Width, m_cBgDC.Height
            End With
        End If
        With m_cTempDc
            m_cRender.Blit .hdc, 0, 0, .Width, .Height, m_cBgDC.hdc, 0, 0, SRCCOPY
        End With
        '/* redraw textbox
        If Not bUseArea Then
            SendMessageLong m_lCtrlHnd, WM_PRINT, m_cTempDc.hdc, PRF_CLIENT Or PRF_CHECKVISIBLE
        End If
        '/* place image in control
        lHdc = GetDC(m_lCtrlHnd)
        If Not m_bFixBackground Then
            With m_tRect
                m_cRender.Blit lHdc, .left, .top, .Right - .left + 1, .Bottom - .top + 1, m_cTempDc.hdc, .left, .top, SRCCOPY
            End With
        Else
            With tRect
                m_cRender.Blit lHdc, 0, 0, .Right - .left + 1, .Bottom - .top + 1, m_cTempDc.hdc, 0, 0, SRCCOPY
            End With
        End If
    
    Case Else
        With tRect
            '/* build offsets
            If m_eWtrMrkPos = WMK_TOPLEFT Then
                lOffsetX = 0
                lOffsetY = 0
            ElseIf m_eWtrMrkPos = WMK_TOPRIGHT Then
                lOffsetX = .Right - m_cBgDC.Width
            ElseIf m_eWtrMrkPos = WMK_BOTTOMLEFT Then
                lOffsetY = .Bottom - m_cBgDC.Height
            ElseIf m_eWtrMrkPos = WMK_BOTTOMRIGHT Then
                lOffsetX = .Right - m_cBgDC.Width
                lOffsetY = .Bottom - m_cBgDC.Height
            ElseIf m_eWtrMrkPos = WMK_CENTER Then
                lOffsetX = (.Right - m_cBgDC.Width) / 2
                lOffsetY = (.Bottom - m_cBgDC.Height) / 2
            End If
        End With
        '/* render image
        If Not m_bFixBackground Then
            With m_tRect
                m_cRender.Blit m_cTempDc.hdc, .left, .top, .Right - .left + 1, .Bottom - .top + 1, m_cBgDC.hdc, m_cBgDC.Width, m_cBgDC.Height, SRCPAINT
            End With
        Else
            With m_cTempDc
                m_cRender.Blit .hdc, 0, 0, .Width, .Height, m_cBgDC.hdc, m_cBgDC.Width, m_cBgDC.Height, SRCCOPY
            End With
        End If
        '/* create the white bg
        With tBg
            .Bottom = m_cTempDc.Height
            .Right = m_cTempDc.Width
        End With
        '/* paint it
        lhBrush = CreateSolidBrush(&HFFFFFF)
        FillRect m_cTempDc.hdc, tBg, lhBrush
        DeleteObject lhBrush
        '/* place image in control
        lHdc = GetDC(m_lCtrlHnd)
        With m_cTempDc
            m_cRender.Blit .hdc, lOffsetX, lOffsetY, .Width, .Height, m_cBgDC.hdc, 0, 0, SRCCOPY
        End With
        '/* redraw textbox
        If Not bUseArea Then
            SendMessageLong m_lCtrlHnd, WM_PRINT, m_cTempDc.hdc, PRF_CLIENT Or PRF_CHECKVISIBLE
        End If
        '/* draw the watermark
        If Not m_bFixBackground Then
            With m_tRect
                m_cRender.Blit lHdc, .left, .top, .Right - .left + 1, .Bottom - .top + 1, m_cTempDc.hdc, .left, .top, SRCCOPY
            End With
        Else
            With tRect
                m_cRender.Blit lHdc, 0, 0, .Right - .left + 1, .Bottom - .top + 1, m_cTempDc.hdc, 0, 0, SRCCOPY
            End With
        End If
    End Select

    ReleaseDC m_lCtrlHnd, lHdc

End Sub

Private Function SetBackdrop() As Boolean

Dim cTempDC As New clsStoreDc

    '/* create the image dc
    If Not m_pWaterMark Is Nothing Then
        With cTempDC
            .UseAlpha = True
            .InjectBitmap m_pWaterMark.handle
            m_cBgDC.Width = .Width
            m_cBgDC.Height = .Height
            m_cRender.Blit m_cBgDC.hdc, 0, 0, m_cBgDC.Width, m_cBgDC.Height, .hdc, 0, 0, SRCCOPY
            .ExtractBitmap
        End With
        SetBackdrop = True
    End If
    
End Function

Private Sub Attach_Control(ByVal lHwnd As Long)

    With m_GSubclass
        If Not m_lCtrlHnd = 0 Then
            .Subclass m_lCtrlHnd, Me
            .AddMessage m_lCtrlHnd, WM_ERASEBKGND, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_PAINT, MSG_BEFORE
        End If
        m_hWndParent = GetParent(m_lCtrlHnd)
        If Not m_hWndParent = 0 Then
            .Subclass m_hWndParent, Me
            .AddMessage m_hWndParent, WM_CTLCOLOREDIT, MSG_BEFORE
            .AddMessage m_hWndParent, WM_COMMAND, MSG_BEFORE
        End If
    End With

End Sub

Private Sub Detach_Control()

    With m_GSubclass
        If Not m_lCtrlHnd = 0 Then
            .DeleteMessage m_lCtrlHnd, WM_ERASEBKGND, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_PAINT, MSG_BEFORE
            .UnSubclass m_lCtrlHnd
            m_lCtrlHnd = 0
        End If
        If Not m_hWndParent = 0 Then
            .DeleteMessage m_hWndParent, WM_CTLCOLOREDIT, MSG_BEFORE
            .DeleteMessage m_hWndParent, WM_COMMAND, MSG_BEFORE
            .UnSubclass m_hWndParent
            m_hWndParent = 0
        End If
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

Dim wCode       As Long

    If IsWindowVisible(lHwnd) = 0 Then
        lReturn = 0
        bHandled = True
        Exit Sub
    End If
        
    Select Case uMsg
    Case WM_PAINT
        GetUpdateRect lHwnd, m_tRect, 0
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, lParam, wParam)
        If m_iScroll = 0 Then
            PostPaint
        Else
            Select Case m_iScroll
            Case 1
                m_iScroll = 1
            Case 2
                m_iScroll = 0
            End Select
        End If

    Case WM_CTLCOLOREDIT
        If m_lCtrlHnd = lParam Then
            SetBkMode wParam, 1
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, lParam, wParam)
            If m_iScroll = 0 Then
                If Not m_bCE Then
                    m_bCE = True
                    PostPaint
                    m_bCE = False
                End If
            Else
                Select Case m_iScroll
                Case 1
                    m_iScroll = 2
                    If Not m_bCE Then
                        m_bCE = True
                        PostPaint True
                        m_bCE = False
                    End If
                Case 2
                    m_iScroll = 0
                End Select
            End If
        Else
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, lParam, wParam)
        End If

    Case WM_COMMAND
        If m_lCtrlHnd = lParam Then
            wCode = wParam \ &H10000
            If wCode = EN_HSCROLL Or wCode = EN_VSCROLL Then
                m_iScroll = 1
            End If
        End If
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, lParam, wParam)

    Case WM_ERASEBKGND
        lReturn = 1
    End Select

End Sub

Private Sub Cleanup()

    If Not m_lCtrlHnd = 0 Then
        Detach_Control
        If Not m_cTempDc Is Nothing Then Set m_cTempDc = Nothing
        If Not m_cBgDC Is Nothing Then Set m_cBgDC = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_pWaterMark Is Nothing Then Set m_pWaterMark = Nothing
        If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
        If Not p_OTextBoxObj Is Nothing Then Set p_OTextBoxObj = Nothing
    End If
    m_lCtrlHnd = 0

End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

