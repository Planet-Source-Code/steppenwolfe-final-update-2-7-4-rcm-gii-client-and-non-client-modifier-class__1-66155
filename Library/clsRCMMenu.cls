VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCMMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'************************************************************************
'*                                                                      *
'* Name:                clsRCMMenu.cls                                  *
'* Original Author:     Steve McMahon steve@vbaccelerator.com           *
'* Original Control:    vbAccelerator IconMenu Component                *
'* Copyright © 2000-2003 Steve McMahon (steve@vbaccelerator.com)        *
'* Original Date:       22 December 1998                                *
'* Rewritten by:        John Underhill (Steppenwolfe), May 4, 2006      *
'*                                                                      *
'************************************************************************


'/~ My Notes (John)
'/~ Intro
'/~ Steve McMahons icon menu control, as rewritten in part by me..
'/~ I changed the class name as per Steve's wishes in his open source licence agreement, (if you
'/~ make serious modifications to his code, he wants you to rename it..). Also so that class and control
'/~ registration would not interfere with applications currently using the original menu control.
'/~ Many small changes made like adding ($, B, Not, ByVal, Chr$, +), added/changed code to Not, With, and
'/~ Exit Sub in many spots. Removed empty routines, shortened many others, complete formatting of code.

'/~ August 11, 2006
'/~ Had to make many small changes to adapt this to RCM. Added transparent menu support and image rollovers.

Implements GXISubclass

Private Const MFT_STRING                            As Long = &H0
' New versions of the names...
Private Const MFS_GRAYED                            As Long = &H3
Private Const MFS_DISABLED                          As Long = MFS_GRAYED
Private Const MIIM_ID                               As Long = &H2
Private Const MIIM_SUBMENU                          As Long = &H4
Private Const MIIM_TYPE                             As Long = &H10
Private Const MIIM_DATA                             As Long = &H20
Private Const MIIM_STRING                           As Long = &H40
Private Const BITSPIXEL                             As Long = 12
Private Const PS_SOLID                              As Long = 0
Private Const DT_CENTER                             As Long = &H1
Private Const DT_CALCRECT                           As Long = &H400
Private Const DT_SINGLELINE                         As Long = &H20
Private Const ST_TRANSPARENT                        As Long = 1
' DrawEdge:
Private Const BDR_SUNKENOUTER                       As Long = &H2
Private Const BDR_RAISEDINNER                       As Long = &H4
Private Const EDGE_ETCHED                           As Double = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const BDR_CUSTOMDRAW                        As Long = &H16
Private Const BF_LEFT                               As Long = &H1
Private Const BF_TOP                                As Long = &H2
Private Const BF_RIGHT                              As Long = &H4
Private Const BF_BOTTOM                             As Long = &H8
Private Const BF_RECT                               As Double = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const CLR_INVALID                           As Long = -1
Private Const WS_EX_LAYERED                         As Long = &H80000
Private Const LWA_ALPHA                             As Long = &H2&

Private Enum WC_EXTENDEDSTYLE
    GWL_EXSTYLE = -20
    GWL_STYLE = -16
    GWL_WNDPROC = -4
    GWL_HINSTANCE = -6
    GWL_HWNDPARENT = -8
    GWL_ID = -12
    GWL_USERDATA = -21
End Enum

Private Enum SWP_ENUM
    SWP_DRAWFRAME = 32
    SWP_FRAMECHANGED = 32
    SWP_HIDEWINDOW = 128
    SWP_NOACTIVATE = 16
    SWP_NOCOPYBITS = 256
    SWP_NOMOVE = 2
    SWP_NOSIZE = 1
    SWP_NOREDRAW = 8
    SWP_NOZORDER = 4
    SWP_SHOWWINDOW = 64
    SWP_NOOWNERZORDER = 512
    SWP_NOREPOSITION = 512
    SWP_NOSENDCHANGING = 1024
    SWP_DEFERERASE = 8192
    SWP_ASYNCWINDOWPOS = 16384
End Enum

Private Enum SWP_WNDMSG
    HWND_BROADCAST = 65535
    HWND_BOTTOM = 1
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
    HWND_DESKTOP = 0
End Enum

Public Enum ECHIGHLIGHTCONST
    ECSTANDARD = 0
    ECGRADIENT = 1
    ECBUTTON = 2
    ECROLLOVER = 3
End Enum

Private Type MEASUREITEMSTRUCT
    CtlType                                         As Long
    CtlID                                           As Long
    itemID                                          As Long
    itemWidth                                       As Long
    itemHeight                                      As Long
    itemData                                        As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType                                         As Long
    CtlID                                           As Long
    itemID                                          As Long
    itemAction                                      As Long
    itemState                                       As Long
    hwndItem                                        As Long
    hdc                                             As Long
    rcItem                                          As RECT
    itemData                                        As Long
End Type

Private Type MENUITEMINFO
    cbSize                                          As Long
    fMask                                           As Long
    fType                                           As Long
    fState                                          As Long
    wID                                             As Long
    hSubMenu                                        As Long
    hbmpChecked                                     As Long
    hbmpUnchecked                                   As Long
    dwItemData                                      As Long
    dwTypeData                                      As Long
    cch                                             As Long
End Type

Private Type VERSION_INFO
    dwOSVersionInfoSize                             As Long
    dwMajorVersion                                  As Long
    dwMinorVersion                                  As Long
    dwBuildNumber                                   As Long
    dwPlatformId                                    As Long
    szCSDVersion                                    As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As VERSION_INFO) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, _
                                                                  ByVal crey As Byte, _
                                                                  ByVal bAlpha As Byte, _
                                                                  ByVal dwFlags As Long) As Long

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, _
                                                     ByVal nPos As Long) As Long

Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
                                                                                ByVal uItem As Long, _
                                                                                ByVal fByPosition As Boolean, _
                                                                                lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, _
                                                                                ByVal un As Long, _
                                                                                ByVal bool As Boolean, _
                                                                                lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function DrawEdgeApi Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, _
                                                                    qrc As RECT, _
                                                                    ByVal edge As Long, _
                                                                    ByVal grfFlags As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                       lpDeviceName As Any, _
                                                                       lpOutput As Any, _
                                                                       lpInitData As Any) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal x As Long, _
                                                ByVal y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
                                                  ByVal nPos As Long) As Long

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" (ByVal hIml As Long, _
                                                                    ByVal i As Long, _
                                                                    prcImage As RECT) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, _
                                                            ByVal i As Long, _
                                                            ByVal hdcDst As Long, _
                                                            ByVal x As Long, _
                                                            ByVal y As Long, _
                                                            ByVal fStyle As Long) As Long

Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIml As Long, _
                                                               ByVal i As Long, _
                                                               ByVal diIgnore As Long) As Long

Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImageList As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImageList As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

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

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long


Public Event Repaint()
Private m_bUseClient                                As Boolean
Private m_bUseIcons                                 As Boolean
Private m_hMenu                                     As Long
Private m_hWnd                                      As Long
Private m_iCount                                    As Long
Private m_iDownOn                                   As Long
Private m_iOver                                     As Long
Private m_hSubMenu()                                As Long
Private m_lCaptionHeight                            As Long
Private m_iRestore                                  As Long
Private m_hMenuRestore()                            As Long
Private m_iMenuPosition()                           As Long
Private m_lMenuTextSize()                           As Long
Private m_lMenuShortCutSize()                       As Long
Private m_iHaveSeenCount                            As Long
Private m_hMenuSeen()                               As Long
Private m_lMenuItemHeight                           As Long
Private m_sTopLevelAccelerator()                    As String
Private m_sCaption()                                As String
Private m_sShortCut()                               As String
Private m_sAccelerator()                            As String
Private m_cMemDC                                    As clsStoreDc
Private m_tR()                                      As RECT
Private m_oActiveMenuColor                          As OLE_COLOR
Private m_oActiveMenuColorOver                      As OLE_COLOR
Private m_oInActiveMenuColor                        As OLE_COLOR
Private m_oMenuBackgroundColor                      As OLE_COLOR
Private m_tMIIS()                                   As MENUITEMINFO
Private m_fnt                                       As StdFont
Private m_fntSymbol                                 As StdFont
Private WithEvents m_cTmr                           As clsTimer
Attribute m_cTmr.VB_VarHelpID = -1
Private m_cMenu                                     As Object
Private m_MenuSubclass                              As GXMSubclass
Private m_eHighlightStyle                           As ECHIGHLIGHTCONST
'/*** skinned menus
Private m_cMenuBg                                   As clsStoreDc
Private m_cMenuBar                                  As clsStoreDc
Private m_pMenuBgPicture                            As IPicture
Private m_pMenuBarPicture                           As IPicture
Private m_bMenuBar                                  As Boolean
Private m_bMenuBg                                   As Boolean
'/*** icons
Private m_colIcons                                  As Collection
Private m_lIconSize                                 As Long
Private m_hIml                                      As Long
Private m_ptrVb6ImageList                           As Long
Private m_lMenuVPtr                                 As Long
Private m_lMenuOffsetY                              As Long
Private m_oEdgeColorDark                            As OLE_COLOR
Private m_oEdgeColorLight                           As OLE_COLOR
Private m_bCustomRollOver                           As Boolean
'/*** added
Private m_cNCM                                      As clsMetrics
Private m_cBrush                                    As clsBrush
Private m_OfficeXPStyle                             As Boolean
Private m_cToolbarMenu                              As clsToolbarMenu
Private m_lFontSymbol                               As Long
Private m_bTransparent                              As Boolean
Private m_btTransIndex                              As Byte
Private m_lMenuRollOverColor                        As Long
Private m_bWin32                                    As Boolean
Private m_cRender                                   As clsRender


Private Sub Class_Initialize()

    Set m_cMemDC = New clsStoreDc
    Set m_MenuSubclass = New GXMSubclass
    Set m_fnt = New StdFont
    m_fnt.Name = "MS Sans Serif"
    Set m_fntSymbol = New StdFont
    m_fntSymbol.Name = "Marlett"
    m_fntSymbol.Size = m_fnt.Size * 1.2
    StoreColors
    Set m_cNCM = New clsMetrics
    m_cNCM.GetMetrics
    m_cMemDC.Width = Screen.Width \ Screen.TwipsPerPixelY
    m_cMemDC.Height = 24
    Set m_fntSymbol = New StdFont
    m_fntSymbol.Name = "Marlett"
    m_fntSymbol.Size = p_MenuFont.Size * 1.2
    Set m_cBrush = New clsBrush
    m_cBrush.Create
    m_oMenuBackgroundColor = -1
    m_lMenuRollOverColor = -1
    Set m_cRender = New clsRender
    m_bWin32 = Compatability_Check

End Sub

Private Function Compatability_Check() As Boolean

Dim tVer  As VERSION_INFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    If tVer.dwMajorVersion >= 5 Then
        Compatability_Check = True
    End If

End Function

Private Sub m_cTmr_ThatTime()
'/* timer event
    MouseMove
End Sub


Private Sub TransparentMenu(ByVal lHwnd As Long, _
                            ByVal lTrsIdx As Integer)

    If Not m_bWin32 Then Exit Sub
    If lTrsIdx > 255 Or lTrsIdx < 128 Then
        lTrsIdx = 220
    End If
    If Not TransState(lHwnd) = True Then
        SetWindowLong lHwnd, GWL_EXSTYLE, GetWindowLong(lHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    End If
    SetLayeredWindowAttributes lHwnd, 0&, lTrsIdx, LWA_ALPHA
    
End Sub

Private Property Get TransState(ByVal lHwnd As Long) As Boolean

Dim lRet    As Long

    If Not m_bWin32 Then Exit Property
    lRet = GetWindowLong(lHwnd, GWL_EXSTYLE)
    If (lRet And WS_EX_LAYERED) = WS_EX_LAYERED Then
        TransState = True
    Else
        TransState = False
    End If
    
End Property

Public Property Let p_ActiveMenuForeColor(ByVal PropVal As OLE_COLOR)
    m_oActiveMenuColor = PropVal
End Property

Public Property Get p_ActiveMenuForeColor() As OLE_COLOR

    If m_oActiveMenuColor = -1 Then
        If m_OfficeXPStyle Then
            p_ActiveMenuForeColor = vbMenuText
        Else
            p_ActiveMenuForeColor = vbHighlightText
        End If
    Else
        p_ActiveMenuForeColor = m_oActiveMenuColor
    End If
   
End Property

Public Property Let p_InActiveMenuForeColor(ByVal PropVal As OLE_COLOR)
    m_oInActiveMenuColor = PropVal
End Property

Private Property Get p_BlendColor(ByVal oColorFrom As OLE_COLOR, _
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
     
   
    p_BlendColor = RGB(((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), _
        ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), _
        ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255))
      
End Property

Public Property Get p_InActiveMenuForeColor() As OLE_COLOR

    If m_oInActiveMenuColor = -1 Then
        p_InActiveMenuForeColor = vbMenuText
    Else
        p_InActiveMenuForeColor = m_oInActiveMenuColor
    End If
   
End Property


Public Property Let p_CaptionHeight(ByVal lHeight As Long)
'/* caption height
    m_lCaptionHeight = lHeight
End Property

Public Property Get p_CustomRollOver() As Boolean
    p_CustomRollOver = m_bCustomRollOver
End Property

Public Property Let p_CustomRollOver(PropVal As Boolean)
    m_bCustomRollOver = PropVal
End Property

Public Property Get p_RollOverColor() As Long
'/* custom caption rollover effect
    p_RollOverColor = m_lMenuRollOverColor
End Property

Public Property Let p_RollOverColor(ByVal PropVal As Long)
    m_lMenuRollOverColor = PropVal
End Property

Friend Property Get p_MenuCount() As Long
'/* top menu count
    p_MenuCount = m_iCount
End Property

Friend Property Get p_MenuFont() As StdFont
'/* menu font
    Set p_MenuFont = m_fnt
End Property

Friend Property Let p_MenuFont(PropVal As StdFont)
    Set m_fnt = PropVal
End Property

Friend Property Set p_MenuFont(PropVal As StdFont)

    Set m_fnt = PropVal
    m_fntSymbol.Name = "Marlett"
    m_fntSymbol.Size = m_fnt.Size * 1.2
    
End Property

Private Property Get p_FontHandle() As Long
'/* font handle
Dim iFn     As IFont

    Set iFn = m_fnt
    p_FontHandle = iFn.hFont

End Property

Private Property Get p_FontSymbol() As Long

Dim iFn         As IFont

    Set iFn = m_fntSymbol
    m_lFontSymbol = iFn.hFont
   
End Property

Public Property Get p_HighlightStyle() As ECHIGHLIGHTCONST
    
    p_HighlightStyle = m_eHighlightStyle
   
End Property

Public Property Let p_HighlightStyle(ByVal eStyle As ECHIGHLIGHTCONST)

    m_eHighlightStyle = eStyle
   
End Property

Public Property Get p_ImageListHndl() As Long
'/* iml handle
    p_ImageListHndl = m_hIml
End Property

Public Property Get p_ImageListPtr() As Long
'/* ptr to iml
    p_ImageListPtr = m_ptrVb6ImageList
End Property

Public Property Get p_IconIndex(ByVal sCaption As String) As Long
    p_IconIndex = m_colIcons.Item(sCaption)
End Property

Public Property Get p_IconIndexForIndex(ByVal lItem As Long) As Long
    p_IconIndexForIndex = m_colIcons.Item(lItem)
End Property

Public Property Set p_IconIndexSet(PropVal As Collection)
'/* iml index

    If Not PropVal Is Nothing Then
        Set m_colIcons = PropVal
        If m_colIcons.Count < 1 Then
            m_bUseIcons = False
        End If
    End If

End Property

Public Property Get p_IconItemCount() As Long

    p_IconItemCount = m_colIcons.Count
   
End Property

Public Property Set p_ImageList(ByRef vImageList As Variant)

Dim o   As Object
Dim rc  As RECT

    m_hIml = 0
    m_ptrVb6ImageList = 0
    If VarType(vImageList) = vbLong Then
        '/* Assume a handle to an image list:
        m_hIml = vImageList
    ElseIf (VarType(vImageList) = vbObject) Then
        '/* Assume a VB image list:
        On Error Resume Next
        '/* Get the image list initialised
        vImageList.ListImages(1).Draw 0&, 0&, 0&, 1
        m_hIml = vImageList.hImageList
        If Err.Number = 0 Then
            '/* Check for VB6 image list:
            If TypeName(vImageList) = "ImageList" Then
                Set o = vImageList
                m_ptrVb6ImageList = ObjPtr(o)
            End If
        End If
        On Error GoTo 0
    End If

    If Not m_hIml = 0 Then
        If Not m_ptrVb6ImageList = 0 Then
            m_lIconSize = vImageList.ImageHeight
            m_bUseIcons = True
        Else
            ImageList_GetImageRect m_hIml, 0&, rc
            m_lIconSize = rc.Bottom - rc.top
            m_bUseIcons = True
        End If
    End If

End Property

Private Property Get p_LighterColour(ByVal oColor As OLE_COLOR) As Long

Dim lC          As Long
Dim lR          As Long
Dim lG          As Long
Dim lB          As Long
Dim h           As Single
Dim s           As Single
Dim l           As Single
Static slCol    As Long
Static slLight  As Long

    lC = TranslateColor(oColor)
    If Not lC = slCol Then
        slCol = lC
        RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
        If (l > 0.99) Then
            l = l * 0.8
        Else
            l = l * 1.2
            If (l > 1) Then
                l = 1
            End If
        End If
        HLSToRGB h, s, l, lR, lG, lB
        slLight = RGB(lR, lG, lB)
    End If
    p_LighterColour = slLight
   
End Property

Public Property Let p_MenuBackgroundColor(ByVal PropVal As OLE_COLOR)

    m_oMenuBackgroundColor = PropVal
   
End Property

Public Property Get p_MenuBackgroundColor() As OLE_COLOR

    If m_oMenuBackgroundColor = -1 Then
        p_MenuBackgroundColor = &H333333
    Else
        p_MenuBackgroundColor = m_oMenuBackgroundColor
    End If
   
End Property

Public Property Get p_MenuBarPicture() As IPicture
'/* rollover bar pic
    Set p_MenuBarPicture = m_pMenuBarPicture
End Property

Public Property Set p_MenuBarPicture(ByRef PropVal As IPicture)

    If Not PropVal Is Nothing Then
        Set m_pMenuBarPicture = PropVal
        Set m_cMenuBar = New clsStoreDc
        m_cMenuBar.CreateFromPicture PropVal
        m_bMenuBar = True
    Else
        Set m_cMenuBar = Nothing
        m_bMenuBar = False
    End If

End Property

Public Property Get p_MenuBgPicture() As IPicture
'/* menu bg pic
    Set p_MenuBgPicture = m_pMenuBgPicture
End Property

Public Property Set p_MenuBgPicture(ByRef PropVal As IPicture)

    If Not PropVal Is Nothing Then
        Set m_pMenuBgPicture = PropVal
        Set m_cMenuBg = New clsStoreDc
        m_cMenuBg.CreateFromPicture PropVal
        m_bMenuBg = True
    Else
        Set m_cMenuBg = Nothing
        m_bMenuBg = False
    End If

End Property

Public Property Get p_MenuOffsetY() As Long
'/* menu draw offset
    p_MenuOffsetY = m_lMenuOffsetY
End Property

Public Property Let p_MenuOffsetY(ByVal PropVal As Long)
    m_lMenuOffsetY = PropVal
End Property

Public Property Get p_MenuHandle() As Long
'/* menu handle
    p_MenuHandle = m_hMenu
End Property

Public Property Let p_MenuHandle(ByVal PropVal As Long)
    m_hMenu = PropVal
End Property

Friend Property Get HotItem() As Long
    HotItem = m_iDownOn
End Property

Friend Property Let HotItem(ByVal iHotItem As Long)
' Set the hotitem
    m_iOver = iHotItem
    RepaintOwner
End Property

Public Property Get p_OfficeXpStyle() As Boolean
    p_OfficeXpStyle = m_OfficeXPStyle
End Property

Public Property Let p_OfficeXpStyle(ByVal PropVal As Boolean)
    m_OfficeXPStyle = PropVal
End Property

Private Property Get p_NoPalette(Optional ByVal bForce As Boolean = False) As Boolean

Static bOnce      As Boolean
Static bNoPalette As Boolean
Dim lHdc          As Long
Dim lBits         As Long

    If bForce Then
        bOnce = False
    End If
    If Not bOnce Then
        lHdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        If lHdc <> 0 Then
            lBits = GetDeviceCaps(lHdc, BITSPIXEL)
            If lBits <> 0 Then
                bOnce = True
            End If
            bNoPalette = (lBits > 8)
            DeleteDC lHdc
        End If
    End If
    p_NoPalette = bNoPalette

End Property

Public Property Let p_PopupMenuObject(PropVal As Variant)
    Set m_cMenu = PropVal
End Property

Private Property Get p_SlightlyLighterColour(ByVal oColor As OLE_COLOR) As Long

Dim lC          As Long
Dim lR          As Long
Dim lG          As Long
Dim lB          As Long
Dim h           As Single
Dim s           As Single
Dim l           As Single
Static slCol    As Long
Static slLight  As Long

    lC = TranslateColor(oColor)
    If Not lC = slCol Then
        slCol = lC
        RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
        If (l > 0.99) Then
            l = l * 0.95
        Else
            l = l * 1.05
            If (l > 1) Then
                l = 1
            End If
        End If
        HLSToRGB h, s, l, lR, lG, lB
        slLight = RGB(lR, lG, lB)
    End If
    p_SlightlyLighterColour = slLight

End Property

Public Property Get p_Transparent() As Boolean
    p_Transparent = m_bTransparent
End Property

Public Property Let p_Transparent(PropVal As Boolean)
    m_bTransparent = PropVal
End Property

Public Property Get p_TransIndex() As Byte
    p_TransIndex = m_btTransIndex
End Property

Public Property Let p_TransIndex(PropVal As Byte)
    m_btTransIndex = PropVal
End Property

Private Sub StoreColors()

    GetEdgeColors m_oMenuBackgroundColor, m_oEdgeColorLight, m_oEdgeColorDark

End Sub

Private Sub DrawEdge(ByVal lHdc As Long, _
                     ByRef tR As RECT, _
                     ByVal borderFlags As Long, _
                     ByVal borderParts As Long, _
                     ByVal bXpStyle As Boolean)

Dim hPen            As Long
Dim hPenOld         As Long
Dim hPenLight       As Long
Dim hPenDark        As Long
Dim tJunk           As POINTAPI

On Error GoTo Handler

    If Not p_NoPalette Then
        DrawEdgeApi lHdc, tR, borderFlags, borderParts
        Exit Sub
    End If

    If bXpStyle Then
        With tR
            If (.Bottom > .top) Then
                hPen = CreatePen(0&, 1, TranslateColor(vbHighlight))
            Else
                hPen = CreatePen(0&, 1, TranslateColor(vb3DShadow))
            End If
            hPenOld = SelectObject(lHdc, hPen)
            MoveToEx lHdc, .left, .top, tJunk
            LineTo lHdc, .Right - 1, .top
            If (.Bottom > .top) Then
                LineTo lHdc, .Right - 1, .Bottom - 1
                LineTo lHdc, .left, .Bottom - 1
                LineTo lHdc, .left, .top
            End If
        End With
        SelectObject lHdc, hPenOld
        DeleteObject hPen
        hPenOld = 0
    Else
        hPenDark = CreatePen(PS_SOLID, 1, m_lMenuRollOverColor)
        hPenLight = CreatePen(PS_SOLID, 1, m_oEdgeColorLight)
        ' we only handle sunken outer & raised inner:
        If borderFlags = BDR_SUNKENOUTER Then
            If borderParts = BF_RECT Then
                hPenOld = SelectObject(lHdc, hPenDark)
                With tR
                    MoveToEx lHdc, .left, .Bottom - 1, tJunk
                    LineTo lHdc, .left, .top
                    LineTo lHdc, .Right - 1, .top
                    SelectObject lHdc, hPenOld
                    DeleteObject hPenDark
                    hPenOld = 0
                    hPenOld = SelectObject(lHdc, hPenLight)
                    LineTo lHdc, .Right - 1, .Bottom - 1
                    LineTo lHdc, .left, .Bottom - 1
                End With
                SelectObject lHdc, hPenOld
                DeleteObject hPenLight
                hPenOld = 0
            Else
                ' single line, darker colour:
                hPenOld = SelectObject(lHdc, hPenDark)
                With tR
                    MoveToEx lHdc, .left, .top, tJunk
                    LineTo lHdc, .Right, .top
                End With
                SelectObject lHdc, hPenOld
                DeleteObject hPenDark
                DeleteObject hPenLight
                hPenOld = 0
            End If
        ElseIf borderFlags = BDR_CUSTOMDRAW Then
            hPenOld = SelectObject(lHdc, hPenDark)
            With tR
                MoveToEx lHdc, .left, .top, tJunk
                LineTo lHdc, .Right, .top
            End With
            SelectObject lHdc, hPenOld
            DeleteObject hPenDark
            DeleteObject hPenLight
            hPenOld = 0
        Else
            hPenOld = SelectObject(lHdc, hPenLight)
            With tR
                MoveToEx lHdc, .left, .Bottom - 1, tJunk
                LineTo lHdc, .left, .top
                LineTo lHdc, .Right - 1, .top
                SelectObject lHdc, hPenOld
                hPenOld = SelectObject(lHdc, hPenDark)
                LineTo lHdc, .Right - 1, .Bottom - 1
                LineTo lHdc, .left, .Bottom - 1
            End With
            SelectObject lHdc, hPenOld
            DeleteObject hPenDark
            DeleteObject hPenLight
            hPenOld = 0
        End If
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub GetEdgeColors(ByVal oColor As OLE_COLOR, _
                         ByRef lighter As OLE_COLOR, _
                         ByRef darker As OLE_COLOR)

Dim h   As Single
Dim s   As Single
Dim l   As Single
Dim lT  As Single
Dim lC  As Long
Dim lR  As Long
Dim lG  As Long
Dim lB  As Long

    lC = TranslateColor(oColor)
    RGBToHLS lC And &HFF&, (lC \ &H100&) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
    s = s * 0.9
    ' If luminance is > 85%, then we need to darken the ligher colour & the darker colour:
    If l > 0.85 Then
        lT = l - 0.2
        HLSToRGB h, s, lT, lR, lG, lB
        lighter = RGB(lR, lG, lB)
        lT = lT - 0.3
        HLSToRGB h, s, lT, lR, lG, lB
        darker = RGB(lR, lG, lB)
    ' else if luminance is < 33%, then we need a lighter dark colour & lighter light colour:
    ElseIf (l < 0.2) Then
        lT = l + 0.2
        HLSToRGB h, s, lT, lR, lG, lB
        darker = RGB(lR, lG, lB)
        lT = lT + 0.2
        HLSToRGB h, s, lT, lR, lG, lB
        lighter = RGB(lR, lG, lB)
    ' else we darken for dark & lighten for light:
    Else
        lT = l - 0.2
        If lT < 0 Then
            lT = 0
        End If
        HLSToRGB h, s, lT, lR, lG, lB
        darker = RGB(lR, lG, lB)
        lT = l + 0.2
        If lT > 1 Then
            lT = 1
        End If
        HLSToRGB h, s, lT, lR, lG, lB
        lighter = RGB(lR, lG, lB)
    End If

End Sub

Friend Function GetMenuHandle(ByVal iNewPopup As Long) As Long
' Returns the popup menu handle for a given top level
' menu item (1 based index)

    If iNewPopup > 0 Then
        If iNewPopup <= m_iCount Then
            GetMenuHandle = m_hSubMenu(iNewPopup)
        End If
    End If

End Function

Friend Function HitTest(tP As POINTAPI) As Long

Dim iMenu   As Long
Dim tR      As RECT

    If m_bUseClient Then
        GetWindowRect m_hWnd, tR
        tP.x = tP.x - tR.left
        tP.y = tP.y - tR.top
    Else
        ScreenToClient m_hWnd, tP
    End If
    For iMenu = 1 To m_iCount
        If Not PtInRect(m_tR(iMenu), tP.x + 2, tP.y) = 0 Then
            HitTest = iMenu
            Exit For
        End If
    Next iMenu

End Function

Public Sub HLSToRGB(ByVal h As Single, _
                    ByVal s As Single, _
                    ByVal l As Single, _
                    ByRef r As Long, _
                    ByRef g As Long, _
                    ByRef b As Long)

Dim rR      As Single
Dim rG      As Single
Dim rB      As Single
Dim Min     As Single
Dim Max     As Single

    If s = 0 Then
        ' Achromatic case:
        rR = l
        rG = l
        rB = l
    Else 'NOT S...
        ' Chromatic case:
        ' delta = Max-Min
        If l <= 0.5 Then
            Min = l * (1 - s)
        Else
            Min = l - s * (1 - l)
        End If
        ' Get the Max value:
        Max = 2 * l - Min
        ' Now depending on sector we can evaluate the h,l,s:
        If h < 1 Then
            rR = Max
            If h < 0 Then
                rG = Min
                rB = rG - h * (Max - Min)
            Else
                rB = Min
                rG = h * (Max - Min) + rB
            End If
        ElseIf (h < 3) Then
            rG = Max
            If h < 2 Then
                rB = Min
                rR = rB - (h - 2) * (Max - Min)
            Else
                rR = Min
                rB = (h - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If h < 4 Then
                rR = Min
                rG = rR - (h - 4) * (Max - Min)
            Else
                rG = Min
                rR = (h - 4) * (Max - Min) + rG
            End If
        End If
    End If
    r = rR * 255
    g = rG * 255
    b = rB * 255

End Sub

Public Sub RGBToHLS(ByVal r As Long, _
                    ByVal g As Long, _
                    ByVal b As Long, _
                    ByRef h As Single, _
                    ByRef s As Single, _
                    ByRef l As Single)

Dim Max     As Single
Dim Min     As Single
Dim delta   As Single
Dim rR      As Single
Dim rG      As Single
Dim rB      As Single

    rR = r / 255
    rG = g / 255
    rB = b / 255

    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) / 2
    '/* saturation}
    If Max = Min Then
        s = 0
        h = 0
    Else
        '/* saturation
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If
        '/* hue
        delta = Max - Min
        If rR = Max Then
            h = (rG - rB) / delta
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta
        End If
    End If

End Sub

Private Function Maximum(ByRef rR As Single, _
                         ByRef rG As Single, _
                         ByRef rB As Single) As Single

    If rR > rG Then
        If rR > rB Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If rB > rG Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If

End Function

Private Function Minimum(ByRef rR As Single, _
                         ByRef rG As Single, _
                         ByRef rB As Single) As Single

    If rR < rG Then
        If rR < rB Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If rB < rG Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If

End Function

Private Function MenuHitTest() As Long

Dim tP  As POINTAPI

    If m_iCount > 0 Then
        GetCursorPos tP
        MenuHitTest = HitTest(tP)
    End If

End Function

Private Sub MouseMove()

Dim iMenu       As Long
Dim iLastOver   As Long

    iMenu = MenuHitTest()
    iLastOver = m_iOver
    m_iOver = iMenu
    
    If Not m_iOver = iLastOver Then
        RepaintOwner
    End If
    If m_cTmr Is Nothing Then
        Set m_cTmr = New clsTimer
    Else
        m_cTmr.Interval = 0
        Set m_cTmr = Nothing
        Set m_cTmr = New clsTimer
    End If
    If m_iOver < 1 And m_iDownOn = 0 Then
        m_cTmr.Interval = 0
    Else
        If m_iDownOn > 0 Then
            If GetAsyncKeyState(vbLeftButton) = 0 Then
                m_iDownOn = 0
                RepaintOwner
            End If
        End If
        m_cTmr.Interval = 50
    End If

End Sub

Private Sub RepaintOwner()

    If Not (m_bUseClient) Then
        SendMessageLong m_hWnd, &H85, 0&, 0&
    Else
        RaiseEvent Repaint
    End If

End Sub

Public Sub Render(ByVal hFnt As Long, _
                  ByVal lHdc As Long, _
                  ByVal lLeft As Long, _
                  ByVal lTop As Long, _
                  ByVal lWidth As Long, _
                  ByVal lHeight As Long, _
                  ByVal lYoffset As Long, _
                  ByVal bActiveWin As Boolean, _
                  ByRef lActualRight As Long)

Dim bProcess    As Boolean
Dim bDraw       As Boolean
Dim bPress      As Boolean
Dim iIdx        As Long
Dim lC          As Long
Dim lhDCC       As Long
Dim hFntOld     As Long
Dim lX          As Long
Dim lID         As Long
Dim lMaxBottom  As Long
Dim lIndex      As Long
Dim iPos        As Long
Dim hBrush      As Long
Dim sCap        As String
Dim tTR         As RECT
Dim tBR         As RECT
Dim tMII        As MENUITEMINFO
Dim tMIIL       As MENUITEMINFO
Dim tWRS        As RECT

On Error Resume Next

    bProcess = True
    If m_hMenu = 0 Then
        If (m_cMenu Is Nothing) Then
            bProcess = False
        End If
    End If
    If bProcess Then
        With m_cMemDC
            .Width = lWidth
            .Height = lHeight
            lhDCC = .hdc
        End With
        hFntOld = SelectObject(lhDCC, hFnt)
        m_iCount = 0
        Erase m_tR
        If m_hMenu = 0 Then
            If m_cMenu.p_MenuCount > 0 Then
                lC = GetMenuItemCount(m_cMenu.p_MenuHandle(1))
            End If
        Else
            lC = GetMenuItemCount(m_hMenu)
        End If
        If lC > 0 Then
            lX = 8
            lTop = lTop + 2
            m_cRender.Blit lhDCC, 0&, 0&, lWidth, lHeight, lHdc, lLeft, lTop, SRCCOPY
            SetBkMode lhDCC, ST_TRANSPARENT
            For iIdx = 0 To lC - 1
                bDraw = False
                If m_hMenu = 0 Then
                    tMIIL.fMask = MIIM_ID Or MIIM_DATA Or MIIM_SUBMENU
                    tMIIL.cbSize = LenB(tMIIL)
                    GetMenuItemInfo m_cMenu.p_MenuHandle(1), iIdx, True, tMIIL
                    lIndex = m_cMenu.ItemForID(tMIIL.wID)
                    If lIndex > 0 Then
                        sCap = m_cMenu.Caption(lIndex)
                        bDraw = True
                    End If
                Else
                    lID = GetMenuItemID(m_hMenu, iIdx)
                    If lID = -1 Then
                        With tMII
                            .fMask = MIIM_TYPE
                            .cch = 127
                            .dwTypeData = StrPtr(String$(128, 0))
                            .cbSize = LenB(tMII)
                        End With
                        GetMenuItemInfo m_hMenu, iIdx, True, tMII
                        With tMII
                            If (.fType And MFT_STRING) = MFT_STRING Then
                                If .cch > 0 Then
                                    sCap = ExtractString(m_hMenu, iIdx)
                                Else
                                    sCap = vbNullString
                                End If
                            End If
                        End With
                        bDraw = True
                    End If
                End If
                If bDraw Then
                    With tTR
                        .top = 0
                        .Bottom = lHeight
                        .left = 0
                        .Right = 0
                    End With
                    DrawText lhDCC, sCap, -1, tTR, DT_CALCRECT
                    If (tTR.Bottom - tTR.top + 1 > lMaxBottom) Then
                        lMaxBottom = tTR.Bottom - tTR.top + 1
                    End If
                    
                    OffsetRect tTR, lX, 2
                    LSet tBR = tTR
                    InflateRect tBR, 2, 2
                    tBR.Right = tBR.Right + 7
                    m_iCount = m_iCount + 1
                    bPress = False
                    
                    If m_iCount = m_iDownOn Then
                        '/*** added rollover effects
                        ' This is the item that was clicked:
                        If m_iDownOn = m_iOver Then
                            ' Draw Pressed
                            bPress = True
                            LSet tWRS = tBR
                            InflateRect tWRS, -1, -1
                            hBrush = CreateSolidBrush(m_lMenuRollOverColor)
                            FillRect lhDCC, tWRS, hBrush
                            DeleteObject hBrush
                            If m_bCustomRollOver Then
                                DrawEdge lhDCC, tWRS, EDGE_ETCHED, BF_RECT, False
                            Else
                                DrawEdge lhDCC, tWRS, EDGE_ETCHED, BF_RECT, True
                            End If
                            SetTextColor lhDCC, TranslateColor(m_oActiveMenuColor)
                        Else
                            ' Draw Raised
                            LSet tWRS = tBR
                            InflateRect tWRS, -1, -1
                            hBrush = CreateSolidBrush(m_lMenuRollOverColor)
                            FillRect lhDCC, tWRS, hBrush
                            DeleteObject hBrush
                            If m_bCustomRollOver Then
                                DrawEdge lhDCC, tWRS, EDGE_ETCHED, BF_RECT, False
                            Else
                                DrawEdge lhDCC, tWRS, BDR_SUNKENOUTER, BF_RECT, True
                            End If
                            SetTextColor lhDCC, TranslateColor(m_oInActiveMenuColor)
                        End If
                    Else
                        ' Not down on, may be over:
                        If m_iCount = m_iOver Then
                            ' Draw Raised
                            LSet tWRS = tBR
                            InflateRect tWRS, -1, -1
                            hBrush = CreateSolidBrush(m_lMenuRollOverColor)
                            If m_bCustomRollOver Then
                                DrawFocusRect lhDCC, tBR
                            End If
                            FillRect lhDCC, tWRS, hBrush
                            DeleteObject hBrush
                            If m_bCustomRollOver Then
                                DrawEdge lhDCC, tWRS, BDR_CUSTOMDRAW, BF_RECT, False
                            Else
                                DrawEdge lhDCC, tWRS, BDR_RAISEDINNER, BF_RECT, True
                            End If
                            SetTextColor lhDCC, TranslateColor(m_oActiveMenuColor)
                        Else
                            ' Draw None
                            If bActiveWin Then
                                SetTextColor lhDCC, TranslateColor(m_oActiveMenuColor)
                            Else
                                SetTextColor lhDCC, TranslateColor(m_oInActiveMenuColor)
                            End If
                        End If
                    End If
                    If bPress Then
                        OffsetRect tTR, 1, 1
                    End If
                    With tTR
                        .Right = .Right + 2
                        .left = .left + 2
                    End With
                    DrawText lhDCC, sCap, -1, tTR, DT_CENTER Or DT_SINGLELINE
                    If bPress Then
                        OffsetRect tTR, -1, -1
                    End If
                    ReDim Preserve m_tR(1 To m_iCount) As RECT
                    ReDim Preserve m_hSubMenu(1 To m_iCount) As Long
                    ReDim Preserve m_sTopLevelAccelerator(1 To m_iCount) As String
                    If m_bUseClient Then
                        With m_tR(m_iCount)
                            .left = lLeft + tBR.left
                            .Right = .left + (tBR.Right - tBR.left)
                            .top = lTop
                            .Bottom = .top + (tBR.Bottom - tBR.top)
                        End With
                    Else
                        OffsetRect tBR, lLeft, lYoffset
                        LSet m_tR(m_iCount) = tBR
                    End If
                    If m_hMenu = 0 Then
                        m_hSubMenu(m_iCount) = tMIIL.wID
                    Else
                        m_hSubMenu(m_iCount) = GetSubMenu(m_hMenu, iIdx)
                    End If
                    iPos = InStr(sCap, "&")
                    If iPos > 0 Then
                        If iPos < Len(sCap) Then
                            m_sTopLevelAccelerator(m_iCount) = UCase$(Mid$(sCap, iPos + 1, 1))
                        End If
                    End If
                    lX = lX + tTR.Right - tTR.left + 1 + 10
                End If
            Next iIdx
            lActualRight = lX
            m_cRender.Blit lHdc, lLeft, lTop, lWidth, lHeight, lhDCC, 0&, 0&, SRCCOPY
        End If
        SelectObject lhDCC, hFntOld
        DeleteObject hFnt
        hFntOld = 0
    End If
    Set m_cMenu = Nothing

On Error GoTo 0

End Sub

Private Function ExtractString(ByVal lMenu As Long, lMenuID As Long) As String
'/* Thanks to Keith (Lavolpe) for spotting api bug that caused crashes, and suggesting this fix..

Dim CaptionBytes()  As Byte
Dim lPtr            As Long
Dim lRes            As Long
Dim MIS             As MENUITEMINFO

On Error GoTo Handler

    With MIS
        .cbSize = Len(MIS)
        .fMask = MIIM_STRING
    End With
    
    If GetMenuItemInfo(lMenu, lMenuID, True, MIS) = 0 Then Exit Function
    ReDim CaptionBytes(0 To MIS.cch)
    MIS.dwTypeData = VarPtr(CaptionBytes(0))
    MIS.cch = MIS.cch + 1
    lRes = GetMenuItemInfo(lMenu, lMenuID, True, MIS)
    If lRes = 0 Then Exit Function
    ExtractString = left$(CStr(StrConv(CaptionBytes(), vbUnicode)), MIS.cch)

Handler:
    On Error GoTo 0

End Function

Private Sub SetColors()

    If m_oMenuBackgroundColor = -1 Then m_oMenuBackgroundColor = m_oActiveMenuColor
    If m_lMenuRollOverColor = -1 Then m_lMenuRollOverColor = m_oMenuBackgroundColor
    If m_bMenuBar Then m_OfficeXPStyle = False
    

End Sub

' Convert Automation color to Windows color
Private Function TranslateColor(ByVal Clr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(Clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function

Private Function MeasureItem(ByVal wParam As Long, _
                             ByVal lParam As Long) As Long

Dim tMIS        As MEASUREITEMSTRUCT
Dim lID         As Long
   
    '/* because we don't get the popup menu handle
    '/* in the tMIS structure, we have to do an internal
    '/* lookup to find info about this menu item.
    '/* poor implementation of MEASUREITEMSTRUCT - it
    '/* should have a .hWndItem field like DRAWITEMSTRUCT
    
    CopyMemory tMIS, ByVal lParam, LenB(tMIS)
    If tMIS.CtlType = 1 Then
        lID = GetInternalID(tMIS.itemID)
        '/* Width
        tMIS.itemWidth = 4 + 22 + m_lMenuTextSize(lID) + 4
        With tMIS
            If m_lMenuShortCutSize(lID) > 0 Then
                .itemWidth = .itemWidth + 4 + m_lMenuShortCutSize(lID) + 4
            End If
            If (m_OfficeXPStyle) Then
                .itemWidth = .itemWidth + 4
            End If
        End With
        
        '/* Height:
        With tMIS
            If lID > 0 And lID <= m_iRestore Then
                If (m_tMIIS(lID).fType And &H800) = &H800 Then
                    If (m_OfficeXPStyle) Then
                        .itemHeight = 3
                    Else
                        .itemHeight = 8
                    End If
                Else
                    '/* menu item height is always the same
                    .itemHeight = m_lMenuItemHeight + 6
                    If m_OfficeXPStyle Then
                        .itemHeight = .itemHeight + 4
                    End If
                End If
            End If
        End With
        CopyMemory ByVal lParam, tMIS, LenB(tMIS)
    Else
        MeasureItem = m_MenuSubclass.CallOldWndProc(m_hWnd, &H2C, wParam, lParam)
    End If

End Function

Friend Function AltKeyAccelerator(ByVal vKey As KeyCodeConstants) As Boolean

Dim iIdx   As Long

    For iIdx = 1 To m_iCount
        If m_sTopLevelAccelerator(iIdx) = UCase$(Chr$(vKey)) Then
            PressButton iIdx, True
            If Not m_cTmr Is Nothing Then
                m_cTmr.Interval = 0
            End If
            m_cToolbarMenu.TrackPopup m_iDownOn, m_lMenuOffsetY
            AltKeyAccelerator = True
        End If
    Next iIdx

End Function

Private Function DrawItem(ByVal wParam As Long, _
                          ByVal lParam As Long) As Long

'/* this function really needs a rewrite (way too proceedural)..

Dim bRadioCheck     As Boolean
Dim bDisabled       As Boolean
Dim bChecked        As Boolean
Dim bHighlighted    As Boolean
Dim bIsTopLevel     As Boolean
Dim hBr             As Long
Dim lHdc            As Long
Dim hFntOld         As Long
Dim hFntsOld        As Long
Dim lID             As Long
Dim lSelLeft        As Long
Dim lIconIndex      As Long
Dim lX              As Long
Dim lY              As Long
Dim hBrush          As Long
Dim lHnd            As Long
Dim sCC             As String
Dim tDIS            As DRAWITEMSTRUCT
Dim tMII            As MENUITEMINFO
Dim tR              As RECT
Dim tTR             As RECT
Dim tWR             As RECT
Dim tSideRect       As RECT
Dim tWRS            As RECT

On Error Resume Next

    CopyMemory tDIS, ByVal lParam, LenB(tDIS)

    If tDIS.CtlType = 1 Then
        '/* tDIS.hWndItem is the menu containing the item, tDIS.itemID is the wID
        With tDIS
            m_cMemDC.Width = .rcItem.Right - .rcItem.left + 1
            m_cMemDC.Height = .rcItem.Bottom - .rcItem.top + 1
        End With
        lHdc = m_cMemDC.hdc
        hFntOld = SelectObject(lHdc, p_FontHandle)
        LSet tR = tDIS.rcItem
        OffsetRect tR, -tR.left, -tR.top
        '/* Fill background:
        With tTR
            .Right = m_cMemDC.Width
            .Bottom = m_cMemDC.Height
        End With

        If Not m_bMenuBg Then
            hBr = CreateSolidBrush(TranslateColor(m_oMenuBackgroundColor))
            FillRect lHdc, tTR, hBr
            DeleteObject hBr
        Else
            With tR
                TileArea lHdc, .left, .top, .Right - .left + 1, .Bottom - .top + 1, m_cMenuBg.hdc, m_cMenuBg.Width, m_cMenuBg.Height, tDIS.rcItem.top
            End With
        End If

        If m_OfficeXPStyle Then
            LSet tSideRect = tTR
            tSideRect.Right = m_lMenuItemHeight + 8
            LighterControlColour lHdc, tSideRect, tDIS.rcItem.top
        End If

        tR.top = tR.top + 1
        SetBkMode lHdc, 1
        '/* Draw the text:
        With tMII
            .cbSize = LenB(tMII)
            .fMask = &H10 Or &H1 Or &H20
            ReDim b(0 To 128) As Byte
            .dwTypeData = VarPtr(b(0))
        End With
        GetMenuItemInfo tDIS.hwndItem, tDIS.itemID, False, tMII

        If (tMII.fType And &H800) = &H800 Then
            '/* Separator:
            LSet tWR = tR
            tWR.top = (tWR.Bottom - tWR.top - 2) \ 2 + tWR.top
            tWR.Bottom = tWR.top + 2
            InflateRect tWR, -12, 0
            If m_OfficeXPStyle Then
                LSet tWRS = tWR
                With tWRS
                    .left = tSideRect.Right + 4
                    .Right = .Right + 20
                    .top = .top + 1
                    .Bottom = .top
                End With
                DrawEdge lHdc, tWRS, &H2, &H2 Or &H8, True
            Else
                DrawEdge lHdc, tWR, &H2, &H2 Or &H8, False
            End If
        Else
            With tMII
                '/* Text item:
                bRadioCheck = ((.fType And &H200) = &H200)
                bDisabled = ((.fState And &H3) = &H3)
                bChecked = ((.fState And &H8) = &H8)
                bHighlighted = ((.fState And &H80) = &H80)
                lID = .dwItemData
            End With
            '/* Icon
            lIconIndex = -1
            On Error Resume Next
            lIconIndex = m_colIcons.Item(m_sCaption(lID))
            On Error GoTo 0
            If bChecked Or lIconIndex > -1 Then
                lSelLeft = 4 + (tR.Bottom - tR.top + 1 - 4)
            End If

            If bHighlighted And Not bDisabled Then
                If m_eHighlightStyle = ECGRADIENT Then
                    '/* Draw a gradient
                    LSet tWR = tR
                    With tWR
                        .left = lSelLeft
                        .Right = .left + 4 + (tR.Bottom - tR.top + 1 - 4)
                        hBr = CreateSolidBrush(TranslateColor(p_InActiveMenuForeColor))
                    End With
                    FillRect lHdc, tWR, hBr
                    DeleteObject hBr
                    LSet tWR = tR
                    With tWR
                        .left = .left + 4 + (tR.Bottom - tR.top + 1 - 4)
                        DrawGradient lHdc, tWR, TranslateColor(p_InActiveMenuForeColor), TranslateColor(m_oMenuBackgroundColor), False
                    End With
                ElseIf m_eHighlightStyle = ECBUTTON Then
                    '/* do nothing
                Else
                    '/* standard:
                    If m_OfficeXPStyle Then
                        LSet tWR = tR
                        With tWR
                            .left = .left + 1
                            .Right = .Right - 2
                        End With
                        LighterSelectedColour lHdc, tWR, tDIS.rcItem.top + tWR.top
                        DrawEdge lHdc, tWR, 0&, 0&, True
                    Else
                        LSet tWR = tR
                        tWR.left = lSelLeft
                        If Not m_bMenuBar Then
                            HighlightBackColor lHdc, tWR, tDIS.rcItem.top + tWR.top
                        Else
                            With tWR
                                If lIconIndex > -1 Then
                                    If .Right > m_cMenuBar.Width Then
                                        m_cRender.Stretch lHdc, .left + 1, .top, .Right - 1, .Bottom, m_cMenuBar.hdc, 0, 0, m_cMenuBar.Width, m_cMenuBar.Height, SRCCOPY
                                    Else
                                        m_cRender.Blit lHdc, .left + 1, .top, .Right - 1, .Bottom, m_cMenuBar.hdc, 0&, 0&, SRCCOPY
                                    End If
                                Else
                                    If .Right > m_cMenuBar.Width Then
                                        m_cRender.Stretch lHdc, .left, .top, .Right, .Bottom, m_cMenuBar.hdc, 0, 0, m_cMenuBar.Width, m_cMenuBar.Height, SRCCOPY
                                    Else
                                        m_cRender.Blit lHdc, .left, .top, .Right, .Bottom, m_cMenuBar.hdc, 0&, 0&, SRCCOPY
                                    End If
                                End If
                            End With
                        End If
                    End If
                End If
            End If

            If bDisabled Then
                SetTextColor lHdc, TranslateColor(vb3DHighlight)
            Else
                If bHighlighted Then
                    SetTextColor lHdc, TranslateColor(m_oActiveMenuColor)
                Else
                    SetTextColor lHdc, TranslateColor(p_InActiveMenuForeColor)
                End If
            End If

            '/* Get the check/icon space:
            LSet tWR = tR
            With tWR
                If m_eHighlightStyle = ECBUTTON Then
                    InflateRect tWR, -2, -2
                Else
                    .left = .left + 1
                End If
                .Right = .left + (.Bottom - .top + 1 - 2)
            End With

            '/* Check:
            If bChecked Then
                '/* Colour in
                If Not bHighlighted Then
                    SetBkMode lHdc, 2
                    If m_OfficeXPStyle Then
                        LSet tWRS = tWR
                        InflateRect tWRS, -1, -1
                        hBrush = CreateSolidBrush(p_BlendColor(vbHighlight, p_BlendColor(m_oMenuBackgroundColor, vbButtonFace, 128), 40))
                        FillRect lHdc, tWRS, hBrush
                        DeleteObject hBrush
                    Else
                        If p_NoPalette Then
                            hBrush = CreateSolidBrush(p_LighterColour(m_oMenuBackgroundColor))
                            FillRect lHdc, tWR, hBrush
                            DeleteObject hBrush
                        Else
                            With tWR
                                m_cBrush.Rectangle lHdc, .left, .top, .Right - .left, .Bottom - .top, 1, PATCOPYEX, True, m_oMenuBackgroundColor, vb3DHighlight
                            End With
                        End If
                    End If
                    SetBkMode lHdc, 1
                    If bDisabled Then
                        SetTextColor lHdc, TranslateColor(vb3DHighlight)
                    End If
                ElseIf m_OfficeXPStyle Then
                    LSet tWRS = tWR
                    InflateRect tWRS, -1, -1
                    hBrush = CreateSolidBrush(p_BlendColor(vbHighlight, m_oMenuBackgroundColor, 128))
                    FillRect lHdc, tWRS, hBrush
                    DeleteObject hBrush
                End If

                If Not bDisabled Then
                    If bHighlighted Then
                        SetTextColor lHdc, TranslateColor(p_ActiveMenuForeColor)
                    Else
                        SetTextColor lHdc, TranslateColor(p_InActiveMenuForeColor)
                    End If
                End If

                If m_OfficeXPStyle Then
                    LSet tWRS = tWR
                    InflateRect tWRS, -1, -1
                    DrawEdge lHdc, tWRS, &H2, BF_RECT, m_OfficeXPStyle
                Else
                    DrawEdge lHdc, tWR, &H2, BF_RECT, m_OfficeXPStyle
                End If

                If lIconIndex = -1 Then
                    '/* Draw the appropriate symbol:
                    SelectObject lHdc, hFntOld
                    hFntsOld = SelectObject(lHdc, p_FontSymbol)
                    If bHighlighted Then
                        SetTextColor lHdc, TranslateColor(p_InActiveMenuForeColor)
                    End If
                    If bRadioCheck Then
                        RenderText lHdc, "h", tWR, &H4 Or &H1 Or &H20, bDisabled
                    Else
                        RenderText lHdc, "b", tWR, &H4 Or &H1 Or &H20, bDisabled
                    End If
                    SelectObject lHdc, hFntsOld
                    hFntOld = SelectObject(lHdc, p_FontHandle)
                    hFntsOld = 0
                    If bHighlighted Then
                        SetTextColor lHdc, TranslateColor(p_ActiveMenuForeColor)
                    End If
                Else
                    With tWR
                        lX = .left + (.Right - .left + 1 - m_lIconSize) \ 2
                        lY = .top + (.Bottom - .top + 1 - m_lIconSize) \ 2
                    End With
                    If bDisabled Then
                        DrawIconDisabled m_ptrVb6ImageList, lHdc, m_hIml, lIconIndex, lX, lY, m_lIconSize
                    Else
                        DrawIcon m_ptrVb6ImageList, lHdc, m_hIml, lIconIndex, lX, lY
                    End If
                End If
            Else
                If lIconIndex > -1 Then
                    If bHighlighted And Not bDisabled Then
                        If (Not (m_eHighlightStyle = ECBUTTON) And Not (m_OfficeXPStyle)) Then
                            DrawEdge lHdc, tWR, &H4, BF_RECT, m_OfficeXPStyle
                        End If
                    End If
                    With tWR
                        lX = .left + (.Right - .left + 1 - m_lIconSize) \ 2
                        lY = .top + (.Bottom - .top + 1 - m_lIconSize) \ 2
                    End With
                    lX = lX + 2 * Abs(m_eHighlightStyle = ECBUTTON)
                    If bDisabled Then
                        DrawIconDisabled m_ptrVb6ImageList, lHdc, m_hIml, lIconIndex, lX, lY, m_lIconSize
                    Else
                        If m_OfficeXPStyle Then
                            If bHighlighted Then
                                DrawIconDisabled m_ptrVb6ImageList, lHdc, m_hIml, lIconIndex, lX + 1, lY + 1, m_lIconSize, True
                                DrawIcon m_ptrVb6ImageList, lHdc, m_hIml, lIconIndex, lX - 1, lY - 1
                            Else
                                DrawIcon m_ptrVb6ImageList, lHdc, m_hIml, lIconIndex, lX, lY
                            End If
                        Else
                            DrawIcon m_ptrVb6ImageList, lHdc, m_hIml, lIconIndex, lX, lY
                        End If
                    End If
                End If
            End If

            '/* Draw text:
            If m_eHighlightStyle = ECBUTTON And Not (bDisabled) Then
                SetTextColor lHdc, TranslateColor(p_InActiveMenuForeColor)
            End If
            LSet tWR = tR
            If Not m_bMenuBar Then
                tWR.left = 4 + (tR.Bottom - tR.top + 1 - 4) + 2
            Else
                tWR.left = 4 + (tR.Bottom - tR.top + 1 - 4) + 5
            End If
            If m_OfficeXPStyle Then
                tWR.left = tWR.left + 4
            End If
            If lID > 0 Then
                If lID <= m_iRestore Then
                    RenderText lHdc, m_sCaption(lID), tWR, &H0 Or &H20 Or &H4, bDisabled
                    If Len(m_sShortCut(lID)) > 0 Then
                        tWR.left = tWR.left + m_lMenuTextSize(lID) + 4 + 4
                        RenderText lHdc, m_sShortCut(lID), tWR, &H0 Or &H20 Or &H4, bDisabled
                    End If
                End If
            End If
            '/* Highlighted:
            If bHighlighted And m_eHighlightStyle = ECBUTTON And Not (bDisabled) Then
                LSet tWR = tR
                InflateRect tWR, -2, 0
                DrawEdge lHdc, tWR, &H4, BF_RECT, m_OfficeXPStyle
            End If
        End If
        SelectObject lHdc, hFntOld
        With tDIS
            m_cRender.Blit .hdc, .rcItem.left, .rcItem.top, .rcItem.Right - .rcItem.left + 1, .rcItem.Bottom - .rcItem.top + 1, lHdc, 0&, 0&, SRCCOPY
        End With
        If m_bTransparent Then
            lHnd = WindowFromDC(tDIS.hdc)
            If Not lHnd = 0 Then
                TransparentMenu lHnd, m_btTransIndex
            End If
        End If
        hFntOld = 0
        hFntsOld = 0
    Else
        DrawItem = m_MenuSubclass.CallOldWndProc(m_hWnd, &H2B, wParam, lParam)
    End If

On Error GoTo 0

End Function

Private Function ParseMenuChar(ByVal hMenu As Long, _
                               ByVal iChar As Integer) As Long

Dim sChar           As String
Dim l               As Long
Dim lH()            As Long
Dim sItems()        As String
    
    sChar = UCase$(Chr$(iChar))
    For l = 1 To m_iRestore
        If m_hMenuRestore(l) = hMenu Then
            If m_sAccelerator(l) = sChar Then
               ParseMenuChar = &H20000 Or m_iMenuPosition(l)
               Exit Function
            End If
        End If
    Next l
    
End Function

Friend Sub GetRect(ByVal iButton As Long, _
                   ByRef tR As RECT)

Dim tRW     As RECT

    If iButton > 0 Then
        If iButton <= m_iCount Then
            LSet tR = m_tR(iButton)
            GetWindowRect m_hWnd, tRW
            If m_bUseClient Then
                OffsetRect tR, tRW.left, tRW.top
            Else
                OffsetRect tR, tRW.left, tRW.top + m_lCaptionHeight
            End If
        End If
    End If

End Sub

Friend Sub PressButton(ByVal iButton As Long, _
                       ByVal bState As Boolean)
    
    If bState Then
        m_iDownOn = iButton
    Else
        If m_iDownOn = iButton Then
            m_iDownOn = -1
        End If
    End If
    RepaintOwner
    
End Sub

Public Sub OwnerDrawMenu(ByVal hMenu As Long)

Dim bHaveSeen               As Boolean
Dim lC                      As Long
Dim iMenu                   As Long
Dim iPos                    As Long
Dim lID                     As Long
Dim hFntOld                 As Long
Dim lMenuTextSize           As Long
Dim lMenuShortCutSize       As Long
Dim i                       As Long
Dim sCap                    As String
Dim sShortCut               As String
Dim tR                      As RECT
Dim tMII                    As MENUITEMINFO
Dim tMIIS                   As MENUITEMINFO
Dim aCaption()              As String

On Error Resume Next

    '/* Set OD flag on the fly...
    bHaveSeen = HaveSeen(hMenu)
    hFntOld = SelectObject(m_cMemDC.hdc, hMenu)
    lC = GetMenuItemCount(hMenu)
    ReDim aCaption(0)
    For iMenu = 0 To lC - 1
        ReDim Preserve aCaption(0 To lC - 1)
        aCaption(iMenu) = String$(128, 0)
        If Not bHaveSeen Then
            With tMIIS
                .fMask = &H10 Or &H20 Or &H2
                .cch = 127
                .dwTypeData = StrPtr(aCaption(iMenu))
                .cbSize = LenB(tMIIS)
            End With
            GetMenuItemInfo hMenu, iMenu, True, tMIIS
            lID = AddToRestoreList(hMenu, iMenu, tMIIS)
            If Not (tMIIS.fType And &H100) = &H100 Then
                '/* Setting this flag causes tMIIS.dwTypeData to be
                '/* overwritten with our own app-defined value:
                With tMII
                    .fType = .fType Or &H100 And Not &H0
                    .dwItemData = lID
                    .cbSize = LenB(tMII)
                    .fMask = &H10 Or &H20
                End With
                SetMenuItemInfo hMenu, iMenu, True, tMII
            End If
        Else
            With tMII
                .fMask = &H10 Or &H20 Or &H2
                .cbSize = Len(tMIIS)
                ReDim b(0 To 128) As Byte
                .dwTypeData = StrPtr(aCaption(iMenu))
                GetMenuItemInfo hMenu, iMenu, True, tMII
                lID = .dwItemData
            End With
            If Not ((tMII.fType And &H100) = &H100) Then
                lID = ReplaceIndex(hMenu, iMenu)
                With tMIIS
                    .fMask = &H10 Or &H20 Or &H2
                    .cch = 127
                    .dwTypeData = StrPtr(aCaption(iMenu))
                    .cbSize = LenB(tMIIS)
                End With
                GetMenuItemInfo hMenu, iMenu, True, tMIIS
                If lID = 0 Then
                    '/* New item,just added:
                    lID = AddToRestoreList(hMenu, iMenu, tMIIS)
                Else
                    '/* replacing existing:
                    ReplaceRestoreList lID, hMenu, iMenu, tMIIS
                End If
                '/* Setting this flag causes tMIIS.dwTypeData to be
                '/* overwritten with our own app-defined value:
                With tMII
                    .fType = .fType Or &H100 And Not &H0
                    .dwItemData = lID
                    .cbSize = LenB(tMII)
                    .fMask = &H10 Or &H20
                End With
                SetMenuItemInfo hMenu, iMenu, True, tMII
            End If
        End If

        If lID > 0 And lID <= m_iRestore Then
            sCap = m_sCaption(lID)
            sShortCut = m_sShortCut(lID)
            DrawText m_cMemDC.hdc, sCap, -1, tR, &H0 Or &H20 Or &H400
            With tR
                If .Right - .left + 1 > lMenuTextSize Then
                    lMenuTextSize = .Right - .left + 1
                End If
                If Len(sShortCut) > 0 Then
                    DrawText m_cMemDC.hdc, sShortCut, -1, tR, &H0 Or &H20 Or &H400
                    If .Right - .left + 1 > lMenuShortCutSize Then
                        lMenuShortCutSize = .Right - .left + 1
                    End If
                End If
                m_lMenuItemHeight = .Bottom - .top + 2
            End With
            
            If m_lMenuItemHeight < m_lIconSize - 1 Then
                m_lMenuItemHeight = m_lIconSize - 1
            End If
        End If
    Next iMenu
   
    For i = 1 To m_iRestore
        If m_hMenuRestore(i) = hMenu Then
            m_lMenuTextSize(i) = lMenuTextSize
            m_lMenuShortCutSize(i) = lMenuShortCutSize
        End If
    Next i

    SelectObject m_cMemDC.hdc, hFntOld
    hFntOld = 0

On Error GoTo 0

End Sub

Private Function HaveSeen(ByVal hMenu As Long) As Boolean
   
'/* When &H117 fires, this may or not be
'/* a new menu.  We use an array to store which menus
'/* we've already worked on:

Dim i       As Long
   
    For i = 1 To m_iHaveSeenCount
        If hMenu = m_hMenuSeen(i) Then
            HaveSeen = True
            Exit Function
        End If
    Next i
   
    m_iHaveSeenCount = m_iHaveSeenCount + 1
    ReDim Preserve m_hMenuSeen(1 To m_iHaveSeenCount) As Long
    m_hMenuSeen(m_iHaveSeenCount) = hMenu

End Function

Private Function AddToRestoreList(ByVal hMenu As Long, _
                                  ByVal iMenu As Long, _
                                  ByRef tMIIS As MENUITEMINFO) As Long

' Here we store information about a menu item.  When the
' menus are closed again we can reset things back to the
' way they were using this struct.

    m_iRestore = m_iRestore + 1
    ReDim Preserve m_hMenuRestore(1 To m_iRestore) As Long
    ReDim Preserve m_iMenuPosition(1 To m_iRestore) As Long
    ReDim Preserve m_tMIIS(1 To m_iRestore) As MENUITEMINFO
    ReDim Preserve m_sCaption(1 To m_iRestore) As String
    ReDim Preserve m_sShortCut(1 To m_iRestore) As String
    ReDim Preserve m_sAccelerator(1 To m_iRestore) As String
    ReDim Preserve m_lMenuTextSize(1 To m_iRestore) As Long
    ReDim Preserve m_lMenuShortCutSize(1 To m_iRestore) As Long
    ReplaceRestoreList m_iRestore, hMenu, iMenu, tMIIS
    AddToRestoreList = m_iRestore
    
End Function

Private Sub ReplaceRestoreList(ByVal lIdx As Long, _
                               ByVal hMenu As Long, _
                               ByVal iMenu As Long, _
                               ByRef tMIIS As MENUITEMINFO)

Dim iPos    As Long
Dim sCap    As String

    m_hMenuRestore(lIdx) = hMenu
    m_iMenuPosition(lIdx) = iMenu
    LSet m_tMIIS(lIdx) = tMIIS
    If tMIIS.cch > 0 Then
        sCap = ExtractString(hMenu, iMenu)
    Else
        sCap = vbNullString
    End If
    iPos = InStr(sCap, vbTab)
    If iPos > 0 Then
        m_sShortCut(lIdx) = Mid$(sCap, iPos + 1)
        m_sCaption(lIdx) = left$(sCap, iPos - 1)
    Else
        m_sCaption(lIdx) = sCap
        m_sShortCut(lIdx) = vbNullString
    End If
    iPos = InStr(m_sCaption(lIdx), "&")
    If iPos > 0 Then
        If iPos < Len(m_sCaption(lIdx)) Then
            m_sAccelerator(lIdx) = UCase$(Mid$(m_sCaption(lIdx), iPos + 1, 1))
        End If
    End If
     
End Sub

Private Function ReplaceIndex(ByVal hMenu As Long, _
                              ByVal iMenu As Long)

Dim i       As Long

    For i = 1 To m_iRestore
        If m_hMenuRestore(i) = hMenu Then
            If m_iMenuPosition(i) = iMenu Then
                ReplaceIndex = i
                Exit Function
            End If
        End If
    Next i

End Function

Private Sub RestoreList()

Dim i       As Long

    '/* erase
    For i = 1 To m_iRestore
        SetMenuItemInfo m_hMenuRestore(i), m_iMenuPosition(i), True, m_tMIIS(i)
    Next i
    
    m_iRestore = 0
    Erase m_hMenuRestore
    Erase m_iMenuPosition
    Erase m_tMIIS
    Erase m_sCaption()
    Erase m_sShortCut()
    Erase m_sAccelerator()
    m_iHaveSeenCount = 0
    Erase m_hMenuSeen()
   
End Sub

Public Sub IconItemCaptionChanged(ByVal sOldCaption As String, _
                                  ByVal sNewCaption As String)

Dim lIdx        As Long

    lIdx = -1
    On Error Resume Next
    lIdx = m_colIcons.Item(sOldCaption)
    If lIdx > -1 Then
        m_colIcons.Remove sOldCaption
        m_colIcons.Add lIdx, sNewCaption
    End If
   
End Sub

Public Sub ClearIcons()

    Set m_colIcons = New Collection
   
End Sub

Private Function GetInternalID(ByVal wID As Long) As Long

Dim i       As Long

    For i = 1 To m_iRestore
        If m_tMIIS(i).wID = wID Then
            GetInternalID = i
            Exit Function
        End If
    Next i

End Function

Private Function RenderText(ByVal lHdc As Long, _
                            ByVal sText As String, _
                            ByRef tR As RECT, _
                            ByVal dtFlags As Long, _
                            ByVal bDisabled As Boolean)

Dim tWR     As RECT

    LSet tWR = tR
    If bDisabled Then
        If m_OfficeXPStyle Then
            SetTextColor lHdc, TranslateColor(vb3DShadow)
        Else
            SetTextColor lHdc, TranslateColor(vb3DHighlight)
            OffsetRect tWR, 1, 1
        End If
    End If
    
    DrawText lHdc, sText, -1, tWR, dtFlags
    If bDisabled Then
        If Not (m_OfficeXPStyle) Then
            SetTextColor lHdc, TranslateColor(vbButtonShadow)
            OffsetRect tWR, -1, -1
            DrawText lHdc, sText, -1, tWR, dtFlags
        End If
    End If

End Function

Private Sub DrawGradient(ByVal hdc As Long, _
                         ByRef rct As RECT, _
                         ByVal lEndColour As Long, _
                         ByVal lStartColour As Long, _
                         ByVal bVertical As Boolean)

Dim dR(1 To 3)              As Double
Dim dPos                    As Double
Dim d                       As Double
Dim lStep                   As Long
Dim lPos                    As Long
Dim lSize                   As Long
Dim hBr                     As Long
Dim bRGB(1 To 3)            As Integer
Dim bRGBStart(1 To 3)       As Integer
Dim tR                      As RECT
   
    LSet tR = rct
    If bVertical Then
        lSize = (tR.Bottom - tR.top)
    Else
        lSize = (tR.Right - tR.left)
    End If
    lStep = lSize \ 255
    If (lStep < 3) Then
        lStep = 3
    End If
       
    bRGB(1) = lStartColour And &HFF&
    bRGB(2) = (lStartColour And &HFF00&) \ &H100&
    bRGB(3) = (lStartColour And &HFF0000) \ &H10000
    bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
    dR(1) = (lEndColour And &HFF&) - bRGB(1)
    dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
    dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)
    
    
    For lPos = lSize To 0 Step -lStep
        With tR
            '/* Draw bar:
            If bVertical Then
                .top = .Bottom - lStep
            Else
                .left = .Right - lStep
            End If
            If .top < rct.top Then
                .top = rct.top
            End If
            If .left < rct.left Then
                .left = rct.left
            End If
        End With
      
        hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
        FillRect hdc, tR, hBr
        DeleteObject hBr
        '/* Adjust colour:
        dPos = ((lSize - lPos) / lSize)
        If bVertical Then
            tR.Bottom = tR.top
            bRGB(1) = bRGBStart(1) + dR(1) * dPos
            bRGB(2) = bRGBStart(2) + dR(2) * dPos
            bRGB(3) = bRGBStart(3) + dR(3) * dPos
        Else
            tR.Right = tR.left
            bRGB(1) = bRGBStart(1) + dR(1) * dPos
            bRGB(2) = bRGBStart(2) + dR(2) * dPos
            bRGB(3) = bRGBStart(3) + dR(3) * dPos
        End If
    Next lPos

End Sub

Private Sub TileArea(ByVal lDHdc As Long, _
                     ByVal lDX As Long, _
                     ByVal lDY As Long, _
                     ByVal lDWidth As Long, _
                     ByVal lDHeight As Long, _
                     ByVal lSDc As Long, _
                     ByVal lSrcWidth As Long, _
                     ByVal lSrcHeight As Long, _
                     ByVal lOffsetY As Long)

Dim lSrcX                   As Long
Dim lSrcY                   As Long
Dim lSrcStartX              As Long
Dim lSrcStartY              As Long
Dim lSrcStartWidth          As Long
Dim lSrcStartHeight         As Long
Dim lDstX                   As Long
Dim lDstY                   As Long
Dim lDstWidth               As Long
Dim lDstHeight              As Long

    lSrcStartX = (lDX Mod lSrcWidth)
    lSrcStartY = ((lDY + lOffsetY) Mod lSrcHeight)
    lSrcStartWidth = (lSrcWidth - lSrcStartX)
    lSrcStartHeight = (lSrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    lDstY = lDY
    lDstHeight = lSrcStartHeight
    
    If lDWidth > lSrcWidth Then
        m_cRender.Stretch lDHdc, lDstX, lDstY, lDWidth, lSrcHeight, lSDc, 0, 0, lSrcWidth, lSrcHeight, SRCCOPY
     '   Exit Sub
    End If
   
    Do While lDstY < (lDY + lDHeight)
        If (lDstY + lDstHeight) > (lDY + lDHeight) Then
            lDstHeight = lDY + lDHeight - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = lDX
        lSrcX = lSrcStartX
        Do While lDstX < (lDX + lDWidth)
            If (lDstX + lDstWidth) > (lDX + lDWidth) Then
                lDstWidth = lDX + lDWidth - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            m_cRender.Blit lDHdc, lDstX, lDstY, lDstWidth, lDstHeight, lSDc, lSrcX, lSrcY, SRCCOPY
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = lSrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = lSrcHeight
    Loop

End Sub

Private Sub DrawIcon(ByVal ptrVb6ImageList As Long, _
                     ByVal hdc As Long, _
                     ByVal hIml As Long, _
                     ByVal iIconIndex As Long, _
                     ByVal lX As Long, _
                     ByVal lY As Long, _
                     Optional ByVal bSelected As Boolean = False, _
                     Optional ByVal bBlend25 As Boolean = False)

Dim lFlags      As Long
Dim lR          As Long

    lFlags = 1
    If (bSelected) Then
        lFlags = lFlags Or 4
    End If
    
    If (bBlend25) Then
        lFlags = lFlags Or 2
    End If
    
    If Not m_ptrVb6ImageList = 0 Then
        Dim o As Object
        On Error Resume Next
        Set o = ObjectFromPtr(m_ptrVb6ImageList)
        If Not (o Is Nothing) Then
            With Screen
                o.ListImages(iIconIndex + 1).Draw hdc, lX * .TwipsPerPixelX, _
                    lY * .TwipsPerPixelY, lFlags
            End With
        End If
        On Error GoTo 0
    Else
        lR = ImageList_Draw(hIml, iIconIndex, hdc, lX, lY, lFlags)
    End If
    
End Sub

Private Sub DrawIconDisabled(ByVal ptrVb6ImageList As Long, _
                             ByVal hdc As Long, _
                             ByVal hIml As Long, _
                             ByVal iIconIndex As Long, _
                             ByVal lX As Long, _
                             ByVal lY As Long, _
                             ByVal lSize As Long, _
                             Optional ByVal asShadow As Boolean)

Dim lR          As Long
Dim hIcon       As Long

    hIcon = 0
    If Not m_ptrVb6ImageList = 0 Then
        Dim o As Object
        On Error Resume Next
        Set o = ObjectFromPtr(m_ptrVb6ImageList)
        If Not o Is Nothing Then
            Dim lhDCDisp As Long
            Dim lHdc As Long
            Dim lhBmp As Long
            Dim lhBmpOld As Long
            Dim lhIml As Long
            lhDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
            lHdc = CreateCompatibleDC(lhDCDisp)
            lhBmp = CreateCompatibleBitmap(lhDCDisp, o.ImageWidth, o.ImageHeight)
            lhBmpOld = SelectObject(lHdc, lhBmp)
            o.ListImages.Item(iIconIndex + 1).Draw lHdc, 0, 0, 0
            lhIml = ImageList_Create(o.ImageWidth, o.ImageHeight, &H1 Or &H20, 1&, 1&)
            ImageList_AddMasked lhIml, lhBmp, TranslateColor(o.BackColor)
            SelectObject lHdc, lhBmpOld
            DeleteObject lhBmp
            DeleteDC lHdc
            hIcon = ImageList_GetIcon(lhIml, 0&, 0&)
            ImageList_Destroy lhIml
        End If
        On Error GoTo 0
    Else
        hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
    End If
    
    If Not hIcon = 0 Then
        If (asShadow) Then
            Dim hBr As Long
            hBr = GetSysColorBrush(vb3DShadow And &H1F)
            lR = DrawState(hdc, hBr, 0&, hIcon, 0&, lX, lY, lSize, lSize, &H3 Or &H80)
            DeleteObject hBr
        Else
            lR = DrawState(hdc, 0&, 0&, hIcon, 0&, lX, lY, lSize, lSize, &H3 Or &H20)
        End If
        DestroyIcon hIcon
    End If
   
End Sub

Private Sub LighterBackColor(ByVal lHdc As Long, _
                             ByRef tR As RECT, _
                             ByVal lOffsetY As Long, _
                             ByVal bInfrequent As Boolean)

Dim hBrush      As Long

    SetBkMode lHdc, 2
    If (p_NoPalette) Then
        If bInfrequent Then
            hBrush = CreateSolidBrush(p_SlightlyLighterColour(p_MenuBackgroundColor))
        Else
            hBrush = CreateSolidBrush(p_LighterColour(p_MenuBackgroundColor))
        End If
        FillRect lHdc, tR, hBrush
        DeleteObject hBrush
    Else
        With tR
            m_cBrush.Rectangle lHdc, .left, .top, .Right - .left + 1, _
                .Bottom - .top + 1, 1, PATCOPYEX, True, p_MenuBackgroundColor, vb3DHighlight
        End With
    End If
    SetBkMode lHdc, 1

End Sub

Private Sub HighlightBackColor(ByVal lHdc As Long, _
                               ByRef tR As RECT, _
                              ByVal lOffsetY As Long)
Dim hBr     As Long

    hBr = CreateSolidBrush(TranslateColor(vbHighlight))
    FillRect lHdc, tR, hBr
    DeleteObject hBr
   
End Sub

Private Sub NormalBackground(ByVal lHdc As Long, _
                             ByRef tR As RECT, _
                             ByVal lOffsetY As Long)

Dim hBrush      As Long

    hBrush = CreateSolidBrush(TranslateColor(m_oMenuBackgroundColor))
    FillRect lHdc, tR, hBrush
    DeleteObject hBrush

End Sub

Private Sub LighterControlColour(ByVal lHdc As Long, _
                                 ByRef tR As RECT, _
                                 ByVal lOffsetY As Long)

Dim hBrush      As Long

    SetBkMode lHdc, 2
    If (p_NoPalette) Then
        hBrush = CreateSolidBrush(p_BlendColor(p_MenuBackgroundColor, vbButtonFace, 128))
        FillRect lHdc, tR, hBrush
        DeleteObject hBrush
    Else
        With tR
            m_cBrush.Rectangle lHdc, .left, .top, .Right - .left + 1, _
                .Bottom - .top + 1, 1, PATCOPYEX, True, p_MenuBackgroundColor, vb3DHighlight
        End With
    End If
    SetBkMode lHdc, 1
    
End Sub

Private Sub LighterSelectedColour(ByVal lHdc As Long, _
                                  ByRef tR As RECT, _
                                  ByVal lOffsetY As Long)

Dim hBrush      As Long

    SetBkMode lHdc, 2
    If (p_NoPalette) Then
        hBrush = CreateSolidBrush(p_BlendColor(vbHighlight, p_MenuBackgroundColor, 80))
        FillRect lHdc, tR, hBrush
        DeleteObject hBrush
    Else
        With tR
            m_cBrush.Rectangle lHdc, .left, .top, .Right - .left + 1, _
                .Bottom - .top + 1, 1, PATCOPYEX, True, p_MenuBackgroundColor, vb3DHighlight
        End With
    End If
   SetBkMode lHdc, 1
   
End Sub

Public Sub Attach(ByVal lHwnd As Long, _
                  ByVal bUseClient As Boolean)

    m_bUseClient = bUseClient
    m_hWnd = lHwnd
    With m_MenuSubclass
        .Subclass lHwnd, Me
        .AddMessage m_hWnd, WM_LBUTTONDOWN, MSG_BEFORE
        .AddMessage m_hWnd, WM_MOUSEMOVE, MSG_BEFORE
        .AddMessage m_hWnd, WM_SETTINGCHANGE, MSG_BEFORE
        .AddMessage m_hWnd, WM_DRAWITEM, MSG_BEFORE
        .AddMessage m_hWnd, WM_MEASUREITEM, MSG_BEFORE
        .AddMessage m_hWnd, WM_MENUCHAR, MSG_BEFORE
        .AddMessage m_hWnd, WM_EXITMENULOOP, MSG_BEFORE
        .AddMessage m_hWnd, WM_DESTROY, MSG_BEFORE
    End With
    Set m_cToolbarMenu = New clsToolbarMenu
    m_cToolbarMenu.CoolMenuAttach m_hWnd, Me
    SetColors
    
End Sub

Public Sub Detach()

    If Not m_cToolbarMenu Is Nothing Then
        m_cToolbarMenu.CoolMenuDetach
        Set m_cToolbarMenu = Nothing
    End If
    
    With m_MenuSubclass
        If Not m_hWnd = 0 Then
            .DeleteMessage m_hWnd, WM_LBUTTONDOWN, MSG_BEFORE
            .DeleteMessage m_hWnd, WM_MOUSEMOVE, MSG_BEFORE
            .DeleteMessage m_hWnd, WM_SETTINGCHANGE, MSG_BEFORE
            .DeleteMessage m_hWnd, WM_DRAWITEM, MSG_BEFORE
            .DeleteMessage m_hWnd, WM_MEASUREITEM, MSG_BEFORE
            .DeleteMessage m_hWnd, WM_MENUCHAR, MSG_BEFORE
            .DeleteMessage m_hWnd, WM_EXITMENULOOP, MSG_BEFORE
            .DeleteMessage m_hWnd, WM_DESTROY, MSG_BEFORE
            .UnSubclass m_hWnd
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

Dim iMenu               As Long
Dim iLastDownOn         As Long
Dim iLastOver           As Long
Dim lR                  As Long
Dim lFlag               As Long
Dim hMenu               As Long
Dim iChar               As Long

    Select Case uMsg
    Case WM_MOUSEMOVE
        lReturn = m_MenuSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        MouseMove
        
    Case WM_LBUTTONDOWN
        lReturn = m_MenuSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        ' If in range, then...
        iMenu = MenuHitTest()
        iLastDownOn = m_iDownOn
        m_iDownOn = iMenu
        If m_iDownOn <> iLastDownOn Then
            RepaintOwner
        End If
        If m_iDownOn > 0 Then
            m_cTmr.Interval = 0
            Set m_cTmr = Nothing
            lR = m_cToolbarMenu.TrackPopup(m_iDownOn, m_lMenuOffsetY, m_bTransparent)
            RestoreList
        End If
        
        Case WM_MENUCHAR
            '/* Check that this is my menu:
            lFlag = wParam \ &H10000
            If ((lFlag And &H2000) <> &H2000) Then
                hMenu = lParam
                iChar = (wParam And &HFFFF&)
                '/* See if this corresponds to an accelerator on the menu:
                lR = ParseMenuChar(hMenu, iChar)
                If lR > 0 Then
                    lReturn = lR
                    Exit Sub
                End If
            End If
            lReturn = m_MenuSubclass.CallOldWndProc(m_hWnd, &H120, wParam, lParam)
        
        Case WM_MEASUREITEM
            lReturn = MeasureItem(wParam, lParam)
            
        Case WM_DRAWITEM
            DrawItem wParam, lParam
            
        Case WM_SETTINGCHANGE
            m_cNCM.GetMetrics
            
        Case WM_EXITMENULOOP
            RestoreList
            
        Case WM_DESTROY
            Detach
    End Select

End Sub

Private Sub Class_Terminate()

On Error Resume Next

    RestoreList
    Detach
    If Not m_cToolbarMenu Is Nothing Then Set m_cToolbarMenu = Nothing
    If Not m_cBrush Is Nothing Then Set m_cBrush = Nothing
    If Not m_cNCM Is Nothing Then Set m_cNCM = Nothing
    If Not m_MenuSubclass Is Nothing Then Set m_MenuSubclass = Nothing
    If Not m_cRender Is Nothing Then Set m_cRender = Nothing
    If Not m_cMenuBar Is Nothing Then Set m_cMenuBar = Nothing
    If Not m_cMenuBg Is Nothing Then Set m_cMenuBg = Nothing
    If Not m_cMemDC Is Nothing Then Set m_cMemDC = Nothing
    If Not m_cTmr Is Nothing Then Set m_cTmr = Nothing
    If Not m_fnt Is Nothing Then Set m_fnt = Nothing
    If Not m_fntSymbol Is Nothing Then Set m_fntSymbol = Nothing
    If Not p_MenuBgPicture Is Nothing Then Set p_MenuBgPicture = Nothing
    If Not p_MenuBarPicture Is Nothing Then Set p_MenuBarPicture = Nothing
    If Not m_colIcons Is Nothing Then Set m_colIcons = Nothing
    Erase m_hSubMenu
    Erase m_hMenuRestore
    Erase m_iMenuPosition
    Erase m_lMenuTextSize
    Erase m_lMenuShortCutSize
    Erase m_hMenuSeen
    Erase m_sTopLevelAccelerator
    Erase m_sCaption
    Erase m_sShortCut
    Erase m_sAccelerator
    Erase m_tR
    Erase m_tMIIS
    m_hMenu = 0
    m_hWnd = 0

End Sub
