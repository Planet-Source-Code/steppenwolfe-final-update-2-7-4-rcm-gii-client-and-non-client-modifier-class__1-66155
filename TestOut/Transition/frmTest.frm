VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   780
      Left            =   3690
      TabIndex        =   2
      Top             =   2205
      Width           =   1725
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "Maximize"
      Height          =   690
      Index           =   1
      Left            =   6705
      TabIndex        =   1
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "Restore"
      Height          =   690
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WS_EX_LAYERED                         As Long = &H80000
Private Const WS_EX_TRANSPARENT                     As Long = &H20&
Private Const LWA_ALPHA                             As Long = &H2&

Private Enum WD_STATE
    WD_MINIMIZED = 1
    WD_NORMAL = 2
    WD_MAXIMIZED = 3
End Enum

Private Enum WC_EXTENDEDSTYLE
    GWL_EXSTYLE = -20
    GWL_STYLE = -16
    GWL_WNDPROC = -4
    GWL_HINSTANCE = -6
    GWL_HWNDPARENT = -8
    GWL_ID = -12
    GWL_USERDATA = -21
End Enum

'/* system metrics
Private Enum SYSTEM_METRICS
    SM_CXSCREEN = 0         '/* X Size of screen
    SM_CYSCREEN = 1         '/* Y Size of Screen
    SM_CXVSCROLL = 2        '/* X Size of arrow in vertical scroll bar.
    SM_CYHSCROLL = 3        '/* Y Size of arrow in horizontal scroll bar
    SM_CYCAPTION = 4        '/* Height of windows caption
    SM_CXBORDER = 5         '/* Width of no-sizable borders
    SM_CYBORDER = 6         '/* Height of non-sizable borders
    SM_CYVTHUMB = 9         '/* Height of scroll box on horizontal scroll bar
    SM_CXHTHUMB = 10        '/* Width of scroll box on horizontal scroll bar
    SM_CXICON = 11          '/* Width of standard icon
    SM_CYICON = 12          '/* Height of standard icon
    SM_CXCURSOR = 13        '/* Width of standard cursor
    SM_CYCURSOR = 14        '/* Height of standard cursor
    SM_CYMENU = 15          '/* Height of menu
    SM_CXFULLSCREEN = 16    '/* Width of client area of maximized window
    SM_CYFULLSCREEN = 17    '/* Height of client area of maximized window
    SM_CYKANJIWINDOW = 18   '/* Height of Kanji window
    SM_MOUSEPRESENT = 19    '/* True is a mouse is present
    SM_CYVSCROLL = 20       '/* Height of arrow in vertical scroll bar
    SM_CXHSCROLL = 21       '/* Width of arrow in vertical scroll bar
    SM_CXMIN = 28           '/* Minimum width of window
    SM_CYMIN = 29           '/* Minimum height of window
    SM_CXSIZE = 30          '/* Width of title bar bitmaps
    SM_CYSIZE = 31          '/* height of title bar bitmaps
    SM_CXFRAME = 32         '/* frame size x
    SM_CYFRAME = 33         '/* frame size bottom
    SM_CXMINTRACK = 34      '/* Minimum tracking width of window
    SM_CYMINTRACK = 35      '/* Minimum tracking height of window
    SM_CYSMCAPTION = 51     '/* height of window small caption
    SM_CXMINIMIZED = 57     '/* width of rectangle into which minimised windows must fit.
    SM_CYMINIMIZED = 58     '/* height of rectangle into which minimised windows must fit.
    SM_CXMAXTRACK = 59      '/* maximum width when resizing window
    SM_CYMAXTRACK = 60      '/* maximum width when resizing window
    SM_CXMAXIMIZED = 61     '/* default width of maximised window
    SM_CYMAXIMIZED = 62     '/* default height of maximised window
End Enum

Private Type VERSION_INFO
    dwOSVersionInfoSize                             As Long
    dwMajorVersion                                  As Long
    dwMinorVersion                                  As Long
    dwBuildNumber                                   As Long
    dwPlatformId                                    As Long
    szCSDVersion                                    As String * 128
End Type

Private Type POINTAPI
    x                                   As Long
    y                                   As Long
End Type

Private Type RECT
    left                                As Long
    top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type


Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal eIndex As SYSTEM_METRICS) As Long

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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long


Private m_bMaximized                    As Boolean
Private m_lHostHwnd                     As Long
Private m_bWin32                        As Boolean
Private m_tSizeState                    As RECT


Private Sub Window_Minimize()

    Window_Store

End Sub

Private Sub Window_Maximize()

Dim lX      As Long
Dim lY      As Long
Dim lW      As Long
Dim lH      As Long
Dim lCt     As Long
Dim tMax    As RECT
Dim tTemp   As RECT

    Window_Store
    Window_Metrics WD_MAXIMIZED, tMax
    
    m_bMaximized = True
    With tMax
        lW = .Right / 8
        lH = .Bottom / 8
    End With
    
    LSet tTemp = m_tSizeState
    With tTemp
        OffsetRect tTemp, -.left, -.top
        .left = m_tSizeState.left
        .top = m_tSizeState.top
    End With
    
    With tTemp
        lX = .left / 8
        lY = .top / 8
        For lCt = 1 To 9
            .top = .top - lY
            If .top < tMax.top Then .top = tMax.top
            .left = .left - lX
            If .left < tMax.left Then .left = tMax.left
            .Right = .Right + lW
            If .Right > tMax.Right Then .Right = tMax.Right
            .Bottom = .Bottom + lH
            If .Bottom > tMax.Bottom Then .Bottom = tMax.Bottom
            MoveWindow m_lHostHwnd, .left, .top, .Right, .Bottom, 1
            DoEvents
            Sleep 50
        Next lCt
        '/* if window is offscreen
        MoveWindow m_lHostHwnd, tMax.left, tMax.top, tMax.Right, tMax.Bottom, 1
    End With
    
End Sub

Private Sub Window_Restore()

Dim lX      As Long
Dim lY      As Long
Dim lW      As Long
Dim lH      As Long
Dim lCt     As Long
Dim tMax    As RECT
Dim tTemp   As RECT

    Window_Metrics WD_MAXIMIZED, tMax
    
    With m_tSizeState
        lX = .left / 8
        lY = .top / 8
    End With
    
    LSet tTemp = m_tSizeState
    With tTemp
        OffsetRect tTemp, -.left, -.top
    End With
    
    With tMax
        lW = .Right / 8
        lH = .Bottom / 8
        For lCt = 9 To 1 Step -1
            .top = .top + lY
            If .top > m_tSizeState.top Then .top = m_tSizeState.top
            .left = .left + lX
            If .left > m_tSizeState.left Then .left = m_tSizeState.left
            .Right = .Right - lW
            If .Right < tTemp.Right Then .Right = tTemp.Right
            .Bottom = .Bottom - lH
            If .Bottom < tTemp.Bottom Then .Bottom = tTemp.Bottom
            MoveWindow m_lHostHwnd, .left, .top, .Right, .Bottom, 1
            DoEvents
            Sleep 50
        Next lCt
        MoveWindow m_lHostHwnd, m_tSizeState.left, m_tSizeState.top, tTemp.Right, tTemp.Bottom, 1
    End With

    m_bMaximized = False

End Sub

Private Sub Window_Store()
    '/* store current size
    GetWindowRect m_lHostHwnd, m_tSizeState
End Sub

Private Sub Window_Metrics(ByRef eSizeState As WD_STATE, _
                           ByRef tRect As RECT)

    Select Case eSizeState
    Case WD_MINIMIZED
        With tRect
            .Bottom = 0
            .left = 0
            .Right = 0
            .top = 0
        End With
    
    Case WD_MAXIMIZED
        With tRect
            .Bottom = GetSystemMetrics(SM_CYFULLSCREEN)
            .left = 0
            .Right = GetSystemMetrics(SM_CXFULLSCREEN)
            .top = 0
        End With
    
    Case WD_NORMAL
        LSet tRect = m_tSizeState
    End Select

End Sub

Private Sub Window_Fade(ByVal lHwnd As Long, _
                        Optional ByVal bFadeOut As Boolean)

Dim lStyle  As Long
Dim lFdIdx  As Long
Dim lCt     As Long

    lStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
    SetWindowLong lHwnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED
    lFdIdx = 255
    For lCt = 9 To 0 Step -1
        SetLayeredWindowAttributes lHwnd, 0, lFdIdx, LWA_ALPHA
        lFdIdx = lFdIdx - 25
        DoEvents
        Sleep 100
    Next lCt
    SetWindowLong lHwnd, GWL_EXSTYLE, lStyle And Not WS_EX_LAYERED
    
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


Private Function Compatability_Check() As Boolean

Dim tVer  As VERSION_INFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    If tVer.dwMajorVersion >= 5 Then
        Compatability_Check = True
    End If

End Function


Private Sub Window_Restore2()

Dim tNrm    As RECT

    Window_Metrics WD_NORMAL, tNrm
    Dim lX, lY As Long
    
    m_bMaximized = False
    With tNrm
        MoveWindow m_lHostHwnd, .left, .top, .Right - .left, .Bottom - .top, 1
    End With
    
End Sub

Private Sub Window_Metrics2(ByRef eSizeState As WD_STATE, _
                           ByRef tRect As RECT)

    Select Case eSizeState
    Case WD_MINIMIZED
        With tRect
            .Bottom = m_lMinFormHeight
            .left = 0
            .Right = m_lMinFormWidth
            .top = 0
        End With
    
    Case WD_MAXIMIZED
        Dim lHDiff, lWDiff As Long
        lHDiff = (m_lCaptionMetric + m_lBottomMetric) - (m_lTopBorderHeight + (m_lBottomBorderHeight * 2))
        lWDiff = (m_lLeftMetric * 2) - (m_lLeftBorderWidth + m_lRightBorderWidth)
        With tRect
            .Bottom = (GetSystemMetrics(SM_CYFULLSCREEN) - lHDiff) + 4
            .left = (lWDiff / 2)
            .Right = GetSystemMetrics(SM_CXFULLSCREEN) - lWDiff
            .top = -4
        End With
    
    Case WD_NORMAL
        LSet tRect = m_tSizeState
    End Select

End Sub

Private Sub cmdSize_Click(Index As Integer)

    Select Case Index
    Case 0
        Window_Restore
    Case 1
        Window_Maximize
    End Select
    
End Sub

Private Sub Command1_Click()
    Window_Fade Me.hwnd
End Sub

Private Sub Form_Load()
    m_bWin32 = Compatability_Check
    m_lHostHwnd = Me.hwnd
End Sub
