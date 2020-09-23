VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStoreDc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DIB_RGB_COLORS        As Long = &H0

Private Type SAFEARRAYBOUND
    cElements                       As Long
    lLbound                         As Long
End Type

Private Type SAFEARRAYID
    cDims                           As Integer
    fFeatures                       As Integer
    cbElements                      As Long
    cLocks                          As Long
    pvData                          As Long
    Bounds                          As SAFEARRAYBOUND
End Type

Private Type BITMAP
    bmType                          As Long
    bmWidth                         As Long
    bmHeight                        As Long
    bmWidthBytes                    As Long
    bmPlanes                        As Integer
    bmBitsPixel                     As Integer
    bmBits                          As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                          As Long
    biWidth                         As Long
    biHeight                        As Long
    biPlanes                        As Integer
    biBitCount                      As Integer
    biCompression                   As Long
    biSizeImage                     As Long
    biXPelsPerMeter                 As Long
    biYPelsPerMeter                 As Long
    biClrUsed                       As Long
    biClrImportant  As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, _
                                                                     ByVal Length As Long)

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                       lpDeviceName As Any, _
                                                                       lpOutput As Any, _
                                                                       lpInitData As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long

Private Declare Function TransparentBlt Lib "Msimg32.dll" (ByVal hdc As Long, _
                                                           ByVal x As Long, _
                                                           ByVal y As Long, _
                                                           ByVal nWidth As Long, _
                                                           ByVal nHeight As Long, _
                                                           ByVal hSrcDC As Long, _
                                                           ByVal xSrc As Long, _
                                                           ByVal ySrc As Long, _
                                                           ByVal nSrcWidth As Long, _
                                                           ByVal nSrcHeight As Long, _
                                                           ByVal crTransparent As Long) As Boolean

Private Declare Function CreateDIBSection32 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, _
                                                                                  lpBitsInfo As BITMAPINFOHEADER, _
                                                                                  ByVal wUsage As Long, _
                                                                                  lpBits As Long, _
                                                                                  ByVal handle As Long, _
                                                                                  ByVal dw As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, _
                                                       pBitmapInfo As Any, _
                                                       ByVal un As Long, _
                                                       lpVoid As Any, _
                                                       ByVal handle As Long, _
                                                       ByVal dw As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long


Private m_bUseAlpha         As Boolean
Private m_bInit             As Boolean
Private m_bMono             As Boolean
Private m_lHdc              As Long
Private m_hBmpOld           As Long
Private m_hBmp              As Long
Private m_lHandle           As Long
Private m_hDIB              As Long
Private m_lpBits            As Long
Private m_lWidth            As Long
Private m_lHeight           As Long
Private m_lSizeX            As Long
Private m_lSizeY            As Long
Private m_tBIH              As BITMAPINFOHEADER


Private Sub Class_Initialize()
    'm_bUseAlpha = True
End Sub

'/* use 32bit image
Public Property Get UseAlpha() As Boolean
    UseAlpha = m_bUseAlpha
End Property

Public Property Let UseAlpha(PropVal As Boolean)
    m_bUseAlpha = PropVal
End Property

Public Property Get hdc() As Long
    hdc = m_lHdc
End Property

Public Property Get Handle32() As Long
    Handle32 = m_lHandle
End Property

Public Property Get Bits() As Long
    Bits = m_lpBits
End Property

Public Property Let Bits(ByVal PropVal As Long)
    m_lpBits = PropVal
End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Property Let Height(ByVal lH As Long)
    If lH > m_lHeight Then
        ImageCreate m_lWidth, lH
    End If
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Let Width(ByVal lW As Long)
    If lW > m_lWidth Then
        ImageCreate lW, m_lHeight
    End If
End Property

Public Property Get Mono() As Boolean
    Mono = m_bMono
End Property

Public Property Let Mono(ByVal bState As Boolean)

    If Not (m_bMono = bState) Then
        m_bInit = True
    End If
    m_bMono = bState
    
End Property


Public Sub CreateFromPicture(oPicture As StdPicture)

Dim lhDCC       As Long
Dim lHdc        As Long
Dim lhBmpOld    As Long
Dim tBmp        As BITMAP

    If oPicture Is Nothing Then Exit Sub
    GetObjectAPI oPicture.handle, Len(tBmp), tBmp
    With tBmp
        Width = .bmWidth
        Height = .bmHeight
        lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        lHdc = CreateCompatibleDC(lhDCC)
        '/* use bitmap with alpha channel
        If m_bUseAlpha Then
            m_lHandle = ConvertTo32(oPicture.handle, tBmp, 255)
            If Not m_lHandle = 0 Then
                lhBmpOld = SelectObject(lHdc, m_lHandle)
            '/* default on failure
            Else
                lhBmpOld = SelectObject(lHdc, oPicture.handle)
            End If
        Else
            lhBmpOld = SelectObject(lHdc, oPicture.handle)
        End If
        '/* blit the image into dc
        BitBlt hdc, 0, 0, .bmWidth, .bmHeight, lHdc, 0, 0, vbSrcCopy
    End With
    
    SelectObject lHdc, lhBmpOld
    DeleteDC lHdc
    DeleteDC lhDCC

End Sub

Private Function ConvertTo32(ByVal lHsource As Long, _
                             ByRef tBmp As BITMAP, _
                             ByRef gAlpha As Byte) As Long

Dim aSBits()    As Byte
Dim i           As Long
Dim lHdc        As Long
Dim lhDIB       As Long
Dim hDIBold     As Long
Dim thDC        As Long
Dim tOldBmp     As Long
Dim uBIH        As BITMAPINFOHEADER
Dim uSSA        As SAFEARRAYID

    With uBIH
        .biBitCount = 32
        .biHeight = tBmp.bmHeight
        .biWidth = tBmp.bmWidth
        .biPlanes = 1
        .biSize = Len(uBIH)
    End With
    
    '/* create dib section
    With tBmp
        .bmWidthBytes = 4 * .bmWidth
    End With
    lHdc = CreateCompatibleDC(0)
    If lHdc = 0 Then Exit Function
    lhDIB = CreateDIBSection(lHdc, uBIH, DIB_RGB_COLORS, tBmp.bmBits, 0, 0)
    If lhDIB = 0 Then Exit Function
    m_lpBits = tBmp.bmBits
    
    thDC = CreateCompatibleDC(0)
    If thDC = 0 Then
        DeleteObject lhDIB
        lhDIB = 0
    Else
        hDIBold = SelectObject(lHdc, lhDIB)
        tOldBmp = SelectObject(thDC, lHsource)
        With tBmp
            BitBlt lHdc, 0, 0, .bmWidth, .bmHeight, thDC, 0, 0, vbSrcCopy
        End With
        '/* clean up
        SelectObject lHdc, hDIBold
        SelectObject thDC, tOldBmp
        DeleteDC thDC
        '/* build to 32bit
        If Not tBmp.bmBitsPixel = 32 Then
            With tBmp
                MapDIBits uSSA, aSBits(), .bmBits, .bmWidthBytes * .bmHeight
            End With
            For i = 3 To UBound(aSBits) Step 4
                aSBits(i) = gAlpha
            Next
            Call UnmapDIBits(aSBits)
            gAlpha = 255
        End If
    End If

    Call DeleteDC(lHdc)
    ConvertTo32 = lhDIB
    
End Function

Private Sub MapDIBits(uSA As SAFEARRAYID, _
                      aBits() As Byte, _
                      ByVal lpData As Long, _
                      ByVal lSize As Long)

    With uSA
        .cbElements = 1
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = lSize
        .pvData = lpData
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
    
End Sub

Private Sub UnmapDIBits(aBits() As Byte)
    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub

Private Sub ImageCreate(ByVal lW As Long, _
                        ByVal lH As Long)

Dim lHdc    As Long

    ImageDestroy
    lHdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    m_lHdc = CreateCompatibleDC(lHdc)
    m_hBmp = CreateCompatibleBitmap(lHdc, lW, lH)
    m_hBmpOld = SelectObject(m_lHdc, m_hBmp)
    If m_hBmpOld = 0 Then
        ImageDestroy
    Else
        m_lWidth = lW
        m_lHeight = lH
    End If
    DeleteDC lHdc

End Sub

Public Sub InjectBitmap(ByVal lhBmp As Long)

Dim tBM As BITMAP

    ImageDestroy
    GetObjectAPI lhBmp, Len(tBM), tBM
    With tBM
        Width = .bmWidth
        Height = .bmHeight
    End With
    If m_bUseAlpha Then
        m_lHandle = ConvertTo32(lhBmp, tBM, 255)
    End If
    Init lhBmp

    m_bInit = False

End Sub

Private Function Init(Optional ByVal hBmp As Long = 0) As Boolean

Dim hDCDisp     As Long
Dim lWidth      As Long
Dim lHeight     As Long

    If m_bMono Then
        If m_lHdc = 0 Then
            m_lHdc = CreateCompatibleDC(0)
        End If
    Else
        hDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        If Not hDCDisp = 0 Then
            If m_lHdc = 0 Then
                m_lHdc = CreateCompatibleDC(hDCDisp)
                If m_lHdc = 0 Then
                    Exit Function
                End If
            End If
        Else
            Exit Function
        End If
    End If

    If Not m_lHdc = 0 Then
        If Not m_hBmpOld = 0 Then
            SelectObject m_lHdc, m_hBmpOld
            m_hBmpOld = 0
        End If
        If Not m_hBmp = 0 Then
            DeleteObject m_hBmp
            m_hBmp = 0
        End If
        If hBmp = 0 Then
            If m_bMono Then
                m_hBmp = CreateCompatibleBitmap(m_lHdc, m_lSizeX, m_lSizeY)
            Else
                m_hBmp = CreateCompatibleBitmap(hDCDisp, m_lSizeX, m_lSizeY)
            End If
        Else
            m_hBmp = hBmp
        End If
        If m_hBmp = 0 Then
            DeleteDC hDCDisp
            hDCDisp = 0
        Else
            m_hBmpOld = SelectObject(m_lHdc, m_hBmp)
            Init = True
        End If
    Else
        DeleteDC hDCDisp
        hDCDisp = 0
    End If

    If Not (hDCDisp = 0) Then
        DeleteDC hDCDisp
    End If

End Function

Public Function ExtractBitmap() As Long

    If Not m_hBmpOld = 0 Then
        SelectObject m_lHdc, m_hBmpOld
        m_hBmpOld = 0
    End If
    If Not m_lHdc = 0 Then
        DeleteDC m_lHdc
        m_lHdc = 0
    End If
    ExtractBitmap = m_hBmp
    m_hBmp = 0
    m_bInit = True

End Function

Private Sub SetSize(ByVal lX As Long, _
                    ByVal lY As Long)

Dim bInit As Boolean
    If lX > m_lSizeX Or lY > m_lSizeY Then
        m_bInit = True
    End If
    m_lSizeX = lX
    m_lSizeY = lY

End Sub

Private Sub ImageDestroy()

    If Not m_hBmpOld = 0 Then
        SelectObject m_lHdc, m_hBmpOld
        m_hBmpOld = 0
    End If
    If Not m_hBmp = 0 Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    m_lWidth = 0
    m_lHeight = 0
    If Not m_lHdc = 0 Then
        DeleteDC m_lHdc
        m_lHdc = 0
    End If
    
    If Not m_lHandle = 0 Then
        DeleteObject m_lHandle
    End If
    m_lpBits = 0
    Call ZeroMemory(m_tBIH, Len(m_tBIH))
    
End Sub

Private Sub Class_Terminate()
    ImageDestroy
End Sub



