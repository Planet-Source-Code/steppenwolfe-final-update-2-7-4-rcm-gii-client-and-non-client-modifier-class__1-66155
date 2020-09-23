VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/* vbAccelerator
'/* http://vbaccelerator.com/
'/* Copyright © 1999 Steve McMahon (steve@vbaccelerator.com)

Private Const CLR_INVALID     As Integer = -1

Private Type BITMAP '24 bytes
    bmType                        As Long
    bmWidth                       As Long
    bmHeight                      As Long
    bmWidthBytes                  As Long
    bmPlanes                      As Integer
    bmBitsPixel                   As Integer
    bmBits                        As Long
End Type

Private Type POINTAPI
    x                             As Long
    y                             As Long
End Type

Private Type RECT
    left                          As Long
    top                           As Long
    Right                         As Long
    Bottom                        As Long
End Type

Public Enum EPBRasterOperations
    PATCOPYEX = &HF00021
    PATINVERTEX = &H5A0049
    PATPAINTEX = &HFB0A09
End Enum


Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nXOrg As Long, _
                                                    ByVal nYOrg As Long, _
                                                    lppt As Any) As Long

Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal dwRop As Long) As Long

Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long

Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal crColor As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private m_hBrush              As Long
Private m_lPattern(0 To 3)    As Long


Private Sub Class_Initialize()

Dim i       As Long

    For i = 0 To 3
        m_lPattern(i) = &HAAAA5555
    Next i

End Sub

Private Sub Class_Terminate()

    Destroy
    Erase m_lPattern

End Sub

Public Function Create() As Boolean

Dim tbm         As BITMAP
Dim hBm         As Long

On Error GoTo Handler

    Destroy
    '/* Create a monochrome bitmap containing the desired pattern:
    With tbm
        .bmType = 0
        .bmWidth = 16
        .bmHeight = 8
        .bmWidthBytes = 2
        .bmPlanes = 1
        .bmBitsPixel = 1
        .bmBits = VarPtr(m_lPattern(0))
    End With
    hBm = CreateBitmapIndirect(tbm)
    '/* Make a brush from the bitmap bits
    m_hBrush = CreatePatternBrush(hBm)
    '// Delete the useless bitmap
    DeleteObject hBm

Handler:
    On Error GoTo 0

End Function

Public Sub Destroy()

    If Not m_hBrush = 0 Then
        DeleteObject m_hBrush
        m_hBrush = 0
    End If

End Sub

Public Sub DrawLine(ByVal lngHdc As Long, _
                    ByVal xPixels As Long, _
                    ByVal yPixels As Long, _
                    ByVal widthPixels As Long, _
                    ByVal heightPixels As Long, _
                    Optional ByVal eRop As EPBRasterOperations = PATCOPYEX, _
                    Optional ByVal oBackColor As OLE_COLOR = -1, _
                    Optional ByVal oForeColor As OLE_COLOR = -1)

Dim hOldBrush       As Long
Dim lBkColor        As Long
Dim lForeColor      As Long

    hOldBrush = SelectObject(lngHdc, m_hBrush)
    If Not (oBackColor = -1) Then
        lBkColor = SetBkColor(lngHdc, TranslateColor(oBackColor))
    End If
    
    If Not (oForeColor = -1) Then
        lForeColor = SetTextColor(lngHdc, TranslateColor(oForeColor))
    End If
    
    PatBlt lngHdc, xPixels, yPixels, widthPixels, heightPixels, eRop
    If Not (oBackColor = -1) Then
        SetBkColor lngHdc, lBkColor
    End If
    
    If Not (oForeColor = -1) Then
        SetTextColor lngHdc, lForeColor
    End If
    SelectObject lngHdc, hOldBrush

End Sub

Public Sub Rectangle(ByVal lngHdc As Long, _
                     ByVal xPixels As Long, _
                     ByVal yPixels As Long, _
                     ByVal widthRectPixels As Long, _
                     ByVal heightRectPixels As Long, _
                     ByVal lineSizePixels As Long, _
                     Optional ByVal eRop As EPBRasterOperations = PATCOPYEX, _
                     Optional ByVal bFill As Boolean = False, _
                     Optional ByVal oBackColor As OLE_COLOR = -1, _
                     Optional ByVal oForeColor As OLE_COLOR = -1)

Dim lBkColor        As Long
Dim lForeColor      As Long
Dim tR              As RECT
Dim hOldBrush       As Long

    If bFill Then
        If Not (oBackColor = -1) Then
            lBkColor = SetBkColor(lngHdc, TranslateColor(oBackColor))
        End If
        If Not (oForeColor = -1) Then
            lForeColor = SetTextColor(lngHdc, TranslateColor(oForeColor))
        End If
        With tR
            .left = xPixels
            .top = xPixels
            .Right = .left + widthRectPixels
            .Bottom = .top + heightRectPixels
        End With
        FillRect lngHdc, tR, m_hBrush
    Else
        hOldBrush = SelectObject(lngHdc, m_hBrush)
        If Not (oBackColor = -1) Then
            lBkColor = SetBkColor(lngHdc, TranslateColor(oBackColor))
        End If
        If Not (oForeColor = -1) Then
            lForeColor = SetTextColor(lngHdc, TranslateColor(oForeColor))
        End If
        PatBlt lngHdc, xPixels, yPixels, widthRectPixels, lineSizePixels, eRop
        PatBlt lngHdc, xPixels + widthRectPixels, yPixels, lineSizePixels, heightRectPixels, eRop
        PatBlt lngHdc, xPixels, yPixels + heightRectPixels, widthRectPixels, lineSizePixels, eRop
        PatBlt lngHdc, xPixels, yPixels, lineSizePixels, heightRectPixels, eRop
        If Not (oBackColor = -1) Then
            SetBkColor lngHdc, lBkColor
        End If
        If Not (oForeColor = -1) Then
            SetTextColor lngHdc, lForeColor
        End If
        SelectObject lngHdc, hOldBrush
    End If

End Sub

Public Sub SetBrushOrigin(ByVal lngHdc As Long, _
                          ByVal xPixels As Long, _
                          ByVal yPixels As Long)


    SetBrushOrgEx lngHdc, -xPixels And &H7&, -yPixels And &H7&, ByVal 0&

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

'/* Convert Automation color to Windows color

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function


