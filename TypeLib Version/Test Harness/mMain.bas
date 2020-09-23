Attribute VB_Name = "mMain"
Option Explicit

Private Const RDW_ALLCHILDREN           As Long = &H80
Private Const RDW_ERASE                 As Long = &H4
Private Const RDW_ERASENOW              As Long = &H200
Private Const RDW_FRAME                 As Long = &H400
Private Const RDW_INTERNALPAINT         As Long = &H2
Private Const RDW_INVALIDATE            As Long = &H1
Private Const RDW_NOCHILDREN            As Long = &H40
Private Const RDW_NOERASE               As Long = &H20
Private Const RDW_NOFRAME               As Long = &H800
Private Const RDW_NOINTERNALPAINT       As Long = &H10
Private Const RDW_UPDATENOW             As Long = &H100
Private Const RDW_VALIDATE              As Long = &H8
Private Const RDW_NRPGROUP              As Long = RDW_INVALIDATE Or _
    RDW_ERASE Or RDW_UPDATENOW Or RDW_ALLCHILDREN
    
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, _
                                                    lprcUpdate As Any, _
                                                    ByVal hrgnUpdate As Long, _
                                                    ByVal fuRedraw As Long) As Long

Private m_lIndex As Long

Public Property Get p_Skin() As Long
    p_Skin = m_lIndex
End Property

Public Property Let p_Skin(PropVal As Long)
    m_lIndex = PropVal
End Property

Public Sub Window_Redraw(lHwnd As Long)
'/* forced repaint
    RedrawWindow lHwnd, ByVal 0&, 0&, RDW_NRPGROUP
End Sub

Public Sub Reload(oFrm As Object)

    Unload oFrm
    oFrm.Show
    
End Sub

