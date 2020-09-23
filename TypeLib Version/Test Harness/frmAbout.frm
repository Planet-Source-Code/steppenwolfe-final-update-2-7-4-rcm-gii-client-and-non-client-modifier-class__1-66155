VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RCM GII"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbout 
      Height          =   3690
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAbout.frx":57E2
      Top             =   315
      Width           =   5910
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   405
      Picture         =   "frmAbout.frx":57E8
      ScaleHeight     =   405
      ScaleWidth      =   3750
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   3750
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cRCMA   As New clsRCM


Private Sub Form_Load()
    LoadAbout
    LoadSkin p_Skin
End Sub

Public Sub LoadSkin(ByVal iStyle As Long)
'/* load skin

Dim lWidth As Long

On Error GoTo Handler

    '/* recreate the instance
    '/* needed to flush message queue
    If Not m_cRCMA Is Nothing Then
        Set m_cRCMA = Nothing
    End If
    Set m_cRCMA = New clsRCM
    
    Select Case iStyle
    '/* halo
    Case 0
        With m_cRCMA
            Set .p_ICaption = LoadResPicture("BAR-HALO", vbResBitmap)
            '/* bottom of frame
            Set .p_ICBottom = LoadResPicture("BOTTOM-HALO", vbResBitmap)
            '/* left side
            Set .p_ICLeft = LoadResPicture("LEFT-HALO", vbResBitmap)
            '/* right side
            Set .p_ICRight = LoadResPicture("RIGHT-HALO", vbResBitmap)
            '/* minimum btn
            Set .p_ICBoxMin = LoadResPicture("MINIMIZE-HALO", vbResBitmap)
            '/* maximum btn
            Set .p_ICBoxMax = LoadResPicture("MAXIMIZE-HALO", vbResBitmap)
            '/* restore btn
            Set .p_ICBoxRst = LoadResPicture("RESTORE-HALO", vbResBitmap)
            '/* close btn
            Set .p_ICBoxCls = LoadResPicture("CLOSE-HALO", vbResBitmap)
            '/* menu bar bg image
            Set .p_IMenuBarBg = LoadResPicture("MENUBARBG-HALO", vbResBitmap)
            '/* menu bg image
            Set .p_IMenuBg = LoadResPicture("MENUBG-HALO", vbResBitmap)
            '/* highlight bar image
            Set .p_IMenuRollover = LoadResPicture("MENUHIGHLITE-HALO", vbResBitmap)
            '/* host form
            Set .p_OParentObj = Me
            '/* custom shape
            .p_CustomCaption = True
            '/* tile start offset
            .p_CustomStartPos = 50
            '/* tile end offset
            .p_CustomEndPos = 160
            '/* form minimum sizes
            .p_MinFormWidth = 310
            .p_MinFormHeight = 74
            '/* use forms icon & caption text
            .p_UseFormCaption = True
            .p_UseFormIcon = True
            .p_CaptionFntClr = &HFDFDFD
            .p_CenterCaption = False
            .p_CaptionOffsetY = 3
            '/* form button offsets
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -14
            .p_ButtonOffsetY = 11
            '/* sizing handle borders
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 1

            '/* start
            .Attach
        End With
    '/* gt
    Case 1
        With m_cRCMA
            Set .p_ICaption = LoadResPicture("BAR-GT", vbResBitmap)
            '/* bottom of frame
            Set .p_ICBottom = LoadResPicture("BOTTOM-GT", vbResBitmap)
            '/* left side
            Set .p_ICLeft = LoadResPicture("LEFT-GT", vbResBitmap)
            '/* right side
            Set .p_ICRight = LoadResPicture("RIGHT-GT", vbResBitmap)
            '/* minimum btn
            Set .p_ICBoxMin = LoadResPicture("MINIMIZE-GT", vbResBitmap)
            '/* maximum btn
            Set .p_ICBoxMax = LoadResPicture("MAXIMIZE-GT", vbResBitmap)
            '/* restore btn
            Set .p_ICBoxRst = LoadResPicture("RESTORE-GT", vbResBitmap)
            '/* close btn
            Set .p_ICBoxCls = LoadResPicture("CLOSE-GT", vbResBitmap)
            '/* caption box
            Set .p_ICCapFrame = LoadResPicture("CAPTIONFRAME-GT", vbResBitmap)
            '/* menu bg image
            Set .p_IMenuBg = LoadResPicture("MENUBG-GT", vbResBitmap)
            '/* highlight bar image
            Set .p_IMenuBg = LoadResPicture("MENUHIGHLITE-GT", vbResBitmap)
            '/* menu bar bg image
            Set .p_IMenuBarBg = LoadResPicture("MENUBARBG-GT", vbResBitmap)
            '/* use caption box
            .p_CaptionFrame = True
            '/* host form
            Set .p_OParentObj = Me
            '/* custom shape
            .p_CustomCaption = True
            '/* tile start offset
            .p_CustomStartPos = 50
            '/* tile end offset
            .p_CustomEndPos = 200
            '/* tile start offset
            .p_MinFormWidth = 310
            .p_MinFormHeight = 74
            '/* use forms icon & caption text
            .p_UseFormCaption = True
            .p_UseFormIcon = True
            .p_CaptionFntClr = &HFDFDFD
            .p_CenterCaption = False
            .p_CaptionOffsetY = 3
            '/* form button offsets
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -3
            .p_ButtonOffsetY = 0
            '/* sizing handle borders
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 1
        
            '/* start
            .Attach
        End With
    '/* lime
    Case 2
        With m_cRCMA
            Set .p_ICaption = frmMain.picBar(0).Picture
            '/* bottom of frame
            Set .p_ICBottom = frmMain.picBottom(0).Picture
            '/* left side
            Set .p_ICLeft = frmMain.picLeft(0).Picture
            '/* right side
            Set .p_ICRight = frmMain.picRight(0).Picture
            '/* minimum btn
            Set .p_ICBoxMin = frmMain.picMin(0).Picture
            '/* maximum btn
            Set .p_ICBoxMax = frmMain.picMax(0).Picture
            '/* restore btn
            Set .p_ICBoxRst = frmMain.picRst(0).Picture
            '/* close btn
            Set .p_ICBoxCls = frmMain.picCls(0).Picture
            '/* host form
            Set .p_OParentObj = Me
            '/* use irregular frame shape /*
            .p_CustomCaption = True
            '/* tile start offset
            .p_CustomStartPos = 150
            '/* tile end offset
            .p_CustomEndPos = 160
            '/* use forms icon & caption text
            .p_UseFormCaption = True
            .p_UseFormIcon = True
            .p_CaptionFntClr = 438102
            .p_CaptionOffsetX = 25
            .p_CaptionOffsetY = 5
            '/* form button offsets
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -9
            .p_ButtonOffsetY = 4
            '/* sizing handle borders
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 1
            .p_FadeOutEffect = True
            
            .Attach
        End With
    '/* lg style
    Case 3
        With m_cRCMA
            Set .p_ICaption = LoadResPicture("BAR-LGESQUE", vbResBitmap)
            '/* bottom of frame
            Set .p_ICBottom = LoadResPicture("BOTTOM-LGESQUE", vbResBitmap)
            '/* left side
            Set .p_ICLeft = LoadResPicture("LEFT-LGESQUE", vbResBitmap)
            '/* right side
            Set .p_ICRight = LoadResPicture("RIGHT-LGESQUE", vbResBitmap)
            '/* minimum btn
            Set .p_ICBoxMin = LoadResPicture("MINIMIZE-LGESQUE", vbResBitmap)
            '/* maximum btn
            Set .p_ICBoxMax = LoadResPicture("MAXIMIZE-LGESQUE", vbResBitmap)
            '/* restore btn
            Set .p_ICBoxRst = LoadResPicture("RESTORE-LGESQUE", vbResBitmap)
            '/* close btn
            Set .p_ICBoxCls = LoadResPicture("CLOSE-LGESQUE", vbResBitmap)
            '/* menu bg image
            Set .p_IMenuBg = LoadResPicture("MENUBG-LGESQUE", vbResBitmap)
            '/* highlight bar image
            Set .p_IMenuRollover = LoadResPicture("MENUHIGHLITE-LGESQUE", vbResBitmap)
            '/* menu bar bg image
            Set .p_IMenuBarBg = LoadResPicture("MENUBARBG-LGESQUE", vbResBitmap)
            '/* host form
            Set .p_OParentObj = Me
            '/* standard shape
            .p_CustomCaption = False
            '/* tile start offset
            .p_MinFormWidth = 310
            '/* use forms icon & caption text
            .p_UseFormCaption = True
            .p_UseFormIcon = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CenterCaption = True
            .p_CaptionOffsetY = 5
            '/* form button offsets
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -8
            .p_ButtonOffsetY = 6
            '/* sizing handle borders
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 1

            '/* start
            .Attach
        End With
    End Select

On Error GoTo 0
Exit Sub

Handler:
    Set m_cRCMA = Nothing
    Debug.Print Err.Number & " " & Err.Description

End Sub

Private Sub LoadAbout()

Dim FF          As Long
Dim sBuffer     As String
Dim sPath       As String

On Error GoTo Handler

    sPath = App.Path & "\properties.txt"
    FF = FreeFile
    Open sPath For Binary Access Read As #FF
    sBuffer = Space$(LOF(FF))
    Get #FF, , sBuffer
    Close #FF
    txtAbout.Text = sBuffer

Exit Sub
Handler:
    Close #FF
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_cRCMA = Nothing
End Sub
