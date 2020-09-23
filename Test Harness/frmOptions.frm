VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   Caption         =   "More Controls.."
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   5625
      ScaleHeight     =   1665
      ScaleWidth      =   2385
      TabIndex        =   5
      Top             =   2295
      Width           =   2445
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   2970
      TabIndex        =   4
      Top             =   2295
      Width           =   2445
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   315
      TabIndex        =   3
      Top             =   2340
      Width           =   2445
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1725
      Left            =   5580
      TabIndex        =   2
      Top             =   405
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3043
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   2970
      TabIndex        =   1
      Top             =   405
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      Height          =   1725
      Left            =   315
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   405
      Width           =   2445
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cRCMO     As New clsRCM
Private m_cFrames   As clsRCMFrameStyle


Private Sub Form_Load()
    LoadSkin p_Skin
End Sub

Public Sub LoadSkin(ByVal iStyle As Long)
'/* load skin

Dim lWidth As Long

On Error GoTo Handler

    '/* recreate the instance
    '/* needed to flush message queue
    If Not m_cRCMO Is Nothing Then
        Set m_cRCMO = Nothing
    End If
    Set m_cRCMO = New clsRCM
    
    If Not m_cFrames Is Nothing Then
        Set m_cFrames = Nothing
    End If
    Set m_cFrames = New clsRCMFrameStyle
    
    With m_cFrames
        Set .p_OParentObj = Me
        .p_ColorFocused = vbGreen
        .p_ColorHover = vbBlue
        .p_ColorNormal = &H666666
        .p_FrameDirListBox = True
        .p_FrameFileListBox = True
        .p_FrameListBox = True
        .p_FramePictureBox = True
        .p_FrameTreeView = True
        .Attatch_Frame
    End With
    
    Select Case iStyle
    '/* halo
    Case 0
        With m_cRCMO
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
            '/* host form
            Set .p_OParentObj = Me
            '/* custom shape
            .p_CustomCaption = True
            '/* tile start offset
            .p_CustomStartPos = 50
            '/* tile end offset
            .p_CustomEndPos = 160
            '/* tile start offset
            .p_MinFormWidth = 310
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

            .Attach
        End With
    '/* gt
    Case 1
        With m_cRCMO
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
             '/* use caption title frame
            .p_CaptionFrame = True
            Set .p_ICCapFrame = LoadResPicture("CAPTIONFRAME-GT", vbResBitmap)
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

            .Attach
        End With
    '/* lime
    Case 2
        With m_cRCMO
            '/~ this one uses picture boxes so you can see
            '/~ the format.. Uses similar skin to WB4, very
            '/~ easy to make, and most of the metrics
            '/~ are calculated automatically..
            '/* caption bar <- using picturebox method, (res file is cleaner coding)
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

            .Attach
        End With
    '/* lg style
    Case 3
        With m_cRCMO
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
            '/* host form
            Set .p_OParentObj = Me
            '/* standard shape (false) paints with bitblt, custom with transparentblt
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
            .p_ButtonOffsetX = -9
            .p_ButtonOffsetY = 4
            '/* sizing handle borders
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 1

            '/* start
            .Attach
        End With
    End Select

Handler:
    Debug.Print Err.Number & " " & Err.Description
    On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_cFrames = Nothing
    Set m_cRCMO = Nothing
End Sub
