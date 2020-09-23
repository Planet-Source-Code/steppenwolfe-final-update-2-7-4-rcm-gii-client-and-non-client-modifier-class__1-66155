VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwTest 
      Height          =   2535
      Left            =   4275
      TabIndex        =   12
      Top             =   810
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   420
      Left            =   6525
      TabIndex        =   11
      Top             =   5355
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   5175
      TabIndex        =   1
      Top             =   5310
      Width           =   1275
   End
   Begin VB.TextBox txtTest 
      Height          =   2535
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmTest.frx":0000
      Top             =   810
      Width           =   3840
   End
   Begin VB.PictureBox picvttrk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1755
      Picture         =   "frmTest.frx":0006
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   90
      Width           =   240
   End
   Begin VB.PictureBox picvtthumb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1440
      Picture         =   "frmTest.frx":04C8
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   90
      Width           =   240
   End
   Begin VB.PictureBox picvtbtnup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   855
      Picture         =   "frmTest.frx":098A
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   90
      Width           =   480
   End
   Begin VB.PictureBox picvtbtndwn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   270
      Picture         =   "frmTest.frx":0FCC
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   90
      Width           =   480
   End
   Begin VB.PictureBox picbtnlft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2160
      Picture         =   "frmTest.frx":160E
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   90
      Width           =   480
   End
   Begin VB.PictureBox picbtnrgt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2745
      Picture         =   "frmTest.frx":1C50
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   90
      Width           =   480
   End
   Begin VB.PictureBox pichzthumb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3375
      Picture         =   "frmTest.frx":2292
      ScaleHeight     =   240
      ScaleWidth      =   450
      TabIndex        =   8
      Top             =   90
      Width           =   450
   End
   Begin VB.PictureBox pichxtrk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3915
      Picture         =   "frmTest.frx":2894
      ScaleHeight     =   240
      ScaleWidth      =   360
      TabIndex        =   9
      Top             =   90
      Width           =   360
   End
   Begin VB.PictureBox picsizer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4410
      Picture         =   "frmTest.frx":2D56
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cScrollTxt    As clsRCMScrollBars

Private Sub Form_Load()

    LoadAbout
    LoadList
    Set m_cScrollTxt = New clsRCMScrollBars
    With m_cScrollTxt
        '/* vertical scrollbar
        Set .p_IVBtDwn = picvtbtndwn.Picture
        Set .p_IVBtUp = picvtbtnup.Picture
        Set .p_IVThumb = picvtthumb.Picture
        Set .p_IVTrack = picvttrk.Picture
        '/* horizontal scrollbar
        Set .p_IHBtLft = picbtnlft.Picture
        Set .p_IHBtRgt = picbtnrgt.Picture
        Set .p_IHThumb = pichzthumb.Picture
        Set .p_IHTrack = pichxtrk.Picture
        Set .p_ISizer = picsizer.Picture
        '.p_CtrlHnd = txtTest.hwnd
        Set .p_OScrollBarObj = lvwTest
        .Scrollbar_Attach
    End With

End Sub

Private Sub LoadList()

    With lvwTest
        .AllowColumnReorder = True
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Test Column 1", .Width - 100
        .ListItems.Add 1, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 2, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 3, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 4, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 5, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 6, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 7, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 8, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 9, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 10, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 11, , "testlistwidth testlistwidth testlistwidth testlistwidth "
        .ListItems.Add 12, , "testlistwidth testlistwidth testlistwidth testlistwidth "
    End With
    
End Sub

Private Sub LoadAbout()

Dim FF          As Long
Dim sBuffer     As String
Dim sPath       As String

On Error GoTo Handler

    sPath = App.Path & "\clsRCMScrollBars.cls"
    FF = FreeFile
    Open sPath For Binary Access Read As #FF
    sBuffer = Space$(LOF(FF))
    Get #FF, , sBuffer
    Close #FF
    txtTest.Text = sBuffer

Exit Sub
Handler:
    Close #FF
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_cScrollTxt = Nothing
End Sub
