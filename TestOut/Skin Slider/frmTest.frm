VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   2790
      TabIndex        =   6
      Top             =   2295
      Width           =   1365
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   1455
      Left            =   360
      TabIndex        =   5
      Top             =   1125
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   2566
      _Version        =   393216
      Orientation     =   1
   End
   Begin VB.PictureBox pichztrack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   630
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   75
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   765
      Width           =   300
   End
   Begin VB.PictureBox picvttrack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1395
      Picture         =   "frmTest.frx":016E
      ScaleHeight     =   300
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   585
      Width           =   75
   End
   Begin VB.PictureBox picvtthumb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   1395
      Picture         =   "frmTest.frx":02F0
      ScaleHeight     =   165
      ScaleWidth      =   630
      TabIndex        =   2
      Top             =   405
      Width           =   630
   End
   Begin VB.PictureBox pichzthumb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   630
      Picture         =   "frmTest.frx":08B2
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   1
      Top             =   405
      Width           =   330
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   405
      Left            =   1350
      TabIndex        =   0
      Top             =   1305
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_cSlider As clsRCMSlider


Private Sub Command1_Click()
Slider1.Visible = Not Slider1.Visible
End Sub

Private Sub Form_Load()

    Set m_cSlider = New clsRCMSlider
    With m_cSlider
        Set .p_OParentObj = Me
        Set .p_ISldHThumb = pichzthumb.Picture
        Set .p_ISldHTrack = pichztrack.Picture
        Set .p_ISldVThumb = picvtthumb.Picture
        Set .p_ISldVTrack = picvttrack.Picture
        .Attatch_Slider
    End With
    
End Sub
