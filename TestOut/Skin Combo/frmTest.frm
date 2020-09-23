VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   2340
      TabIndex        =   4
      Top             =   270
      Width           =   1320
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   855
      TabIndex        =   3
      Top             =   1665
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   855
      TabIndex        =   2
      Top             =   2205
      Width           =   2310
   End
   Begin VB.PictureBox picCombo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   1080
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTest.frx":0F72
      Left            =   855
      List            =   "frmTest.frx":0F74
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1125
      Width           =   2310
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_RCMCombo As clsRCMCombo


Private Sub Command1_Click()
Combo1.FontSize = 12
End Sub

Private Sub Form_Load()

    With Combo1
        .AddItem "test1"
        .AddItem "test2"
        .AddItem "test3"
        .AddItem "test4"
    End With
    
    With ImageCombo1
        .ComboItems.Add 1, , "test1"
        .ComboItems.Add 2, , "test2"
        .ComboItems.Add 3, , "test3"
        .ComboItems.Add 4, , "test4"
    End With
    
    Set m_RCMCombo = New clsRCMCombo
    With m_RCMCombo
        Set .p_IComboImg = picCombo.Picture
        Set .p_OParentObj = Me
        .p_FrameColor = &HFCAC65
        .p_FrameHighLite = &HFFE9D6
        .p_FrameStyle = FrameFlat
        .Attatch_ComboBox
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_RCMCombo = Nothing
End Sub

