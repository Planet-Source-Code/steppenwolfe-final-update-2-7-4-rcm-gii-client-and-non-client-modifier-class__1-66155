VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   2610
      TabIndex        =   7
      Top             =   1710
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   2385
      TabIndex        =   6
      Top             =   675
      Width           =   1050
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000D&
      Caption         =   "Check2"
      Height          =   735
      Left            =   270
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   4
      Top             =   2025
      Value           =   2  'Grayed
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Check1xxxxxx"
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   3
      Top             =   1710
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000C000&
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   1395
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H000000FF&
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox picChk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   2340
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_NeoChk As clsRCMChkBox


Private Sub Form_Load()

    Set m_NeoChk = New clsRCMChkBox
    With m_NeoChk
        Set .p_IChkImg = picChk.Picture
        Set .p_OParentObj = Me
        .Attatch_ChkBox
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_NeoChk = Nothing
End Sub

