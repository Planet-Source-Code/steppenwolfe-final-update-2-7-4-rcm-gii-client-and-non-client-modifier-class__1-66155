VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2970
      TabIndex        =   9
      Top             =   1125
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   2565
      TabIndex        =   8
      Top             =   315
      Width           =   870
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Index           =   7
      Left            =   1395
      TabIndex        =   7
      Top             =   1440
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Index           =   5
      Left            =   1395
      TabIndex        =   6
      Top             =   900
      Width           =   870
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option2"
      Height          =   195
      Index           =   4
      Left            =   1395
      TabIndex        =   5
      Top             =   1170
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option2"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1710
      Width           =   915
   End
   Begin VB.PictureBox picOpt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   45
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   1560
      TabIndex        =   2
      Top             =   135
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option2"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1170
      Width           =   870
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   900
      Width           =   915
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_NeoOpt As clsRCMOptBtn


Private Sub Command1_Click()

    If Option1(1).Caption = "test" Then
        'm_NeoOpt.Set_Caption "return", Option1(1).hwnd
        Option1(1).Caption = "return"
    Else
        'm_NeoOpt.Set_Caption "test", Option1(1).hwnd
        Option1(1).Caption = "test"
    End If
    Option1(1).FontBold = Not Option1(1).FontBold

End Sub

Private Sub Form_Load()

    Set m_NeoOpt = New clsRCMOptBtn
    With m_NeoOpt
        Set .p_IOptImg = picOpt.Picture
        Set .p_OParentObj = Me
        .Attatch_OptBtn
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_NeoOpt = Nothing
End Sub
