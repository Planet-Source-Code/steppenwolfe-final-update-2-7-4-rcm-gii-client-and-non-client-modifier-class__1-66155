VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1545
      Left            =   4770
      ScaleHeight     =   1485
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   810
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   510
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   675
         Width           =   1410
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Enabled         =   0   'False
      Height          =   465
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1710
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1635
      Left            =   2295
      TabIndex        =   2
      Top             =   720
      Width           =   2085
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   900
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   645
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   810
      Width           =   1545
   End
   Begin VB.PictureBox picCmd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   90
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   5700
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   5700
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_NeoCmd As clsRCMCommand


Private Sub Form_Load()

    Set m_NeoCmd = New clsRCMCommand
    With m_NeoCmd
        Set .p_ICmdImg = picCmd.Picture
        Set .p_OParentObj = Me
        .p_RenderOffsetX = 6
        .p_ForeColor = vbWhite
        .p_TextAntiAliased = True
        .Attatch_Command
    End With

End Sub
