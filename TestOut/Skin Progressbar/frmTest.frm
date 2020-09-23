VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   2475
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Command1"
      Height          =   465
      Left            =   3150
      TabIndex        =   2
      Top             =   2295
      Width           =   1095
   End
   Begin VB.PictureBox picprogress 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   405
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   135
      ScaleWidth      =   90
      TabIndex        =   1
      Top             =   855
      Width           =   90
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   195
      Left            =   1215
      TabIndex        =   0
      Top             =   990
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cProgress As clsRCMProgressbar


Private Sub Form_Load()

    Set m_cProgress = New clsRCMProgressbar
    With m_cProgress
        Set .p_OProgressBarObj = pb1
        Set .p_IPBar = picprogress.Picture
        .p_BackColor = &HFFFFFF
        .p_BarColor = &H8000000D
        '.p_BarStyle = XPSTYLE
        .p_BarStyle = GRAPHICAL
        .Progressbar_Attach
    End With
    
    Timer1.Interval = 250
    
End Sub

Private Sub Command1_Click()
    Timer1.Enabled = Not Timer1.Enabled
    pb1.Visible = Not pb1.Visible
End Sub

Private Sub Timer1_Timer()

    With pb1
        If Not .Value = .Max Then
            .Value = .Value + 1
        Else
            .Value = .Min
        End If
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_cProgress = Nothing
End Sub
