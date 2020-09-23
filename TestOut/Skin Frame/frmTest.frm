VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   6975
      ScaleHeight     =   1035
      ScaleWidth      =   1980
      TabIndex        =   3
      Top             =   4905
      Width           =   2040
   End
   Begin VB.TextBox txtTest 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4140
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmTest.frx":0000
      Top             =   225
      Width           =   5745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   600
      Left            =   1170
      TabIndex        =   1
      Top             =   5310
      Width           =   1770
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   270
      Picture         =   "frmTest.frx":0006
      ScaleHeight     =   3300
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   225
      Width           =   3300
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cFrame      As clsRCMFrameStyle
Private cCtrlBg     As clsRCMWaterMark

Private Sub Command1_Click()
Picture2.Visible = Not Picture2.Visible
End Sub

Private Sub Form_Load()

    LoadText

    Set cFrame = New clsRCMFrameStyle
    With cFrame
        .p_FramePictureBox = True
        .p_ColorNormal = vbBlue
        .p_ColorHover = vbRed
        .p_ColorFocused = vbGreen
        Set .p_OParentObj = Me
        .Attatch_Frame
    End With
    
    Set cCtrlBg = New clsRCMWaterMark
    With cCtrlBg
        Set .p_IWaterMark = Picture1.Picture
        .p_WaterMarkPosition = WMK_CENTER
        Set .p_OTextBoxObj = txtTest
        .Attach
    End With

End Sub

Private Sub LoadText()

Dim FF          As Long
Dim sBuffer     As String

On Error GoTo Handler

    FF = FreeFile
    Open App.Path & "\clsRCMWaterMark.cls" For Binary Access Read As #FF
    sBuffer = Space$(LOF(FF))
    Get #FF, , sBuffer
    Close #FF
    txtTest.Text = sBuffer

Exit Sub
Handler:
    Close #FF

End Sub

Private Sub Form_Resize()

On Error Resume Next
   ' txtTest.Left = 10
   ' txtTest.Width = Me.ScaleWidth - 20
   ' txtTest.Top = 10
   ' txtTest.Height = Me.ScaleHeight - 1200
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not cCtrlBg Is Nothing Then Set cCtrlBg = Nothing
    If Not cFrame Is Nothing Then Set cFrame = Nothing
End Sub
