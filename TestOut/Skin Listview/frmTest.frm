VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   315
      TabIndex        =   3
      Top             =   4005
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   4860
      TabIndex        =   2
      Top             =   3915
      Width           =   1410
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   3570
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   6297
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   4365
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0000
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0E54
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0FAE
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1E02
            Key             =   "DELETE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":239E
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":293A
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2A94
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2BEE
            Key             =   "PASTE"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   405
      Picture         =   "frmTest.frx":2D48
      ScaleHeight     =   300
      ScaleWidth      =   2400
      TabIndex        =   1
      Top             =   405
      Width           =   2400
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cHeader   As clsRCMListview


Private Sub Form_Load()
    LoadAbout
    LoadSkin
End Sub

Private Sub Command1_Click()

    With lvwData
        .ListItems.Add Text:="1) Checkboxes, option buttons, command buttons, imagecombo, and comboboxes are now skinned"
        .ListItems.Add Text:="2) Added transparent menus for 2K/XP"
        .ListItems.Add Text:="3) Menus now support many different styles, with additional properties added"
        .ListItems.Add Text:="4) Added transparency support for 98/ME"
        .ListItems.Add Text:="5) Added support for 32b bitmaps, and alpha rendering"
        .ListItems.Add Text:="6) Removed multi-threading, (crashes), for the time being"
        .ListItems.Add Text:="7) Images are now auto sampled for transparency color, (x-0, y-0 pixel)"
        .ListItems.Add Text:="8) Alt keyboard hook added, menu accelerators should operational in next rev"
        .ListItems.Add Text:="9) Transparency bug, alt key bug, and many others are addressed in this version"
        .ListItems.Add Text:="9) All paint operations moved to clsRender wrapper in preparation for type library in 2.5"
                .ListItems.Add Text:="1) Checkboxes, option buttons, command buttons, imagecombo, and comboboxes are now skinned"
        .ListItems.Add Text:="2) Added transparent menus for 2K/XP"
        .ListItems.Add Text:="3) Menus now support many different styles, with additional properties added"
        .ListItems.Add Text:="4) Added transparency support for 98/ME"
        .ListItems.Add Text:="5) Added support for 32b bitmaps, and alpha rendering"
        .ListItems.Add Text:="6) Removed multi-threading, (crashes), for the time being"
        .ListItems.Add Text:="7) Images are now auto sampled for transparency color, (x-0, y-0 pixel)"
        .ListItems.Add Text:="8) Alt keyboard hook added, menu accelerators should operational in next rev"
        .ListItems.Add Text:="9) Transparency bug, alt key bug, and many others are addressed in this version"
        .ListItems.Add Text:="9) All paint operations moved to clsRender wrapper in preparation for type library in 2.5"
                .ListItems.Add Text:="1) Checkboxes, option buttons, command buttons, imagecombo, and comboboxes are now skinned"
        .ListItems.Add Text:="2) Added transparent menus for 2K/XP"
        .ListItems.Add Text:="3) Menus now support many different styles, with additional properties added"
        .ListItems.Add Text:="4) Added transparency support for 98/ME"
        .ListItems.Add Text:="5) Added support for 32b bitmaps, and alpha rendering"
        .ListItems.Add Text:="6) Removed multi-threading, (crashes), for the time being"
        .ListItems.Add Text:="7) Images are now auto sampled for transparency color, (x-0, y-0 pixel)"
        .ListItems.Add Text:="8) Alt keyboard hook added, menu accelerators should operational in next rev"
        .ListItems.Add Text:="9) Transparency bug, alt key bug, and many others are addressed in this version"
        .ListItems.Add Text:="9) All paint operations moved to clsRender wrapper in preparation for type library in 2.5"
    End With

End Sub

Private Sub Command2_Click()
lvwData.Visible = Not lvwData.Visible
End Sub

Private Sub LoadSkin()

    Set m_cHeader = New clsRCMListview
    With m_cHeader
        Set .p_IHeader = picHeader.Picture
        Set .p_OListViewObj = Me.lvwData
        .p_TextForeColor = vbWhite
        .p_TextHighLite = vbRed
        .Attatch_Listview
    End With
    
End Sub

Private Sub LoadAbout()

    With lvwData
        .View = lvwReport
        .AllowColumnReorder = True
        Set .ColumnHeaderIcons = iml16
        .ColumnHeaders.Add 1, , "TestColumn 1", .Width / 3
        .ColumnHeaders.Add 2, , "TestColumn 2", .Width / 3
        .ColumnHeaders.Add 3, , "TestColumn 3", (.Width / 3) - 100
        .ColumnHeaders.Item(1).Icon = iml16.ListImages.Item("OPEN").Index
        .ColumnHeaders.Item(2).Icon = iml16.ListImages.Item("DELETE").Index
        .ColumnHeaders.Item(3).Icon = iml16.ListImages.Item("HELP").Index
        .ListItems.Add Text:="1) Checkboxes, option buttons, command buttons, imagecombo, and comboboxes are now skinned"
        .ListItems.Add Text:="2) Added transparent menus for 2K/XP"
        .ListItems.Add Text:="3) Menus now support many different styles, with additional properties added"
        .ListItems.Add Text:="4) Added transparency support for 98/ME"
        .ListItems.Add Text:="5) Added support for 32b bitmaps, and alpha rendering"
        .ListItems.Add Text:="6) Removed multi-threading, (crashes), for the time being"
        .ListItems.Add Text:="7) Images are now auto sampled for transparency color, (x-0, y-0 pixel)"
        .ListItems.Add Text:="8) Alt keyboard hook added, menu accelerators should operational in next rev"
        .ListItems.Add Text:="9) Transparency bug, alt key bug, and many others are addressed in this version"
        .ListItems.Add Text:="9) All paint operations moved to clsRender wrapper in preparation for type library in 2.5"
        .AllowColumnReorder = False
    End With
    
End Sub

Private Sub Form_Resize()

On Error Resume Next

    With lvwData
        .left = 100
        .Width = Me.ScaleWidth - 200
        .ColumnHeaders.Item(1).Width = .Width / 3
        .ColumnHeaders.Item(2).Width = .Width / 3
        .ColumnHeaders.Item(3).Width = (.Width / 3) - 100
    End With
    
On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_cHeader Is Nothing Then Set m_cHeader = Nothing
End Sub
