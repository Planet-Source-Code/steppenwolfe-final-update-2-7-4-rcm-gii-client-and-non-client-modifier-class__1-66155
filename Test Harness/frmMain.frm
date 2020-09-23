VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "RCM Generation II Ver 2.7.4"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picControls 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   11415
      TabIndex        =   39
      Tag             =   "CM"
      Top             =   7065
      Width           =   11415
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   7695
         TabIndex        =   76
         Tag             =   "CM"
         Top             =   90
         Width           =   3435
         Begin VB.CommandButton Command1 
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1935
            Picture         =   "frmMain.frx":57E2
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   180
            Width           =   1320
         End
         Begin VB.CommandButton Command2 
            Caption         =   "More.."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   360
            Picture         =   "frmMain.frx":1904C
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   180
            Width           =   1320
         End
      End
      Begin VB.OptionButton optSkin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Halo"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   43
         Tag             =   "CM"
         Top             =   450
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton optSkin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GT 3"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   42
         Tag             =   "CM"
         Top             =   450
         Width           =   825
      End
      Begin VB.OptionButton optSkin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lime"
         Height          =   195
         Index           =   2
         Left            =   1980
         TabIndex        =   41
         Tag             =   "CM"
         Top             =   450
         Width           =   780
      End
      Begin VB.OptionButton optSkin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LG-esque"
         Height          =   195
         Index           =   3
         Left            =   2835
         TabIndex        =   40
         Tag             =   "CM"
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Skins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   44
         Tag             =   "CM"
         Top             =   135
         Width           =   885
      End
   End
   Begin VB.PictureBox picMain 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   0
      ScaleHeight     =   7050
      ScaleWidth      =   11385
      TabIndex        =   0
      Tag             =   "CM"
      Top             =   0
      Width           =   11415
      Begin VB.Frame frDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "5 State Command Buttons"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Tag             =   "CM"
         Top             =   4635
         Width           =   2580
         Begin VB.CommandButton cmdTest 
            Caption         =   "Normal"
            Height          =   420
            Index           =   0
            Left            =   225
            Picture         =   "frmMain.frx":2C8B6
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   405
            Width           =   1410
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "Disabled"
            Enabled         =   0   'False
            Height          =   420
            Index           =   1
            Left            =   225
            Picture         =   "frmMain.frx":2D2B8
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   900
            Width           =   1410
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "Arial Font"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   225
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1395
            Width           =   1410
         End
      End
      Begin VB.Frame frDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Combo and ImageCombo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Index           =   1
         Left            =   8280
         TabIndex        =   24
         Tag             =   "CM"
         Top             =   4635
         Width           =   2670
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   180
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   450
            Width           =   2040
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   180
            TabIndex        =   25
            Top             =   1575
            Width           =   2040
         End
         Begin MSComctlLib.ImageCombo ImageCombo1 
            Height          =   330
            Left            =   180
            TabIndex        =   26
            Top             =   990
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Text            =   "ImageCombo1"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ComboBox"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   180
            TabIndex        =   30
            Tag             =   "CM"
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ImageCombo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   180
            TabIndex        =   29
            Tag             =   "CM"
            Top             =   810
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Drive ListBox"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   180
            TabIndex        =   28
            Tag             =   "CM"
            Top             =   1395
            Width           =   855
         End
      End
      Begin VB.Frame frDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Menu Options (on skin reload)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Index           =   2
         Left            =   225
         TabIndex        =   20
         Tag             =   "CM"
         Top             =   3105
         Width           =   2940
         Begin VB.CheckBox chkMenu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Transparent  (XP/2K)"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   23
            Tag             =   "CM"
            Top             =   315
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkMenu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Office XP Style"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   22
            Tag             =   "CM"
            Top             =   630
            Width           =   2085
         End
         Begin VB.CheckBox chkMenu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Image Rollover"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   21
            Tag             =   "CM"
            Top             =   945
            Value           =   1  'Checked
            Width           =   1725
         End
      End
      Begin VB.Frame frDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "12 State Check Boxes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Index           =   3
         Left            =   5580
         TabIndex        =   15
         Tag             =   "CM"
         Top             =   4635
         Width           =   2580
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Special font"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   19
            Tag             =   "CM"
            Top             =   1350
            Value           =   1  'Checked
            Width           =   1500
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disabled"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   18
            Tag             =   "CM"
            Top             =   1035
            Value           =   1  'Checked
            Width           =   1230
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Greyed"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   17
            Tag             =   "CM"
            Top             =   720
            Value           =   2  'Grayed
            Width           =   1230
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Normal"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Tag             =   "CM"
            Top             =   405
            Width           =   1230
         End
      End
      Begin VB.Frame frDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "8 State Option Buttons"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Index           =   4
         Left            =   2880
         TabIndex        =   10
         Tag             =   "CM"
         Top             =   4635
         Width           =   2580
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Special font"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   14
            Tag             =   "CM"
            Top             =   1395
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disabled"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   13
            Tag             =   "CM"
            Top             =   1080
            Width           =   1185
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Arial"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   12
            Tag             =   "CM"
            Top             =   765
            Width           =   1050
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Normal"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   11
            Tag             =   "CM"
            Top             =   450
            Width           =   1095
         End
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
         Height          =   2400
         Left            =   6030
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   405
         Width           =   5010
      End
      Begin VB.Frame frDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Progress Bar  (4 Styles)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Index           =   5
         Left            =   3285
         TabIndex        =   4
         Tag             =   "CM"
         Top             =   3105
         Width           =   2715
         Begin MSComctlLib.ProgressBar pb1 
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   5
            Top             =   450
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar pb1 
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   6
            Top             =   945
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Graphical"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   8
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "XP Style"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   7
            Top             =   765
            Width           =   645
         End
      End
      Begin VB.Frame frDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Slider Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Index           =   6
         Left            =   6075
         TabIndex        =   1
         Tag             =   "CM"
         Top             =   3105
         Width           =   2580
         Begin MSComctlLib.Slider Slider2 
            Height          =   1095
            Left            =   1935
            TabIndex        =   2
            Top             =   180
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   1931
            _Version        =   393216
            Orientation     =   1
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   555
            Left            =   90
            TabIndex        =   3
            Top             =   360
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   979
            _Version        =   393216
         End
      End
      Begin MSComctlLib.ListView lvwData 
         Height          =   2355
         Left            =   180
         TabIndex        =   35
         Top             =   405
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   4154
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":2DCBA
         Height          =   1500
         Left            =   8775
         TabIndex        =   38
         Tag             =   "CM"
         Top             =   3060
         Width           =   2445
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Textbox Watermark"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6030
         TabIndex        =   37
         Top             =   225
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scrollbar and Header"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   36
         Top             =   225
         Width           =   1530
      End
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   0
      Top             =   5850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DD60
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":415DA
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54E54
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":686CE
            Key             =   "CLOSE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BF48
            Key             =   "TEST1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F7C2
            Key             =   "TEST2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A303C
            Key             =   "TEST3"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8B26
            Key             =   "ABOUT"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BC3A0
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CFC1A
            Key             =   "PROPERTY"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E3494
            Key             =   "DESCRIPTION"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F6D0E
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBackground 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   4590
      Picture         =   "frmMain.frx":10A588
      ScaleHeight     =   2685
      ScaleWidth      =   3240
      TabIndex        =   45
      Top             =   4455
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.PictureBox picMenuBar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4590
      Picture         =   "frmMain.frx":126AE2
      ScaleHeight     =   360
      ScaleWidth      =   2250
      TabIndex        =   46
      Top             =   4050
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picRight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Index           =   0
      Left            =   4095
      Picture         =   "frmMain.frx":129584
      ScaleHeight     =   3465
      ScaleWidth      =   120
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picBottom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   585
      Picture         =   "frmMain.frx":12AB6E
      ScaleHeight     =   135
      ScaleWidth      =   2880
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   4275
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2970
      Picture         =   "frmMain.frx":12BFF0
      ScaleHeight     =   285
      ScaleWidth      =   855
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1170
      Picture         =   "frmMain.frx":12CCF6
      ScaleHeight     =   285
      ScaleWidth      =   855
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2070
      Picture         =   "frmMain.frx":12D9FC
      ScaleHeight     =   285
      ScaleWidth      =   855
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   270
      Picture         =   "frmMain.frx":12E702
      ScaleHeight     =   285
      ScaleWidth      =   855
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   270
      Picture         =   "frmMain.frx":12F408
      ScaleHeight     =   420
      ScaleWidth      =   3750
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox picMenuBarBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   270
      Picture         =   "frmMain.frx":13468A
      ScaleHeight     =   375
      ScaleWidth      =   3750
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox picCmd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4590
      Picture         =   "frmMain.frx":13903C
      ScaleHeight     =   315
      ScaleWidth      =   825
      TabIndex        =   55
      Top             =   315
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox picOpt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4590
      Picture         =   "frmMain.frx":139E46
      ScaleHeight     =   195
      ScaleWidth      =   1560
      TabIndex        =   56
      Top             =   1125
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picChk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4590
      Picture         =   "frmMain.frx":13AE60
      ScaleHeight     =   195
      ScaleWidth      =   2340
      TabIndex        =   57
      Top             =   1440
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.PictureBox picCombo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4590
      Picture         =   "frmMain.frx":13C666
      ScaleHeight     =   225
      ScaleWidth      =   900
      TabIndex        =   58
      Top             =   3600
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picWaterMark 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   8055
      Picture         =   "frmMain.frx":13D134
      ScaleHeight     =   3300
      ScaleWidth      =   3300
      TabIndex        =   59
      Top             =   270
      Width           =   3300
   End
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4590
      Picture         =   "frmMain.frx":1608A6
      ScaleHeight     =   255
      ScaleWidth      =   2400
      TabIndex        =   60
      Top             =   765
      Width           =   2400
   End
   Begin VB.PictureBox picsizer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5805
      Picture         =   "frmMain.frx":1628C8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   61
      Top             =   2970
      Width           =   240
   End
   Begin VB.PictureBox pichxtrk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5175
      Picture         =   "frmMain.frx":162C0A
      ScaleHeight     =   240
      ScaleWidth      =   360
      TabIndex        =   62
      Top             =   3195
      Width           =   360
   End
   Begin VB.PictureBox pichzthumb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5175
      Picture         =   "frmMain.frx":1630CC
      ScaleHeight     =   240
      ScaleWidth      =   450
      TabIndex        =   63
      Top             =   2925
      Width           =   450
   End
   Begin VB.PictureBox picbtnrgt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6255
      Picture         =   "frmMain.frx":1636CE
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   64
      Top             =   2655
      Width           =   480
   End
   Begin VB.PictureBox picbtnlft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5715
      Picture         =   "frmMain.frx":163D10
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   65
      Top             =   2655
      Width           =   480
   End
   Begin VB.PictureBox picvtbtndwn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4590
      Picture         =   "frmMain.frx":164352
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   66
      Top             =   2655
      Width           =   480
   End
   Begin VB.PictureBox picvtbtnup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5130
      Picture         =   "frmMain.frx":164994
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   67
      Top             =   2655
      Width           =   480
   End
   Begin VB.PictureBox picvtthumb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4590
      Picture         =   "frmMain.frx":164FD6
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   68
      Top             =   2925
      Width           =   240
   End
   Begin VB.PictureBox picvttrk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4860
      Picture         =   "frmMain.frx":165498
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   69
      Top             =   2925
      Width           =   240
   End
   Begin VB.PictureBox picprogress 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4635
      Picture         =   "frmMain.frx":16595A
      ScaleHeight     =   195
      ScaleWidth      =   120
      TabIndex        =   70
      Top             =   2340
      Width           =   120
   End
   Begin VB.PictureBox picLeft 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Index           =   0
      Left            =   90
      Picture         =   "frmMain.frx":165AD4
      ScaleHeight     =   3465
      ScaleWidth      =   120
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox pichzsldtrack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   5265
      Picture         =   "frmMain.frx":1670BE
      ScaleHeight     =   75
      ScaleWidth      =   300
      TabIndex        =   72
      Top             =   2115
      Width           =   300
   End
   Begin VB.PictureBox picvtsldtrack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4995
      Picture         =   "frmMain.frx":16722C
      ScaleHeight     =   300
      ScaleWidth      =   75
      TabIndex        =   73
      Top             =   1845
      Width           =   75
   End
   Begin VB.PictureBox picvtsldthmb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   5265
      Picture         =   "frmMain.frx":1673AE
      ScaleHeight     =   165
      ScaleWidth      =   630
      TabIndex        =   74
      Top             =   1890
      Width           =   630
   End
   Begin VB.PictureBox pichzsldthmb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4590
      Picture         =   "frmMain.frx":167970
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   75
      Top             =   1845
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   6480
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mF 
         Caption         =   "New File"
         Index           =   0
      End
      Begin VB.Menu mF 
         Caption         =   "Open"
         Index           =   1
      End
      Begin VB.Menu mF 
         Caption         =   "Save As"
         Index           =   2
      End
      Begin VB.Menu mF 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mF 
         Caption         =   "Exit"
         Index           =   4
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "Edit"
      Begin VB.Menu mEd 
         Caption         =   "Test 1"
         Index           =   0
      End
      Begin VB.Menu mEd 
         Caption         =   "Test 2"
         Index           =   1
      End
      Begin VB.Menu mEd 
         Caption         =   "Test 3"
         Index           =   2
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mH 
         Caption         =   "Contents"
         Index           =   0
      End
      Begin VB.Menu mH 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mH 
         Caption         =   "About"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cRCM              As clsRCM
Private m_cCommand          As clsRCMCommand
Private m_cCheckbox         As clsRCMChkBox
Private m_cOptionbtn        As clsRCMOptBtn
Private m_cCombobox         As clsRCMCombo
Private m_cWatermark        As clsRCMWaterMark
Private m_cListview         As clsRCMListview
Private m_cTxtScrollBars    As clsRCMScrollBars
Private m_cLvwScrollBars    As clsRCMScrollBars
Private m_cPrgBarGrp        As clsRCMProgressbar
Private m_cPrgBarXP         As clsRCMProgressbar
Private m_cSlider           As clsRCMSlider
Private m_lCurrentSkin      As Long


Private Sub Form_Load()

    optSkin(CInt(p_Skin)).Value = True
    LoadAbout
    LoadText
    ControlStyle p_Skin
    LoadSkin p_Skin

End Sub

Private Sub optSkin_Click(Index As Integer)

    If Index = p_Skin Then Exit Sub
    p_Skin = Index
    CleanUp
    ControlStyle p_Skin
    Window_Redraw Me.hwnd
    LoadSkin p_Skin
    Window_Redraw Me.hwnd
    
End Sub

Private Sub Command2_Click()
    frmOptions.Show
End Sub


Private Sub ControlStyle(ByVal lIndex As Long)

Dim lCtrl       As Control
Dim lBkColor    As Long
Dim lFrColor    As Long

On Error Resume Next

    Select Case lIndex
    Case 0
        lBkColor = &HC7C7C7
        lFrColor = &H222222
    Case 1
        lBkColor = &H666666
        lFrColor = &HFFFFFF
    Case 2
        lBkColor = &HAAAAAA
        lFrColor = &H444444
    Case 3
        lBkColor = &HA7A7A7
        lFrColor = &H222222
    End Select
    
    For Each lCtrl In Me
        If lCtrl.Tag = "CM" Then
            lCtrl.BackColor = lBkColor
            lCtrl.ForeColor = lFrColor
        End If
    Next lCtrl

On Error GoTo 0

End Sub

Private Sub Form_Resize()
'/* resize/repaint controls

On Error Resume Next

    picMain.Height = (Me.ScaleHeight - picControls.ScaleHeight)

On Error GoTo 0

End Sub

Private Sub Command1_Click()
    frmAbout.Show
End Sub

Public Sub LoadSkin(ByVal lStyle As Long)
'/* load skin

On Error GoTo Handler

    '/* recreate the instance is
    '/* needed to reset subclassing
    '/* and purge resources.
    '/* adding classes seperately means
    '/* seperate instances and much
    '/* better performance..
    
    If Not m_cRCM Is Nothing Then
        Set m_cRCM = Nothing
    End If
    Set m_cRCM = New clsRCM
    
    If Not m_cCheckbox Is Nothing Then
        Set m_cCheckbox = Nothing
    End If
    Set m_cCheckbox = New clsRCMChkBox
    
    If Not m_cOptionbtn Is Nothing Then
        Set m_cOptionbtn = Nothing
    End If
    Set m_cOptionbtn = New clsRCMOptBtn
    
    If Not m_cCombobox Is Nothing Then
        Set m_cCombobox = Nothing
    End If
    Set m_cCombobox = New clsRCMCombo
    
    If Not m_cWatermark Is Nothing Then
        Set m_cWatermark = Nothing
    End If
    Set m_cWatermark = New clsRCMWaterMark
    
    If Not m_cListview Is Nothing Then
        Set m_cListview = Nothing
    End If
    Set m_cListview = New clsRCMListview
    
    If Not m_cTxtScrollBars Is Nothing Then
        Set m_cTxtScrollBars = Nothing
    End If
    Set m_cTxtScrollBars = New clsRCMScrollBars
    
    If Not m_cLvwScrollBars Is Nothing Then
        Set m_cLvwScrollBars = Nothing
    End If
    Set m_cLvwScrollBars = New clsRCMScrollBars

    If Not m_cSlider Is Nothing Then
        Set m_cSlider = Nothing
    End If
    Set m_cSlider = New clsRCMSlider
    
    If Not m_cCommand Is Nothing Then
        Set m_cCommand = Nothing
    End If
    Set m_cCommand = New clsRCMCommand

    BuildProgressBars lStyle
    
    '/* style global
    m_lCurrentSkin = lStyle
    
    '/* add menu items and use iml image key
    '/* use vbaccel image list for nicer 32b icons..
    '/* menu imagelist
    With m_cRCM
        Set .p_MenuImageList = iml16
        .p_MenuIconIndex(mF(0).Caption) = iml16.ListImages.Item("NEW").Index - 1
        .p_MenuIconIndex(mF(1).Caption) = iml16.ListImages.Item("OPEN").Index - 1
        .p_MenuIconIndex(mF(2).Caption) = iml16.ListImages.Item("SAVE").Index - 1
        .p_MenuIconIndex(mF(4).Caption) = iml16.ListImages.Item("CLOSE").Index - 1
        
        .p_MenuIconIndex(mEd(0).Caption) = iml16.ListImages.Item("TEST1").Index - 1
        .p_MenuIconIndex(mEd(1).Caption) = iml16.ListImages.Item("TEST2").Index - 1
        .p_MenuIconIndex(mEd(2).Caption) = iml16.ListImages.Item("TEST3").Index - 1
        .p_MenuIconIndex(mH(0).Caption) = iml16.ListImages.Item("HELP").Index - 1
        .p_MenuIconIndex(mH(2).Caption) = iml16.ListImages.Item("ABOUT").Index - 1
      '  '/* transparency index (1-255)
        .p_MenuTransIdx = 235
        '/* turn on transparency
        .p_MenuTransparent = (chkMenu(0).Value = 1)
        '/* use custom captionbar rollover style
        .p_MenuRollOver = (chkMenu(1).Value = 1)
        '/* use xp style
        .p_OfficeXpStyle = (chkMenu(2).Value = 1)
    End With
    
    '/* scrollbars
    BuildScrollbars lStyle
    '/* create the watermark
    With m_cWatermark
        Set .p_IWaterMark = picWaterMark.Picture
        Set .p_OTextBoxObj = txtTest
        .p_WaterMarkPosition = WMK_CENTER
        .Attach
    End With
    
    Select Case lStyle
    '/* halo
    Case 0
        '/* listview headers
        With m_cListview
            Set .p_IHeader = LoadResPicture("LVWHDR-HALO", vbResBitmap)
            Set .p_OListViewObj = Me.lvwData
            .p_TextAntiAliased = True
            .Attatch_Listview
        End With
        '/* comboboxes
        With m_cCombobox
            Set .p_OParentObj = Me
            Set .p_IComboImg = LoadResPicture("COMBOBOX-HALO", vbResBitmap)
            .p_FrameColor = &H288FF
            .p_FrameHighLite = &HC7C7C7
            .p_FrameStyle = FrameFlat
            .Attatch_ComboBox
       End With
        '/* option button
        With m_cOptionbtn
            Set .p_OParentObj = Me
            Set .p_IOptImg = LoadResPicture("OPTION-HALO", vbResBitmap)
            .p_TransparentColor = &HFFFFFF
            .Attatch_OptBtn
        End With
        '/* checkbox
        With m_cCheckbox
            Set .p_OParentObj = Me
            Set .p_IChkImg = LoadResPicture("CHKBOX-HALO", vbResBitmap)
            .p_TransparentColor = &HFF00FF
            .Attatch_ChkBox
        End With
        '/* command button
        With m_cCommand
            Set .p_OParentObj = Me
            Set .p_ICmdImg = LoadResPicture("COMMAND-HALO", vbResBitmap)
            .p_RenderOffsetX = 6
            .p_ForeColor = vbWhite
            .p_TextAntiAliased = True
            .p_TransparentColor = -1
            .Attatch_Command
        End With
        '/* sliders
        With m_cSlider
            .p_BackColor = &HC7C7C7
            Set .p_OParentObj = Me
            Set .p_ISldHThumb = LoadResPicture("SLIDERTHUMBHZ-HALO", vbResBitmap)
            Set .p_ISldHTrack = LoadResPicture("SLIDERTRACKHZ-HALO", vbResBitmap)
            Set .p_ISldVThumb = LoadResPicture("SLIDERTHUMBVT-HALO", vbResBitmap)
            Set .p_ISldVTrack = LoadResPicture("SLIDERTRACKVT-HALO", vbResBitmap)
            .Attatch_Slider
        End With
        With m_cRCM
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
            '/* menu bar bg image
            Set .p_IMenuBarBg = LoadResPicture("MENUBARBG-HALO", vbResBitmap)
            '/* menu bg image
            Set .p_IMenuBg = LoadResPicture("MENUBG-HALO", vbResBitmap)
            '/* highlight bar image
            Set .p_IMenuRollover = LoadResPicture("MENUHIGHLITE-HALO", vbResBitmap)

            '/* host form
            Set .p_OParentObj = Me
            '/* custom shape
            .p_CustomCaption = True
            '/* tile start offset
            .p_CustomStartPos = 50
            '/* tile end offset
            .p_CustomEndPos = 160
            '/* form minimum sizes
            .p_MinFormWidth = 310
            .p_MinFormHeight = 74
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
            '/*** menu options ***/
            '/* use skinned menu (main switch)
            .p_MenuCustom = True
            .p_MenuActiveForeColor = &H26EFF
            .p_MenuBackgroundColor = &H424242
            '/* menu item highlte color
            .p_MenuActiveForeColor = &H333333
            '/* menu bg color
            .p_MenuBackgroundColor = &HC7C7C7
            '/* inactive item forecolor
            .p_MenuInActiveForeColor = &H4D4D4D
            '/* offset from left
            .p_MenuOffsetX = 10
            '/* offset from top center
            .p_MenuOffsetY = 4
            '/* caption rollover accent color
            .p_MenuRollOverColor = &H407E60
            '/* rollover style
            .p_MenuRollOverStyle = ECROLLOVEREX
            
            '/* accelerate process thread
            .p_ThreadAccel = True
            '/* start
            .Attach
        End With
        
    '/* gt
    Case 1
        '/* listview headers
        With m_cListview
            Set .p_IHeader = LoadResPicture("LVWHDR-GT", vbResBitmap)
            Set .p_OListViewObj = Me.lvwData
            .p_TextForeColor = &HEEEEEE
            .p_TextHighLite = &HFFFFFF
            .p_TextAntiAliased = True
            .Attatch_Listview
        End With
        '/* comboboxes
       With m_cCombobox
            Set .p_OParentObj = Me
            Set .p_IComboImg = LoadResPicture("COMBOBOX-GT", vbResBitmap)
            .p_FrameColor = &HFED3AC
            .p_FrameHighLite = &H666666
            .p_FrameStyle = FrameFlat
            .Attatch_ComboBox
        End With
        '/* option button
        With m_cOptionbtn
            Set .p_OParentObj = Me
            Set .p_IOptImg = LoadResPicture("OPTION-GT", vbResBitmap)
            .p_TransparentColor = &HFF00FF
            .Attatch_OptBtn
        End With
        '/* checkbox
        With m_cCheckbox
            Set .p_OParentObj = Me
            Set .p_IChkImg = LoadResPicture("CHKBOX-GT", vbResBitmap)
            .p_TransparentColor = &HFF00FF
            .Attatch_ChkBox
        End With
        '/* command buttons
        With m_cCommand
            '/* command button
            Set .p_OParentObj = Me
            Set .p_ICmdImg = LoadResPicture("COMMAND-GT", vbResBitmap)
            .p_RenderOffsetX = 6
            .p_ForeColor = &HDCDCDC
            .p_TextAntiAliased = True
            .p_TransparentColor = -1
            .Attatch_Command
        End With
        '/* slider
        With m_cSlider
            .p_BackColor = &H666666
            Set .p_OParentObj = Me
            Set .p_ISldHThumb = LoadResPicture("SLIDERTHUMBHZ-GT", vbResBitmap)
            Set .p_ISldHTrack = LoadResPicture("SLIDERTRACKHZ-GT", vbResBitmap)
            Set .p_ISldVThumb = LoadResPicture("SLIDERTHUMBVT-GT", vbResBitmap)
            Set .p_ISldVTrack = LoadResPicture("SLIDERTRACKVT-GT", vbResBitmap)
            .Attatch_Slider
        End With
        With m_cRCM
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
            '/* caption box
            Set .p_ICCapFrame = LoadResPicture("CAPTIONFRAME-GT", vbResBitmap)
            '/* menu bg image
            Set .p_IMenuBg = LoadResPicture("MENUBG-GT", vbResBitmap)
            '/* highlight bar image
            Set .p_IMenuBg = LoadResPicture("MENUHIGHLITE-GT", vbResBitmap)
            '/* menu bar bg image
            Set .p_IMenuBarBg = LoadResPicture("MENUBARBG-GT", vbResBitmap)
            '/* use caption box
            .p_CaptionFrame = True
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
            '/* show the menu
            .p_MenuCustom = True
            '/* menu item highlte color
            .p_MenuActiveForeColor = vbWhite
            '/* menu bg color
            .p_MenuBackgroundColor = &HAAAAAA
            '/* inactive item forecolor
            .p_MenuInActiveForeColor = &H111111
            '/* offset from left
            .p_MenuOffsetX = 10
            '/* offset from top center
            .p_MenuOffsetY = 5
            '/* caption rollover accent color
            .p_MenuRollOverColor = &H222222
            '/* rollover style
            .p_MenuRollOverStyle = ECGRADIENTEX
            
            '/* accelerate process thread
            .p_ThreadAccel = True
            '/* start
            .Attach
        End With

    '/* lime
    Case 2
        '/* listview headers
        With m_cListview
            Set .p_IHeader = picHeader.Picture
            Set .p_OListViewObj = Me.lvwData
            .p_TextAntiAliased = True
            .Attatch_Listview
        End With
        '/* comboboxes
        With m_cCombobox
            Set .p_OParentObj = Me
            Set .p_IComboImg = picCombo
            .p_FrameColor = &H25AC68
            .p_FrameHighLite = &HD46E7DB
            .p_FrameStyle = FrameFlat
            .Attatch_ComboBox
        End With
        '/* option buttons
        With m_cOptionbtn
            Set .p_OParentObj = Me
            Set .p_IOptImg = picOpt.Picture
           .p_TransparentColor = 16711935
            .Attatch_OptBtn
        End With
        '/* checkboxes
        With m_cCheckbox
            Set .p_OParentObj = Me
            Set .p_IChkImg = picChk.Picture
            .p_TransparentColor = &HFF00FF
            .Attatch_ChkBox
        End With
        '/* command buttons
        With m_cCommand
            Set .p_OParentObj = Me
            Set .p_ICmdImg = picCmd.Picture
            .p_RenderOffsetX = 6
            .p_ForeColor = &H333333
            .p_TextAntiAliased = True
            .p_TransparentColor = &HFFFFFF
            .Attatch_Command
        End With
        '/* slider
        With m_cSlider
            .p_BackColor = &HAAAAAA
            Set .p_OParentObj = Me
            Set .p_ISldHThumb = pichzsldthmb.Picture
            Set .p_ISldHTrack = pichzsldtrack.Picture
            Set .p_ISldVThumb = picvtsldthmb.Picture
            Set .p_ISldVTrack = picvtsldtrack.Picture
            .Attatch_Slider
        End With
        With m_cRCM
            '/~ this one uses picture boxes so you can see
            '/~ the format.. Uses similar skin to WB4, very
            '/~ easy to make, and most of the metrics
            '/~ are calculated automatically..
            '/* caption bar <- using picturebox method, (res file is cleaner coding)
            Set .p_ICaption = picBar(0).Picture
            '/* bottom of frame
            Set .p_ICBottom = picBottom(0).Picture
            '/* left side
            Set .p_ICLeft = picLeft(0).Picture
            '/* right side
            Set .p_ICRight = picRight(0).Picture
            '/* minimum btn
            Set .p_ICBoxMin = picMin(0).Picture
            '/* maximum btn
            Set .p_ICBoxMax = picMax(0).Picture
            '/* restore btn
            Set .p_ICBoxRst = picRst(0).Picture
            '/* close btn
            Set .p_ICBoxCls = picCls(0).Picture
            '/* menu bar bg image
            Set .p_IMenuBarBg = picMenuBarBg(0).Picture
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
            
            '/* show the menu
            .p_MenuCustom = True
            '/* menu item highlte color
            .p_MenuActiveForeColor = &HAAAAAA
            '/* menu bg color
            .p_MenuBackgroundColor = &H474747
            '/* inactive item forecolor
            .p_MenuInActiveForeColor = &HFFFFFF
            '/* offset from left
            .p_MenuOffsetX = 15
            '/* offset from top center
            .p_MenuOffsetY = 5
            '/* caption rollover accent color
            .p_MenuRollOverColor = &HCCCCCC
            '/* rollover style
            .p_MenuRollOverStyle = ECBUTTONEX
            .p_OfficeXpStyle = True
            
            '/* accelerate process thread
            .p_ThreadAccel = True
            '/* start
            .Attach
        End With

    '/* lg style
    Case 3
        '/* listview headers
        With m_cListview
            Set .p_IHeader = LoadResPicture("LVWHDR-LGESQUE", vbResBitmap)
            Set .p_OListViewObj = Me.lvwData
            .p_TextHighLite = &H333333
            .p_TextForeColor = &H999999
            .p_TextAntiAliased = True
            .Attatch_Listview
        End With
        '/* comboboxes
        With m_cCombobox
            Set .p_OParentObj = Me
            Set .p_IComboImg = LoadResPicture("COMBOBOX-LGESQUE", vbResBitmap)
            .p_FrameColor = &HFFE9D6
            .p_FrameHighLite = &HFED3AC
            .p_FrameStyle = FrameFlat
            .Attatch_ComboBox
        End With
        '/* option buttons
        With m_cOptionbtn
            Set .p_OParentObj = Me
            Set .p_IOptImg = LoadResPicture("OPTION-LGESQUE", vbResBitmap)
            .p_TransparentColor = &HFFFFFF
            .Attatch_OptBtn
        End With
        '/* checkboxes
        With m_cCheckbox
            Set .p_OParentObj = Me
            Set .p_IChkImg = LoadResPicture("CHKBOX-LGESQUE", vbResBitmap)
            .p_TransparentColor = &HFF00FF
            .Attatch_ChkBox
        End With
        '/* command buttons
        With m_cCommand
            Set .p_OParentObj = Me
            Set .p_ICmdImg = LoadResPicture("COMMAND-LGESQUE", vbResBitmap)
            .p_RenderOffsetX = 6
            .p_ForeColor = &H333333
            .p_ColorHiLite = &H444444
            .p_TextAntiAliased = True
            .p_TransparentColor = -1
            .Attatch_Command
        End With
        '/* slider
        With m_cSlider
            .p_BackColor = &HA7A7A7
            Set .p_OParentObj = Me
            Set .p_ISldHThumb = LoadResPicture("SLIDERTHUMBHZ-LGESQUE", vbResBitmap)
            Set .p_ISldHTrack = LoadResPicture("SLIDERTRACKHZ-LGESQUE", vbResBitmap)
            Set .p_ISldVThumb = LoadResPicture("SLIDERTHUMBVT-LGESQUE", vbResBitmap)
            Set .p_ISldVTrack = LoadResPicture("SLIDERTRACKVT-LGESQUE", vbResBitmap)
            .Attatch_Slider
        End With
        With m_cRCM
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
            '/* menu bg image
            Set .p_IMenuBg = LoadResPicture("MENUBG-LGESQUE", vbResBitmap)
            '/* highlight bar image
            Set .p_IMenuRollover = LoadResPicture("MENUHIGHLITE-LGESQUE", vbResBitmap)
            '/* menu bar bg image
            Set .p_IMenuBarBg = LoadResPicture("MENUBARBG-LGESQUE", vbResBitmap)
            
            '/* host form
            Set .p_OParentObj = Me
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
            .p_ButtonOffsetX = -8
            .p_ButtonOffsetY = 6
            '/* sizing handle borders
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 1

            '/* show the menu
            .p_MenuCustom = True
            '/* menu item highlte color
            .p_MenuActiveForeColor = vbWhite
            '/* menu bg color
            .p_MenuBackgroundColor = &HAAAAAA
            '/* inactive item forecolor
            .p_MenuInActiveForeColor = &H222222
            '/* offset from left
            .p_MenuOffsetX = 10
            '/* offset from top center
            .p_MenuOffsetY = 5
            '/* caption rollover accent color
            .p_MenuRollOverColor = &HA7A7A7
            '/* rollover style
            .p_MenuRollOverStyle = ECROLLOVEREX

            '/* accelerate process thread
            .p_ThreadAccel = True
            '/* start
            .Attach
        End With
    End Select

Exit Sub
Handler:
    Debug.Print Err.Number & " " & Err.Description
    On Error GoTo 0

End Sub

Private Sub BuildScrollbars(ByVal lIndex As Long)

On Error Resume Next

    Select Case lIndex
    '/* halo
    Case 0
        With m_cTxtScrollBars
            '/* vertical scrollbar
            Set .p_IVBtDwn = LoadResPicture("SCRLBTDWNVT-HALO", vbResBitmap)
            Set .p_IVBtUp = LoadResPicture("SCRLBTUPVT-HALO", vbResBitmap)
            Set .p_IVThumb = LoadResPicture("SCRLTHUMBVT-HALO", vbResBitmap)
            Set .p_IVTrack = LoadResPicture("SCRLTRACKVT-HALO", vbResBitmap)
            '/* horizontal scrollbar
            Set .p_IHBtLft = LoadResPicture("SCRLBTLFTHZ-HALO", vbResBitmap)
            Set .p_IHBtRgt = LoadResPicture("SCRLBTRGTHZ-HALO", vbResBitmap)
            Set .p_IHThumb = LoadResPicture("SCRLTHUMBHZ-HALO", vbResBitmap)
            Set .p_IHTrack = LoadResPicture("SCRLTRACKHZ-HALO", vbResBitmap)
            Set .p_ISizer = LoadResPicture("SCRLSIZER-HALO", vbResBitmap)
            Set .p_OScrollBarObj = txtTest
            .Scrollbar_Attach
        End With
        With m_cLvwScrollBars
            '/* vertical scrollbar
            Set .p_IVBtDwn = LoadResPicture("SCRLBTDWNVT-HALO", vbResBitmap)
            Set .p_IVBtUp = LoadResPicture("SCRLBTUPVT-HALO", vbResBitmap)
            Set .p_IVThumb = LoadResPicture("SCRLTHUMBVT-HALO", vbResBitmap)
            Set .p_IVTrack = LoadResPicture("SCRLTRACKVT-HALO", vbResBitmap)
            '/* horizontal scrollbar
            Set .p_IHBtLft = LoadResPicture("SCRLBTLFTHZ-HALO", vbResBitmap)
            Set .p_IHBtRgt = LoadResPicture("SCRLBTRGTHZ-HALO", vbResBitmap)
            Set .p_IHThumb = LoadResPicture("SCRLTHUMBHZ-HALO", vbResBitmap)
            Set .p_IHTrack = LoadResPicture("SCRLTRACKHZ-HALO", vbResBitmap)
            Set .p_ISizer = LoadResPicture("SCRLSIZER-HALO", vbResBitmap)
            Set .p_OScrollBarObj = lvwData
            .Scrollbar_Attach
        End With
    '/* gt
    Case 1
        With m_cTxtScrollBars
            '/* vertical scrollbar
            Set .p_IVBtDwn = LoadResPicture("SCRLBTDWNVT-GT", vbResBitmap)
            Set .p_IVBtUp = LoadResPicture("SCRLBTUPVT-GT", vbResBitmap)
            Set .p_IVThumb = LoadResPicture("SCRLTHUMBVT-GT", vbResBitmap)
            Set .p_IVTrack = LoadResPicture("SCRLTRACKVT-GT", vbResBitmap)
            '/* horizontal scrollbar
            Set .p_IHBtLft = LoadResPicture("SCRLBTLFTHZ-GT", vbResBitmap)
            Set .p_IHBtRgt = LoadResPicture("SCRLBTRGTHZ-GT", vbResBitmap) '
            Set .p_IHThumb = LoadResPicture("SCRLTHUMBHZ-GT", vbResBitmap)
            Set .p_IHTrack = LoadResPicture("SCRLTRACKHZ-GT", vbResBitmap)
            Set .p_ISizer = LoadResPicture("SCRLSIZER-GT", vbResBitmap)
            Set .p_OScrollBarObj = txtTest
            .Scrollbar_Attach
        End With
        With m_cLvwScrollBars
            '/* vertical scrollbar
            Set .p_IVBtDwn = LoadResPicture("SCRLBTDWNVT-GT", vbResBitmap)
            Set .p_IVBtUp = LoadResPicture("SCRLBTUPVT-GT", vbResBitmap)
            Set .p_IVThumb = LoadResPicture("SCRLTHUMBVT-GT", vbResBitmap)
            Set .p_IVTrack = LoadResPicture("SCRLTRACKVT-GT", vbResBitmap)
            '/* horizontal scrollbar
            Set .p_IHBtLft = LoadResPicture("SCRLBTLFTHZ-GT", vbResBitmap)
            Set .p_IHBtRgt = LoadResPicture("SCRLBTRGTHZ-GT", vbResBitmap)
            Set .p_IHThumb = LoadResPicture("SCRLTHUMBHZ-GT", vbResBitmap)
            Set .p_IHTrack = LoadResPicture("SCRLTRACKHZ-GT", vbResBitmap)
            Set .p_ISizer = LoadResPicture("SCRLSIZER-GT", vbResBitmap)
            Set .p_OScrollBarObj = lvwData
            .Scrollbar_Attach
        End With
    '/* lime
    Case 2
        With m_cTxtScrollBars
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
            Set .p_OScrollBarObj = txtTest
            .Scrollbar_Attach
        End With
        With m_cLvwScrollBars
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
            Set .p_OScrollBarObj = lvwData
            .Scrollbar_Attach
        End With
    '/* lg
    Case 3
        With m_cTxtScrollBars
            '/* vertical scrollbar
            Set .p_IVBtDwn = LoadResPicture("SCRLBTDWNVT-LGESQUE", vbResBitmap)
            Set .p_IVBtUp = LoadResPicture("SCRLBTUPVT-LGESQUE", vbResBitmap)
            Set .p_IVThumb = LoadResPicture("SCRLTHUMBVT-LGESQUE", vbResBitmap)
            Set .p_IVTrack = LoadResPicture("SCRLTRACKVT-LGESQUE", vbResBitmap)
            '/* horizontal scrollbar
            Set .p_IHBtLft = LoadResPicture("SCRLBTLFTHZ-LGESQUE", vbResBitmap)
            Set .p_IHBtRgt = LoadResPicture("SCRLBTRGTHZ-LGESQUE", vbResBitmap)
            Set .p_IHThumb = LoadResPicture("SCRLTHUMBHZ-LGESQUE", vbResBitmap)
            Set .p_IHTrack = LoadResPicture("SCRLTRACKHZ-LGESQUE", vbResBitmap)
            Set .p_ISizer = LoadResPicture("SCRLSIZER-LGESQUE", vbResBitmap)
            Set .p_OScrollBarObj = txtTest
            .Scrollbar_Attach
        End With
        With m_cLvwScrollBars
            '/* vertical scrollbar
            Set .p_IVBtDwn = LoadResPicture("SCRLBTDWNVT-LGESQUE", vbResBitmap)
            Set .p_IVBtUp = LoadResPicture("SCRLBTUPVT-LGESQUE", vbResBitmap)
            Set .p_IVThumb = LoadResPicture("SCRLTHUMBVT-LGESQUE", vbResBitmap)
            Set .p_IVTrack = LoadResPicture("SCRLTRACKVT-LGESQUE", vbResBitmap)
            '/* horizontal scrollbar
            Set .p_IHBtLft = LoadResPicture("SCRLBTLFTHZ-LGESQUE", vbResBitmap)
            Set .p_IHBtRgt = LoadResPicture("SCRLBTRGTHZ-LGESQUE", vbResBitmap)
            Set .p_IHThumb = LoadResPicture("SCRLTHUMBHZ-LGESQUE", vbResBitmap)
            Set .p_IHTrack = LoadResPicture("SCRLTRACKHZ-LGESQUE", vbResBitmap)
            Set .p_ISizer = LoadResPicture("SCRLSIZER-LGESQUE", vbResBitmap)
            Set .p_OScrollBarObj = lvwData
            .Scrollbar_Attach
        End With
    End Select
    
End Sub

Private Sub BuildProgressBars(ByVal lIndex As Long)

    Set m_cPrgBarGrp = New clsRCMProgressbar
    Set m_cPrgBarXP = New clsRCMProgressbar
    On Error Resume Next
    Select Case lIndex
    '/* halo
    Case 0
        With m_cPrgBarGrp
            Set .p_OProgressBarObj = pb1(0)
            Set .p_IPBar = LoadResPicture("PROGRESS-HALO", vbResBitmap)
            .p_BackColor = &HC7C7C7
            .p_BarStyle = GRAPHICAL
            .Progressbar_Attach
        End With
        With m_cPrgBarXP
            Set .p_OProgressBarObj = pb1(1)
            .p_BackColor = &HFFFFFF
            .p_BarColor = &H28D22B
            .p_BarStyle = XPSTYLE
            .Progressbar_Attach
        End With
    '/* gt
    Case 1
        With m_cPrgBarGrp
            Set .p_OProgressBarObj = pb1(0)
            Set .p_IPBar = LoadResPicture("PROGRESS-GT", vbResBitmap)
            .p_BackColor = &H666666
            .p_BarStyle = GRAPHICAL
            .Progressbar_Attach
        End With
        With m_cPrgBarXP
            Set .p_OProgressBarObj = pb1(1)
            .p_BackColor = &HFFFFFF
            .p_BarColor = &H28D22B
            .p_BarStyle = XPSTYLE
            .Progressbar_Attach
        End With
    '/* lime
    Case 2
        With m_cPrgBarGrp
            Set .p_OProgressBarObj = pb1(0)
            Set .p_IPBar = picprogress.Picture
            .p_BackColor = &HFFFFFF
            .p_BarStyle = GRAPHICAL
            .Progressbar_Attach
        End With
        With m_cPrgBarXP
            Set .p_OProgressBarObj = pb1(1)
            .p_BackColor = &HAAAAAA
            .p_BarColor = &H28D22B
            .p_BarStyle = XPSTYLE
            .Progressbar_Attach
        End With
    '/* lg
    Case 3
        With m_cPrgBarGrp
            Set .p_OProgressBarObj = pb1(0)
            Set .p_IPBar = LoadResPicture("PROGRESS-LGESQUE", vbResBitmap)
            .p_BackColor = &HA7A7A7
            .p_BarStyle = GRAPHICAL
            .Progressbar_Attach
        End With
        With m_cPrgBarXP
            Set .p_OProgressBarObj = pb1(1)
            .p_BackColor = &HFFFFFF
            .p_BarColor = &H28D22B
            .p_BarStyle = XPSTYLE
            .Progressbar_Attach
        End With
    End Select

    With Timer1
        .Interval = 250
        .Enabled = True
    End With
    
End Sub

Private Sub DestroyProgressBars()

    Timer1.Enabled = False
    If Not m_cPrgBarGrp Is Nothing Then Set m_cPrgBarGrp = Nothing
    If Not m_cPrgBarXP Is Nothing Then Set m_cPrgBarXP = Nothing
    
End Sub

Private Sub Timer1_Timer()

    With pb1(0)
        If Not .Value = .Max Then
            .Value = .Value + 1
        Else
            .Value = .Min
        End If
    End With
    With pb1(1)
        If Not .Value = .Max Then
            .Value = .Value + 1
        Else
            .Value = .Min
        End If
    End With
    
End Sub

Private Sub LoadAbout()
'/* properties list

Dim lItem As ListItem

    With lvwData
        .View = lvwReport
        .AllowColumnReorder = True
        Set .ColumnHeaderIcons = iml16
        .ColumnHeaders.Add 1, , "Property", .Width / 4
        .ColumnHeaders.Add 2, , "Desc", .Width / 2
        .ColumnHeaders.Add 3, , "Val", (.Width / 4) - 150
        .ColumnHeaders.Item(1).Icon = iml16.ListImages.Item("PROPERTY").Index
        .ColumnHeaders.Item(2).Icon = iml16.ListImages.Item("DESCRIPTION").Index
        .ColumnHeaders.Item(3).Icon = iml16.ListImages.Item("TYPE").Index
        .AllowColumnReorder = False
        Set lItem = .ListItems.Add(Text:="p_BottomSizingBorder")
        lItem.SubItems(1) = "bottom sizing border"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_ButtonOffsetX")
        lItem.SubItems(1) = "control button offset vert"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_ButtonOffsetY")
        lItem.SubItems(1) = "control button offset vert"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CustomCaption")
        lItem.SubItems(1) = "use owner drawn caption text"
        lItem.SubItems(2) = "str"
        Set lItem = .ListItems.Add(Text:="p_CaptionFntClr")
        lItem.SubItems(1) = "caption text font color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CaptionFrame")
        lItem.SubItems(1) = "use a caption bar text frame image"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_CaptionOffset")
        lItem.SubItems(1) = "caption title offset"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CenterCaption")
        lItem.SubItems(1) = "center caption on form"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_UseFormIcon")
        lItem.SubItems(1) = "use forms icon"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_UseFormCaption")
        lItem.SubItems(1) = "use forms caption text"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_TopSizingBorder")
        lItem.SubItems(1) = "top sizing border height"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_BottomSizingBorder")
        lItem.SubItems(1) = "bottom sizing border"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MinFormHeight")
        lItem.SubItems(1) = "minimum form height dimension"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MinFormWidth")
        lItem.SubItems(1) = "minimum form width dimension"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CaptionOffsetX")
        lItem.SubItems(1) = "caption offset position X"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CaptionOffsetY")
        lItem.SubItems(1) = "caption offset position Y"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_ThreadAccel")
        lItem.SubItems(1) = "thread accleration switch"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_MenuCustom")
        lItem.SubItems(1) = "use skinned menus"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_MenuActiveForeColor")
        lItem.SubItems(1) = "active menu item selected color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MenuInActiveForeColor")
        lItem.SubItems(1) = "active menu item color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MenuBackgroundColor")
        lItem.SubItems(1) = "menu bg color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MenuOffsetX")
        lItem.SubItems(1) = "menu start position X"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MenuOffsetY")
        lItem.SubItems(1) = "menu start position Y"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MenuImageList")
        lItem.SubItems(1) = "menu image list"
        lItem.SubItems(2) = "var"
        Set lItem = .ListItems.Add(Text:="p_MenuIconIndex")
        lItem.SubItems(1) = "icon index collection"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MenuRollOver")
        lItem.SubItems(1) = "custom caption rollover effect"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_MenuRollOverColor")
        lItem.SubItems(1) = "captionbar rollover accent color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_MenuFont")
        lItem.SubItems(1) = "specify menu font"
        lItem.SubItems(2) = "str"
        Set lItem = .ListItems.Add(Text:="p_MenuRollOverStyle")
        lItem.SubItems(1) = "rollover bar style"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_OfficeXpStyle")
        lItem.SubItems(1) = "office xp style menu"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_MenuTransparent")
        lItem.SubItems(1) = "use transparent menus (2K/XP)"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_MenuTransIdx")
        lItem.SubItems(1) = "transparency level"
        lItem.SubItems(2) = "byte"
        Set lItem = .ListItems.Add(Text:="p_SkinCommand")
        lItem.SubItems(1) = "skin command buttons"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_CmdTransColor")
        lItem.SubItems(1) = "transparent color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CmdCaption")
        lItem.SubItems(1) = "control caption"
        lItem.SubItems(2) = "str"
        Set lItem = .ListItems.Add(Text:="p_CmdRenderOffsetX")
        lItem.SubItems(1) = "stretchblt offset X"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CmdRenderOffsetY")
        lItem.SubItems(1) = "stretchblt offset Y"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CmdForeColor")
        lItem.SubItems(1) = "button text color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CmdColorHiLite")
        lItem.SubItems(1) = "button over text color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_CmdDrawFocused")
        lItem.SubItems(1) = "draw focus rect"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_CmdTextAntiAliased")
        lItem.SubItems(1) = "use anti-aliased caption"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_CmdIconPosition")
        lItem.SubItems(1) = "icon position"
        lItem.SubItems(2) = "enum"
        Set lItem = .ListItems.Add(Text:="p_SkinOptionButton")
        lItem.SubItems(1) = "skin option buttons"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_OptTransColor")
        lItem.SubItems(1) = "transparent color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_SkinCheckbox")
        lItem.SubItems(1) = "skin checkboxes"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_ChkTransColor")
        lItem.SubItems(1) = "transparent color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_SkinComboBox")
        lItem.SubItems(1) = "skin checkboxes"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_ComboTransColor")
        lItem.SubItems(1) = "transparent color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_ComboFrameColor")
        lItem.SubItems(1) = "base frame color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_ComboFrameHighLite")
        lItem.SubItems(1) = "frame accent color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_ComboFrameStyle")
        lItem.SubItems(1) = "frame style"
        lItem.SubItems(2) = "enum"
        Set lItem = .ListItems.Add(Text:="p_WaterMarkPosition")
        lItem.SubItems(1) = "position options"
        lItem.SubItems(2) = "enum"
        Set lItem = .ListItems.Add(Text:="p_WaterMarkCtrlHnd")
        lItem.SubItems(1) = "control handle"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_FrameColorNormal")
        lItem.SubItems(1) = "frame normal color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_FrameColorHover")
        lItem.SubItems(1) = "frame mouse over color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_FrameColorFocused")
        lItem.SubItems(1) = "frame focused color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_FrameListView")
        lItem.SubItems(1) = "apply to listview"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_FrameTextBox")
        lItem.SubItems(1) = "apply to textbox"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_FrameTreeView")
        lItem.SubItems(1) = "apply to treeview"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_FrameFileListBox")
        lItem.SubItems(1) = "apply to file listbox"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_FrameDirListBox")
        lItem.SubItems(1) = "apply to dir listbox"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_FrameListBox")
        lItem.SubItems(1) = "apply to listbox"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_FramePictureBox")
        lItem.SubItems(1) = "apply to picturebox"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_SkinListView")
        lItem.SubItems(1) = "skin listview headers"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_OListViewObj")
        lItem.SubItems(1) = "control object"
        lItem.SubItems(2) = "obj"
        Set lItem = .ListItems.Add(Text:="p_ListViewFlatPanel")
        lItem.SubItems(1) = "use flat header bg"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_ListViewTextAntiAliased")
        lItem.SubItems(1) = "use anti-aliased caption"
        lItem.SubItems(2) = "bool"
        Set lItem = .ListItems.Add(Text:="p_ListViewForeColor")
        lItem.SubItems(1) = "header font color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_ListViewHighLite")
        lItem.SubItems(1) = "header font highlite color"
        lItem.SubItems(2) = "long"
        Set lItem = .ListItems.Add(Text:="p_ScrollBarAdd")
        lItem.SubItems(1) = "add scrollbars to a control"
        lItem.SubItems(2) = "long"
    End With
    
End Sub

Private Sub LoadText()

Dim FF          As Long
Dim sBuffer     As String
Dim sPath       As String

On Error GoTo Handler

    sPath = Left$(App.Path, InStrRev(App.Path, "\")) & "Library\"
    FF = FreeFile
    Open sPath & "clsRCMWaterMark.cls" For Binary Access Read As #FF
    sBuffer = Space$(LOF(FF))
    Get #FF, , sBuffer
    Close #FF
    txtTest.Text = sBuffer

Exit Sub
Handler:
    Close #FF

End Sub

Private Sub CleanUp()

    
    If Not m_cOptionbtn Is Nothing Then
        m_cOptionbtn.CleanUp
        Set m_cOptionbtn = Nothing
    End If
    DestroyProgressBars
    If Not m_cCommand Is Nothing Then Set m_cCommand = Nothing
    If Not m_cLvwScrollBars Is Nothing Then Set m_cLvwScrollBars = Nothing
    If Not m_cTxtScrollBars Is Nothing Then Set m_cTxtScrollBars = Nothing
    If Not m_cSlider Is Nothing Then Set m_cSlider = Nothing
    If Not m_cListview Is Nothing Then Set m_cListview = Nothing
    If Not m_cWatermark Is Nothing Then Set m_cWatermark = Nothing
    If Not m_cCombobox Is Nothing Then Set m_cCombobox = Nothing
    If Not m_cCheckbox Is Nothing Then Set m_cCheckbox = Nothing
    If Not m_cRCM Is Nothing Then Set m_cRCM = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CleanUp
End Sub

