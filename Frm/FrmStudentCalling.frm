VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmStudentCalling 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13830
   Icon            =   "FrmStudentCalling.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9210
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13830
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmStudentCalling.frx":6852
      Left            =   15480
      List            =   "FrmStudentCalling.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   22
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Visible         =   0   'False
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   15480
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15600
      Top             =   3720
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
            Picture         =   "FrmStudentCalling.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentCalling.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentCalling.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentCalling.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentCalling.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentCalling.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentCalling.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentCalling.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmStudentCalling.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmStudentCalling.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
      BackColor       =   14871017
      FontSize        =   9.75
      FontName        =   "Arial"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmStudentCalling.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   9210
      Left            =   0
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Width           =   13830
      _cx             =   24395
      _cy             =   16245
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame2 
         Height          =   3765
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   2820
         Visible         =   0   'False
         Width           =   8715
         Begin VB.PictureBox Picture1 
            Height          =   2805
            Left            =   120
            ScaleHeight     =   2745
            ScaleWidth      =   8355
            TabIndex        =   97
            Top             =   840
            Width           =   8415
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   31
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   30
               Left            =   6180
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   29
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   28
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   27
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   26
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   25
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   24
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   23
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   1320
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   22
               Left            =   6180
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   1320
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   21
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   1320
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   20
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   1320
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   19
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   1320
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   18
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   1320
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   17
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   1320
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   16
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   1320
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   15
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   14
               Left            =   6180
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   13
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   12
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   11
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   10
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   9
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   8
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   7
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   6
               Left            =   6180
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   5
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   4
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   3
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   2
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   1
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblTimeFree 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1:00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   0
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lbDay 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   105
               Top             =   720
               Width           =   660
            End
            Begin VB.Label lbCount 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   104
               Top             =   360
               Width           =   660
            End
            Begin VB.Label lbDayName 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   103
               Top             =   25
               Width           =   735
            End
            Begin VB.Label lbDayName 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   255
               Index           =   2
               Left            =   960
               TabIndex        =   102
               Top             =   25
               Width           =   735
            End
            Begin VB.Label lbDayName 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   255
               Index           =   3
               Left            =   1800
               TabIndex        =   101
               Top             =   25
               Width           =   735
            End
            Begin VB.Label lbDayName 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   255
               Index           =   5
               Left            =   3480
               TabIndex        =   100
               Top             =   25
               Width           =   735
            End
            Begin VB.Label lbDayName 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   255
               Index           =   6
               Left            =   4320
               TabIndex        =   99
               Top             =   25
               Width           =   735
            End
            Begin VB.Label lbDayName 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   255
               Index           =   7
               Left            =   5160
               TabIndex        =   98
               Top             =   25
               Width           =   735
            End
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„Ê«⁄Ìœ «·„ «Õ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   450
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   270
            Width           =   5325
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   6960
            TabIndex        =   95
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5115
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   2130
         Visible         =   0   'False
         Width           =   12075
         Begin VSFlex8UCtl.VSFlexGrid FG5 
            Height          =   3255
            Left            =   90
            TabIndex        =   83
            Top             =   1080
            Width           =   11955
            _cx             =   21087
            _cy             =   5741
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmStudentCalling.frx":15BA9
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   -1  'True
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   10530
            TabIndex        =   85
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ÕÃÊ“«  ”«»ﬁ… ··„ÊŸ›"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   240
            Width           =   5325
         End
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   14130
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   13845
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   2955
         Left            =   -90
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1710
         Width           =   13860
         _cx             =   24448
         _cy             =   5212
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.TextBox txtNoteSerialCash 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   300
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   -30
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.TextBox txtNoteSerialCash 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1635
            Width           =   1380
         End
         Begin VB.CommandButton cmdPrintCash 
            Caption         =   "ÿ»«⁄… ”‰œ «·ﬁ»÷"
            Height          =   495
            Left            =   4050
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1530
            Width           =   990
         End
         Begin VB.CommandButton Command2 
            Caption         =   "⁄—÷ ”‰œ «·ﬁ»÷"
            Height          =   495
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   1530
            Width           =   1380
         End
         Begin VB.TextBox txtCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   435
            Left            =   4515
            TabIndex        =   87
            Top             =   525
            Width           =   2115
         End
         Begin VB.CommandButton cmdAddCustomer 
            Caption         =   "«÷«›… ⁄„Ì· ÃœÌœ"
            Height          =   435
            Left            =   3390
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   525
            Width           =   1125
         End
         Begin VB.TextBox txtArboun 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            TabIndex        =   75
            Top             =   1740
            Width           =   1545
         End
         Begin VB.TextBox TxtStudentEmail 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   8790
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   2310
            Width           =   3540
         End
         Begin VB.TextBox TxtMobile 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   8790
            MaxLength       =   14
            TabIndex        =   12
            Top             =   1845
            Width           =   3540
         End
         Begin VB.TextBox TxtStudentPhone 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   8790
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1455
            Width           =   3540
         End
         Begin VB.TextBox TxtSudCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8790
            TabIndex        =   8
            Top             =   1050
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.TextBox TxtUQama 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   10785
            TabIndex        =   7
            Top             =   1050
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   10785
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   120
            Width           =   1545
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   2250
            Width           =   7710
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   10785
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   525
            Width           =   1545
         End
         Begin MSDataListLib.DataCombo DcbCompany 
            Height          =   315
            Left            =   7335
            TabIndex        =   5
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   585
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker EnterDate 
            Height          =   390
            Left            =   60
            TabIndex        =   6
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   688
            _Version        =   393216
            Format          =   122486785
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbEmployee 
            Height          =   315
            Left            =   3645
            TabIndex        =   3
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   120
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker EnterTime 
            Height          =   360
            Left            =   60
            TabIndex        =   10
            Top             =   1050
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            Format          =   122486786
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbStudent 
            Height          =   315
            Left            =   3645
            TabIndex        =   9
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   1050
            Visible         =   0   'False
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal EnterDateH 
            Height          =   375
            Left            =   540
            TabIndex        =   78
            Top             =   585
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Index           =   56
            Left            =   6570
            TabIndex        =   93
            Top             =   1920
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   375
            Index           =   76
            Left            =   6420
            TabIndex        =   88
            Top             =   585
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄—»Ê‰"
            Height          =   345
            Index           =   3
            Left            =   1620
            TabIndex        =   76
            Top             =   1770
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»—Ìœ «·«·ﬂ —Ê‰Ì"
            Height          =   360
            Index           =   9
            Left            =   12420
            TabIndex        =   68
            Top             =   2310
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÃÊ«·"
            Height          =   390
            Index           =   17
            Left            =   12420
            TabIndex        =   67
            Top             =   1875
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Â« ›"
            Height          =   330
            Index           =   5
            Left            =   12420
            TabIndex        =   66
            Top             =   1485
            Width           =   1440
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÂÊÌ…"
            Height          =   345
            Index           =   12
            Left            =   12420
            TabIndex        =   65
            Top             =   990
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ œ—»"
            Height          =   345
            Index           =   1
            Left            =   9615
            TabIndex        =   64
            Top             =   990
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «—ÌŒ «·„Ê⁄œ ÂÃ—Ì"
            Height          =   405
            Index           =   2
            Left            =   1695
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   585
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Êﬁ  «·„Ê⁄œ "
            Height          =   345
            Index           =   1
            Left            =   1695
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1050
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ÊŸ›"
            Height          =   375
            Index           =   0
            Left            =   12420
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   390
            Index           =   15
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2325
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„Ê⁄œ „Ì·«œÌ"
            Height          =   375
            Index           =   12
            Left            =   1695
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   120
            Width           =   1965
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„Ì·"
            Height          =   435
            Index           =   5
            Left            =   12420
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   525
            Width           =   1440
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1350
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   7860
         Width           =   13770
         _cx             =   24289
         _cy             =   2381
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ImpulseButton.ISButton btnNew 
            Height          =   390
            Left            =   11865
            TabIndex        =   34
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   735
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":15D33
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   390
            Left            =   10050
            TabIndex        =   35
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   735
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":1C595
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   390
            Left            =   8220
            TabIndex        =   16
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   735
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":22DF7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   390
            Left            =   6105
            TabIndex        =   36
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   735
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":23191
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   390
            Left            =   4980
            TabIndex        =   37
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   735
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":2352B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   525
            Left            =   3780
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   735
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   926
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":23AC5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   390
            Left            =   2385
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   720
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":2A327
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   390
            Left            =   390
            TabIndex        =   40
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   735
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":2A6C1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8670
            TabIndex        =   41
            Top             =   90
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   315
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   285
            Width           =   645
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2445
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   285
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   255
            Index           =   1
            Left            =   1110
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   255
            Index           =   0
            Left            =   3345
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   285
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   450
            Index           =   14
            Left            =   12630
            TabIndex        =   42
            Top             =   90
            Width           =   1185
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   840
         Index           =   18
         Left            =   0
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   0
         Width           =   13845
         _cx             =   24421
         _cy             =   1482
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   120
            TabIndex        =   51
            Top             =   270
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":2AA5B
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   690
            TabIndex        =   52
            Top             =   270
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":2ADF5
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1395
            TabIndex        =   53
            Top             =   270
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":2B18F
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2010
            TabIndex        =   54
            Top             =   270
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentCalling.frx":2B529
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   690
            Left            =   12750
            Picture         =   "FrmStudentCalling.frx":2B8C3
            Stretch         =   -1  'True
            Top             =   105
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  «·« ’«·/ «·„Ê«⁄Ìœ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   2
            Left            =   7815
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   270
            Width           =   4785
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   660
         Left            =   0
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   750
         Width           =   13860
         _cx             =   24448
         _cy             =   1164
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   405
            Left            =   10515
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   135
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   405
            Left            =   7650
            TabIndex        =   0
            Top             =   135
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   714
            _Version        =   393216
            Format          =   137363457
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   135
            Width           =   4770
            _ExtentX        =   8414
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   375
            Left            =   6570
            TabIndex        =   77
            Top             =   135
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÕÃ“"
            Height          =   300
            Index           =   4
            Left            =   12750
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   135
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   375
            Index           =   25
            Left            =   9270
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   135
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   375
            Index           =   0
            Left            =   4710
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   135
            Width           =   1440
         End
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   3480
         Left            =   0
         TabIndex        =   69
         Top             =   4425
         Width           =   13770
         _cx             =   24289
         _cy             =   6138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "»Ì«‰«  «·ÕÃ“|»Ì«‰« "
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   1
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   3105
            Index           =   1
            Left            =   45
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   45
            Width           =   13680
            _cx             =   24130
            _cy             =   5477
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   3045
               Index           =   0
               Left            =   19050
               TabIndex        =   71
               Top             =   360
               Width           =   13485
               _cx             =   23786
               _cy             =   5371
               Appearance      =   2
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmStudentCalling.frx":2CCC8
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   7
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   -1  'True
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VSFlex8UCtl.VSFlexGrid FG4 
               Height          =   2310
               Left            =   0
               TabIndex        =   14
               Top             =   330
               Width           =   13605
               _cx             =   23998
               _cy             =   4075
               Appearance      =   2
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmStudentCalling.frx":2CD88
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   7
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   -1  'True
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin ImpulseButton.ISButton Cmd_DeleteRow 
               Height          =   345
               Index           =   4
               Left            =   1815
               TabIndex        =   80
               Top             =   2700
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " Õ–› ”ÿ—"
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmStudentCalling.frx":2CFC2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd_DeleteAll 
               Height          =   345
               Index           =   4
               Left            =   0
               TabIndex        =   81
               Top             =   2685
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " Õ–› «·ﬂ·"
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmStudentCalling.frx":2D55C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Label3"
               Height          =   240
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   2775
               Width           =   225
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   3105
            Index           =   0
            Left            =   14415
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   45
            Width           =   13680
            _cx             =   24130
            _cy             =   5477
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   3060
               Index           =   1
               Left            =   19050
               TabIndex        =   73
               Top             =   330
               Width           =   13500
               _cx             =   23812
               _cy             =   5397
               Appearance      =   2
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmStudentCalling.frx":2DAF6
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   7
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   -1  'True
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   2835
               Index           =   11
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   105
               Width           =   13515
            End
            Begin VB.Shape Shape1 
               BorderWidth     =   2
               Height          =   3030
               Left            =   0
               Top             =   0
               Width           =   13665
            End
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   15480
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmStudentCalling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim ii As Long
Public LngRow As Long


Public Function print_reportCash(Optional NoteSerial As String, Optional NoteSerial1 As String, Optional BankName As String, Optional PaymentType As String, Optional Box As String, Optional Custcode As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

   ' MySQL = "Select * From payment_voucher  where NoteID=" & val(XPTxtID.Text)
MySQL = "SELECT BillMaintNo, Notes.paydes,    dbo.Notes.Note_Value, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.BanksData.BankName, dbo.Notes.NoteType, dbo.Notes.BoxID, dbo.TblBoxesData.BoxName, "
MySQL = MySQL & "                        dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Notes.Remark, dbo.Notes.NoteSerial, dbo.Notes.NoteDate,"
MySQL = MySQL & "                                 dbo.Notes.note_value_by_characters, dbo.Notes.NoteID, dbo.Notes.general_des_notes, dbo.Notes.person, dbo.TblCustemers.Fullcode, dbo.Notes.PreVAT,"
MySQL = MySQL & "                                 dbo.Notes.Vat , dbo.Notes.NoteSerial1, dbo.Notes.ManulaNO, dbo.Notes.ManualNO"
MySQL = MySQL & "           FROM         dbo.Notes LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
MySQL = MySQL & "           Where (dbo.Notes.NoteType = 4)"
MySQL = MySQL & "           and NoteID=" & val(txtNoteSerialCash(1).Text)


    If SystemOptions.UserInterface = ArabicInterface Then
    '    StrFileName = App.path & "\Reports\" & "Payment_voucher.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\Payment_voucher.rpt"
    Else
     '   StrFileName = App.path & "\Reports\" & "Payment_voucherE.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\Payment_voucher.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(5).AddCurrentValue "" '''DcboCreditSide.Text
   
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(5).AddCurrentValue "" 'DcboCreditSide.Text
        StrReportTitle = ""
 
    End If
Dim i As Integer
Dim str As String
'With Grid5
'str = ""
'For i = 1 To .Rows - 1
'If (.TextMatrix(i, .ColIndex("NoteSerial1"))) <> "" And .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
'str = str & .TextMatrix(i, .ColIndex("NoteSerial1"))
'If i <> (.Rows - 1) Then
'str = str & ","
'End If
'End If
'Next i

    xReport.ParameterFields(3).AddCurrentValue user_name
    '
    xReport.ParameterFields(6).AddCurrentValue NoteSerial1

    xReport.ParameterFields(7).AddCurrentValue BankName
    xReport.ParameterFields(8).AddCurrentValue PaymentType
    xReport.ParameterFields(9).AddCurrentValue Box
    xReport.ParameterFields(10).AddCurrentValue Custcode
    xReport.ParameterFields(11).AddCurrentValue str
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
 

Private Sub cmdPrintCash_Click()
  
 
 If txtNoteSerialCash(0) <> "" Then
                print_reportCash txtNoteSerialCash(0), txtNoteSerialCash(0), "", "", "", DcbCompany.Text
    End If
End Sub

Private Sub Command2_Click()
    FrmCashing.show
    FrmCashing.Retrive val(txtNoteSerialCash(1).Text)
 
End Sub


Private Sub cmdAddCustomer_Click()
    Dim Dcombos As New ClsDataCombos
If SystemOptions.DontShowMoreDetailsCompItem Then
    
    FrmCustemers.show
    FrmCustemers.Retrive val(DcbCompany.BoundText), Me.Name
    FrmCustemers.FormNamee = Me.Name
    
   ' Dcombos.GetCustomersSuppliers 1, Me.DcbCompany, True
    If DcbCompany.Text = "" Then
   '     DcbCompany.BoundText = mCustId
    End If
    Exit Sub
End If
           
Dim CUSTID As Double
Dim mCode As String

If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(txtCustomerName) = "" Then MsgBox "«œŒ· «”„ «·⁄„Ì·": Exit Sub
    If Trim(TxtMobile) = "" Then TxtMobile.locked = False: MsgBox "«œŒ· —ﬁ„ «·Â« ›/«·ÃÊ«·  ": Exit Sub
Else
    If Trim(txtCustomerName) = "" Then MsgBox "Enter the customer name": Exit Sub
    If Trim(TxtMobile) = "" Then TxtMobile.locked = False: MsgBox "Enter your phone / mobile number  ": Exit Sub

End If

Dim s As String
Dim rsDummy As New ADODB.Recordset

s = "Select * from TblCustemers WHere 1=1 "
If Trim(TxtMobile) <> "" Then
    s = s & " And Cus_mobile = N'" & Trim(TxtMobile) & "' "
End If
If Trim(txtCustomerName) <> "" Then
    'If Trim(TxtMobile) <> "" Then
    '    s = s & " Or CusName = '" & Trim(txtCustomerName.Text) & "'"
    'Else
    '    s = s & " and CusName = '" & Trim(txtCustomerName.Text) & "'"
    'End If
End If
rsDummy.Open s, Cn, adOpenStatic
If Not rsDummy.EOF Then
    Text15.Text = rsDummy!Fullcode & ""
    
    DcbCompany.BoundText = val(rsDummy!CusID & "")
   
    txtCustomerName.backcolor = vbGreen
    'TxtMobile.backcolor = vbGreen
    Exit Sub
Else
    txtCustomerName.backcolor = vbWhite
    'TxtMobile.backcolor = vbWhite
End If

    createCustomer txtCustomerName.Text, txtCustomerName.Text, val(DcbBranch.BoundText), CUSTID, TxtMobile.Text, mCode
    Text15.Text = mCode
    
    Dcombos.GetCustomersSuppliers 1, Me.DcbCompany, True
    DcbCompany.BoundText = CUSTID
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „ «÷«›… «·⁄„Ì·"
    Else
        MsgBox "Customer added"
    End If
    'txtCustomerName = ""

End Sub

Private Sub Fg4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.mRow = Fg4.Row
        FrmItemSearch.RetrunType = 9878
        FrmItemSearch.show vbModal
    End If
End Sub

Private Sub Label5_Click()
Frame2.Visible = False
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(Index).Caption) <> 0 Then
        lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
    End If
    'ff

End Sub


Private Sub SendMessage()
            Dim subject As String
            Dim Msg As String
            Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
           Dim companyphone As String
           Dim optIsResponsible As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "")
    
    If Trim(TxtMobile) = "" Then Exit Sub
     
      
    'If Not SystemOptions.UserInterface = EnglishInterface Then
      
        
    'Else
   
    
    'End If
    Dim isFound As Boolean
    isFound = True
    Dim txtCodeSend As String
    txtCodeSend = "+966"
    
    If Not FindString(TxtMobile, "+966", 1) Then
        If Not FindString(TxtMobile, "00966", 1) Then
            isFound = False
        End If
        isFound = False
    End If
    If Not isFound Then
        txtCodeSend = "+966"
    Else
        txtCodeSend = ""
        'txtPhoneCust = "+966" & val(txtPhoneCust)
    End If
    Dim mTxt As String
    mTxt = txtCodeSend & val(TxtMobile)
    lbl(56).Caption = mTxt
    t = sendMessageM("user", "password", Msg, "", mTxt)
    
        Msg = "⁄“Ì“ ‰« «·⁄„Ì·…   „  √ﬂÌœ ÕÃ“ „Ê⁄œﬂ " & CHR(13)
        Msg = Msg & Trim(Fg4.TextMatrix(1, Fg4.ColIndex("ItemName"))) & CHR(13)
        Msg = Msg & " » «—ÌŒ " & EnterDate & CHR(13)
        Msg = Msg & " «·”«⁄… " & Trim(Fg4.TextMatrix(1, Fg4.ColIndex("HoursT"))) & CHR(13)
        Msg = Msg & " —ﬁ„ «·ÕÃ“ " & TxtSerial1
        Msg = Msg & " ‘ﬂ—« ·«Œ Ì«—ﬂ ⁄«·„ Ê”«„ ··√“Ì«¡ Ê«· Ã„Ì· "
    
        t = sendMessageM("user", "password", Msg, "", mTxt)
    DoEvents


End Sub
Private Function FindString(Control As Control, FindStr As String, Optional StartPos As Integer = 1) As Boolean
Dim a As Integer
    a = InStr(StartPos, LCase$(Control.Text), LCase$(FindStr))
    If a = 0 Then
        FindString = False
    Else
        FindString = True
        Control.SetFocus
        Control.SelStart = a - 1
        Control.SelLength = Len(FindStr)
    End If
End Function

 

Private Sub RemoveGridRowAll()
    
    
        Fg4.Rows = 1
   
    
End Sub


Private Sub RemoveGridRow()
    
   
        With Me.Fg4
    'MsgBox .Row
            If .Row <= 0 Then
                    .Rows = 2
            Exit Sub
            Else
            .RemoveItem .Row
            End If
        End With
    
End Sub

Private Sub Cmd_DeleteRow_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then

    RemoveGridRow

End If
End Sub
Private Sub DcbCompany_Change()
DcbCompany_Click (0)
End Sub

Private Sub DcbCompany_Click(Area As Integer)

        If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
            
            If val(DcbCompany.BoundText) = 0 Then Exit Sub
            Dim EmpCode  As String
            GetTblCustemersCode , , DcbCompany.BoundText, EmpCode
            Me.Text15.Text = EmpCode
            Dim Dcombos As New ClsDataCombos
            Dcombos.GetStudent Me.DcbStudent, 0, val(DcbCompany.BoundText)
            
            GetCustomerNamebyPhone , , DcbCompany.BoundText
            
        End If
        

  
End Sub
Sub GetStudentCalling()
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     StudentEmail, Mobile, StudentPhone"
sql = sql & " From dbo.TblStudent"
sql = sql & " where id =" & val(DcbStudent.BoundText) & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Me.TxtMobile.Text = IIf(IsNull(Rs3("Mobile").value), "", Rs3("Mobile").value)
Me.TxtStudentEmail.Text = IIf(IsNull(Rs3("StudentEmail").value), "", Rs3("StudentEmail").value)
Me.TxtStudentPhone.Text = IIf(IsNull(Rs3("StudentPhone").value), "", Rs3("StudentPhone").value)
Else
TxtMobile = ""
TxtStudentEmail = ""
TxtStudentPhone = ""
End If
End Sub

Private Sub DcbCompany_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
   FrmCustemerSearch.SearchType = 27
        FrmCustemerSearch.show vbModal
  End If
End Sub

Private Sub DcbEmployee_Change()
DcbEmployee_Click (0)
End Sub
Private Sub DcbStudent_Change()
DcbStudent_Click (0)
End Sub
Private Sub DcbStudent_Click(Area As Integer)
Dim UQama As String
  If val(DcbStudent.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetStudentCode val(DcbStudent.BoundText), EmpCode, 0, UQama
    TxtUQama.Text = UQama
    Me.TxtSudCode.Text = EmpCode
    GetStudentCalling
End Sub

Private Sub DcbStudent_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearStudent.inde = 102
Load FrmSearStudent
FrmSearStudent.show vbModal
End If
End Sub




Private Sub fg4_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String, LngRow As Long
Dim rsDummy As New ADODB.Recordset
Dim s As String
With Fg4
 '   If val(.TextMatrix(Row, .ColIndex("EmpID"))) = 0 Then
        
   Select Case .ColKey(Col)
    Case "ItemID", "ItemName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
                .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
                s = "Select PeriodService from tblItems Where ItemId = " & val(StrAccountCode)
                
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("PeriodT")) = rsDummy!PeriodService & ""
                    '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                End If
                s = "Select FullCode,ItemName From TblItems "
                s = s & " WHERE ItemId = " & val(StrAccountCode)
                rsDummy.Close
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("Code")) = rsDummy!Fullcode & ""
                    '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                End If
                
                
  Case "Code"
        s = "Select ItemId,ItemName From TblItems "
        s = s & " WHERE "
        s = s & " Code = '" & Trim(.TextMatrix(Row, .ColIndex("Code"))) & "' Or FullCode ='" & Trim(.TextMatrix(Row, .ColIndex("Code"))) & "'"
       ' s = s & " and  TblItems.ItemID IN (SELECT ItemID FROM TblEmpItemsTrans2 WHERE TblEmpItemsTrans2.EmpID = " & val(.TextMatrix(Row, .ColIndex("EmpID"))) & ")"
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            .TextMatrix(Row, .ColIndex("ItemID")) = val(rsDummy!ItemID & "")
            .TextMatrix(Row, .ColIndex("ItemName")) = (rsDummy!ItemName & "")
        Else
            .TextMatrix(Row, .ColIndex("ItemID")) = 0
            .TextMatrix(Row, .ColIndex("ItemName")) = ""
            .TextMatrix(Row, .ColIndex("Code")) = ""
            MsgBox "Â–« «·ﬂÊœ €Ì— „”Ã· „‰ ﬁ»·"
            Exit Sub
        End If
   Case "Emp_code"
        s = "Select Emp_code,Emp_Name,Emp_id  FROM TblEmployee"
        s = s & " WHERE "
        s = s & " Emp_code = '" & Trim(.TextMatrix(Row, .ColIndex("Emp_code"))) & "'"
       
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            .TextMatrix(Row, .ColIndex("EmpID")) = val(rsDummy!Emp_id & "")
            .TextMatrix(Row, .ColIndex("EmpName")) = (rsDummy!emp_name & "")
            fg4_AfterEdit Row, Fg4.ColIndex("EmpID")
        Else
            .TextMatrix(Row, .ColIndex("EmpID")) = 0
            .TextMatrix(Row, .ColIndex("EmpName")) = ""
            .TextMatrix(Row, .ColIndex("Emp_code")) = ""
            MsgBox "Â–« «·ﬂÊœ €Ì— „”Ã· „‰ ﬁ»·"
            Exit Sub
        End If
        
    Case "EmpID", "EmpName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpName"), False, True)
                .TextMatrix(Row, .ColIndex("EmpID")) = StrAccountCode
                  s = "Select Emp_code,Emp_Name,Emp_id  FROM TblEmployee"
                 s = s & " WHERE "
                 s = s & " Emp_id = " & val(StrAccountCode)
                
                 rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("Emp_code")) = Trim(rsDummy!emp_code & "")
                End If
                rsDummy.Close
'                s = "Select PercentV from TblTasks Where Id = " & val(StrAccountCode)
                If Trim(.TextMatrix(Row, .ColIndex("HoursT"))) <> "" And Trim(.TextMatrix(Row, .ColIndex("EmpID"))) <> "" Then
                     s = "select * from tblRestsSiftTrans2 Where  EmpId  = " & val(.TextMatrix(Row, .ColIndex("EmpID")))
                     
                     s = s & " AND CAST(FromTime as time) >='" & FormatDateTime(.TextMatrix(Row, .ColIndex("HoursT")), vbShortTime) & "'"
                     s = s & " AND CAST(ToTime as time) >='" & FormatDateTime(.TextMatrix(Row, .ColIndex("HoursT")), vbShortTime) & "'"
                     s = s & " And tblRestsSiftTrans2.FromDate =" & SQLDate(Me.EnterDate.value, True) & ""
                     rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                     If Not rsDummy.EOF Then
                         .TextMatrix(Row, .ColIndex("Status")) = "Â–Â «·› —… ÂÏ —«Õ… ··„ÊŸ›"
                         '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                    Else
                        .TextMatrix(Row, .ColIndex("Status")) = ""
                     End If
                    s = "select * from TblStudCalling2 "
                    s = s & " Inner join TblStudCalling On TblStudCalling.Id = TblStudCalling2.MasterId"
                    s = s & " Where  TblStudCalling2.EmpId  = " & val(.TextMatrix(Row, .ColIndex("EmpID")))
                    s = s & " AND CAST(HoursT as time) ='" & FormatDateTime(.TextMatrix(Row, .ColIndex("HoursT")), vbShortTime) & "'"
                    s = s & " And TblStudCalling.EnterDate =" & SQLDate(Me.EnterDate.value, True) & ""
                    If Trim(TxtSerial1) <> "" Then
                        s = s & " and MasterId <> " & val(TxtSerial1)
                    End If
                    Set rsDummy = New ADODB.Recordset
                    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummy.EOF Then
                        MsgBox "Â–« «·„ÊŸ› ·Â ÕÃÊ“«  ›Ï ‰›” «·› —…"
                        
                        s = " SELECT TblStudCalling2.*,"
                        s = s & "       TblEmployee.Emp_Name        EmpName,"
                        s = s & "       TblItems.ItemName,"
                        s = s & "       tblReservationType.Name  AS ReservationTypeName"
                        s = s & " From TblStudCalling2"
                        s = s & "       INNER JOIN tblReservationType"
                        s = s & "            ON  tblReservationType.ID = TblStudCalling2.ReservationTypeCode"
                        s = s & "       INNER JOIN TblEmployee"
                        s = s & "            ON  TblEmployee.Emp_ID = TblStudCalling2.EmpID"
                        s = s & "       INNER JOIN TblItems"
                        s = s & "            ON  TblItems.ItemID = TblStudCalling2.ItemID"
                        s = s & " Where TblStudCalling2.MasterId = " & val(rsDummy!MasterID & "")
                        s = s & " and TblStudCalling2.EmpId = " & val(.TextMatrix(Row, .ColIndex("EmpID")))
                        loadgrid s, FG5, True, True
                        Frame1.Visible = True
                        .TextMatrix(Row, .ColIndex("EmpID")) = 0
                        .TextMatrix(Row, .ColIndex("EmpName")) = ""
                        '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                    End If
                    
                    
                End If
    Case "ReservationTypeCode", "ReservationTypeName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ReservationTypeName"), False, True)
                .TextMatrix(Row, .ColIndex("ReservationTypeCode")) = StrAccountCode
'                s = "Select PercentV from TblTasks Where Id = " & val(StrAccountCode)
'                s = "select TblTasks.ID,TblTasks.PercentV from TblTasks Where TblTasks.Id  = " & val(.TextMatrix(Row, .ColIndex("TasksID")))
'                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'                If Not rsDummy.EOF Then
'                    .TextMatrix(Row, .ColIndex("TasksID")) = rsDummy!ID & ""
'                    '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
'                End If
    Case "HoursT"

        If Trim(.TextMatrix(Row, .ColIndex("HoursT"))) <> "" Then
                     s = "select * from tblRestsSiftTrans2 Where  EmpId  = " & val(.TextMatrix(Row, .ColIndex("EmpID")))
                    
                    s = s & " And ('" & FormatDateTime(.TextMatrix(Row, .ColIndex("HoursT")), vbShortTime) & "'   BETWEEN CAST(FromTime AS TIME) and CAST(ToTime AS TIME) )"
'                       s = s & " AND CAST(FromTime as time) >='" & FormatDateTime(.TextMatrix(Row, .ColIndex("HoursT")), vbShortTime) & "'"
'                     s = s & " AND CAST(ToTime as time) <='" & FormatDateTime(.TextMatrix(Row, .ColIndex("HoursT")), vbShortTime) & "'"
                     s = s & " And tblRestsSiftTrans2.FromDate =" & SQLDate(Me.EnterDate.value, True) & ""
                     Set rsDummy = New ADODB.Recordset
                     rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                     If Not rsDummy.EOF Then
                         .TextMatrix(Row, .ColIndex("Status")) = "Â–Â «·› —… ÂÏ —«Õ… ··„ÊŸ›"
                         '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                         MsgBox "Â–Â «·› —… ÂÏ —«Õ… ··„ÊŸ›"
                    Else
                        .TextMatrix(Row, .ColIndex("Status")) = ""
                     End If
                       s = "select * from TblStudCalling2 "
                    s = s & " Inner join TblStudCalling On TblStudCalling.Id = TblStudCalling2.MasterId"
                    s = s & " Where  TblStudCalling2.EmpId  = " & val(.TextMatrix(Row, .ColIndex("EmpID")))
                    'Salim here********************************
                    s = s & " AND CAST(HoursT as time) ='" & FormatDateTime(.TextMatrix(Row, .ColIndex("HoursT")), vbShortTime) & "'"
                    'Salim here********************************
                    
                    s = s & " And TblStudCalling.EnterDate =" & SQLDate(Me.EnterDate.value, True) & ""
                    If Trim(TxtSerial1) <> "" Then
                        s = s & " and MasterId <> " & val(TxtSerial1)
                    End If
                    
                    Set rsDummy = New ADODB.Recordset
                    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummy.EOF Then
                        MsgBox "Â–« «·„ÊŸ› ·Â ÕÃÊ“«  ›Ï ‰›” «·› —…"
                        
                        s = " SELECT TblStudCalling2.*,"
                        s = s & "       TblEmployee.Emp_Name        EmpName,"
                        s = s & "       TblItems.ItemName,"
                        s = s & "       tblReservationType.Name  AS ReservationTypeName"
                        s = s & " From TblStudCalling2"
                        s = s & "       INNER JOIN tblReservationType"
                        s = s & "            ON  tblReservationType.ID = TblStudCalling2.ReservationTypeCode"
                        s = s & "       INNER JOIN TblEmployee"
                        s = s & "            ON  TblEmployee.Emp_ID = TblStudCalling2.EmpID"
                        s = s & "       INNER JOIN TblItems"
                        s = s & "            ON  TblItems.ItemID = TblStudCalling2.ItemID"
                        s = s & " Where TblStudCalling2.MasterId = " & val(rsDummy!MasterID & "")
                        s = s & " and TblStudCalling2.EmpId = " & val(.TextMatrix(Row, .ColIndex("EmpID")))
                        loadgrid s, FG5, True, True
                        Frame1.Visible = True
                        .TextMatrix(Row, .ColIndex("EmpID")) = 0
                        .TextMatrix(Row, .ColIndex("EmpName")) = ""
                        '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                    End If
                End If
    End Select
  '  CalcAmount
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
End With
End Sub


Private Sub FG4_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim s As String
Dim rsDummy As New ADODB.Recordset
  With Me.Fg4

        Select Case .ColKey(Col)
        Case "ShowTime"
                s = " SELECT * FROM TbLSheft"
            s = s & " Inner Join"
            s = s & " TblShiftWorker"
            s = s & " ON TbLSheft.SeftCode = TblShiftWorker.ShiftID"
            s = s & " Where TblShiftWorker.EmpID = " & val(.TextMatrix(Row, .ColIndex("EmpID")))
            
            FillTimeFree val(.TextMatrix(Row, .ColIndex("EmpID")))
            Frame2.Visible = True
                 Case "HoursT"
                 ' LngRow = Row

 'LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                ItemProductionDate2.Index = 1
                Load ItemProductionDate2
                LngRow = Row
                ItemProductionDate2.show 1
                fg4_AfterEdit Row, Col
                Case "CMD"
           
                        
                        s = " SELECT TblStudCalling2.*,"
                        s = s & "       TblEmployee.Emp_Name        EmpName,"
                        s = s & "       TblItems.ItemName,"
                        s = s & "       tblReservationType.Name  AS ReservationTypeName"
                        s = s & " From TblStudCalling2"
                        s = s & "       INNER JOIN tblReservationType"
                        s = s & "            ON  tblReservationType.ID = TblStudCalling2.ReservationTypeCode"
                        s = s & "       INNER JOIN TblEmployee"
                        s = s & "            ON  TblEmployee.Emp_ID = TblStudCalling2.EmpID"
                        s = s & "       INNER JOIN TblItems"
                        s = s & "            ON  TblItems.ItemID = TblStudCalling2.ItemID"
                        s = s & "       INNER JOIN TblStudCalling"
                        s = s & "            ON  TblStudCalling.iD = TblStudCalling2.MasterId"
                        
                        s = s & " Where "
                        s = s & "  TblStudCalling.EnterDate =" & SQLDate(Me.EnterDate.value, True) & ""
                        s = s & " and TblStudCalling2.EmpId = " & val(.TextMatrix(Row, .ColIndex("EmpID")))
                        loadgrid s, FG5, True, True
                        Frame1.Visible = True
             
                        '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                    
                End Select
                End With
End Sub
Private Sub FillTimeFree(ByVal mEmpId As Long)
Dim rsDummy As New ADODB.Recordset
rsDummy.Open GetFeildDay(mEmpId), Cn, adOpenStatic, adLockReadOnly
Dim i As Long
Dim mHours As Long
Dim mHoursTo As Long
Dim mMin As Long
Dim mMinTo As Long
For i = 0 To lblTimeFree.count - 1
    lblTimeFree(i).Visible = False
Next
i = 0
Dim mHours2 As Long
Dim mHoursTo2 As Long
Dim mMin2 As Long
Dim mMinTo2 As Long

Dim mCaption As String
Dim mIndexTime As Long
mIndexTime = 0
Do While Not rsDummy.EOF
        
    mHours = val(rsDummy!FromTime1 & "")
    mHoursTo = val(rsDummy!TOTime1 & "")
    
    mMin = val(Replace(rsDummy!FromTime1 & "", mHours & ":", ""))
    mMinTo = val(Replace(rsDummy!TOTime1 & "", mHoursTo & ":", ""))
    
    mHours2 = val(rsDummy!FromTime2 & "")
    mHoursTo2 = val(rsDummy!TOTime2 & "")
    mMin2 = val(Replace(rsDummy!FromTime2 & "", mHours2 & ":", ""))
    mMinTo2 = val(Replace(rsDummy!TOTime2 & "", mHoursTo2 & ":", ""))
    
    i = 0
    For i = mHours To mHoursTo
        If i > 12 Then
            mCaption = i - 12 & ":" & mMin & "PM"
        Else
            mCaption = i & ":" & mMin & "AM"
        End If
        
        mCaption = Format(mCaption, "HH:mm AM/PM")
        If ChkTimeForEmp(mCaption) Then
            lblTimeFree(mIndexTime) = mCaption
            lblTimeFree(mIndexTime).Visible = True
            mIndexTime = mIndexTime + 1
        End If
        
    Next
    
  
    i = 0
    For i = mHours2 To mHoursTo2
        If i > 12 Then
            mCaption = i - 12 & ":" & mMin2 & "PM"
        Else
            mCaption = i & ":" & mMin2 & "AM"
        End If
        mCaption = Format(mCaption, "HH:mm AM/PM")
        If ChkTimeForEmp(mCaption) Then
            lblTimeFree(mIndexTime) = mCaption
            lblTimeFree(mIndexTime).Visible = True
            mIndexTime = mIndexTime + 1
        End If
    Next
    
    
    
    
    rsDummy.MoveNext
Loop

End Sub

Private Function ChkTimeForEmp(ByVal mCaption As String) As Boolean
 Dim s As String
    s = "select * from TblStudCalling2 "
    s = s & " Inner join TblStudCalling On TblStudCalling.Id = TblStudCalling2.MasterId"
    s = s & " Where  TblStudCalling2.EmpId  = " & val(Fg4.TextMatrix(Fg4.Row, Fg4.ColIndex("EmpID")))
    s = s & " AND CAST(HoursT as time) ='" & FormatDateTime(mCaption, vbShortTime) & "'"
    s = s & " And TblStudCalling.EnterDate =" & SQLDate(Me.EnterDate.value, True) & ""
    If Trim(TxtSerial1) <> "" Then
        s = s & " and MasterId <> " & val(TxtSerial1)
    End If
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic
    If rsDummy.EOF Then
        ChkTimeForEmp = True
    Else
        ChkTimeForEmp = False
    End If
End Function

Private Function GetFeildDay(mEmpId As Long) As String
Dim s As String
s = ""
Select Case Weekday(EnterDate, 0)

Case 1
    s = s & " Select * from ("
    s = s & " SELECT TbLSheft.SatWoVo,"
    s = s & "    Shiftfrom AS FromTime1,"
    s = s & "    ShiftTO  as ToTime1,"
    s = s & "        ShfitFromW AS FromTime2,"
    s = s & "        ShfitToW  as ToTime2,"
    s = s & "        SeftCode"
    s = s & " From TbLSheft"
    s = s & " INNER JOIN TblShiftWorker"
    s = s & "             ON  TbLSheft.SeftCode = TblShiftWorker.ShiftID"
    s = s & " Where SatWoVo = 0"
    s = s & " and TblShiftWorker.EmpID = " & mEmpId
    s = s & ") T "
Case 2
    s = s & " Select * from ("
    s = s & " SELECT TbLSheft.SunWoVo,"
    s = s & "    FromSun AS FromTime1,"
    s = s & "    ToSun  as ToTime1,"
    s = s & "        FromSunW AS FromTime2,"
    s = s & "        ToSunW  as ToTime2,"
    s = s & "        SeftCode"
    s = s & " From TbLSheft"
    s = s & " INNER JOIN TblShiftWorker"
    s = s & "             ON  TbLSheft.SeftCode = TblShiftWorker.ShiftID"
    s = s & " Where SunWoVo = 0"
    s = s & " and TblShiftWorker.EmpID = " & mEmpId
    s = s & ") T "

Case 3

    s = s & " Select * from ("
    s = s & " SELECT TbLSheft.MonWoVo,"
    s = s & "    FromMon AS FromTime1,"
    s = s & "    ToMon  as ToTime1,"
    s = s & "        FromMonW AS FromTime2,"
    s = s & "        ToMonW  as ToTime2,"
    s = s & "        SeftCode"
    s = s & " From TbLSheft"
    s = s & " INNER JOIN TblShiftWorker"
    s = s & "             ON  TbLSheft.SeftCode = TblShiftWorker.ShiftID"
    s = s & " Where MonWoVo = 0"
    s = s & " and TblShiftWorker.EmpID = " & mEmpId
    s = s & ") T "
Case 4

    s = s & " Select * from ("
    s = s & " SELECT TbLSheft.TuesWoVo,"
    s = s & "    FromTues AS FromTime1,"
    s = s & "    ToTues  as ToTime1,"
    s = s & "        FromTuesW AS FromTime2,"
    s = s & "        ToTuesW  as ToTime2,"
    s = s & "        SeftCode"
    s = s & " From TbLSheft"
    s = s & " INNER JOIN TblShiftWorker"
    s = s & "             ON  TbLSheft.SeftCode = TblShiftWorker.ShiftID"
    s = s & " Where TuesWoVo = 0"
    s = s & " and TblShiftWorker.EmpID = " & mEmpId
    s = s & ") T "
Case 5

    s = s & " Select * from ("
    s = s & " SELECT TbLSheft.WedWoVo,"
    s = s & "    FromWed AS FromTime1,"
    s = s & "    ToWed  as ToTime1,"
    s = s & "        FromWedW AS FromTime2,"
    s = s & "        ToWedW  as ToTime2,"
    s = s & "        SeftCode"
    s = s & " From TbLSheft"
    s = s & " INNER JOIN TblShiftWorker"
    s = s & "             ON  TbLSheft.SeftCode = TblShiftWorker.ShiftID"
    s = s & " Where WedWoVo = 0"
    s = s & " and TblShiftWorker.EmpID = " & mEmpId
    s = s & ") T "
Case 6

    s = s & " Select * from ("
    s = s & " SELECT TbLSheft.ThurWoVo,"
    s = s & "    FromThru  AS FromTime1,"
    s = s & "   ToThru as ToTime1,"
    s = s & "        FromThruW AS FromTime2,"
    s = s & "        ToThruW  as ToTime2,"
    s = s & "        SeftCode"
    s = s & " From TbLSheft"
    s = s & " INNER JOIN TblShiftWorker"
    s = s & "             ON  TbLSheft.SeftCode = TblShiftWorker.ShiftID"
    s = s & " Where ThurWoVo = 0"
    s = s & " and TblShiftWorker.EmpID = " & mEmpId
    s = s & ") T "

Case 7
    s = s & " Select * from ("
    s = s & " SELECT TbLSheft.FrirWoVo,"
    s = s & "    FromFri AS FromTime1,"
    s = s & "    ToFri  as ToTime1,"
    s = s & "        FromFriW AS FromTime2,"
    s = s & "        ToFriW  as ToTime2,"
    s = s & "        SeftCode"
    s = s & " From TbLSheft"
    s = s & " INNER JOIN TblShiftWorker"
    s = s & "             ON  TbLSheft.SeftCode = TblShiftWorker.ShiftID"
    s = s & " Where FrirWoVo = 0"
    s = s & " and TblShiftWorker.EmpID = " & mEmpId
    s = s & ") T "
End Select
'Select case
GetFeildDay = s
End Function
Private Sub fg4_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With Fg4

   Select Case .ColKey(Col)
        Case "Amount0", "Amount2", "Amount3", "PercentV", "Amount", "DateStart", "DateEnd", "JobOrdersNo", "Hours", "PeriodT", "Hours", "Code", "Emp_Code", "Emp_code"
            .ComboList = ""
        Case "NoteNo"
            .ComboList = ""
        Case "DayMeter"
            .ComboList = ""
        Case "CustName", "PercentV"
            Cancel = True
        End Select
        
    End With
 
End Sub

Private Sub fg4_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg4

        Select Case .ColKey(Col)
 
            Case "ItemName"
             .TextMatrix(Row, .ColIndex("ItemName")) = ""
                StrSQL = "select ItemID,ItemName,ItemNamee from TblItems "
                If Trim(.TextMatrix(Row, .ColIndex("EmpName"))) = "" Then
                Else
                    StrSQL = StrSQL & " WHERE TblItems.ItemID IN (SELECT ItemID FROM TblEmpItemsTrans2 WHERE TblEmpItemsTrans2.EmpID = " & val(.TextMatrix(Row, .ColIndex("EmpID"))) & ")"
                End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg4.BuildComboList(rs, "ItemName", "ItemID")
                Else
                    StrComboList = Fg4.BuildComboList(rs, "ItemNamee", "ItemID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "EmpName"
             .TextMatrix(Row, .ColIndex("EmpName")) = ""
                StrSQL = "SELECT Emp_Id,Emp_Name,Emp_Namee FROM TblEmployee  "
                If Trim(.TextMatrix(Row, .ColIndex("ItemName"))) <> "" Then
                    StrSQL = StrSQL & " WHERE   TblEmployee.Emp_ID IN (SELECT EmpID FROM TblEmpItemsTrans2 WHERE TblEmpItemsTrans2.ItemID = " & val(.TextMatrix(Row, .ColIndex("ItemID"))) & ")"
                Else
                End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg4.BuildComboList(rs, "Emp_Name", "Emp_Id")
                Else
                    StrComboList = Fg4.BuildComboList(rs, "Emp_Namee", "Emp_Id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "ReservationTypeName"
                .TextMatrix(Row, .ColIndex("ReservationTypeName")) = ""
                StrSQL = "SELECT Id,Name,Namee FROM tblReservationType "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg4.BuildComboList(rs, "Name", "Id")
                Else
                    StrComboList = Fg4.BuildComboList(rs, "Namee", "Id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
                 
            End Select
        End With
End Sub


Private Sub ISButton8_Click()
FrmSearStudent.inde = 9
Load FrmSearStudent
FrmSearStudent.show vbModal
End Sub

Private Sub Label20_Click()
Frame1.Visible = False
End Sub

Private Sub Pic_Days_Click(Index As Integer)

End Sub

Private Sub Pic_Empty_Click(Index As Integer)

End Sub

Private Sub lblTimeFree_Click(Index As Integer)
'    lblTimeFree(Index).backcolor = vbRed
    Fg4.TextMatrix(Fg4.Row, Fg4.ColIndex("HoursT")) = lblTimeFree(Index)
    Frame2.Visible = False
'    lblTimeFree(Index).backcolor = vbBlue
End Sub

Private Sub TxtMobile_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMobile.Text, 1)
If KeyAscii = vbKeyReturn Then
    GetCustomerNamebyPhone (TxtMobile.Text)
End If
End Sub


Public Sub GetCustomerNamebyPhone(Optional ByVal phone As String = "", Optional ByVal Name As String = "", Optional ByVal CUSTID As String = "", Optional ByVal SearchCode As String = "")
            If phone = "" And Name = "" And CUSTID = "" And SearchCode = "" Then Exit Sub
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

        If phone <> "" Then
            sql = "SELECT     Cus_mobile , CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (Cus_mobile = '" & phone & "')"
        ElseIf Name <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (CusName = '" & Name & "')"
        ElseIf CUSTID <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (CusID = " & val(CUSTID) & ")"
        ElseIf SearchCode <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     Fullcode ='" & SearchCode & "'"
        Else
        Exit Sub
        End If
  
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        TxtMobile = rs!Cus_mobile & ""
        
        Text15.Text = rs!Fullcode & ""
        DcbCompany.BoundText = val(rs!CusID & "")
        'DcboEmp.BoundText = val(rs!EmpID & "")
        txtCustomerName.Text = IIf(IsNull(rs!CusName), "", rs!CusName)

    Else
         TxtMobile = ""
         Text15 = ""
         DcbCompany.BoundText = ""
          txtCustomerName.Text = ""
              If Me.TxtModFlg <> "R" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·⁄„Ì· €Ì— „ÊÃÊœ", vbCritical
        Else
            MsgBox "This client does not exist", vbCritical
        End If
End If
    End If

    rs.Close

End Sub
Private Sub TxtSudCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
 Dim UQama As String
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, TxtSudCode.Text, 1, UQama
        DcbStudent.BoundText = EmpID
        TxtUQama.Text = UQama
        GetStudentCalling
    End If
End Sub

Private Sub TxtUQama_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
 Dim Fullcode As String
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, Fullcode, 2, TxtUQama.Text
        DcbStudent.BoundText = EmpID
        TxtSudCode.Text = Fullcode
        GetStudentCalling
    End If
End Sub
Private Sub DcbEmployee_Click(Area As Integer)
If val(Me.DcbEmployee.BoundText) = 0 Then Exit Sub
           Me.Text1.Text = get_EMPLOYEE_Data(val(Me.DcbEmployee.BoundText), "Fullcode")
End Sub

Private Sub EnterDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         EnterDateH.value = ToHijriDate(EnterDate.value)
End If
End Sub
Private Sub EnterDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 EnterDate.value = ToGregorianDate(EnterDateH.value)
End If
End Sub
Function GetEmpID() As Double
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = "Select * from TblUsers where UserID =" & user_id & ""
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
GetEmpID = IIf(IsNull(Rs2("Empid").value), 0, Rs2("Empid").value)
Else
GetEmpID = 0
End If
End Function
 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from  TblStudCalling  "
    conection = conection & "  where  (BranchID=0 or BranchID is null or         BranchID in(" & Current_branchSql & "))"
    conection = conection & " Order By ID"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany
   Dcombos.GetEmployees Me.DcbEmployee
    BtnLast_Click
    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
   Me.Refresh
   FiLLTXT
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
  Dim StrSQL As String
 StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(txtNoteSerialCash(1).Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblMultuPayment Where NoteID=" & val(txtNoteSerialCash(1).Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "Delete From Notes Where NoteID=" & val(txtNoteSerialCash(1).Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
    '            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
    '            Cn.Execute StrSQL, , adExecuteNoRecords
    
    
                StrSQL = " delete   notes where   NoteId=" & val(txtNoteSerialCash(1).Text)
  Cn.Execute StrSQL
    Dim sql As String
    Dim ID As Double
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("RecordDateH").value = RecordDateH.value
   RsSavRec.Fields("RecordDate").value = RecordDate.value
   RsSavRec.Fields("EmpID").value = val(Me.DcbEmployee.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("CompID").value = val(Me.DcbCompany.BoundText)
   RsSavRec.Fields("Remarks").value = txtRemarks.Text
   RsSavRec.Fields("EnterDateH").value = EnterDateH.value
   RsSavRec.Fields("EnterDate").value = EnterDate.value
   RsSavRec.Fields("EnterTime").value = FormatDateTime(EnterTime.value, vbShortTime)
   RsSavRec.Fields("StudID").value = val(Me.DcbStudent.BoundText)
   RsSavRec.Fields("Phone").value = Me.TxtStudentPhone.Text
   RsSavRec.Fields("Mobile").value = Me.TxtMobile.Text
   RsSavRec.Fields("Email").value = Me.TxtStudentEmail.Text
   RsSavRec.Fields("Arboun").value = val(Me.txtArboun.Text)
   
       
   RsSavRec.update
   
   
   
   Dim s As String
    
    s = " Delete From TblStudCalling2 Where MasterID = " & val(TxtSerial1.Text)
    
        
        
    
    Cn.Execute s
    
    s = "Select * from TblStudCalling2 Where Id = -1"
    saveGrid s, Fg4, "ItemID", "SerID", "MasterID", val(TxtSerial1.Text)

    
       If val(txtArboun) <> 0 Then
            If Not CreateCash Then GoTo ErrTrap
        End If
   RsSavRec.Resync adAffectCurrent
    Dim Msg As String
      Select Case Me.TxtModFlg.Text
        Case "N"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
               ' Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
     
Private Function CreateCash() As Boolean
CreateCash = False
         Dim rsCash As New ADODB.Recordset
         Dim StrSQL As String
    'StrSQL = "select * From Notes where NoteType=4 and   displayed is null Order By NoteID"
    StrSQL = "select * From Notes where NoteType=-1"
'StrSQL = StrSQL & " and CashingType<=11 and akarid is Null"

    'If SystemOptions.usertype <> UserAdminAll Then
    '    StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    'End If
   On Error GoTo Err

    rsCash.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText


        If TxtModFlg.Text = "N" Then
            txtNoteSerialCash(1).Text = CStr(new_id("Notes", "NoteID", "", True))
            'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            rsCash.AddNew
       
            rsCash("NoteID").value = val(txtNoteSerialCash(1).Text)
            'Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
         
        ElseIf TxtModFlg.Text = "E" Then
    
               txtNoteSerialCash(1).Text = CStr(new_id("Notes", "NoteID", "", True))
            'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            rsCash.AddNew
       
            rsCash("NoteID").value = val(txtNoteSerialCash(1).Text)
            
         End If


            Dim Current_case As Integer, s As String, mBoxID As Long
            Dim rsOut As New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & Me.DcbEmployee.BoundText



            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                mBoxID = val(rsOut!BoxID & "")
            End If
            If mBoxID = 0 Then
                rsOut.Close
                
                s = " SELECT tu.BoxID FROM TblUsers AS tu where UserId = " & user_id
                rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsOut.EOF Then
                    mBoxID = val(rsOut!BoxID & "")
                End If
            End If
        If mBoxID = 0 Then
            MsgBox "ÌÃ»  ”ÃÌ· Œ“Ì‰… ··„” Œœ„ «Ê ··»«∆⁄"
            Exit Function
        End If

        rsCash("branch_no").value = val(Me.DcbBranch.BoundText)
        rsCash("EmpId").value = IIf(Me.DcbEmployee.BoundText = "", Null, (Me.DcbEmployee.BoundText))
        'rsCash("foxy_no").value = val(Text1.Text)
        'rsCash("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        'rsCash("Prefix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)

        'rsCash("CarId").value = IIf(Me.Dccar.BoundText = "", Null, (Me.Dccar.BoundText))
        'rsCash("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    
        If val(txtNoteSerialCash(0).Text) = 0 Then
            txtNoteSerialCash(0).Text = Voucher_coding(val(DcbBranch.BoundText), RecordDate.value, 2, 4, , , "")
        End If
        Dim mNoteSerial As String
        
            mNoteSerial = Notes_coding(val(DcbBranch.BoundText), RecordDate.value)
       
        
'        If CboStatus.ListIndex <> 0 Then
'        TxtNoteSerial.Text = ""
'
'        End If
       
    'If Option1.value = True Then
  '     rsCash("NCashingType").value = 1
   'ElseIf optIsEmp.value = True Then
   '     rsCash("NCashingType").value = 2
   'ElseIf optCash.value = True Then
   '     rsCash("NCashingType").value = 3
   '    ElseIf Option7.value = True Then
   '     rsCash("NCashingType").value = 7
        
   ' Else
    
         rsCash("NCashingType").value = 0
  ' End If
       
    
        'rsCash("ContainerNo").value = IIf(Trim(Me.txtContainerNo.Text) = "", Null, Trim(Me.txtContainerNo.Text))
        'rsCash("ManulaNO").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
        'rsCash("ManualNo").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
        'rsCash("BookNo").value = IIf(Trim(Me.TxtBookNo.Text) = "", Null, Trim(Me.TxtBookNo.Text))
        
        '
        rsCash("NoteSerial").value = mNoteSerial
        rsCash("NoteSerial1").value = IIf(Trim(Me.txtNoteSerialCash(0).Text) = "", Null, Trim(Me.txtNoteSerialCash(0).Text))
        'rsCash("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
        rsCash("NCashingType").value = 2
    
        'rsCash("person").value = IIf(TXTperson.Text = "", "", Trim(TXTperson.Text))
        rsCash("Note_Value").value = IIf(txtArboun.Text = "", Null, val(txtArboun.Text))
        'rsCash("Adv_payment_value").value = IIf(txtAdv_payment_value.Text = "", Null, val(txtAdv_payment_value.Text))
        'rsCash("VAT").value = IIf(TxtVATValue.Text = "", Null, val(TxtVATValue.Text))
    
        '    Rs("Remark").value = IIf(dcproject.BoundText = "", "", Trim(dcproject.BoundText))
        'If lblinvoices.Caption = "" Then
        rsCash("Remark").value = "”‰œ ﬁ»÷ ¬·Ï „‰ ›« Ê—… ”‰œ ÕÃ“ ⁄—»Ê‰ —ﬁ„" & TxtSerial1
        'Else
        'rsCash("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text)) & vbEnter & lblinvoices.Caption
        'End If
        
        'rsCash("BankName").value = IIf(TXTBankName.Text = "", "", Trim(TXTBankName.Text))
        rsCash("NoteType").value = 4
        rsCash("NoteDate").value = RecordDate.value
        rsCash("BillTransNo").value = TxtSerial1.Text
        rsCash("BillTransID").value = val(TxtSerial1.Text)
        rsCash("Transaction_ID").value = Null  'val(TxtSerial1.Text)
        
        'rsCash("BillMaintNo").value = TxtBillMaintNo.Text
        'rsCash("BillMaintID").value = val(TxtBillMaintID.Text)
        'rsCash("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        'rsCash("NoteDateH").value = Me.Txt_DateHigri.value


        rsCash("CashingType").value = 0
        
        '
        rsCash("TotalNotesValue").value = 0
        
        rsCash("CurrentBalance").value = val(txtArboun)
        rsCash("PaymentValue").value = val(txtArboun)
        'rsCash("Percentage").value = val(TxtPercentage.Text)
        'rsCash("PercentageValue").value = val(TxtPercentageValue.Text)
        
        
        rsCash("CusID").value = IIf(DcbCompany.Text = "", Null, DcbCompany.BoundText)
     
       

        '--------------------------------------------------------------------------
        'ÿ—Ìﬁ… «·œ›⁄ «·‰ﬁœÏ «Ê «·‘Ìﬂ
        
        rsCash("NoteCashingType").value = 0
        rsCash("BoxID").value = mBoxID
        rsCash("BankID").value = Null
        rsCash("ChqueNum").value = Null
        rsCash("DueDate").value = Null
    
       

        '--------------------------------------------------------------------------
        rsCash("UserID").value = user_id
        rsCash("numbering_type").value = sand_numbering_type(0)   '”‰œ «·ﬁÌœ
        rsCash("numbering_type1").value = sand_numbering_type(2) '”‰œ «·ﬁ»÷
    
      
    
     '  If DCboCashType.ListIndex = 8 Then
     '       rsCash("ContractNo").value = IIf(TxtContractNo.Text = "", Null, TxtContractNo.Text)
     '       rsCash("ContNo").value = IIf(TXTContNo.Text = "", Null, TXTContNo.Text)
     '       Else
     '        rsCash("ContractNo").value = Null
     '        rsCash("ContNo").value = Null
     '   End If
        
        
   '  If DCboCashType.ListIndex = 9 Then
   ' rsCash("akarid").value = IIf(val(Me.DcbIqara.BoundText) <> 0, val(DcbIqara.BoundText), Null)
   '  rsCash.Fields("UnitType").value = IIf(Me.DcbUnitType.BoundText <> "", val(DcbUnitType.BoundText), Null)
   '  rsCash.Fields("UnitNo").value = IIf(Me.DcbUnitNo.BoundText <> "", val(DcbUnitNo.BoundText), Null)
  '   rsCash("interval").value = IIf(txtinterval.Text = "", Null, val(txtinterval.Text))
  '   rsCash("intervaltype").value = val(cbointervaltype.ListIndex)
  '   rsCash("renterName").value = IIf(txtrenterName.Text = "", Null, txtrenterName.Text)
  '            If cbointervaltype.ListIndex = 0 Then
  '            rsCash("allowdate").value = DateAdd("d", val(txtinterval), XPDtbTrans.value)
  '            ElseIf cbointervaltype.ListIndex = 1 Then
  '            rsCash("allowdate").value = DateAdd("M", val(txtinterval), XPDtbTrans.value)
  '
  '          ElseIf cbointervaltype.ListIndex = 2 Then
  '            rsCash("allowdate").value = DateAdd("YYYY", val(txtinterval), XPDtbTrans.value)
  '
  '           End If
  '                rsCash("allowdateH").value = ToHijriDate(rsCash("allowdate").value)
  '
  '          Else
  '        rsCash("akarid").value = Null
  '   rsCash.Fields("UnitType").value = Null
  '   rsCash.Fields("UnitNo").value = Null
  '   rsCash("interval").value = Null
  '   rsCash("intervaltype").value = Null
  '   rsCash("renterName").value = Null
          
  '      End If
              
              
              
        
        rsCash("sanad_year").value = year(RecordDate.value)
        rsCash("sanad_month").value = Month(RecordDate.value)
    
       
        rsCash("note_value_by_characters").value = Trim$(val(txtArboun))
       

        
            rsCash("cus_or_sub").value = 0 '⁄„Ì· ‰Â«∆Ì
       
    
        rsCash.update
saveBillBuy2

CmdCreateV2_Click
s = "Update TblStudCalling Set  NoteIDCash = " & val(Me.txtNoteSerialCash(1).Text) & ",NoteSerialCash = " & Trim(Me.txtNoteSerialCash(0).Text) & " Where ID = " & val(val(TxtSerial1.Text))
    
                    
Cn.Execute s
CreateCash = True
Exit Function
Err:
CreateCash = False
End Function


Private Sub CmdCreateV2_Click()
Dim s As String
'CHECKaCCOUNTS

     


'END CHECK

If Not createVoucher2 Then Exit Sub
       'FindRec val(TXTLCNO.Text)
       
    

'Me.TxtModFlg2(mIndex).Text = "R"
End Sub
Function createVoucher2() As Boolean

'ee
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    Õ”«» «·" '& TxtNoteSerial.Text


Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double
tablename = "Notes"

Filedname = "NoteID"
'NoteSerial1 = CInt(val(txtNoteSerialCash(0).Text))

BranchID = val(DcbBranch.BoundText)
mRate = 1

'



notytype = 4
Notevalue = val(txtArboun)

'mAccNO = val(DboParentAccount.BoundText)
NoteDate = (RecordDate.value)
 
If Notevalue > 0 Then
   

    If Not CREATE_VOUCHER_GE2(val(txtNoteSerialCash(1).Text), BranchID, val(DCboUserName.BoundText), NoteDate) Then createVoucher2 = False Else createVoucher2 = True
    RsSavRec.Resync adAffectCurrent

    updateNotesValueAndNobytext val(txtNoteSerialCash(0).Text), Format(txtArboun.Text, "###.00")
'
'
'    StrSQL = "update  " & tablename & "   set NoteID=" & NoteID & ",NoteSerial='" & NoteSerial & "'"
'
'    StrSQL = StrSQL & " Where " & Filedname & " = " & NoteSerial1 & ""
'    Cn.Execute StrSQL
     
     
 
End If
End Function


Public Function CREATE_VOUCHER_GE2(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date) As Boolean
Dim StrSQL As String

    Dim Current_case As Integer, s As String, mBoxID As Long
            Dim rsOut As New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & Me.DcbEmployee.BoundText



            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                mBoxID = val(rsOut!BoxID & "")
            End If
                        If mBoxID = 0 Then
                rsOut.Close
                
                s = " SELECT tu.BoxID FROM TblUsers AS tu where UserId = " & user_id
                rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsOut.EOF Then
                    mBoxID = val(rsOut!BoxID & "")
                End If
            End If



'Dim StrAccountCodeDebt As String
Dim StrAccountCodeCridet As String
Dim StrAccountCodeDebt As String
StrAccountCodeDebt = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)   '«·„»Ì⁄« 
StrAccountCodeCridet = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbCompany.BoundText))

     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    'Dim StrAccountCodeDebt As String
    'Dim StrAccountCodeCridet As String
    Dim X As Integer
   
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtSerial1.Text
    notes_id = general_noteid
    my_branch = val(DcbBranch.BoundText)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
    
    'Dim s As String
    Dim mRate As Double
    mRate = 1
    ' „‰ Õ”«» «·⁄„Ì·
    
    

   
    Notevalue = val(txtArboun.Text)
    If Notevalue > 0 Then
        
       ' StrAccountCodeDebt = Trim(DboParentAccount.BoundText)
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  «·’‰œÊﬁ  ", val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
        val(DcbBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , DcbCompany.BoundText) = False Then
            GoTo ErrTrap
        End If
       ' «·Ï Õ”«» «·ﬁÌ„… «·„÷«›…
      
        
        line_no = line_no + 1

    End If

    
    ' «·«ÿ—«›
    
     ' «·Ï Õ”«» «·⁄„Ì·
         
  '  Notevalue = val(txtTotal.Text)
    If Notevalue > 0 Then
    
              

        
        
 
        
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«» «·⁄„Ì·  ", val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
        val(DcbBranch.BoundText)) = False Then
            GoTo ErrTrap
        End If

        line_no = line_no + 1
    End If
    

    updateNotesValueAndNobytext (val(notes_id))
    CREATE_VOUCHER_GE2 = True
    Exit Function
ErrTrap:
CREATE_VOUCHER_GE2 = False
txtNoteSerialCash(1) = ""
txtNoteSerialCash(0) = ""

 

     
 
    '
 




 

End Function



Function saveBillBuy2()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
Dim RsDetails As ADODB.Recordset
    
    StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.txtNoteSerialCash(1).Text) & " and TransType is null"
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.txtNoteSerialCash(1).Text) & " and TransType is null"
    Cn.Execute StrSQL, , adExecuteNoRecords
   
    Dim mTotal As Double
    mTotal = val(txtArboun)
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblNotesBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
  
    'TxtValueTemp.Text = val(XPTxtVal.Text)
    
            RsDetails.AddNew
           ' val (Me.txtNoteSerialCash(1).Text)
            
            RsDetails("NoteID1").value = val(Me.txtNoteSerialCash(1).Text)
            RsDetails("NoteID").value = val(TxtSerial1.Text)
            RsDetails("branch_no").value = val(DcbBranch.BoundText)
            RsDetails("NoteSerial1").value = val(TxtSerial1)
            RsDetails("Note_Value").value = val(mTotal)
            Note_Value1 = val(txtArboun)
            Diff = 0
'            If val(TxtValueTemp.Text) > 0 Then
'          If val(TxtValueTemp.Text) <= Note_Value1 Then
'          Diff = val(TxtValueTemp.Text)
'          TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
'          Else
'          Diff = Note_Value1
'          TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
'          End If
'            End If
          ' .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("RemainingValue")))
            '.TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
            
            'RsDetails("PayedValue").value = val(XPTxtValue(3)) ' val(.TextMatrix(i, .ColIndex("PayedValue")))
            
            'RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
            RsDetails("NoteDate").value = RecordDate.value
           
            RsDetails("DueDate").value = Null
          
            RsDetails("TransPayedValue").value = val(txtArboun)
           '.TextMatrix(i, .ColIndex("NetValue")) = val(XPTxtValue(3))
            RsDetails("NetValue").value = val(txtArboun)
            RsDetails("RemainingValue").value = val(mTotal)
            RsDetails.update
                
          
      

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    

            RsDetails.AddNew
            RsDetails("NoteID").value = val(txtNoteSerialCash(1).Text)
            RsDetails("RecDate").value = RecordDate.value
            RsDetails("Serial").value = txtNoteSerialCash(0).Text
            RsDetails("Transaction_ID").value = val(TxtSerial1.Text)
            RsDetails("Note_Value").value = val(mTotal)
            RsDetails("PayedValue").value = val(txtArboun)
            RsDetails.update



End Function

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
    Dim Shifttime As Date
    EnterDate.Enabled = False
    EnterTime.Enabled = False
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbCompany.BoundText = IIf(IsNull(RsSavRec.Fields("CompID").value), "", RsSavRec.Fields("CompID").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DcbEmployee.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Me.txtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    EnterDateH.value = IIf(IsNull(RsSavRec.Fields("EnterDateH").value), ToHijriDate(Date), RsSavRec.Fields("EnterDateH").value)
    EnterDate.value = IIf(IsNull(RsSavRec.Fields("EnterDate").value), Date, RsSavRec.Fields("EnterDate").value)
    
    Me.TxtMobile.Text = IIf(IsNull(RsSavRec.Fields("Mobile").value), "", RsSavRec.Fields("Mobile").value)
    Me.TxtStudentPhone.Text = IIf(IsNull(RsSavRec.Fields("Phone").value), "", RsSavRec.Fields("Phone").value)
    Me.TxtStudentEmail.Text = IIf(IsNull(RsSavRec.Fields("Email").value), "", RsSavRec.Fields("Email").value)
    Me.DcbStudent.BoundText = IIf(IsNull(RsSavRec.Fields("StudID").value), "", RsSavRec.Fields("StudID").value)
    txtArboun.Text = IIf(IsNull(RsSavRec.Fields("Arboun").value), "", RsSavRec.Fields("Arboun").value)
    
    txtNoteSerialCash(1) = IIf(IsNull(RsSavRec("NoteIDCash").value), "", (RsSavRec("NoteIDCash").value))
    txtNoteSerialCash(0) = IIf(IsNull(RsSavRec("NoteSerialCash").value), "", (RsSavRec("NoteSerialCash").value))
        
    If Not IsNull(RsSavRec("EnterTime").value) Then
        Shifttime = FormatDateTime(RsSavRec("EnterTime").value, vbShortTime)
        Me.EnterTime.value = Shifttime
    End If
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60


Dim s As String



s = " SELECT TblStudCalling2.*,TblEmployee.Emp_Code,TblItems.FullCode Code,"
s = s & "       TblEmployee.Emp_Name        EmpName,"
s = s & "       TblItems.ItemName,"
s = s & "       tblReservationType.Name  AS ReservationTypeName"
s = s & " From TblStudCalling2"
s = s & "       Left Outer JOIN tblReservationType"
s = s & "            ON  tblReservationType.ID = TblStudCalling2.ReservationTypeCode"
s = s & "       INNER JOIN TblEmployee"
s = s & "            ON  TblEmployee.Emp_ID = TblStudCalling2.EmpID"
s = s & "       INNER JOIN TblItems"
s = s & "            ON  TblItems.ItemID = TblStudCalling2.ItemID"
s = s & " Where TblStudCalling2.MasterId = " & val(TxtSerial1)
loadgrid s, Fg4, True, True

ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    
    
If SystemOptions.CustMobNoMandatory Then
        If Trim(TxtMobile) = "" Or Len(Trim(TxtMobile)) < 10 Or mId(Trim(TxtMobile), 1, 1) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·ÃÊ«· ’ÕÌÕ „‰ 10 Œ«‰«  Ì»œ√ »’›—"
        Else
            MsgBox "Please Enter Mob No."
        End If
        Exit Sub
    Else
    
    
    
    
    
   ' txtCodeSend = "+966"
   Dim isFound As Boolean
    If Not FindString(TxtMobile, "+966", 1) Then
        If Not FindString(TxtMobile, "00966", 1) And Not FindString(TxtMobile, "966", 1) Then
            isFound = False
        Else
            isFound = True
        End If
        If Not isFound Then
            isFound = False
                      TxtMobile = "00966" & mId(TxtMobile, 2, Len(TxtMobile))
        End If
    End If
    'TxtMobile = CDbl(TxtMobile)
    If Len(TxtMobile) <> 14 Then
        MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·ÃÊ«· ’ÕÌÕ „‰ 10 Œ«‰«  Ì»œ√ »’›—"
        'txtCodeSend = "+966"
         Exit Sub
    Else
      '  txtCodeSend = ""
        'txtPhoneCust = "+966" & val(txtPhoneCust)
    End If
    
    
    End If
End If
If val(Me.DcbBranch.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
Else
MsgBox "Please Select Branch"
End If
DcbBranch.SetFocus
Exit Sub
End If
If val(Me.DcbEmployee.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„ ’·"
Else
MsgBox "Please Select Caller"
End If
'DcbEmployee.SetFocus
Exit Sub
End If

Dim i As Long
For i = 1 To Fg4.Rows - 1
    If Trim(Fg4.TextMatrix(i, Fg4.ColIndex("ItemName"))) <> "" Or Trim(Fg4.TextMatrix(i, Fg4.ColIndex("EmpName"))) <> "" Or Trim(Fg4.TextMatrix(i, Fg4.ColIndex("ReservationTypeName"))) <> "" Then
        If Trim(Fg4.TextMatrix(i, Fg4.ColIndex("ItemName"))) = "" Then
            MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· «œŒ«· «·„Â„…"
            Exit Sub
        End If
        
        If Trim(Fg4.TextMatrix(i, Fg4.ColIndex("EmpName"))) = "" Then
            MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· «œŒ«· «·„ÊŸ›…"
            Exit Sub
        End If
'        If Trim(FG4.TextMatrix(i, FG4.ColIndex("ReservationTypeName"))) = "" Then
'            MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· «œŒ«· «·ÕÃ“"
'            Exit Sub
'        End If
        
        
    End If
    
Next

If Trim(TxtMobile) = "" Or Trim(DcbCompany.Text) = "" Then
            MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· «œŒ«· «·⁄„Ì· Ê—ﬁ„ «·ÃÊ«·"
            Exit Sub
    
End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    SendMessage
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblStudCalling", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Private Sub RecordDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecordDateH.value = ToHijriDate(RecordDate.value)
End If
End Sub
Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 RecordDate.value = ToGregorianDate(RecordDateH.value)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
    Me.DcbEmployee.BoundText = GeTEmpIDByEmpCode(Text1.Text, True)
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID
        DcbCompany.BoundText = EmpID
    End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim sql As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
         
                RsSavRec.Find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            Dim s As String
    
            s = " Delete From TblStudCalling2 Where MasterID = " & val(TxtSerial1.Text)
            
                
                
            
            Cn.Execute s
                            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(txtNoteSerialCash(1).Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblMultuPayment Where NoteID=" & val(txtNoteSerialCash(1).Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "Delete From Notes Where NoteID=" & val(txtNoteSerialCash(1).Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
    '            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
    '            Cn.Execute StrSQL, , adExecuteNoRecords
    
    
                StrSQL = " delete   notes where   NoteId=" & val(txtNoteSerialCash(1).Text)
                
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
             
               '''''''''''''''''''''''''''''''

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
     LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select

End Sub

' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
        Select Case Me.TxtModFlg.Text
            Case "N"
                    If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
                    Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                        End If
                    Case "E"
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
                    Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                 End If
        End Select
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)
        Select Case IntResult
            Case vbYes
               Cancel = True
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    'XPDtbTrans.Enabled = True
      '  Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
   ' XPDtbTrans.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
  ' XPDtbTrans.Enabled = True
  '     Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
 
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        
        Me.DCboUserName.BoundText = user_id
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
            Msg = Msg & "It was being edited by another user on the network"
           
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    clear_all Me
    TxtModFlg.Text = "N"
    Me.DcbBranch.BoundText = Current_branch
    Me.DCboUserName.BoundText = user_id
    If GetEmpID() <> 0 Then
    Me.DcbEmployee.BoundText = GetEmpID()
    End If
    EnterDate.value = Date
    EnterDateH.value = ToHijriDate(EnterDate.value)
    EnterTime.value = Time
    EnterDate.Enabled = True
    EnterTime.Enabled = True
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      
        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
       
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub


Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
  Label1(2).Caption = "Calling Data"
lbl(4).Caption = "No"
lbl(25).Caption = "Date"
lbl(0).Caption = "Branch"
Label1(0).Caption = "Caller"
Label1(5).Caption = "Company"
lbl(15).Caption = "Remarks"
lbl(1).Caption = "Entry Time"
lbl(2).Caption = "Entry Date"
lbl(12).Caption = "Entry Date"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    Label1(12).Caption = "ID"
    lbl(5).Caption = "Telephone"
    lbl(17).Caption = "Mobile"
    lbl(9).Caption = "Email"
    Label1(1).Caption = "Student"
   ' C1Tab1.Caption = "Data"

    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    Me.lbl(14).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblStudCalling"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

