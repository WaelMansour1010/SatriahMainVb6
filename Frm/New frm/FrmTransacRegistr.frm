VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTransacRegistr 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11745
   ClientLeft      =   1410
   ClientTop       =   2970
   ClientWidth     =   18960
   Icon            =   "FrmTransacRegistr.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11745
   ScaleWidth      =   18960
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
      Left            =   19320
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmTransacRegistr.frx":6852
      Left            =   18960
      List            =   "FrmTransacRegistr.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   19320
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   19320
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   19080
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   19320
      TabIndex        =   5
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
      Left            =   19200
      TabIndex        =   6
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
      Left            =   19320
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
            Picture         =   "FrmTransacRegistr.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegistr.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegistr.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegistr.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegistr.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegistr.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegistr.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegistr.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   18960
      TabIndex        =   7
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
      ButtonImage     =   "FrmTransacRegistr.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18960
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   6480
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
      ButtonImage     =   "FrmTransacRegistr.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   18960
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   6960
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
      ButtonImage     =   "FrmTransacRegistr.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   11745
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   18960
      _cx             =   33443
      _cy             =   20717
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
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   0
         Width           =   18945
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   13
            Top             =   240
            Width           =   405
            _ExtentX        =   714
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
            ButtonImage     =   "FrmTransacRegistr.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   14
            Top             =   240
            Width           =   405
            _ExtentX        =   714
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
            ButtonImage     =   "FrmTransacRegistr.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   15
            Top             =   240
            Width           =   405
            _ExtentX        =   714
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
            ButtonImage     =   "FrmTransacRegistr.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   16
            Top             =   240
            Width           =   405
            _ExtentX        =   714
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
            ButtonImage     =   "FrmTransacRegistr.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ”ÃÌ· „⁄«„·…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   2
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   120
            Width           =   6600
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   16680
            Picture         =   "FrmTransacRegistr.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   960
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   930
         Width           =   18930
         _cx             =   33390
         _cy             =   1693
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
         Begin VB.TextBox TxtNoID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5670
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   -750
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   390
            Left            =   15480
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   360
            Width           =   1530
         End
         Begin VB.TextBox Txtbarcode 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   5745
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   360
            Width           =   2145
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   390
            Left            =   12600
            TabIndex        =   21
            Top             =   360
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   688
            _Version        =   393216
            Format          =   89391105
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmTransacRegistr.frx":17E16
            Height          =   315
            Left            =   630
            TabIndex        =   22
            Top             =   360
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   390
            Left            =   11055
            TabIndex        =   23
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   688
         End
         Begin MSComCtl2.DTPicker RecordTime 
            Height          =   390
            Left            =   9000
            TabIndex        =   24
            Top             =   360
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   688
            _Version        =   393216
            Format          =   89391106
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   300
            Index           =   7
            Left            =   4710
            TabIndex        =   29
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„⁄«„·…"
            Height          =   300
            Index           =   4
            Left            =   17280
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   300
            Index           =   2
            Left            =   14220
            TabIndex        =   27
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Êﬁ "
            Height          =   300
            Index           =   22
            Left            =   10155
            TabIndex        =   26
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»«—ﬂÊœ"
            Height          =   300
            Index           =   9
            Left            =   7920
            TabIndex        =   25
            Top             =   360
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   750
         Left            =   0
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   10185
         Width           =   18960
         _cx             =   33443
         _cy             =   1323
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   450
            Left            =   495
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   150
            Width           =   4290
            _cx             =   7567
            _cy             =   794
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
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   120
               Width           =   900
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   870
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   120
               Width           =   1080
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   120
               Width           =   825
            End
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   210
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   120
               Width           =   870
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8535
            TabIndex        =   32
            Top             =   150
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   345
            Index           =   8
            Left            =   15555
            TabIndex        =   33
            Top             =   150
            Width           =   1710
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8310
         Left            =   0
         TabIndex        =   34
         Top             =   1890
         Width           =   18945
         _cx             =   33417
         _cy             =   14658
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   128
         FrontTabColor   =   14871017
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "«·»Ì«‰«  «·«”«”Ì…|Õ«·… «·«⁄ „«œ|»Ì«‰« "
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   7890
            Left            =   45
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   45
            Width           =   18855
            _cx             =   33258
            _cy             =   13917
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   1725
               Left            =   0
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   0
               Width           =   18990
               _cx             =   33496
               _cy             =   3043
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
               WordWrap        =   0   'False
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
               Begin VB.ComboBox DcbImportExport 
                  Height          =   315
                  ItemData        =   "FrmTransacRegistr.frx":17E2B
                  Left            =   12765
                  List            =   "FrmTransacRegistr.frx":17E2D
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   150
                  Width           =   4170
               End
               Begin VB.TextBox TxtNoImpExp 
                  Alignment       =   1  'Right Justify
                  Height          =   480
                  Left            =   9240
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   150
                  Width           =   2145
               End
               Begin VB.TextBox TxtSummary 
                  Alignment       =   1  'Right Justify
                  Height          =   975
                  Left            =   405
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   37
                  Top             =   675
                  Width           =   16530
               End
               Begin MSComCtl2.DTPicker ImpExpDate 
                  Height          =   480
                  Left            =   2910
                  TabIndex        =   40
                  Top             =   150
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   847
                  _Version        =   393216
                  Format          =   89391105
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal ImpExpDateH 
                  Height          =   480
                  Left            =   405
                  TabIndex        =   41
                  Top             =   150
                  Width           =   1860
                  _ExtentX        =   3281
                  _ExtentY        =   847
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‰Ê⁄"
                  Height          =   435
                  Index           =   11
                  Left            =   17055
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   150
                  Width           =   1740
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—ﬁ„"
                  Height          =   435
                  Index           =   1
                  Left            =   11670
                  TabIndex        =   44
                  Top             =   150
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒÂ"
                  Height          =   435
                  Index           =   0
                  Left            =   5055
                  TabIndex        =   43
                  Top             =   150
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·Œ’ «·”‰œ"
                  Height          =   360
                  Index           =   3
                  Left            =   17055
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   870
                  Width           =   1740
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   1695
               Left            =   0
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1710
               Width           =   18990
               _cx             =   33496
               _cy             =   2990
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
               Begin VB.TextBox TxtProcedureReq 
                  Alignment       =   1  'Right Justify
                  Height          =   960
                  Left            =   405
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   49
                  Top             =   660
                  Width           =   16530
               End
               Begin VB.ComboBox DcbMHDID 
                  Height          =   315
                  ItemData        =   "FrmTransacRegistr.frx":17E2F
                  Left            =   6060
                  List            =   "FrmTransacRegistr.frx":17E31
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   150
                  Width           =   1665
               End
               Begin VB.TextBox TxtMHD 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   150
                  Width           =   1020
               End
               Begin MSComCtl2.DTPicker ExitTime 
                  Height          =   420
                  Left            =   405
                  TabIndex        =   50
                  Top             =   855
                  Visible         =   0   'False
                  Width           =   2310
                  _ExtentX        =   4075
                  _ExtentY        =   741
                  _Version        =   393216
                  Format          =   89391106
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker RequerTime 
                  Height          =   420
                  Left            =   3885
                  TabIndex        =   51
                  Top             =   855
                  Visible         =   0   'False
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   741
                  _Version        =   393216
                  Format          =   89391106
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker EnterTime 
                  Height          =   450
                  Left            =   6795
                  TabIndex        =   52
                  Top             =   660
                  Visible         =   0   'False
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   794
                  _Version        =   393216
                  Format          =   89391106
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker EnterDate 
                  Height          =   465
                  Left            =   9135
                  TabIndex        =   53
                  Top             =   150
                  Width           =   2250
                  _ExtentX        =   3969
                  _ExtentY        =   820
                  _Version        =   393216
                  CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
                  Format          =   89391107
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcbTypTrans 
                  Bindings        =   "FrmTransacRegistr.frx":17E33
                  Height          =   315
                  Left            =   13035
                  TabIndex        =   54
                  Top             =   150
                  Width           =   3900
                  _ExtentX        =   6879
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  BackColor       =   16777215
                  ListField       =   "account_name"
                  BoundColumn     =   "code"
                  Text            =   ""
                  RightToLeft     =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MSComCtl2.DTPicker ExitDate 
                  Height          =   465
                  Left            =   405
                  TabIndex        =   55
                  Top             =   150
                  Width           =   2445
                  _ExtentX        =   4313
                  _ExtentY        =   820
                  _Version        =   393216
                  Enabled         =   0   'False
                  CustomFormat    =   "dd/MM/yyyy, hh:mm:tt"
                  Format          =   89391107
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  «·œŒÊ·"
                  Height          =   420
                  Index           =   6
                  Left            =   11310
                  TabIndex        =   61
                  Top             =   150
                  Width           =   1620
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·„⁄«„·…"
                  Height          =   420
                  Index           =   12
                  Left            =   17055
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   150
                  Width           =   1740
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  «·œŒÊ·"
                  Height          =   405
                  Index           =   5
                  Left            =   8430
                  TabIndex        =   59
                  Top             =   660
                  Visible         =   0   'False
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Êﬁ  «··«“„"
                  Height          =   420
                  Index           =   10
                  Left            =   7875
                  TabIndex        =   58
                  Top             =   150
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Êﬁ  «·„ Êﬁ⁄ ··Œ—ÊÃ"
                  Height          =   600
                  Index           =   13
                  Left            =   2760
                  TabIndex        =   57
                  Top             =   150
                  Width           =   2190
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã—«¡ «·„ÿ·Ê»"
                  Height          =   360
                  Index           =   14
                  Left            =   17055
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   855
                  Width           =   1740
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic9 
               Height          =   825
               Left            =   0
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   3390
               Width           =   18990
               _cx             =   33496
               _cy             =   1455
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
               Begin VB.TextBox TxtRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   675
                  Left            =   405
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   63
                  Top             =   75
                  Width           =   16530
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   240
                  Index           =   15
                  Left            =   17055
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   300
                  Width           =   1740
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic11 
               Height          =   3945
               Left            =   0
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   4050
               Width           =   18990
               _cx             =   33496
               _cy             =   6959
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
               Begin VB.CommandButton Command1 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "≈—”«·"
                  Height          =   435
                  Left            =   16935
                  Style           =   1  'Graphical
                  TabIndex        =   88
                  Top             =   3375
                  Width           =   1725
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   3105
                  Left            =   180
                  TabIndex        =   70
                  Top             =   225
                  Width           =   18630
                  _cx             =   32861
                  _cy             =   5477
                  Appearance      =   1
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
                  BackColorAlternate=   16777088
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483642
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   1
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
                  FormatString    =   $"FrmTransacRegistr.frx":17E48
                  ScrollTrack     =   -1  'True
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
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Height          =   3915
               Left            =   15
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   4050
               Width           =   19035
               Begin VB.CommandButton Command5 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "„Êﬁ› «·”‰œ"
                  Height          =   375
                  Left            =   15840
                  Style           =   1  'Graphical
                  TabIndex        =   121
                  Top             =   2280
                  Width           =   1575
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00E2E9E9&
                  Height          =   2175
                  Left            =   9000
                  TabIndex        =   80
                  Top             =   0
                  Width           =   8685
                  Begin VB.ListBox ListDeptSelect 
                     Height          =   1620
                     ItemData        =   "FrmTransacRegistr.frx":17FD4
                     Left            =   120
                     List            =   "FrmTransacRegistr.frx":17FDB
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   450
                     Width           =   3825
                  End
                  Begin VB.ListBox ListDeptAll 
                     Height          =   1620
                     ItemData        =   "FrmTransacRegistr.frx":17FF0
                     Left            =   4605
                     List            =   "FrmTransacRegistr.frx":17FF7
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   450
                     Width           =   3705
                  End
                  Begin VB.Label Label13 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õœœ «·«ﬁ”«„"
                     ForeColor       =   &H00800000&
                     Height          =   240
                     Left            =   3570
                     TabIndex        =   87
                     Top             =   120
                     Width           =   1470
                  End
                  Begin VB.Label Label10 
                     Alignment       =   2  'Center
                     Caption         =   ">"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   345
                     Left            =   4035
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   540
                     Width           =   495
                  End
                  Begin VB.Label Label9 
                     Alignment       =   2  'Center
                     Caption         =   ">>"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   360
                     Left            =   4035
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   870
                     Width           =   495
                  End
                  Begin VB.Label Label4 
                     Alignment       =   2  'Center
                     Caption         =   "<<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   360
                     Left            =   4035
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   1545
                     Width           =   495
                  End
                  Begin VB.Label Label3 
                     Alignment       =   2  'Center
                     Caption         =   "<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   345
                     Left            =   4035
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   1215
                     Width           =   495
                  End
               End
               Begin VB.Frame Frame3 
                  BackColor       =   &H00E2E9E9&
                  Height          =   2175
                  Left            =   0
                  TabIndex        =   72
                  Top             =   0
                  Width           =   9045
                  Begin VB.ListBox ListAllUsers 
                     Height          =   1620
                     ItemData        =   "FrmTransacRegistr.frx":18009
                     Left            =   4965
                     List            =   "FrmTransacRegistr.frx":18010
                     RightToLeft     =   -1  'True
                     TabIndex        =   74
                     Top             =   450
                     Width           =   3945
                  End
                  Begin VB.ListBox ListUserSelect 
                     Height          =   1620
                     ItemData        =   "FrmTransacRegistr.frx":18022
                     Left            =   120
                     List            =   "FrmTransacRegistr.frx":18029
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   450
                     Width           =   4185
                  End
                  Begin VB.Label Label5 
                     Alignment       =   2  'Center
                     Caption         =   "<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   345
                     Left            =   4395
                     RightToLeft     =   -1  'True
                     TabIndex        =   79
                     Top             =   1215
                     Width           =   495
                  End
                  Begin VB.Label Label6 
                     Alignment       =   2  'Center
                     Caption         =   "<<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   360
                     Left            =   4395
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   1545
                     Width           =   495
                  End
                  Begin VB.Label Label7 
                     Alignment       =   2  'Center
                     Caption         =   ">>"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   360
                     Left            =   4395
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   870
                     Width           =   495
                  End
                  Begin VB.Label Label8 
                     Alignment       =   2  'Center
                     Caption         =   ">"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   345
                     Left            =   4395
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
                     Top             =   540
                     Width           =   495
                  End
                  Begin VB.Label Label12 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õœœ «·„ÊŸ›Ì‰"
                     ForeColor       =   &H00800000&
                     Height          =   240
                     Left            =   3930
                     TabIndex        =   75
                     Top             =   120
                     Width           =   1470
                  End
               End
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   345
                  Left            =   240
                  TabIndex        =   89
                  Top             =   2280
                  Width           =   1710
                  _ExtentX        =   3016
                  _ExtentY        =   609
                  ButtonPositionImage=   1
                  Caption         =   " «ﬂÌœ «·«—”«·"
                  BackColor       =   -2147483635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ColorButton     =   -2147483635
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   7890
            Left            =   19590
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   45
            Width           =   18855
            _cx             =   33258
            _cy             =   13917
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   6930
               Left            =   150
               TabIndex        =   66
               Tag             =   "1"
               Top             =   0
               Width           =   18705
               _cx             =   32994
               _cy             =   12224
               Appearance      =   1
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
               BackColorFixed  =   14871017
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
               Rows            =   3
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmTransacRegistr.frx":1803D
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
               ExplorerBar     =   0
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
            Begin VB.Label Label1100 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   555
               Left            =   11940
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   9750
               Width           =   3690
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   375
               Left            =   7005
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   7230
               Width           =   3615
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic12 
            Height          =   7890
            Left            =   19890
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   45
            Width           =   18855
            _cx             =   33258
            _cy             =   13917
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic13 
               Height          =   1290
               Left            =   0
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   -165
               Width           =   18945
               _cx             =   33417
               _cy             =   2275
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
               Begin VB.TextBox TxtSummary2 
                  Alignment       =   1  'Right Justify
                  Height          =   900
                  Left            =   270
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   110
                  Top             =   270
                  Width           =   16785
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·Œ’ «·”‰œ"
                  Height          =   300
                  Index           =   19
                  Left            =   17145
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   465
                  Width           =   1680
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic14 
               Height          =   1155
               Left            =   0
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   2010
               Width           =   18945
               _cx             =   33417
               _cy             =   2037
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
               Begin VB.TextBox TxtProcedureReq2 
                  Alignment       =   1  'Right Justify
                  Height          =   900
                  Left            =   270
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   113
                  Top             =   135
                  Width           =   16755
               End
               Begin MSComCtl2.DTPicker DTPicker2 
                  Height          =   180
                  Left            =   465
                  TabIndex        =   114
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2235
                  _ExtentX        =   3942
                  _ExtentY        =   318
                  _Version        =   393216
                  Format          =   89391106
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTPicker3 
                  Height          =   180
                  Left            =   3930
                  TabIndex        =   115
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   318
                  _Version        =   393216
                  Format          =   89391106
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTPicker4 
                  Height          =   150
                  Left            =   6735
                  TabIndex        =   116
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   265
                  _Version        =   393216
                  Format          =   89391106
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã—«¡ «·„ÿ·Ê»"
                  Height          =   195
                  Index           =   26
                  Left            =   17145
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   435
                  Width           =   1590
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  «·œŒÊ·"
                  Height          =   90
                  Index           =   23
                  Left            =   8310
                  TabIndex        =   117
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   1200
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic16 
               Height          =   4770
               Left            =   0
               TabIndex        =   134
               TabStop         =   0   'False
               Top             =   3240
               Width           =   18855
               _cx             =   33258
               _cy             =   8414
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
               Begin VB.CommandButton Command3 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "≈—”«·"
                  Height          =   435
                  Left            =   16440
                  Style           =   1  'Graphical
                  TabIndex        =   137
                  Top             =   3480
                  Width           =   1725
               End
               Begin VB.CommandButton Command33 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "≈—”«·"
                  Height          =   585
                  Left            =   20070
                  Style           =   1  'Graphical
                  TabIndex        =   135
                  Top             =   165
                  Width           =   1875
               End
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   3000
                  Left            =   165
                  TabIndex        =   136
                  Top             =   300
                  Width           =   18555
                  _cx             =   32729
                  _cy             =   5292
                  Appearance      =   1
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
                  BackColorAlternate=   16777088
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483642
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   1
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
                  FormatString    =   $"FrmTransacRegistr.frx":18180
                  ScrollTrack     =   -1  'True
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
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   4770
               Left            =   -180
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   3060
               Width           =   19080
               Begin VB.CommandButton Command2 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "„Êﬁ› «·”‰œ"
                  Height          =   375
                  Left            =   15840
                  Style           =   1  'Graphical
                  TabIndex        =   122
                  Top             =   3120
                  Width           =   1695
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E2E9E9&
                  Height          =   2895
                  Left            =   240
                  TabIndex        =   100
                  Top             =   120
                  Width           =   8805
                  Begin VB.ListBox ListUserSelect2 
                     Height          =   2205
                     ItemData        =   "FrmTransacRegistr.frx":1830C
                     Left            =   120
                     List            =   "FrmTransacRegistr.frx":18313
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   450
                     Width           =   4065
                  End
                  Begin VB.ListBox ListAllUsers2 
                     Height          =   2205
                     ItemData        =   "FrmTransacRegistr.frx":18327
                     Left            =   4845
                     List            =   "FrmTransacRegistr.frx":1832E
                     RightToLeft     =   -1  'True
                     TabIndex        =   101
                     Top             =   450
                     Width           =   3825
                  End
                  Begin VB.Label Label23 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õœœ «·„ÊŸ›Ì‰"
                     ForeColor       =   &H00800000&
                     Height          =   240
                     Left            =   3570
                     TabIndex        =   107
                     Top             =   120
                     Width           =   2070
                  End
                  Begin VB.Label Label22 
                     Alignment       =   2  'Center
                     Caption         =   ">"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   345
                     Left            =   4275
                     RightToLeft     =   -1  'True
                     TabIndex        =   106
                     Top             =   780
                     Width           =   495
                  End
                  Begin VB.Label Label21 
                     Alignment       =   2  'Center
                     Caption         =   ">>"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   360
                     Left            =   4275
                     RightToLeft     =   -1  'True
                     TabIndex        =   105
                     Top             =   1110
                     Width           =   495
                  End
                  Begin VB.Label Label20 
                     Alignment       =   2  'Center
                     Caption         =   "<<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   360
                     Left            =   4275
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   1785
                     Width           =   495
                  End
                  Begin VB.Label Label19 
                     Alignment       =   2  'Center
                     Caption         =   "<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   345
                     Left            =   4275
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   1455
                     Width           =   495
                  End
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H00E2E9E9&
                  Height          =   2895
                  Left            =   9120
                  TabIndex        =   92
                  Top             =   120
                  Width           =   8565
                  Begin VB.ListBox ListDeptAll2 
                     Height          =   2205
                     ItemData        =   "FrmTransacRegistr.frx":18340
                     Left            =   4605
                     List            =   "FrmTransacRegistr.frx":18347
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   450
                     Width           =   3825
                  End
                  Begin VB.ListBox ListDeptSelect2 
                     Height          =   2205
                     ItemData        =   "FrmTransacRegistr.frx":18359
                     Left            =   120
                     List            =   "FrmTransacRegistr.frx":18360
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   450
                     Width           =   3825
                  End
                  Begin VB.Label Label18 
                     Alignment       =   2  'Center
                     Caption         =   "<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   345
                     Left            =   4035
                     RightToLeft     =   -1  'True
                     TabIndex        =   99
                     Top             =   1455
                     Width           =   495
                  End
                  Begin VB.Label Label17 
                     Alignment       =   2  'Center
                     Caption         =   "<<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   360
                     Left            =   4035
                     RightToLeft     =   -1  'True
                     TabIndex        =   98
                     Top             =   1785
                     Width           =   495
                  End
                  Begin VB.Label Label16 
                     Alignment       =   2  'Center
                     Caption         =   ">>"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   360
                     Left            =   4035
                     RightToLeft     =   -1  'True
                     TabIndex        =   97
                     Top             =   1110
                     Width           =   495
                  End
                  Begin VB.Label Label15 
                     Alignment       =   2  'Center
                     Caption         =   ">"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   345
                     Left            =   4035
                     RightToLeft     =   -1  'True
                     TabIndex        =   96
                     Top             =   780
                     Width           =   495
                  End
                  Begin VB.Label Label14 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õœœ «·«ﬁ”«„"
                     ForeColor       =   &H00800000&
                     Height          =   240
                     Left            =   3570
                     TabIndex        =   95
                     Top             =   120
                     Width           =   1470
                  End
               End
               Begin ImpulseButton.ISButton ISButton6 
                  Height          =   345
                  Left            =   2880
                  TabIndex        =   108
                  Top             =   3120
                  Width           =   1710
                  _ExtentX        =   3016
                  _ExtentY        =   609
                  ButtonPositionImage=   1
                  Caption         =   " «ﬂÌœ «·«—”«·"
                  BackColor       =   -2147483635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ColorButton     =   -2147483635
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
               Begin ImpulseButton.ISButton ISButton7 
                  Height          =   345
                  Left            =   720
                  TabIndex        =   120
                  Top             =   3120
                  Width           =   1710
                  _ExtentX        =   3016
                  _ExtentY        =   609
                  ButtonPositionImage=   1
                  Caption         =   "«—”«· ··Õ›Ÿ"
                  BackColor       =   -2147483635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ColorButton     =   -2147483635
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic15 
               Height          =   930
               Left            =   0
               TabIndex        =   138
               TabStop         =   0   'False
               Top             =   1080
               Width           =   18855
               _cx             =   33258
               _cy             =   1640
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
               Begin MSDataListLib.DataCombo DCboUserName2 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   139
                  Top             =   240
                  Width           =   16785
                  _ExtentX        =   29607
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„—”· „‰"
                  Height          =   330
                  Index           =   16
                  Left            =   17130
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   345
                  Width           =   1590
               End
            End
         End
      End
      Begin ImpulseButton.ISButton btnNew 
         Height          =   450
         Left            =   16440
         TabIndex        =   123
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   11100
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   794
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
         ButtonImage     =   "FrmTransacRegistr.frx":18375
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   450
         Left            =   13380
         TabIndex        =   124
         ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   11100
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   794
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
         ButtonImage     =   "FrmTransacRegistr.frx":1EBD7
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   450
         Left            =   14895
         TabIndex        =   125
         ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
         Top             =   11100
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   794
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
         ButtonImage     =   "FrmTransacRegistr.frx":1EF71
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   450
         Left            =   11925
         TabIndex        =   126
         ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
         Top             =   11100
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   794
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
         ButtonImage     =   "FrmTransacRegistr.frx":257D3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   450
         Left            =   10500
         TabIndex        =   127
         ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
         Top             =   11100
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   794
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
         ButtonImage     =   "FrmTransacRegistr.frx":25B6D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   450
         Left            =   855
         TabIndex        =   128
         ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
         Top             =   11100
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   794
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
         ButtonImage     =   "FrmTransacRegistr.frx":26107
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton5 
         Height          =   525
         Left            =   9180
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
         Top             =   11100
         Width           =   1185
         _ExtentX        =   2090
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
         ButtonImage     =   "FrmTransacRegistr.frx":264A1
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton8 
         Height          =   450
         Left            =   7890
         TabIndex        =   130
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   11100
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   794
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
         ButtonImage     =   "FrmTransacRegistr.frx":2CD03
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   450
         Left            =   6585
         TabIndex        =   131
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   11100
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   794
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "„—›ﬁ« "
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
         ButtonImage     =   "FrmTransacRegistr.frx":2D09D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Accredit 
         Height          =   465
         Left            =   4545
         TabIndex        =   132
         Top             =   11100
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   820
         ButtonPositionImage=   1
         Caption         =   "«—”«· ··«⁄ „«œ"
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorButton     =   -2147483635
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton ISButton3 
         Height          =   465
         Left            =   2640
         TabIndex        =   133
         Top             =   11100
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   820
         ButtonPositionImage=   1
         Caption         =   "«—”«· ··ﬁ”„"
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorButton     =   -2147483635
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
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
      Left            =   19200
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmTransacRegistr"
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
 Dim II As Long
 Public Ind As Integer

Private Sub Accredit_Click()
Dim BeginTrans As Boolean
 
    SendTopost Me.Name, "TblTransacRegistr", "ID", 0, val(Dcbranch.BoundText), val(TxtSerial1.Text), TxtSerial1.Text

 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
   fillapprovData
End Sub
Function FillMylist()
    Dim sql As String
    Dim Rs2 As ADODB.Recordset
    Dim i As Integer
    Set Rs2 = New ADODB.Recordset
    sql = " SELECT * from  TblEmpDepartments "
    Rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Me.ListAllUsers.Clear
    Me.ListDeptAll.Clear
    Me.ListDeptSelect.Clear
    Me.ListUserSelect.Clear
    Me.ListAllUsers2.Clear
    Me.ListDeptAll2.Clear
    Me.ListDeptSelect2.Clear
    Me.ListUserSelect2.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListDeptAll.AddItem IIf(IsNull(Rs2("DepartmentName").value), "", Rs2("DepartmentName").value)
                ListDeptAll2.AddItem IIf(IsNull(Rs2("DepartmentName").value), "", Rs2("DepartmentName").value)
            Else
                ListDeptAll.AddItem IIf(IsNull(Rs2("DepartmentNamee").value), "", Rs2("DepartmentNamee").value)
                ListDeptAll2.AddItem IIf(IsNull(Rs2("DepartmentNamee").value), "", Rs2("DepartmentNamee").value)
            End If
            ListDeptAll.ItemData(ListDeptAll.NewIndex) = IIf(IsNull(Rs2("DeparmentID").value), 0, Rs2("DeparmentID").value)
            ListDeptAll2.ItemData(ListDeptAll2.NewIndex) = IIf(IsNull(Rs2("DeparmentID").value), 0, Rs2("DeparmentID").value)
            Rs2.MoveNext
        Next i
    End If
    Rs2.Close
End Function

Private Sub Command1_Click()
Frame1.Visible = True
C1Elastic11.Visible = False
End Sub

Private Sub Command2_Click()
Frame4.Visible = False
C1Elastic16.Visible = True
FillGrid val(TxtSerial1.Text)
End Sub

Private Sub Command3_Click()
Frame4.Visible = True
C1Elastic16.Visible = False
End Sub

Private Sub Command5_Click()
Frame1.Visible = False
C1Elastic11.Visible = True
FillGrid val(TxtSerial1.Text)
End Sub

Private Sub DcbMHDID_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbMHDID.ListIndex) = 2 Then
ExitDate.value = DateAdd("d", val(TxtMHD.Text), EnterDate.value)
ElseIf val(DcbMHDID.ListIndex) = 1 Then
ExitDate.value = DateAdd("h", val(TxtMHD.Text), EnterDate.value)
ElseIf val(DcbMHDID.ListIndex) = 0 Then
ExitDate.value = DateAdd("n", val(TxtMHD.Text), EnterDate.value)
ElseIf val(DcbMHDID.ListIndex) = 3 Then
ExitDate.value = DateAdd("M", val(TxtMHD.Text), EnterDate.value)
End If
End If
End Sub
Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.TxtSerial1.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
Accredit.Enabled = False
Else
Accredit.Enabled = True
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
End If
 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.Rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.Rows = 1
    End If
RsDetails.Close
End Function
Private Sub DcbMHDID_Click()
DcbMHDID_Change
End Sub

Private Sub DcbTypTrans_Change()
DcbTypTrans_Click (0)
End Sub

Private Sub DcbTypTrans_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTypTrans.BoundText) > 0 Then
Rerive val(DcbTypTrans.BoundText)
End If
End If
End Sub
Sub Rerive(Optional ID As Double)
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Dim sql As String
sql = "Select * from  TblXXArchDocType where ID=" & ID & " "
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
DcbMHDID.ListIndex = IIf(IsNull(Rs2("TimeUnitID").value), -1, Rs2("TimeUnitID").value)
TxtMHD.Text = IIf(IsNull(Rs2("Time").value), "", Rs2("Time").value)
Else
TxtMHD.Text = ""
DcbMHDID.ListIndex = -1
End If
End Sub

Private Sub EnterDate_Change()
DcbMHDID_Change
End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub
Sub LoadMain()
    Dim conection As String
    Dim My_SQL As String
    C1Elastic5.Visible = True
    C1Tab1.TabVisible(0) = True
    C1Tab1.TabVisible(1) = True
    'C1Elastic17.Visible = True
    
    btnNew.Visible = True
    btnModify.Visible = True
    btnSave.Visible = True
    BtnUndo.Visible = True
    btnDelete.Visible = True
    ISButton5.Visible = True
    ISButton8.Visible = True
    ISButton2.Visible = True
    Accredit.Visible = True
    ISButton3.Visible = True
    btnCancel.Visible = True
        
    Frame1.Visible = False
    C1Elastic11.Visible = True
    C1Elastic2.Enabled = True
    C1Elastic13.Enabled = True
    C1Elastic15.Enabled = True

    conection = "select * from TblTransacRegistr order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    C1Elastic5.Visible = False
    C1Tab1.TabVisible(0) = False
    C1Tab1.TabVisible(1) = False
    C1Tab1.TabVisible(2) = False
    'C1Elastic17.Visible = False
    
    btnNew.Visible = False
    btnModify.Visible = False
    btnSave.Visible = False
    BtnUndo.Visible = False
    btnDelete.Visible = False
    ISButton5.Visible = False
    ISButton8.Visible = False
    ISButton2.Visible = False
    Accredit.Visible = False
    ISButton3.Visible = False
    btnCancel.Visible = False
    
    C1Elastic2.Enabled = False
    C1Elastic13.Enabled = False
    C1Elastic15.Enabled = False
    If SystemOptions.UserInterface = ArabicInterface Then
    With DcbImportExport
    .Clear
    .AddItem "’«œ—"
    .AddItem "Ê«—œ"
    End With
    With DcbMHDID
    .Clear
    .AddItem "œﬁÌﬁ…"
    .AddItem "”«⁄…"
    .AddItem "ÌÊ„"
    .AddItem "‘Â—"
    End With
    Else
        With DcbMHDID
    .Clear
    .AddItem "Minute"
    .AddItem "Hour"
    .AddItem "Day"
    .AddItem "Month"
    End With
    
    With DcbImportExport
    .Clear
    .AddItem "Inbox"
    .AddItem "Outbox"
  End With
    End If
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.ColComboList(FG.ColIndex("FlgTrans")) = "#1;   „ «·«—”«·|#2;  „ «·«” ·«„|#3;  „ «·Õ›Ÿ"
                VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("FlgTrans")) = "#1;   „ «·«—”«·|#2;  „ «·«” ·«„|#3;  „ «·Õ›Ÿ"
                ElseIf SystemOptions.UserInterface = EnglishInterface Then
               FG.ColComboList(FG.ColIndex("FlgTrans")) = "#1;Sent  |#2;Received |#3;Saved"
               VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("FlgTrans")) = "#1;Sent  |#2;Received |#3;Saved "
           End If
    If Ind = 0 Then
    C1Tab1.CurrTab = 0
     LoadMain
    Else
    C1Tab1.TabVisible(2) = True
    C1Tab1.CurrTab = 2
    End If
    FillMylist
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetUsers Me.DCboUserName2
    Dcombos.GetArchDocType DcbTypTrans
       If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If Ind = 0 Then
    BtnLast_Click
       If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
    End If
    ShowTip
    
    Me.TxtModFlg.Text = "R"
   Me.Refresh
ErrTrap:
End Sub


Public Sub FiLLRec()

  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
    RsSavRec.Fields("MHD").value = val(Me.TxtMHD.Text)
    RsSavRec.Fields("MHDID").value = val(Me.DcbMHDID.ListIndex)
    RsSavRec.Fields("BrnchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("ExitDate").value = ExitDate.value
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("RecordDateH").value = RecordDateH.value
    RsSavRec.Fields("barcode").value = Txtbarcode.Text
    RsSavRec.Fields("ImportExport").value = val(Me.DcbImportExport.ListIndex)
    RsSavRec.Fields("TypTrans").value = val(DcbTypTrans.BoundText)
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("NoImpExp").value = TxtNoImpExp.Text
    RsSavRec.Fields("ImpExpDate").value = ImpExpDate.value
    RsSavRec.Fields("ImpExpDateH").value = ImpExpDateH.value
    RsSavRec.Fields("Summary").value = TxtSummary.Text
    RsSavRec.Fields("EnterDate").value = EnterDate.value
    RsSavRec.Fields("ProcedureReq").value = TxtProcedureReq.Text
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("RecordTime").value = FormatDateTime(RecordTime.value, vbShortTime)
    RsSavRec.Fields("EnterTime").value = FormatDateTime(EnterTime.value, vbShortTime)
    RsSavRec.Fields("RequerTime").value = FormatDateTime(RequerTime.value, vbShortTime)
    RsSavRec.Fields("ExitTime").value = FormatDateTime(ExitTime.value, vbShortTime)
    RsSavRec.update
  ''///////////////////
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
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
                FiLLTXT
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
       fillapprovData
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub


' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
   Dim Rec As Date
   Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    TxtMHD.Text = IIf(IsNull(RsSavRec.Fields("MHD").value), "", RsSavRec.Fields("MHD").value)
    Me.DcbMHDID.ListIndex = IIf(IsNull(RsSavRec.Fields("MHDID").value), -1, RsSavRec.Fields("MHDID").value)
    ExitDate.value = IIf(IsNull(RsSavRec.Fields("ExitDate").value), Date, RsSavRec.Fields("ExitDate").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BrnchID").value), "", RsSavRec.Fields("BrnchID").value)
    Me.Txtbarcode.Text = IIf(IsNull(RsSavRec.Fields("barcode").value), "", RsSavRec.Fields("barcode").value)
    Me.DcbTypTrans.BoundText = IIf(IsNull(RsSavRec.Fields("TypTrans").value), "", RsSavRec.Fields("TypTrans").value)
    Me.DcbImportExport.ListIndex = IIf(IsNull(RsSavRec.Fields("ImportExport").value), -1, RsSavRec.Fields("ImportExport").value)
    Me.TxtNoImpExp.Text = IIf(IsNull(RsSavRec.Fields("NoImpExp").value), "", RsSavRec.Fields("NoImpExp").value) ': ProgressBar1.value = 90
    ImpExpDate.value = IIf(IsNull(RsSavRec.Fields("ImpExpDate").value), Date, RsSavRec.Fields("ImpExpDate").value)
    ImpExpDateH.value = IIf(IsNull(RsSavRec.Fields("ImpExpDateH").value), ToHijriDate(Date), RsSavRec.Fields("ImpExpDateH").value)   ': ProgressBar1.value = 10
    Me.TxtSummary.Text = IIf(IsNull(RsSavRec.Fields("Summary").value), "", RsSavRec.Fields("Summary").value)
    EnterDate.value = IIf(IsNull(RsSavRec.Fields("EnterDate").value), Date, RsSavRec.Fields("EnterDate").value)
    TxtProcedureReq.Text = IIf(IsNull(RsSavRec.Fields("ProcedureReq").value), "", RsSavRec.Fields("ProcedureReq").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
If Not IsNull(RsSavRec.Fields("RecordTime")) Then
Rec = FormatDateTime(RsSavRec.Fields("RecordTime"), vbShortTime)
RecordTime.value = Rec
End If
If Not IsNull(RsSavRec.Fields("EnterTime")) Then
Rec = FormatDateTime(RsSavRec.Fields("EnterTime"), vbShortTime)
EnterTime.value = Rec
End If
If Not IsNull(RsSavRec.Fields("RequerTime")) Then
Rec = FormatDateTime(RsSavRec.Fields("RequerTime"), vbShortTime)
RequerTime.value = Rec
End If
If Not IsNull(RsSavRec.Fields("ExitTime")) Then
Rec = FormatDateTime(RsSavRec.Fields("ExitTime"), vbShortTime)
ExitTime.value = Rec
End If
fillapprovData
FillGrid val(TxtSerial1.Text)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.Text = "" And val(Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            Dcbranch.SetFocus
            Exit Sub
     End If

If DcbImportExport.Text = "" Or val(DcbImportExport.ListIndex) = -1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·‰Ê⁄"
Else
MsgBox "Please Select  Type"
End If
DcbImportExport.SetFocus
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
    StrRecID = new_id("TblTransacRegistr", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Private Sub ImpExpDate_Change()
If Me.TxtModFlg.Text <> "R" Then
ImpExpDateH.value = ToHijriDate(ImpExpDate.value)
End If
End Sub




Private Sub ImpExpDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
ImpExpDate.value = ToGregorianDate(ImpExpDateH.value)
End If
End Sub

Private Sub ISButton2_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.Text, "10082017111"
End Sub

Private Sub ISButton3_Click()
Frame1.Visible = True
C1Elastic11.Visible = False
End Sub

Private Sub ISButton4_Click()
If val(TxtSerial1.Text) <> 0 Then
SaveAprove Me.TxtProcedureReq.Text, val(TxtSerial1.Text)
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·«—”«·"
Else
MsgBox "Sent Succesfully"
End If
End If
End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub ISButton6_Click()
If val(TxtSerial1.Text) <> 0 Then
SaveAprove2 Me.TxtProcedureReq2.Text, val(TxtSerial1.Text)
Cn.Execute "Update TblTransacRegistrDet set FlgTrans=1 where ID=" & val(Me.TxtNoID.Text) & ""
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·«—”«·"
Else
MsgBox "Sent Succesfully"
End If
End If
End Sub

Private Sub ISButton7_Click()
Dim sql As String
Dim ArchID As Double
Dim DepID As Double
Dim RoomID As Double
Dim BoxID As Double
Dim ShelfID As Double
Dim TimeID As Double
Dim TimeUnitID As Double

If val(TxtSerial1.Text) <> 0 Then
Cn.Execute "Update TblTransacRegistrDet set FlgTrans=2 where ID=" & val(Me.TxtNoID.Text) & ""
RetriveOper val(DcbTypTrans.BoundText), DepID, ArchID, RoomID, BoxID, ShelfID, TimeID, TimeUnitID
sql = " Update TblTransacRegistrDet set DepID=" & DepID & ",ArchID=" & ArchID & ",RoomID=" & RoomID & ",BoxID=" & BoxID & ",ShelfID=" & ShelfID & ",Time=" & TimeID & ",TimeUnitID=" & TimeUnitID & ""
sql = sql & " Where ID=" & val(Me.TxtNoID.Text) & ""
Cn.Execute sql
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·Õ›Ÿ"
Else
MsgBox "Saved Succesfully"
End If
End If
End Sub
Sub RetriveOper(Optional ID As Double, Optional ByRef DepID As Double, Optional ByRef ArchID As Double, Optional ByRef RoomID As Double, Optional ByRef BoxID As Double, Optional ByRef ShelfID As Double, Optional ByRef TimeID As Double, Optional ByRef TimeUnitID As Double)
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = " SELECT     *"
sql = sql & " FROM         dbo.TblXXArchDocType where ID =" & ID & ""
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
DepID = IIf(IsNull(Rs2("DepID").value), 0, Rs2("DepID").value)
ArchID = IIf(IsNull(Rs2("ArchID").value), 0, Rs2("ArchID").value)
RoomID = IIf(IsNull(Rs2("RoomID").value), 0, Rs2("RoomID").value)
BoxID = IIf(IsNull(Rs2("BoxID").value), 0, Rs2("BoxID").value)
ShelfID = IIf(IsNull(Rs2("ShelfID").value), 0, Rs2("ShelfID").value)
TimeID = IIf(IsNull(Rs2("Time").value), 0, Rs2("Time").value)
TimeUnitID = IIf(IsNull(Rs2("TimeUnitID").value), -1, Rs2("TimeUnitID").value)
Else
ArchID = 0
DepID = 0
RoomID = 0
BoxID = 0
ShelfID = 0
TimeID = 0
TimeUnitID = -1
End If
End Sub

Private Sub ISButton8_Click()
    FrmSearchinvestment.inde = 31
    FrmSearchinvestment.show
End Sub

Private Sub Label10_Click()
 If Me.ListDeptAll.ListIndex > -1 Then
    Me.ListDeptSelect.AddItem ListDeptAll.List(ListDeptAll.ListIndex)
    ListDeptSelect.ItemData(ListDeptSelect.NewIndex) = ListDeptAll.ItemData(ListDeptAll.ListIndex)
End If
FillMylist2
End Sub

Private Sub Label15_Click()
 If Me.ListDeptAll2.ListIndex > -1 Then
    Me.ListDeptSelect2.AddItem ListDeptAll2.List(ListDeptAll2.ListIndex)
    ListDeptSelect2.ItemData(ListDeptSelect2.NewIndex) = ListDeptAll2.ItemData(ListDeptAll2.ListIndex)
End If
FillMylist3
End Sub

Private Sub Label16_Click()
    Dim i As Integer
    Me.ListDeptSelect2.Clear
    For i = 0 To Me.ListDeptAll2.ListCount - 1
        Me.ListDeptSelect2.AddItem ListDeptAll2.List(i)
        ListDeptSelect2.ItemData(i) = ListDeptAll2.ItemData(i)
    Next i
  
   FillMylist3
End Sub

Private Sub Label17_Click()
Me.ListDeptSelect2.Clear
Me.ListAllUsers2.Clear
Me.ListUserSelect2.Clear
End Sub

Private Sub Label18_Click()
If Me.ListDeptSelect2.ListIndex > -1 Then
Me.ListDeptSelect2.RemoveItem (ListDeptSelect2.ListIndex)
End If
FillMylist3
End Sub

Private Sub Label19_Click()
If Me.ListUserSelect2.ListIndex > -1 Then
ListUserSelect2.RemoveItem (ListUserSelect2.ListIndex)
End If
End Sub

Private Sub Label20_Click()
Me.ListUserSelect2.Clear
End Sub

Private Sub Label21_Click()
    Dim i As Integer
    Me.ListUserSelect2.Clear
    For i = 0 To Me.ListAllUsers2.ListCount - 1
        Me.ListUserSelect2.AddItem ListAllUsers2.List(i)
        ListUserSelect2.ItemData(i) = ListAllUsers2.ItemData(i)
    Next i
End Sub

Private Sub Label22_Click()
 If Me.ListAllUsers2.ListIndex > -1 Then
    Me.ListUserSelect2.AddItem ListAllUsers2.List(ListAllUsers2.ListIndex)
    ListUserSelect2.ItemData(ListUserSelect2.NewIndex) = ListAllUsers2.ItemData(ListAllUsers2.ListIndex)
End If
End Sub

Private Sub Label3_Click()
If Me.ListDeptSelect.ListIndex > -1 Then
Me.ListDeptSelect.RemoveItem (ListDeptSelect.ListIndex)
End If
FillMylist2
End Sub

Private Sub Label4_Click()
Me.ListDeptSelect.Clear
Me.ListAllUsers.Clear
Me.ListUserSelect.Clear
End Sub

Private Sub Label5_Click()
If Me.ListUserSelect.ListIndex > -1 Then
ListUserSelect.RemoveItem (ListUserSelect.ListIndex)
End If
End Sub

Private Sub Label6_Click()
Me.ListUserSelect.Clear
End Sub

Private Sub Label7_Click()
    Dim i As Integer
    Me.ListUserSelect.Clear
    For i = 0 To Me.ListAllUsers.ListCount - 1
        Me.ListUserSelect.AddItem ListAllUsers.List(i)
        ListUserSelect.ItemData(i) = ListAllUsers.ItemData(i)
    Next i
End Sub

Private Sub Label8_Click()
 If Me.ListAllUsers.ListIndex > -1 Then
    Me.ListUserSelect.AddItem ListAllUsers.List(ListAllUsers.ListIndex)
    ListUserSelect.ItemData(ListUserSelect.NewIndex) = ListAllUsers.ItemData(ListAllUsers.ListIndex)
End If
End Sub

Private Sub Label9_Click()
    Dim i As Integer
    Me.ListDeptSelect.Clear
    For i = 0 To Me.ListDeptAll.ListCount - 1
        Me.ListDeptSelect.AddItem ListDeptAll.List(i)
        ListDeptSelect.ItemData(i) = ListDeptAll.ItemData(i)
    Next i
  
   FillMylist2
End Sub
Function CheckUserISExist(Optional TransRegID As Double, Optional ToUser As Double) As Boolean
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = "Select * from TblTransacRegistrDet where TransRegID=" & TransRegID & " and ToUser=" & ToUser & ""
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
CheckUserISExist = True
Else
CheckUserISExist = False
End If
End Function
Sub FillGrid(Optional TransRegID As Double)
If TransRegID = 0 Then Exit Sub
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
sql = " SELECT     dbo.TblTransacRegistrDet.ID, dbo.TblTransacRegistrDet.TransRegID, dbo.TblTransacRegistr.RecordDate, dbo.TblTransacRegistr.RecordDateH, "
sql = sql & "                       dbo.TblTransacRegistr.RecordTime, dbo.TblTransacRegistr.BrnchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
sql = sql & "                       dbo.TblTransacRegistr.barcode, dbo.TblTransacRegistr.ImportExport, dbo.TblTransacRegistr.NoImpExp, dbo.TblTransacRegistr.ImpExpDate,"
sql = sql & "                       dbo.TblTransacRegistr.ImpExpDateH, dbo.TblTransacRegistr.Summary, dbo.TblTransacRegistr.EnterDate, dbo.TblTransacRegistr.Remarks,"
sql = sql & "                       dbo.TblTransacRegistr.MHD, dbo.TblTransacRegistr.MHDID, dbo.TblTransacRegistr.ExitDate, dbo.TblTransacRegistr.TypTrans, dbo.TblXXArchDocType.Name,"
sql = sql & "                       dbo.TblXXArchDocType.Namee, dbo.TblXXArchDocType.Code, dbo.TblTransacRegistrDet.FromUser, dbo.TblUsers.UserName, dbo.TblTransacRegistrDet.ToUser,"
sql = sql & "                       TblUsers_1.UserName AS ToUserName, dbo.TblTransacRegistrDet.FlgTrans, dbo.TblTransacRegistrDet.RecDate, dbo.TblTransacRegistrDet.ProcedureReq"
sql = sql & "  FROM         dbo.TblUsers TblUsers_1 RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblTransacRegistrDet ON TblUsers_1.UserID = dbo.TblTransacRegistrDet.ToUser LEFT OUTER JOIN"
sql = sql & "                       dbo.TblUsers ON dbo.TblTransacRegistrDet.FromUser = dbo.TblUsers.UserID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblXXArchDocType RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblTransacRegistr ON dbo.TblXXArchDocType.ID = dbo.TblTransacRegistr.TypTrans LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.TblTransacRegistr.BrnchID = dbo.TblBranchesData.branch_id ON"
sql = sql & "                       dbo.TblTransacRegistrDet.TransRegID = dbo.TblTransacRegistr.ID"
sql = sql & "   Where (dbo.TblTransacRegistrDet.TransRegID = " & TransRegID & ")"
Rs3.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 FG.Clear flexClearScrollable, flexClearEverything
           FG.Rows = 1
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
With FG
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("FlgTrans")) = IIf(IsNull(Rs3("FlgTrans").value), 0, Rs3("FlgTrans").value) + 1
.TextMatrix(i, .ColIndex("TransRegID")) = IIf(IsNull(Rs3("TransRegID").value), "", Rs3("TransRegID").value)
.TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs3("RecordDate").value), "", Rs3("RecordDate").value)
.TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(Rs3("RecordDateH").value), "", Rs3("RecordDateH").value)
.TextMatrix(i, .ColIndex("ProcedureReq")) = IIf(IsNull(Rs3("ProcedureReq").value), "", Rs3("ProcedureReq").value)
.TextMatrix(i, .ColIndex("Summary")) = IIf(IsNull(Rs3("Summary").value), "", Rs3("Summary").value)
.TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(Rs3("UserName").value), "", Rs3("UserName").value)
.TextMatrix(i, .ColIndex("ToUserName")) = IIf(IsNull(Rs3("ToUserName").value), "", Rs3("ToUserName").value)
.TextMatrix(i, .ColIndex("RecDate")) = IIf(IsNull(Rs3("RecDate").value), "", Rs3("RecDate").value)
Rs3.MoveNext
Next i
End With
End If
Set Rs3 = New ADODB.Recordset

sql = " SELECT     dbo.TblTransacRegistrDet.ID, dbo.TblTransacRegistrDet.TransRegID, dbo.TblTransacRegistr.RecordDate, dbo.TblTransacRegistr.RecordDateH, "
sql = sql & "                       dbo.TblTransacRegistr.RecordTime, dbo.TblTransacRegistr.BrnchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
sql = sql & "                       dbo.TblTransacRegistr.barcode, dbo.TblTransacRegistr.ImportExport, dbo.TblTransacRegistr.NoImpExp, dbo.TblTransacRegistr.ImpExpDate,"
sql = sql & "                       dbo.TblTransacRegistr.ImpExpDateH, dbo.TblTransacRegistr.Summary, dbo.TblTransacRegistr.EnterDate, dbo.TblTransacRegistr.Remarks,"
sql = sql & "                       dbo.TblTransacRegistr.MHD, dbo.TblTransacRegistr.MHDID, dbo.TblTransacRegistr.ExitDate, dbo.TblTransacRegistr.TypTrans, dbo.TblXXArchDocType.Name,"
sql = sql & "                       dbo.TblXXArchDocType.Namee, dbo.TblXXArchDocType.Code, dbo.TblTransacRegistrDet.FromUser, dbo.TblUsers.UserName, dbo.TblTransacRegistrDet.ToUser,"
sql = sql & "                       TblUsers_1.UserName AS ToUserName, dbo.TblTransacRegistrDet.FlgTrans, dbo.TblTransacRegistrDet.RecDate, dbo.TblTransacRegistrDet.ProcedureReq"
sql = sql & "  FROM         dbo.TblUsers TblUsers_1 RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblTransacRegistrDet ON TblUsers_1.UserID = dbo.TblTransacRegistrDet.ToUser LEFT OUTER JOIN"
sql = sql & "                       dbo.TblUsers ON dbo.TblTransacRegistrDet.FromUser = dbo.TblUsers.UserID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblXXArchDocType RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblTransacRegistr ON dbo.TblXXArchDocType.ID = dbo.TblTransacRegistr.TypTrans LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.TblTransacRegistr.BrnchID = dbo.TblBranchesData.branch_id ON"
sql = sql & "                       dbo.TblTransacRegistrDet.TransRegID = dbo.TblTransacRegistr.ID"
sql = sql & "   Where (dbo.TblTransacRegistrDet.TransRegID = " & TransRegID & ")"
Rs3.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
           VSFlexGrid1.Rows = 1
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
With VSFlexGrid1
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("FlgTrans")) = IIf(IsNull(Rs3("FlgTrans").value), 0, Rs3("FlgTrans").value) + 1
.TextMatrix(i, .ColIndex("TransRegID")) = IIf(IsNull(Rs3("TransRegID").value), "", Rs3("TransRegID").value)
.TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs3("RecordDate").value), "", Rs3("RecordDate").value)
.TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(Rs3("RecordDateH").value), "", Rs3("RecordDateH").value)
.TextMatrix(i, .ColIndex("ProcedureReq")) = IIf(IsNull(Rs3("ProcedureReq").value), "", Rs3("ProcedureReq").value)
.TextMatrix(i, .ColIndex("Summary")) = IIf(IsNull(Rs3("Summary").value), "", Rs3("Summary").value)
.TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(Rs3("UserName").value), "", Rs3("UserName").value)
.TextMatrix(i, .ColIndex("ToUserName")) = IIf(IsNull(Rs3("ToUserName").value), "", Rs3("ToUserName").value)
.TextMatrix(i, .ColIndex("RecDate")) = IIf(IsNull(Rs3("RecDate").value), "", Rs3("RecDate").value)
Rs3.MoveNext
Next i
End With
End If
End Sub
Sub SaveAprove(Optional ProcedureReq As String, Optional TransRegID As Double)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
sql = "select * from TblTransacRegistrDet where 1=-1"
Rs3.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
For i = 0 To Me.ListUserSelect.ListCount - 1
If CheckUserISExist(TransRegID, ListUserSelect.ItemData(i)) = False Then
Rs3.AddNew
Rs3("RecDate") = Now
Rs3("FromUser") = user_id
Rs3("ProcedureReq") = ProcedureReq
Rs3("TransRegID") = TransRegID
Rs3("ToUser") = ListUserSelect.ItemData(i)
Rs3.update
End If
Next i
End Sub

Sub SaveAprove2(Optional ProcedureReq As String, Optional TransRegID As Double)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
sql = "select * from TblTransacRegistrDet where 1=-1"
Rs3.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
For i = 0 To Me.ListUserSelect2.ListCount - 1
If CheckUserISExist(TransRegID, ListUserSelect2.ItemData(i)) = False Then
Rs3.AddNew
Rs3("RecDate") = Now
Rs3("FromUser") = user_id
Rs3("ProcedureReq") = ProcedureReq
Rs3("TransRegID") = TransRegID
Rs3("ToUser") = ListUserSelect2.ItemData(i)
Rs3.update
End If
Next i
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
XPDtbTrans.value = ToGregorianDate(RecordDateH.value)
End If
End Sub

Private Sub TxtMHD_Change()
DcbMHDID_Change
End Sub

Private Sub TxtMHD_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtMHD.Text, 1)
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
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    'If RsSavRec.EditMode <> adEditNone Then
    '    RsSavRec.CancelUpdate
    '    BtnUndo_Click
    'End If
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
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim x As Integer
    Dim i As Integer
    Dim ID As Double
If ChectRecor() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› ÌÊÃœ ⁄·ÌÂ Õ—ﬂ«  «—”«·"
Else
MsgBox "You Can Not Delete"
End If
Exit Sub
End If
    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If x = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
      Dim StrSQL As String
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                                          RsSavRec.delete
                 If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If

     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
               LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
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
           'Cn.Errors.Clear
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
    If SystemOptions.UserInterface = ArabicInterface Then
        If Ind = 0 Then
            Label1(2).Caption = " ”ÃÌ· «·„⁄«„·…"
        Else
            Label1(2).Caption = "≈Ã—«¡«  «·„⁄«„·…"
        End If
    Else
        Label1(2).Caption = Me.Caption
    End If
        Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    XPDtbTrans.Enabled = True
        
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
     XPDtbTrans.Enabled = False
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
    XPDtbTrans.Enabled = True
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
Function ChectRecor() As Boolean
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = "Select * from TblTransacRegistrDet where TransRegID =" & val(TxtSerial1.Text) & ""
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
ChectRecor = True
Else
ChectRecor = False
End If
End Function
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
If ChectRecor() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ÌÊÃœ ⁄·ÌÂ Õ—ﬂ«  «—”«·"
Else
MsgBox "You Can Not Edit"
End If
Exit Sub
End If
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
        Me.Dcbranch.SetFocus
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
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
    XPDtbTrans.value = Date
    EnterDate.value = Now 'Format(Now, "dd/MM/yyyy, hh:mm:tt")
    EnterTime.value = Time
    ImpExpDate.value = Date
      GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.Rows = 1
            Accredit.Caption = ""
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
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
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

Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  MySQL = " SELECT     dbo.TblTransacRegistr.ID, dbo.TblTransacRegistr.BrnchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
  MySQL = MySQL & "                     dbo.TblTransacRegistr.RecordDate, dbo.TblTransacRegistr.RecordDateH, dbo.TblTransacRegistr.RecordTime, dbo.TblTransacRegistr.barcode,"
  MySQL = MySQL & "                     dbo.TblTransacRegistr.NoImpExp, dbo.TblTransacRegistr.ImportExport, dbo.TblTransacRegistr.ImpExpDate, dbo.TblTransacRegistr.ImpExpDateH,"
  MySQL = MySQL & "                     dbo.TblTransacRegistr.Summary, dbo.TblTransacRegistr.EnterDate, dbo.TblTransacRegistr.EnterTime, dbo.TblTransacRegistr.RequerTime,"
  MySQL = MySQL & "                     dbo.TblTransacRegistr.ExitTime, dbo.TblTransacRegistr.ProcedureReq, dbo.TblTransacRegistr.Remarks, dbo.TblTransacRegistr.MHD, dbo.TblTransacRegistr.MHDID,"
  MySQL = MySQL & "                     dbo.TblTransacRegistr.ExitDate, dbo.TblTransacRegistr.Posted, dbo.TblTransacRegistr.PostedDate, dbo.TblTransacRegistr.Approved,"
  MySQL = MySQL & "                     dbo.TblTransacRegistr.TypTrans , dbo.TblXXArchDocType.code, dbo.TblXXArchDocType.Name, dbo.TblXXArchDocType.NameE"
  MySQL = MySQL & "     FROM         dbo.TblTransacRegistr LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblXXArchDocType ON dbo.TblTransacRegistr.TypTrans = dbo.TblXXArchDocType.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblTransacRegistr.BrnchID = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & "       Where (dbo.TblTransacRegistr.ID =" & val(TxtSerial1.Text) & ")"
 
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransacRegistr.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransacRegistrE.rpt"
        End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub ChangeLang()
On Error GoTo ErrTrap

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Accredit.Caption = "Send For Approval"
    C1Tab1.Caption = "Data|Approval"
    With GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With
    
    lbl(3).Caption = "Summary"
    If Ind = 0 Then
       Me.Caption = "Documents Entry"
    Else
       Me.Caption = "Documents Procedures"
    End If
    
    Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "ID"
    Me.lbl(2).Caption = "Date"
    lbl(22).Caption = "Time"
    lbl(9).Caption = "Barcode"
    lbl(7).Caption = "Branch"
    lbl(11).Caption = "Type"
    lbl(1).Caption = "No. "
    lbl(0).Caption = "Date"
    lbl(12).Caption = "Transaction Type"
    lbl(6).Caption = "Date of Entry"
    lbl(5).Caption = "Time of Entry"
    lbl(10).Caption = "Time Req."
    lbl(13).Caption = "Time To Go Out"
    lbl(14).Caption = "Procedure Req."
    lbl(15).Caption = "Remarks"
    ISButton2.Caption = "Attachments"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
   
    With FG
    
    .TextMatrix(0, .ColIndex("Serial")) = "I"
        .TextMatrix(0, .ColIndex("TransRegID")) = "Doc No."
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("RecordDateH")) = "Hijri Date"
        .TextMatrix(0, .ColIndex("RecDate")) = "Send Date"
        .TextMatrix(0, .ColIndex("Summary")) = "Summary"
        .TextMatrix(0, .ColIndex("ProcedureReq")) = "Procedure"
        .TextMatrix(0, .ColIndex("UserName")) = "Sent From"
        .TextMatrix(0, .ColIndex("ToUserName")) = "Sent To"
        .TextMatrix(0, .ColIndex("FlgTrans")) = "Status"
    End With
    
    Command1.Caption = "Send"
    ISButton3.Caption = "Send to department"
    Label13.Caption = "Select departments"
    Label12.Caption = "Select Employees"
    Command5.Caption = "Document status"
    ISButton4.Caption = "Confirm Send"
    
    Label11.Caption = "Waiting for approval by "
    
    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("TransRegID")) = "Doc No."
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("RecordDateH")) = "Hijri Date"
        .TextMatrix(0, .ColIndex("RecDate")) = "Send Date"
        .TextMatrix(0, .ColIndex("Summary")) = "Summary"
        .TextMatrix(0, .ColIndex("ProcedureReq")) = "Procedure"
        .TextMatrix(0, .ColIndex("UserName")) = "Sent From"
        .TextMatrix(0, .ColIndex("ToUserName")) = "Sent To"
        .TextMatrix(0, .ColIndex("FlgTrans")) = "Status"
    End With
    
    lbl(19).Caption = "Document Summary"
    lbl(16).Caption = "Sent From"
    lbl(26).Caption = "Required procedure"
    Label14.Caption = "Select departments"
    Label23.Caption = "Select Employees"
    
    Command2.Caption = "Document status"
    ISButton6.Caption = "Confirm Send"
    ISButton7.Caption = "Save"
    Command3.Caption = "Send"
    
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblTransacRegistr"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
Function FillMylist3()
    Dim sql As String
    Dim Rs2 As ADODB.Recordset
    Dim i As Integer
    Dim ActivID As String
    ActivID = "0"
    For i = 0 To Me.ListDeptSelect2.ListCount - 1
    ActivID = ActivID & "," & Me.ListDeptSelect2.ItemData(i)
    Next i
    Me.ListAllUsers2.Clear
    Me.ListUserSelect2.Clear
    If ActivID = "0" Then Exit Function
    Set Rs2 = New ADODB.Recordset
    sql = " SELECT     dbo.TblUsers.UserName, dbo.TblUsers.UserID"
    sql = sql & "  FROM         dbo.TblUsers LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
    sql = sql & "   WHERE     dbo.TblEmployee.DepartmentID  in(" & ActivID & ") "
    Rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
            Me.ListAllUsers2.AddItem IIf(IsNull(Rs2("UserName").value), "", Rs2("UserName").value)
            ListAllUsers2.ItemData(ListAllUsers2.NewIndex) = IIf(IsNull(Rs2("UserID").value), 0, Rs2("UserID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
End Function

Function FillMylist2()
    Dim sql As String
    Dim Rs2 As ADODB.Recordset
    Dim i As Integer
    Dim ActivID As String
    ActivID = "0"
    For i = 0 To Me.ListDeptSelect.ListCount - 1
    ActivID = ActivID & "," & Me.ListDeptSelect.ItemData(i)
    Next i
    Me.ListAllUsers.Clear
    Me.ListUserSelect.Clear
    If ActivID = "0" Then Exit Function
    Set Rs2 = New ADODB.Recordset
    sql = " SELECT     dbo.TblUsers.UserName, dbo.TblUsers.UserID"
    sql = sql & "  FROM         dbo.TblUsers LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
    sql = sql & "   WHERE     dbo.TblEmployee.DepartmentID  in(" & ActivID & ") "
    Rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
            Me.ListAllUsers.AddItem IIf(IsNull(Rs2("UserName").value), "", Rs2("UserName").value)
            ListAllUsers.ItemData(ListAllUsers.NewIndex) = IIf(IsNull(Rs2("UserID").value), 0, Rs2("UserID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
End Function
Sub Retrive2(Optional ID As Double)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim RecDate As Date
sql = " SELECT     dbo.TblTransacRegistr.BrnchID, dbo.TblTransacRegistr.RecordDate, dbo.TblTransacRegistr.RecordDateH, dbo.TblTransacRegistr.RecordTime,"
sql = sql & "                       dbo.TblTransacRegistr.barcode, dbo.TblTransacRegistr.Summary, dbo.TblTransacRegistrDet.TransRegID, dbo.TblTransacRegistrDet.FromUser,"
sql = sql & "                      dbo.TblTransacRegistrDet.ToUser , dbo.TblTransacRegistrDet.ID,dbo.TblTransacRegistr.TypTrans"
sql = sql & " FROM         dbo.TblTransacRegistr RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblTransacRegistrDet ON dbo.TblTransacRegistr.ID = dbo.TblTransacRegistrDet.TransRegID"
sql = sql & " Where (dbo.TblTransacRegistrDet.ID = " & ID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
TxtSerial1.Text = IIf(IsNull(Rs3("TransRegID").value), 0, Rs3("TransRegID").value)
XPDtbTrans.value = IIf(IsNull(Rs3("RecordDate").value), Date, Rs3("RecordDate").value)
RecordDateH.value = IIf(IsNull(Rs3("RecordDateH").value), ToHijriDate(Date), Rs3("RecordDateH").value)
If Not IsNull(Rs3("RecordTime").value) Then
RecDate = FormatDateTime(Rs3("RecordTime").value, vbShortTime)
RecordTime.value = RecDate
End If
Txtbarcode.Text = IIf(IsNull(Rs3("barcode").value), "", Rs3("barcode").value)
Dcbranch.BoundText = IIf(IsNull(Rs3("BrnchID").value), "", Rs3("BrnchID").value)
Me.DcbTypTrans.BoundText = IIf(IsNull(Rs3("TypTrans").value), "", Rs3("TypTrans").value)
Me.DCboUserName2.BoundText = IIf(IsNull(Rs3("FromUser").value), "", Rs3("FromUser").value)
TxtSummary2.Text = IIf(IsNull(Rs3("Summary").value), "", Rs3("Summary").value)
TxtNoID.Text = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
End If
End Sub


Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
RecordDateH.value = ToHijriDate(XPDtbTrans.value)
End If
End Sub
