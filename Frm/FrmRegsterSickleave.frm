VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRegsterSickleave 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15120
   Icon            =   "FrmRegsterSickleave.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   15120
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
      TabIndex        =   16
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmRegsterSickleave.frx":6852
      Left            =   15480
      List            =   "FrmRegsterSickleave.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   17
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
      TabIndex        =   18
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
            Picture         =   "FrmRegsterSickleave.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegsterSickleave.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegsterSickleave.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegsterSickleave.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegsterSickleave.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegsterSickleave.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegsterSickleave.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegsterSickleave.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   19
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
      ButtonImage     =   "FrmRegsterSickleave.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   21
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
      ButtonImage     =   "FrmRegsterSickleave.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   22
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
      ButtonImage     =   "FrmRegsterSickleave.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   9990
      Left            =   0
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   15120
      _cx             =   26670
      _cy             =   17621
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
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1620
         Left            =   21450
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   18870
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   26
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   25
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1620
         Index           =   18
         Left            =   0
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   0
         Width           =   15075
         _cx             =   26591
         _cy             =   2858
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
            Height          =   705
            Left            =   150
            TabIndex        =   28
            Top             =   465
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmRegsterSickleave.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   705
            Left            =   750
            TabIndex        =   29
            Top             =   465
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmRegsterSickleave.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   705
            Left            =   1470
            TabIndex        =   30
            Top             =   465
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmRegsterSickleave.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   705
            Left            =   2145
            TabIndex        =   31
            Top             =   465
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmRegsterSickleave.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   1365
            Left            =   13920
            Picture         =   "FrmRegsterSickleave.frx":16A11
            Stretch         =   -1  'True
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ”ÃÌ· «·«Ã«“«  «·„—÷Ì…"
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
            Height          =   825
            Index           =   2
            Left            =   6315
            TabIndex        =   32
            Top             =   465
            Width           =   5340
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   1395
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1440
         Width           =   15075
         _cx             =   26591
         _cy             =   2461
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
            Height          =   750
            Left            =   11205
            TabIndex        =   0
            Top             =   255
            Width           =   2250
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   750
            Left            =   5340
            TabIndex        =   2
            Top             =   255
            Width           =   2310
            _extentx        =   2381
            _extenty        =   873
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   750
            Left            =   7770
            TabIndex        =   1
            Top             =   255
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   1323
            _Version        =   393216
            Format          =   107479041
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   195
            TabIndex        =   3
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   255
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   675
            Index           =   11
            Left            =   4470
            TabIndex        =   49
            Top             =   255
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   675
            Index           =   25
            Left            =   9600
            TabIndex        =   48
            Top             =   255
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   750
            Index           =   4
            Left            =   13545
            TabIndex        =   34
            Top             =   255
            Width           =   1410
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1485
         Left            =   0
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   8505
         Width           =   15120
         _cx             =   26670
         _cy             =   2619
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
         Align           =   2
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
            Height          =   315
            Left            =   13500
            TabIndex        =   36
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   855
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
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
            ButtonImage     =   "FrmRegsterSickleave.frx":17E16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   315
            Left            =   11655
            TabIndex        =   37
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   855
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
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
            ButtonImage     =   "FrmRegsterSickleave.frx":1E678
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   315
            Left            =   9615
            TabIndex        =   38
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   855
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
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
            ButtonImage     =   "FrmRegsterSickleave.frx":24EDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   315
            Left            =   6990
            TabIndex        =   39
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   855
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
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
            ButtonImage     =   "FrmRegsterSickleave.frx":25274
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   315
            Left            =   5055
            TabIndex        =   40
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   855
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
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
            ButtonImage     =   "FrmRegsterSickleave.frx":2560E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   315
            Left            =   480
            TabIndex        =   41
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   855
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
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
            ButtonImage     =   "FrmRegsterSickleave.frx":25BA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8175
            TabIndex        =   42
            Top             =   90
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ButPrient 
            Height          =   750
            Left            =   2565
            TabIndex        =   66
            Top             =   660
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   1323
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            BackStyle       =   0
            ButtonImage     =   "FrmRegsterSickleave.frx":25F42
            ColorButton     =   14871017
            DisplayPersistentHover=   0   'False
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   450
            Index           =   14
            Left            =   13020
            TabIndex        =   47
            Top             =   90
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   0
            Left            =   5190
            TabIndex        =   46
            Top             =   315
            Width           =   2010
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   1
            Left            =   1710
            TabIndex        =   45
            Top             =   315
            Width           =   1860
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   3825
            TabIndex        =   44
            Top             =   315
            Width           =   1230
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   510
            TabIndex        =   43
            Top             =   315
            Width           =   960
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   5340
         Left            =   0
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3045
         Width           =   15075
         _cx             =   26591
         _cy             =   9419
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
         Begin VB.TextBox txtTotalSickDays 
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
            Height          =   375
            Left            =   5160
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   2220
            Width           =   1215
         End
         Begin VB.TextBox txtCurrentSickDays 
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
            Height          =   375
            Left            =   8100
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   2220
            Width           =   1215
         End
         Begin VB.TextBox txtPrevSickDays 
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
            Height          =   375
            Left            =   11370
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   2250
            Width           =   1215
         End
         Begin VB.TextBox TxtLastNoDay 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   645
            Left            =   5010
            TabIndex        =   57
            Top             =   -570
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.TextBox TxtRemarks 
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
            Height          =   1005
            Left            =   300
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1200
            Width           =   13305
         End
         Begin VB.TextBox TxtCode 
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
            Height          =   375
            Left            =   12285
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   90
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo DcbEployee 
            Height          =   315
            Left            =   6705
            TabIndex        =   5
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   90
            Width           =   5475
            _ExtentX        =   9657
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSick 
            Height          =   315
            Left            =   195
            TabIndex        =   6
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   90
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal FrmDateH 
            Height          =   390
            Left            =   7740
            TabIndex        =   8
            Top             =   630
            Width           =   2160
            _extentx        =   2381
            _extenty        =   714
         End
         Begin MSComCtl2.DTPicker FrmDate 
            Height          =   390
            Left            =   9975
            TabIndex        =   7
            Top             =   630
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   688
            _Version        =   393216
            Format          =   107479041
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal ToDateH 
            Height          =   300
            Left            =   195
            TabIndex        =   10
            Top             =   630
            Width           =   2115
            _extentx        =   2381
            _extenty        =   714
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   300
            Left            =   2580
            TabIndex        =   9
            Top             =   630
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            _Version        =   393216
            Format          =   107479041
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker FirstDate 
            Height          =   285
            Left            =   8490
            TabIndex        =   63
            Top             =   4800
            Visible         =   0   'False
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   503
            _Version        =   393216
            Format          =   107479041
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker LastDate 
            Height          =   285
            Left            =   5775
            TabIndex        =   64
            Top             =   4800
            Visible         =   0   'False
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   503
            _Version        =   393216
            Format          =   107479041
            CurrentDate     =   38784
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   1875
            Left            =   300
            TabIndex        =   73
            Top             =   2730
            Width           =   13410
            _cx             =   23654
            _cy             =   3307
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmRegsterSickleave.frx":262DC
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·«Ã«“« "
            Height          =   315
            Index           =   13
            Left            =   6300
            TabIndex        =   69
            Top             =   2340
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «·«Ã«“… «·Õ«·Ì…"
            Height          =   495
            Index           =   12
            Left            =   9390
            TabIndex        =   68
            Top             =   2310
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «Ì«„ «·«Ã«“… «·”«»ﬁ…"
            Height          =   495
            Index           =   10
            Left            =   12990
            TabIndex        =   67
            Top             =   2310
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "”»» «·«Ã«“…"
            Height          =   615
            Index           =   9
            Left            =   11370
            TabIndex        =   65
            Top             =   4770
            Visible         =   0   'False
            Width           =   2955
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   615
            Index           =   6
            Left            =   5160
            TabIndex        =   56
            Top             =   630
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   615
            Index           =   5
            Left            =   12750
            TabIndex        =   55
            Top             =   630
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   615
            Index           =   3
            Left            =   14025
            TabIndex        =   54
            Top             =   630
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "”»» «·«Ã«“…"
            Height          =   495
            Index           =   2
            Left            =   13590
            TabIndex        =   53
            Top             =   1380
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«Ã«“…"
            Height          =   735
            Index           =   1
            Left            =   4635
            TabIndex        =   52
            Top             =   90
            Width           =   2010
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊŸ›"
            Height          =   735
            Index           =   0
            Left            =   13485
            TabIndex        =   51
            Top             =   90
            Width           =   1395
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1275
         Index           =   3
         Left            =   -4560
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   8340
         Visible         =   0   'False
         Width           =   7065
         _cx             =   12462
         _cy             =   2249
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
         Caption         =   " Õœœ «·› —…"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
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
         Begin VB.ComboBox CboYear 
            Height          =   315
            ItemData        =   "FrmRegsterSickleave.frx":26467
            Left            =   2355
            List            =   "FrmRegsterSickleave.frx":26469
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   180
            Width           =   1005
         End
         Begin VB.ComboBox CmbMonth 
            Height          =   315
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰…"
            Height          =   240
            Index           =   8
            Left            =   2955
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘Â—"
            Height          =   195
            Index           =   7
            Left            =   1425
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   270
            Width           =   645
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
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmRegsterSickleave"
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
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblRegsterSickleave", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub




Private Sub ButPrient_Click()
 Dim MySQL As String, StrFileName As String
        Dim RsData As New ADODB.Recordset
        Dim xApp As New CRAXDRT.Application
        Dim xReport As CRAXDRT.Report
        Dim CViewer As ClsReportViewer
        Dim StrReportTitle As String
        Dim Msg As String
       
        MySQL = "        SELECT TblRegsterSickleave.*,tbd2.branch_name, te.Fullcode,te.Emp_Name CusName,ts.Name  Sickleave"
        MySQL = MySQL & "  From TblRegsterSickleave"
        MySQL = MySQL & "                LEFT OUTER JOIN TblEmployee AS te ON te.Emp_ID = TblRegsterSickleave.EmpID"
        MySQL = MySQL & "                LEFT OUTER JOIN TblBranchesData AS tbd  ON tbd.branch_id = TblRegsterSickleave.BranchID"
        MySQL = MySQL & "                LEFT OUTER JOIN TblSickleave AS ts  ON ts.ID= TblRegsterSickleave.SickID, TblBranchesData AS tbd2"
        MySQL = MySQL & " Where 1 = 1"
        
       ' MySQL = MySQL & " Where (dbo.TblEmployee.Emp_ID =" & Me.DCEmP7.BoundText & ")"
        
        If Trim(TxtSerial1) <> "" Then
            MySQL = MySQL & " And (dbo.TblRegsterSickleave.id =N'" & Me.TxtSerial1 & "')"
        End If
        
'        If Trim(DboParentAccount.BoundText) <> "" Then
'            MySQL = MySQL & " And (dbo.TblCustemers.cusid =" & DboParentAccount.BoundText & ")"
'        End If
        
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RegsterSickleave.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RegsterSickleave.rpt"
        End If
        
         If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Sub
        End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        
        xReport.Database.SetDataSource RsData
    
        Dim cCompanyInfo As New ClsCompanyInfo
    
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
            StrReportTitle = "" '& StrAccountName
        Else
            xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
            xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
            StrReportTitle = ""
        End If
    
        xReport.ParameterFields(3).AddCurrentValue user_name
    
        Dim Total As String
        Dim dif As String
        Dim totl As Double
    
        xReport.reporttitle = StrReportTitle
        xReport.EnableParameterPrompting = False
        xReport.ApplicationName = App.title
        xReport.ReportAuthor = App.title
        Set CViewer = New ClsReportViewer
        CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        
End Sub

Private Sub DcbEployee_Change()
DcbEployee_Click (0)
End Sub

Private Sub DcbEployee_Click(Area As Integer)
 Dim EmpCode  As String
 Dim Balance As String
 Dim Account_code2 As String
If Me.TxtModFlg.Text = "N" Then
If val(DcbEployee.BoundText) <> 0 Then
TxtLastNoDay.Text = GetMaxNoday(val(DcbEployee.BoundText))
      
End If
End If
 If val(DcbEployee.BoundText) = 0 Then Exit Sub
 
 lbl(9).Caption = GetEmployeeSalaryAccordingToComponent(val(DcbEployee.BoundText), "", 0)
      GetEmployeeIDFromCode , , DcbEployee.BoundText, EmpCode
      Me.TxtCode.Text = EmpCode
End Sub

Private Sub DisplaySickDays()
Dim rsDummy As New ADODB.Recordset
Dim rsDummy2 As New ADODB.Recordset
Dim mPrevSickDays As Double
Dim mCurrentSickDays As Double
Dim mTotalSickDays As Double
Dim mTotalSickDaysBalnce As Double
Dim s As String
Dim mRow As Long
mCurrentSickDays = DateDiff("D", FrmDate.value, ToDate.value)
Dim mFromDate As Date
Dim mToDate As Date
s = " Select Sum(DATEDIFF(d, FrmDate,ToDate)) as Days from TblRegsterSickleave "
s = s & "  Where EmpID = " & val(DcbEployee.BoundText)
s = s & " AND ToDate  <=" & SQLDate(FrmDate.value, True) & ""
'rsDummy.Close
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not rsDummy.EOF Then
    mPrevSickDays = val(rsDummy!days & "")
End If

mTotalSickDays = mPrevSickDays + mCurrentSickDays
txtCurrentSickDays = mCurrentSickDays
txtPrevSickDays = mPrevSickDays
txtTotalSickDays = mTotalSickDays
rsDummy.Close
Dim i As Integer
Dim mLastDateMonth As Date
Dim mBeginFromDate As Date

  FG.Rows = 1
  mFromDate = FrmDate.value
  mToDate = ToDate.value
                
    Dim mDayH As Long
    mDayH = DateDiff("D", mFromDate, mToDate)

    If Month(DateAdd("D", mDayH, mFromDate)) = Month(mFromDate) And year(DateAdd("D", mDayH, mFromDate)) = year(mFromDate) Then
        mCurrentSickDays = val(txtCurrentSickDays)
        mRow = FG.FindRow(Month(mFromDate), FG.FixedRows, FG.ColIndex("MonthNo"), False, True)
        If mCurrentSickDays = 0 Then i = i + 1: GoTo NextRow
        'If mRow = -1 Then
            FG.Rows = FG.Rows + 1
            mRow = FG.Rows - 1
        'End If
                mCurrentSickDays = val(FG.TextMatrix(mRow, FG.ColIndex("SickDays"))) + mCurrentSickDays
        FG.TextMatrix(mRow, FG.ColIndex("SickDays")) = mCurrentSickDays
        FG.TextMatrix(mRow, FG.ColIndex("MonthName")) = GetMonthName(Month(mFromDate))
        FG.TextMatrix(mRow, FG.ColIndex("MonthNo")) = (Month(mFromDate))
        FG.TextMatrix(mRow, FG.ColIndex("SickYear")) = year(mFromDate)
        
        If FG.Rows - 1 = 1 Then
            mTotalSickDaysBalnce = mCurrentSickDays + val(txtPrevSickDays)
        Else
            mTotalSickDaysBalnce = mTotalSickDaysBalnce + mCurrentSickDays
        End If
        FG.TextMatrix(mRow, FG.ColIndex("SickBalance")) = mTotalSickDaysBalnce
       
       s = " Select Top 1 TblSickleaveDet.PerstageSal,* from TblSickleaveDet"
       s = s & " WHERE SickLID = " & val(DcbSick.BoundText) & "  AND " & mTotalSickDaysBalnce & "  BETWEEN FROMNo and ToNo"
       If rsDummy2.State = 1 Then rsDummy2.Close
       rsDummy2.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
       If Not rsDummy2.EOF Then
            FG.TextMatrix(mRow, FG.ColIndex("SickDiscPercent")) = rsDummy2!PerstageSal & ""
            FG.TextMatrix(mRow, FG.ColIndex("SalryValueDay")) = Round(GetEmployeeSalaryAccordingToComponent(val(DcbEployee.BoundText), "") / GetMonthDaysCount(Month(mFromDate), year(mFromDate)), 2)
            FG.TextMatrix(mRow, FG.ColIndex("SickTotalDisc")) = val(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay"))) * val(rsDummy2!PerstageSal & "") / 100
            
            FG.TextMatrix(mRow, FG.ColIndex("SickDisc")) = val(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay"))) * val(rsDummy2!PerstageSal & "") / 100
            FG.TextMatrix(mRow, FG.ColIndex("SickTotalDisc")) = (val(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay"))) * val(rsDummy2!PerstageSal & "") / 100) * mCurrentSickDays
          
       End If
       Exit Sub
    End If
    

    
RetestAgain:
    If (DateAdd("D", mDayH, mFromDate)) >= (mFromDate) Then
        Dim dd As Date
        Dim MonthCount As Long
        Dim mDay As Long
        MonthCount = GetMonthDaysCount(Month(mFromDate), year(mFromDate))
        mDay = MonthCount - IIf(Day(mFromDate) = 1, 0, Day(mFromDate))
      

        dd = DateAdd("D", -MonthCount, CDate(mFromDate))
         dd = DateAdd("D", mDayH, CDate(mFromDate))
       If Month(DateAdd("D", mDayH, mFromDate)) = Month(mFromDate) And year(mFromDate) = year(DateAdd("D", mDayH, mFromDate)) Then
            mCurrentSickDays = val(mDayH)
        Else
            mCurrentSickDays = val(mDay)
       End If
     
       
      '  mRow = FG.FindRow(Month(mFromDate), FG.FixedRows, FG.ColIndex("MonthNo"), False, True)
        If mCurrentSickDays = 0 Then i = i + 1: GoTo NextRow
      '  If mRow = -1 Then
            FG.Rows = FG.Rows + 1
            mRow = FG.Rows - 1
      '  End If
      
      '  mCurrentSickDays = val(FG.TextMatrix(mRow, FG.ColIndex("SickDays"))) + mCurrentSickDays
        FG.TextMatrix(mRow, FG.ColIndex("SickDays")) = mCurrentSickDays
        FG.TextMatrix(mRow, FG.ColIndex("MonthName")) = GetMonthName(Month(mFromDate))
        FG.TextMatrix(mRow, FG.ColIndex("MonthNo")) = (Month(mFromDate))
        FG.TextMatrix(mRow, FG.ColIndex("SickYear")) = year(mFromDate)
        
        If FG.Rows - 1 = 1 Then
            mTotalSickDaysBalnce = mCurrentSickDays + val(txtPrevSickDays)
        Else
            mTotalSickDaysBalnce = mTotalSickDaysBalnce + mCurrentSickDays
        End If
        FG.TextMatrix(mRow, FG.ColIndex("SickBalance")) = mTotalSickDaysBalnce
       
       s = " Select Top 1 TblSickleaveDet.PerstageSal,* from TblSickleaveDet"
       s = s & " WHERE SickLID = " & val(DcbSick.BoundText) & "  AND " & mTotalSickDaysBalnce & "  BETWEEN FROMNo and ToNo"
       If rsDummy2.State = 1 Then rsDummy2.Close
       rsDummy2.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
       If Not rsDummy2.EOF Then
            FG.TextMatrix(mRow, FG.ColIndex("SickDiscPercent")) = rsDummy2!PerstageSal & ""
            FG.TextMatrix(mRow, FG.ColIndex("SalryValueDay")) = Round(GetEmployeeSalaryAccordingToComponent(val(DcbEployee.BoundText), "") / GetMonthDaysCount(Month(mFromDate), year(mFromDate)), 2)
           ' FG.TextMatrix(mRow, FG.ColIndex("SickTotalDisc")) = val(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay"))) * val(rsDummy2!PerstageSal & "") / 100
            
            FG.TextMatrix(mRow, FG.ColIndex("SickDisc")) = val(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay"))) * val(rsDummy2!PerstageSal & "") / 100
            FG.TextMatrix(mRow, FG.ColIndex("SickTotalDisc")) = (val(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay"))) * val(rsDummy2!PerstageSal & "") / 100) * mCurrentSickDays
          
       End If
        'MonthCount = GetMonthDaysCount(Month(mFromDate), year(mFromDate))
        'mDay = MonthCount - (Day(mFromDate) - 1)
        
        If Day(mFromDate) = 1 Then
            mFromDate = DateAdd("D", mCurrentSickDays, mFromDate)
        Else
            mFromDate = DateAdd("D", mCurrentSickDays + 1, mFromDate)
        End If
        mDayH = mDayH - mCurrentSickDays
        
        If DateAdd("D", mDayH, mFromDate) = mFromDate Or mDayH < 0 Then
            Exit Sub
        End If
        GoTo RetestAgain
        If DateDiff("D", mFromDate, ToDate) < val(txtCurrentSickDays) Then
        End If
        
        mToDate = DateAdd("D", mDay, mFromDate)
        
        If Month(DateAdd("D", mDayH, mFromDate)) > Month(mFromDate) Then
        End If
    End If

NextRow:
      Exit Sub

End Sub

Public Function GetMonthName(ByVal mMonthNo As Long) As String
    
    Dim mLang As Boolean
    mLang = SystemOptions.UserInterface = ArabicInterface
    If SystemOptions.UserInterface = ArabicInterface Then
        Select Case mMonthNo
        Case 1
            GetMonthName = IIf(mLang, "Ì‰«Ì—", "january")
        Case 2
            GetMonthName = IIf(mLang, "›»—«Ì—", "February")
        Case 3
            GetMonthName = IIf(mLang, "„«—”", "March")
        Case 4
            GetMonthName = IIf(mLang, "«»—Ì·", "April")
        Case 5
            GetMonthName = IIf(mLang, "„«ÌÊ", "May")
        Case 6
            GetMonthName = IIf(mLang, "ÌÊ‰ÌÊ", "june")
        Case 7
            GetMonthName = IIf(mLang, "ÌÊ·ÌÊ", "July")
        Case 8
            GetMonthName = IIf(mLang, "√€”ÿ”", "August")
        Case 9
            GetMonthName = IIf(mLang, "”» „»—", "September")
        Case 10
            GetMonthName = IIf(mLang, "«ﬂ Ê»—", "Oct")
        Case 11
            GetMonthName = IIf(mLang, "‰Ê›„»—", "Nov")
        Case 12
            GetMonthName = IIf(mLang, "œÌ”„»—", "Dec")
       End Select
    End If

End Function

Private Sub DcbEployee_Validate(Cancel As Boolean)
DisplaySickDays
End Sub

Private Sub DcbSick_Validate(Cancel As Boolean)
DisplaySickDays
     
     
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from  TblRegsterSickleave  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.DcbBranch
    Dcombos.GetSickleave Me.DcbSick
    Dcombos.GetEmployees Me.DcbEployee
    YearMonth
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
Function GetPerstagSal(Optional SickLID As Double, Optional NoDay As Double) As Double
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim sql As String
sql = " select * from TblSickleaveDet where SickLID=" & SickLID & " "
sql = sql & "and  FromNo <= " & NoDay & " "
sql = sql & "and  ToNo >= " & NoDay & " "
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
GetPerstagSal = IIf(IsNull(Rs5("PerstageSal").value), -1, Rs5("PerstageSal").value)
Else
GetPerstagSal = -1
End If
End Function
Function GetPerstagSalMax(Optional SickLID As Double) As Double
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim sql As String
sql = " select max(ToNo)as MaxToNo from TblSickleaveDet where SickLID=" & SickLID & " "
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
GetPerstagSalMax = IIf(IsNull(Rs5("MaxToNo").value), 0, Rs5("MaxToNo").value)
Else
GetPerstagSalMax = 0
End If
End Function
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Function ChekPeriod() As Boolean
Dim rs2 As ADODB.Recordset
Dim sql As String
Set rs2 = New ADODB.Recordset
sql = "Select * from TblRegsterSickleaveDet where RegSickLID <>" & val(TxtSerial1.Text) & " and EmpID=" & val(Me.DcbEployee.BoundText) & " "
sql = sql & "and  FrmDate<= " & SQLDate(FrmDate.value, True) & ""
sql = sql & "and  ToDate >= " & SQLDate(FrmDate.value, True) & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
ChekPeriod = True
Else
ChekPeriod = False
End If
End Function
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
    Dim i As Integer
      If Me.TxtModFlg.Text = "E" Then
        Cn.Execute " delete from TblRegsterSickleaveDet where RegSickLID=" & val(TxtSerial1.Text) & "  "
        Cn.Execute " delete from TblRegsterSickleave2 where RegSickLID=" & val(TxtSerial1.Text) & "  "
                   
          
      End If
RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
RsSavRec.Fields("EmpID").value = val(Me.DcbEployee.BoundText)
RsSavRec.Fields("SickID").value = val(Me.DcbSick.BoundText)
RsSavRec.Fields("Remarks").value = Me.TxtRemarks.Text
RsSavRec.Fields("RecordDate").value = RecordDate.value
RsSavRec.Fields("RecordDateH").value = RecordDateH.value
RsSavRec.Fields("FrmDate").value = FrmDate.value
RsSavRec.Fields("FrmDateH").value = FrmDateH.value
RsSavRec.Fields("ToDate").value = ToDate.value
RsSavRec.Fields("ToDateH").value = ToDateH.value
RsSavRec.Fields("LastNoDay").value = val(Me.TxtLastNoDay.Text)
   RsSavRec.update
  ''//////////////////////////
  Dim DiffMonth As Integer
  Dim PerstageSal As Double
  Dim str As String
DiffMonth = DateDiff("m", Me.FrmDate.value, Me.ToDate.value)
  Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblRegsterSickleaveDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    Dim ContDay As Double
    ContDay = val(TxtLastNoDay.Text)
       For i = 1 To DiffMonth + 1
       If i = 1 Then
       FirstDate.value = FrmDate.value
       Else
       FirstDate.value = DateAdd("m", i - 1, FrmDate.value)
       CboYear.Text = year(FirstDate.value)
       CmbMonth.ListIndex = Month(FirstDate.value) - 1
       str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.Text
       FirstDate.value = CDate(str)
       End If
       If i = (DiffMonth + 1) Then
       LastDate.value = ToDate.value
       Else
       LastDate.value = MonthLastDay(FirstDate.value)
       End If
       ContDay = ContDay + DateDiff("d", Me.FirstDate.value, Me.LastDate.value) + 1
       If GetPerstagSal(val(DcbSick.BoundText), ContDay) = -1 Then
       If GetPerstagSalMax(val(DcbSick.BoundText)) = 0 Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox " «ﬂœ „‰ «⁄œ«œ«  «·«Ã«“« "
       Else
       MsgBox "Make sure the settings of your vacations"
       End If
       Exit Sub
       Else
       PerstageSal = 100 - GetPerstagSalMax(val(DcbSick.BoundText))
       End If
       Else
       PerstageSal = 100 - GetPerstagSal(val(DcbSick.BoundText), ContDay)
       End If
       CboYear.Text = year(FirstDate.value)
       CmbMonth.ListIndex = Month(FirstDate.value) - 1
       RsDevsub.AddNew
                RsDevsub("RegSickLID").value = val(Me.TxtSerial1.Text)
                RsDevsub("EmpID").value = val(DcbEployee.BoundText)
                RsDevsub("YearID").value = val(Me.CboYear.Text)
                RsDevsub("MonthID").value = val(CmbMonth.ListIndex)
                RsDevsub("BranchID").value = val(Me.DcbBranch.BoundText)
                RsDevsub("FrmDate").value = FirstDate.value
                RsDevsub("ToDate").value = LastDate.value
                RsDevsub("TotalNoDay").value = ContDay
                RsDevsub("PerstageSal").value = PerstageSal
                RsDevsub("NewValSalar").value = (val(lbl(9).Caption) * PerstageSal) / 100
       RsDevsub.update
     Next i
    Dim s As String
              DB_CreateField "TblRegsterSickleave2", "SickID", adInteger, adColNullable, , , " ???    ", False, True
    s = "Select * from TblRegsterSickleave2 Where Id = 0 "
    saveGrid s, FG, "MonthNo", "", "RegSickLID", val(Me.TxtSerial1.Text), "SickID", val(Me.DcbSick.BoundText), "EmpID", val(Me.DcbEployee.BoundText)
     
    
'''///////////////
   
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
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
   Function GetMaxNoday(Optional EmpID As Double) As Double
   Dim sql As String
   Dim Rs4 As ADODB.Recordset
   Set Rs4 = New ADODB.Recordset
   sql = "Select max(TotalNoDay) as MaxNo from TblRegsterSickleaveDet where EmpID=" & EmpID & ""
   Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If Rs4.RecordCount > 0 Then
   GetMaxNoday = IIf(IsNull(Rs4("MaxNo").value), 0, Rs4("MaxNo").value)
   Else
   GetMaxNoday = 0
   End If
   End Function
   Public Function MonthLastDay(ByVal dCurrDate As Date) As Date
    Dim dFirstDayNextMonth As Date
  
    MonthLastDay = Empty
    dCurrDate = Format(dCurrDate, "DD/MM/YYYY")
  
    dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
    MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  
    Exit Function
 
End Function
  Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

End Sub

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()

   On Error GoTo ErrTrap
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DcbEployee.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    Me.DcbSick.BoundText = IIf(IsNull(RsSavRec.Fields("SickID").value), "", RsSavRec.Fields("SickID").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    FrmDateH.value = IIf(IsNull(RsSavRec.Fields("FrmDateH").value), ToHijriDate(Date), RsSavRec.Fields("FrmDateH").value)
    FrmDate.value = IIf(IsNull(RsSavRec.Fields("FrmDate").value), Date, RsSavRec.Fields("FrmDate").value)
    FrmDateH.value = IIf(IsNull(RsSavRec.Fields("FrmDateH").value), ToHijriDate(Date), RsSavRec.Fields("FrmDateH").value)
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value)
    ToDateH.value = IIf(IsNull(RsSavRec.Fields("ToDateH").value), ToHijriDate(Date), RsSavRec.Fields("ToDateH").value)
    TxtLastNoDay.Text = IIf(IsNull(RsSavRec.Fields("LastNoDay").value), 0, RsSavRec.Fields("LastNoDay").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
     DisplaySickDays
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
 If val(lbl(9).Caption) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "·„ Ì „  ÕœÌœ «·—« »"
 Else
 MsgBox "Salary is not specified"
 End If
 Exit Sub
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
If val(Me.DcbEployee.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÊŸ›"
Else
MsgBox "Please Select Employee"
End If
DcbEployee.SetFocus
Exit Sub
End If
If val(Me.DcbSick.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·«Ã«“…"
Else
MsgBox "Please Select Type"
End If
DcbSick.SetFocus
Exit Sub
End If
If ToDate.value < Me.FrmDate.value Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ ‰Â«Ì… «·› —… «ﬁ· „‰ »œ«Ì… «·› —…"
Else
MsgBox "It can not be the end of the period less than the beginning of the period"
End If
Exit Sub
End If
If ChekPeriod() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  €ÌÌ— «·› —… .Â–Â «·› —… „ÊÃÊœ… „”»ﬁ«"
Else
MsgBox "Please Make Sure The Period .This Is Period Already Exists"
End If
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



Private Sub FrmDate_Change()
If Me.TxtModFlg.Text <> "R" Then
If Not IsNull(FrmDate.value) Then
   FrmDateH.value = ToHijriDate(FrmDate.value)
  ' DisplaySickDays
End If
End If
End Sub


Private Sub FrmDate_Validate(Cancel As Boolean)
If Not IsNull(FrmDate.value) Then
   FrmDateH.value = ToHijriDate(FrmDate.value)
   DisplaySickDays
End If
End Sub

Private Sub FrmDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 FrmDate.value = ToGregorianDate(FrmDateH.value)
End If
End Sub

Private Sub RecordDate_Change()
If Me.TxtModFlg.Text <> "R" Then
If Not IsNull(RecordDate.value) Then
         RecordDateH.value = ToHijriDate(RecordDate.value)
End If
End If
End Sub

Private Sub RecordDate_Validate(Cancel As Boolean)
DisplaySickDays
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 RecordDate.value = ToGregorianDate(RecordDateH.value)
End If
End Sub

Private Sub ToDate_Change()
If Me.TxtModFlg.Text <> "R" Then
If Not IsNull(ToDate.value) Then
   ToDateH.value = ToHijriDate(ToDate.value)
   'DisplaySickDays
End If
End If
End Sub


Private Sub ToDate_Validate(Cancel As Boolean)
DisplaySickDays
End Sub

Private Sub ToDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 ToDate.value = ToGregorianDate(ToDateH.value)
End If
End Sub


Private Sub txtcode_KeyPress(KeyAscii As Integer)
  Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtCode.Text, EmpID
        Me.DcbEployee.BoundText = EmpID
    End If
End Sub

Function ChePayment() As Boolean
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = "Select * from  TblRegsterSickleaveDet where RegSickLID= " & val(TxtSerial1.Text) & " and payed=1 "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChePayment = True
Else
ChePayment = False
End If
End Function

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
    Dim i As Integer
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim ID As Double
    If ChePayment() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–›  „ ⁄„· „”Ì— —Ê« »"
    Else
    MsgBox "Can not be delete. Linked To Salary"
    End If
    Exit Sub
    End If
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
         Cn.Execute "Delete from TblRegsterSickleaveDet where RegSickLID=" & val(TxtSerial1.Text) & " "
                RsSavRec.Find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.Delete
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
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
    If ChePayment() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·  „ ⁄„· „”Ì— —Ê« »"
    Else
    MsgBox "Can not be edited Linked To Salary"
    End If
    Exit Sub
    End If
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


Private Sub ChangeLang()
On Error GoTo ErrTrap
   On Error GoTo ErrTrap
    



    Label1(2).Caption = "Seak Leave"
    lbl(14).Caption = "By"
    lbl(4).Caption = "No"
    lbl(25).Caption = "Date"
    lbl(11).Caption = "Branch"
    
    lbl(0).Caption = "Emp."
    
    lbl(1).Caption = "Type"
    lbl(3).Caption = "Period"
    
    lbl(5).Caption = "From"
   lbl(6).Caption = "To"
   lbl(2).Caption = "Remarks"
   
   
    
'    Cmd(3).Caption = "Delete"
'    Cmd(4).Caption = "Delete All"
'    ISButton5.Caption = "Print"
    'ISButton8.Caption = "Search"
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
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
    lbl(1).Caption = "Date"

'    SelectDept.RightToLeft = False
'    SelectDept.Caption = "Management"
     
ErrTrap:
 
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblRegsterSickleave"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub





Private Sub TestByWael()
'    Dim ff As Date
''    s = " Select (DATEDIFF(d, FrmDate,ToDate)) as Days,FrmDate,ToDate, "
''    s = s & " ,SELECT DAY(DATEADD(DD,-1,DATEADD(MM,DATEDIFF(MM,-1,FrmDate),0)))"
''    s = s & " from TblRegsterSickleave "
''    s = s & "  Where EmpID = " & val(DcbEployee.BoundText)
''    s = s & " AND ToDate  <=" & SQLDate(FrmDate.value, True) & ""
''    s = "SELECT * FROM TblRegsterSickleave"
''    s = s & "  Where EmpID = " & val(DcbEployee.BoundText)
''    s = s & "  and SickID= " & val(DcbSick.BoundText)
'
'    s = " SELECT Days,Days2,frmDate,ToDate,FrD,ToD,"
'    s = s & " frmDate22,ToDate22,FrD2,ToD2,"
'    s = s & " DATEadd(DD,FrD - DAY(frmDate ) ,frmDate ) FF,DATEadd(DD,ToD - DAY(ToDate ) ,ToDate ), "
'    s = s & " DATEadd(DD,FrD2 - DAY(frmDate22 ) ,frmDate22 ) FF2,DATEadd(DD,ToD2 - DAY(ToDate22 ) ,ToDate22 ) "
'    s = s & " TT"
'    s = s & " FROM ("
'
'   ' s = s & " Select  (DATEDIFF(d, FrmDate,ToDate)) as Days,frmDate,ToDate"
'   ' s = s & " ,FrD = (SELECT DAY(DATEADD(DD,-1,DATEADD(MM,DATEDIFF(MM,-1,FrmDate),0))))"
'   ' s = s & " ,ToD = (SELECT DAY(DATEADD(DD,-1,DATEADD(MM,DATEDIFF(MM,-1,ToDate),0))))"
''
''    s = s & "     SELECT TblEmployee.Emp_ID,
''    s = s & " (DATEDIFF(d, isNull(frmDate," & SQLDate(FrmDate.value, True) & "), isNull(ToDate," & SQLDate(ToDate.value, True) & ") )) AS Days,"
''    s = s & "                    isNull(frmDate," & SQLDate(FrmDate.value, True) & ") frmDate,"
''    s = s & "                       isNull(ToDate," & SQLDate(ToDate.value, True) & ") ToDate"
''    s = s & "                       ,"
''    s = s & "                       FrD = ("
''    s = s & "                           SELECT DAY(DATEADD(DD, -1, DATEADD(MM, DATEDIFF(MM, -1,  isNull(frmDate," & SQLDate(FrmDate.value, True) & ")), 0)))"
''    s = s & "                       ),"
''    s = s & "                       ToD = ("
''    s = s & "                           SELECT DAY(DATEADD(DD, -1, DATEADD(MM, DATEDIFF(MM, -1, isNull(ToDate," & SQLDate(ToDate.value, True) & ")), 0)))"
''    s = s & "                       )"
'
'    s = s & "     SELECT TblEmployee.Emp_ID, (DATEDIFF(d, '" & DisplayDate(FrmDate.value) & "', '" & DisplayDate(ToDate.value) & "')) AS Days,"
'    s = s & "                    '" & DisplayDate(FrmDate.value) & "' frmDate,"
'    s = s & "                       '" & DisplayDate(ToDate.value) & "' ToDate "
'    s = s & "                       ,"
'    s = s & "                       FrD = ("
'    s = s & "                           SELECT DAY(DATEADD(DD, -1, DATEADD(MM, DATEDIFF(MM, -1,  '" & DisplayDate(FrmDate.value) & "'), 0)))"
'    s = s & "                       ),"
'    s = s & "                       ToD = ("
'    s = s & "                           SELECT DAY(DATEADD(DD, -1, DATEADD(MM, DATEDIFF(MM, -1, '" & DisplayDate(ToDate.value) & "'), 0)))"
'    s = s & "                       )"
'
'    s = s & "                       ,(DATEDIFF(d, isNull(frmDate," & SQLDate(FrmDate.value, True) & "), isNull(ToDate," & SQLDate(ToDate.value, True) & ") )) AS Days2,"
'    s = s & "                    isNull(frmDate," & SQLDate(FrmDate.value, True) & ") frmDate22,"
'    s = s & "                       isNull(ToDate," & SQLDate(ToDate.value, True) & ") ToDate22"
'    s = s & "                       ,"
'    s = s & "                       FrD2 = ("
'    s = s & "                           SELECT DAY(DATEADD(DD, -1, DATEADD(MM, DATEDIFF(MM, -1,  isNull(frmDate," & SQLDate(FrmDate.value, True) & ")), 0)))"
'    s = s & "                       ),"
'    s = s & "                       ToD2 = ("
'    s = s & "                           SELECT DAY(DATEADD(DD, -1, DATEADD(MM, DATEDIFF(MM, -1, isNull(ToDate," & SQLDate(ToDate.value, True) & ")), 0)))"
'    s = s & "                       )"
'
'
'    s = s & "              From"
'    s = s & "               TblEmployee  Left OUTER JOIN"
'
'
'    s = s & "               TblRegsterSickleave"
'    s = s & "    ON TblEmployee.Emp_ID =TblRegsterSickleave.EmpID"
'    s = s & "  and SickID= " & val(DcbSick.BoundText)
'    s = s & "  Where TblEmployee.Emp_ID = " & val(DcbEployee.BoundText)
'
'    s = s & " ) T"
'
'
'    FG.Rows = 1
'
' '   rsDummy.Close
'    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    Do While Not rsDummy.EOF
'      '  For i = 1 To 12
'
'
'
'                mBeginFromDate = rsDummy!FrmDate
'
'                ff = rsDummy!ff
'
'
'TryTest:
'            If DateDiff("D", mBeginFromDate, rsDummy!ToDate) < DateDiff("D", mBeginFromDate, ff) Then
'              ' mCurrentSickDays = val(rsDummy!days & "")
'               mCurrentSickDays = DateDiff("D", mBeginFromDate, rsDummy!ToDate)
'   mRow = FG.FindRow(Month(rsDummy!ToDate), FG.FixedRows, FG.ColIndex("MonthNo"), False, True)
'               If mCurrentSickDays = 0 Then i = i + 1: GoTo NextRow
'               If mRow = -1 Then
'                    FG.Rows = FG.Rows + 1
'                    mRow = FG.Rows - 1
'                End If
'               mCurrentSickDays = val(FG.TextMatrix(mRow, FG.ColIndex("SickDays"))) + mCurrentSickDays
'               FG.TextMatrix(mRow, FG.ColIndex("SickDays")) = mCurrentSickDays
'               FG.TextMatrix(mRow, FG.ColIndex("MonthName")) = GetMonthName(Month(rsDummy!ToDate))
'               FG.TextMatrix(mRow, FG.ColIndex("MonthNo")) = (Month(rsDummy!ToDate))
'               FG.TextMatrix(mRow, FG.ColIndex("SickYear")) = year(rsDummy!ToDate)
'
'                If FG.Rows - 1 = 1 Then
'                    mTotalSickDaysBalnce = mCurrentSickDays + val(txtPrevSickDays)
'                Else
'                    mTotalSickDaysBalnce = mTotalSickDaysBalnce + mCurrentSickDays
'                End If
'                FG.TextMatrix(mRow, FG.ColIndex("SickBalance")) = mTotalSickDaysBalnce
'
'               s = " Select Top 1 TblSickleaveDet.PerstageSal,* from TblSickleaveDet"
'               s = s & " WHERE SickLID = " & val(DcbSick.BoundText) & "  AND " & mTotalSickDaysBalnce & "  BETWEEN FROMNo and ToNo"
'               If rsDummy2.State = 1 Then rsDummy2.Close
'               rsDummy2.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'               If Not rsDummy2.EOF Then
'                    FG.TextMatrix(mRow, FG.ColIndex("SickDiscPercent")) = rsDummy2!PerstageSal & ""
'                    FG.TextMatrix(mRow, FG.ColIndex("SalryValueDay")) = GetSalaryEmployee(val(DcbEployee.BoundText)) / GetMonthDaysCount(Month(rsDummy!ToDate), year(rsDummy!ToDate))
'                    FG.TextMatrix(mRow, FG.ColIndex("SickTotalDisc")) = val(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay"))) * val(rsDummy2!PerstageSal & "") / 100
'
'               End If
'               i = i + 1
'               GoTo NextRow
'
'
'            ElseIf DateDiff("D", rsDummy!FrmDate, rsDummy!ToDate) > DateDiff("D", mBeginFromDate, ff) Then
'                mCurrentSickDays = DateDiff("D", mBeginFromDate, ff)
'                 mRow = FG.FindRow(Month(ff), FG.FixedRows, FG.ColIndex("MonthNo"), False, True)
'               If mCurrentSickDays = 0 Then i = i + 1: GoTo NextRow
'               If mRow = -1 Then
'                    FG.Rows = FG.Rows + 1
'                    mRow = FG.Rows - 1
'                End If
'                mCurrentSickDays = val(FG.TextMatrix(mRow, FG.ColIndex("SickDays"))) + mCurrentSickDays
'
'                FG.TextMatrix(mRow, FG.ColIndex("SickDays")) = mCurrentSickDays
'                FG.TextMatrix(mRow, FG.ColIndex("MonthName")) = GetMonthName(Month(ff))
'                FG.TextMatrix(mRow, FG.ColIndex("MonthNo")) = (Month(ff))
'                FG.TextMatrix(mRow, FG.ColIndex("SickYear")) = year(ff)
'
'                If FG.Rows - 1 = 1 Then
'                    mTotalSickDaysBalnce = mCurrentSickDays '+ val(txtPrevSickDays)
'                Else
'                    mTotalSickDaysBalnce = mTotalSickDaysBalnce + mCurrentSickDays
'                End If
'                FG.TextMatrix(mRow, FG.ColIndex("SickBalance")) = mTotalSickDaysBalnce
'
'
'
'                s = " Select  PerstageSal from TblSickleaveDet"
'                s = s & " WHERE SickLID = " & val(DcbSick.BoundText) & "  AND " & val(mTotalSickDaysBalnce) & "  BETWEEN FROMNo and ToNo"
'                If rsDummy2.State = 1 Then rsDummy2.Close
'                rsDummy2.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                If Not rsDummy2.EOF Then
'                    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SickDiscPercent")) = rsDummy2!PerstageSal & ""
'
'                    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay")) = GetSalaryEmployee(val(DcbEployee.BoundText)) / GetMonthDaysCount(Month(ff), year(ff))
'                    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SickTotalDisc")) = val(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("SalryValueDay"))) * val(rsDummy2!PerstageSal & "") / 100
'                End If
'                Dim DayMonth2 As Integer
'
'               mBeginFromDate = DateAdd("D", 0, ff)
'               DayMonth2 = GetMonthDaysCount(Month(DateAdd("M", 1, mBeginFromDate)), year(DateAdd("M", 1, mBeginFromDate)))
'               ff = DateAdd("d", DayMonth2 - Day(DateAdd("M", 1, mBeginFromDate)), DateAdd("M", 1, mBeginFromDate))
'               i = i + 1
'               'FF = Day(DateAdd("D", -1, DateAdd("M", DateDiff("M", -1, FrmDate), 0)))
'               GoTo TryTest
'            Else
'                GoTo NextRow
'            End If
'
'
'
'
'       ' Next
'NextRow22:
'        rsDummy.MoveNext
'    Loop


End Sub
