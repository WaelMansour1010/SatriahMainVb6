VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmAproveComponYear 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmAproveComponYear.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   14550
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
      TabIndex        =   5
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmAproveComponYear.frx":6852
      Left            =   15480
      List            =   "FrmAproveComponYear.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   6
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
      TabIndex        =   7
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
            Picture         =   "FrmAproveComponYear.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAproveComponYear.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAproveComponYear.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAproveComponYear.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAproveComponYear.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAproveComponYear.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAproveComponYear.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAproveComponYear.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   8
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
      ButtonImage     =   "FrmAproveComponYear.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   10
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
      ButtonImage     =   "FrmAproveComponYear.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   11
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
      ButtonImage     =   "FrmAproveComponYear.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8385
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   14550
      _cx             =   25665
      _cy             =   14790
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
         Height          =   765
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   14535
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
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
            ButtonImage     =   "FrmAproveComponYear.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   17
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
            ButtonImage     =   "FrmAproveComponYear.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   18
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
            ButtonImage     =   "FrmAproveComponYear.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   19
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
            ButtonImage     =   "FrmAproveComponYear.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈À»«  «·»œ·«  «·„ﬁœ„…"
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
            Height          =   375
            Index           =   2
            Left            =   8880
            TabIndex        =   20
            Top             =   240
            Width           =   4080
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "FrmAproveComponYear.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1230
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   7185
         Width           =   14580
         _cx             =   25718
         _cy             =   2170
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
            Height          =   210
            Left            =   13125
            TabIndex        =   22
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   765
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   370
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
            ButtonImage     =   "FrmAproveComponYear.frx":17E16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   210
            Left            =   11310
            TabIndex        =   23
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   765
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   370
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
            ButtonImage     =   "FrmAproveComponYear.frx":1E678
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   210
            Left            =   9675
            TabIndex        =   24
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   765
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   370
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
            ButtonImage     =   "FrmAproveComponYear.frx":24EDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   210
            Left            =   7965
            TabIndex        =   25
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   765
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   370
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
            ButtonImage     =   "FrmAproveComponYear.frx":25274
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   210
            Left            =   6225
            TabIndex        =   26
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   765
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   370
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
            ButtonImage     =   "FrmAproveComponYear.frx":2560E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   270
            Left            =   5220
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   765
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   476
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
            ButtonImage     =   "FrmAproveComponYear.frx":25BA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   210
            Left            =   1665
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   765
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   370
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
            ButtonImage     =   "FrmAproveComponYear.frx":2C40A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   210
            Left            =   3435
            TabIndex        =   29
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   765
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   370
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
            ButtonImage     =   "FrmAproveComponYear.frx":2C7A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10620
            TabIndex        =   35
            Top             =   75
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   555
            Index           =   11
            Left            =   2760
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   90
            Width           =   5295
            _cx             =   9340
            _cy             =   979
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
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   315
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   120
               Width           =   945
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   1455
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   120
               Width           =   1875
            End
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1350
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   120
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   35
               Left            =   3525
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   150
            Left            =   240
            TabIndex        =   40
            Top             =   150
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   135
            Left            =   1785
            TabIndex        =   39
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   150
            Index           =   1
            Left            =   810
            TabIndex        =   38
            Top             =   150
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   150
            Index           =   0
            Left            =   2505
            TabIndex        =   37
            Top             =   150
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   195
            Index           =   8
            Left            =   13605
            TabIndex        =   36
            Top             =   75
            Width           =   900
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5820
         Left            =   0
         TabIndex        =   30
         Top             =   1395
         Width           =   14535
         _cx             =   25638
         _cy             =   10266
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
         Caption         =   "»Ì«‰«  «”«”Ì…|ÌœÊÌ|«·„—›ﬁ« "
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
         Flags(1)        =   2
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   5400
            Left            =   45
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9525
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
               Height          =   300
               Left            =   210
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               Top             =   105
               Width           =   6615
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   3960
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   1365
               Width           =   14445
               _cx             =   25479
               _cy             =   6985
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
               Begin XtremeSuiteControls.CheckBox CheckBox1 
                  Height          =   270
                  Left            =   12840
                  TabIndex        =   81
                  Top             =   120
                  Width           =   1335
                  _Version        =   786432
                  _ExtentX        =   2355
                  _ExtentY        =   476
                  _StockProps     =   79
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  ForeColor       =   8388608
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   360
                  Index           =   3
                  Left            =   13095
                  TabIndex        =   41
                  Top             =   3585
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   635
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› ”ÿ— "
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
                  ButtonImage     =   "FrmAproveComponYear.frx":2CB3E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   360
                  Index           =   4
                  Left            =   11775
                  TabIndex        =   42
                  Top             =   3585
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   635
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ·"
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
                  ButtonImage     =   "FrmAproveComponYear.frx":2D0D8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                  Height          =   3240
                  Left            =   0
                  TabIndex        =   76
                  Top             =   375
                  Width           =   14340
                  _cx             =   25294
                  _cy             =   5715
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
                  BackColor       =   16777215
                  ForeColor       =   -2147483640
                  BackColorFixed  =   14871017
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   16776960
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483642
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   12
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmAproveComponYear.frx":2D672
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   1
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   9
                  Left            =   360
                  TabIndex        =   88
                  Top             =   3720
                  Width           =   3885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   285
                  Index           =   7
                  Left            =   4320
                  TabIndex        =   87
                  Top             =   3720
                  Width           =   1005
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   795
               Left            =   0
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   495
               Width           =   14445
               _cx             =   25479
               _cy             =   1402
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
               Begin VB.TextBox Text1 
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
                  Height          =   300
                  Left            =   10755
                  MaxLength       =   50
                  TabIndex        =   62
                  Top             =   465
                  Width           =   750
               End
               Begin XtremeSuiteControls.CheckBox SelectBranch 
                  Height          =   225
                  Left            =   11595
                  TabIndex        =   61
                  Top             =   120
                  Width           =   1020
                  _Version        =   786432
                  _ExtentX        =   1799
                  _ExtentY        =   397
                  _StockProps     =   79
                  Caption         =   "›—⁄ „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdAll 
                  Height          =   270
                  Left            =   12735
                  TabIndex        =   63
                  Top             =   105
                  Width           =   1380
                  _Version        =   786432
                  _ExtentX        =   2434
                  _ExtentY        =   476
                  _StockProps     =   79
                  Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdEmp 
                  Height          =   225
                  Left            =   12735
                  TabIndex        =   64
                  Top             =   465
                  Width           =   1380
                  _Version        =   786432
                  _ExtentX        =   2434
                  _ExtentY        =   397
                  _StockProps     =   79
                  Caption         =   "„ÊŸ› „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbEmployee1 
                  Height          =   315
                  Left            =   6930
                  TabIndex        =   65
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   465
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbBranch1 
                  Height          =   315
                  Left            =   6930
                  TabIndex        =   66
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   120
                  Width           =   4575
                  _ExtentX        =   8070
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcpDept1 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   67
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   120
                  Width           =   4380
                  _ExtentX        =   7726
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   735
                  Left            =   120
                  TabIndex        =   68
                  ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                  Top             =   15
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   1296
                  Caption         =   "«÷«›…"
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
                  ButtonImage     =   "FrmAproveComponYear.frx":2D959
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin XtremeSuiteControls.CheckBox SelectDept 
                  Height          =   225
                  Left            =   5550
                  TabIndex        =   69
                  Top             =   120
                  Width           =   1230
                  _Version        =   786432
                  _ExtentX        =   2170
                  _ExtentY        =   397
                  _StockProps     =   79
                  Caption         =   "«œ«—… „Õœœ…"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbProject1 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   70
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   465
                  Width           =   4380
                  _ExtentX        =   7726
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox SelectProject 
                  Height          =   225
                  Left            =   5550
                  TabIndex        =   71
                  Top             =   465
                  Width           =   1230
                  _Version        =   786432
                  _ExtentX        =   2170
                  _ExtentY        =   397
                  _StockProps     =   79
                  Caption         =   "„‘—Ê⁄ „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin MSComCtl2.DTPicker FromDate 
               Height          =   300
               Left            =   11520
               TabIndex        =   77
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox        =   -1  'True
               Format          =   97976321
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker ToDate 
               Height          =   300
               Left            =   8040
               TabIndex        =   79
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox        =   -1  'True
               Format          =   97976321
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï  «—ÌŒ"
               Height          =   270
               Index           =   6
               Left            =   9960
               TabIndex        =   80
               Top             =   135
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰  «—ÌŒ"
               Height          =   270
               Index           =   0
               Left            =   13200
               TabIndex        =   78
               Top             =   135
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   195
               Index           =   5
               Left            =   6480
               TabIndex        =   75
               Top             =   105
               Width           =   1230
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   5400
            Left            =   15480
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9525
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
               Height          =   1650
               Left            =   10740
               MaxLength       =   50
               TabIndex        =   47
               Top             =   3405
               Width           =   765
            End
            Begin XtremeSuiteControls.CheckBox BranchSelect 
               Height          =   1320
               Left            =   11595
               TabIndex        =   46
               Top             =   1320
               Width           =   1005
               _Version        =   786432
               _ExtentX        =   1773
               _ExtentY        =   2328
               _StockProps     =   79
               Caption         =   "›—⁄ „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton SelectAll 
               Height          =   1620
               Left            =   12735
               TabIndex        =   48
               Top             =   1230
               Width           =   1395
               _Version        =   786432
               _ExtentX        =   2461
               _ExtentY        =   2857
               _StockProps     =   79
               Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton EmpSelect 
               Height          =   1320
               Left            =   12735
               TabIndex        =   49
               Top             =   3405
               Width           =   1395
               _Version        =   786432
               _ExtentX        =   2461
               _ExtentY        =   2328
               _StockProps     =   79
               Caption         =   "„ÊŸ› „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbEmployee 
               Height          =   315
               Left            =   6945
               TabIndex        =   50
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   3405
               Width           =   3840
               _ExtentX        =   6773
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBranch 
               Height          =   315
               Left            =   6945
               TabIndex        =   51
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1320
               Width           =   4560
               _ExtentX        =   8043
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDepatment 
               Height          =   315
               Left            =   1080
               TabIndex        =   52
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1320
               Width           =   4380
               _ExtentX        =   7726
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   4320
               Left            =   135
               TabIndex        =   53
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   735
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   7620
               Caption         =   "«÷«›…"
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
               ButtonImage     =   "FrmAproveComponYear.frx":341BB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin XtremeSuiteControls.CheckBox DeptSelect 
               Height          =   1320
               Left            =   5550
               TabIndex        =   54
               Top             =   1320
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   2328
               _StockProps     =   79
               Caption         =   "«œ«—… „Õœœ…"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbProject 
               Height          =   315
               Left            =   1080
               TabIndex        =   55
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   3405
               Width           =   4380
               _ExtentX        =   7726
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ProjSelect 
               Height          =   1320
               Left            =   5550
               TabIndex        =   56
               Top             =   3405
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   2328
               _StockProps     =   79
               Caption         =   "„‘—Ê⁄ „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
               ForeColor       =   &H00800000&
               Height          =   1500
               Index           =   3
               Left            =   12870
               TabIndex        =   57
               Top             =   0
               Width           =   1530
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   5400
            Left            =   15180
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9525
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
            Begin C1SizerLibCtl.C1Tab C1Tab2 
               Height          =   5700
               Left            =   0
               TabIndex        =   59
               Top             =   0
               Width           =   14535
               _cx             =   25638
               _cy             =   10054
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
               Caption         =   "»Ì«‰«  «”«”Ì…|ÌœÊÌ|«·„—›ﬁ« "
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
               Flags(2)        =   2
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   690
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   720
         Width           =   14565
         _cx             =   25691
         _cy             =   1217
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
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11565
            TabIndex        =   1
            Top             =   240
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   8640
            TabIndex        =   43
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   97976321
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBrnch2 
            Height          =   315
            Left            =   1080
            TabIndex        =   72
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   2
            Left            =   6960
            TabIndex        =   73
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   10200
            TabIndex        =   44
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   285
            Index           =   4
            Left            =   13440
            TabIndex        =   33
            Top             =   240
            Width           =   975
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
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmAproveComponYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecID As String
 Public LngRow As Long
 Public LngCol As Long
 Dim II As Long
 Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "«À»«  «·»œ·«  «·„ﬁœ„… —ﬁ„" & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim Sql As String
tablename = "TblApproveCompoYear"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1.Text)
Notevalue = 0
 notytype = 9054
Notevalue = val(Lbl(9).Caption)
 

 BranchID = val(DcbBrnch2.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                    Else
                                                 If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                         CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                              TxtNoteID.Text = NoteID
                                                             TxtNoteSerial.Text = NoteSerial
                                                 Else
                                                              Sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                              Sql = Sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                 Sql = Sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                 Cn.Execute Sql
                                        
                                                End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function

Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempCustomerCode As String
    Dim StrTempCustomerCodeInsuranceAccount  As String
    
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim I As Integer
 
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
 LngDevNO = 0
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
    my_branch = BranchID
  With GridInstallments
   StrTempAccountCode = get_account_code_branch(64, my_branch)
            StrTempDes = "«À»«  «·»œ·«  «·„” Õﬁ…    " & TxtSerial1.Text
       
      For I = 1 To .Rows - 1
      If .Cell(flexcpChecked, I, .ColIndex("Selct")) = flexChecked And val(.TextMatrix(I, .ColIndex("MordValue"))) <> 0 And val(.TextMatrix(I, .ColIndex("EmpID"))) <> 0 Then
             LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, val(.TextMatrix(I, .ColIndex("MordValue"))), 0, StrTempDes & "      »œ·«  „ﬁœ„… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
          StrTempAccountCode = get_EMPLOYEE_Account(val(.TextMatrix(I, .ColIndex("EmpID"))), "Account_code1")
          LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, val(.TextMatrix(I, .ColIndex("MordValue"))), 1, StrTempDes & "     «ÃÊ— „” Õﬁ… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
     End If
      Next I
  End With
    

ErrTrap:
End Function

Private Sub CheckBox1_Click()
Relin
Dim I As Integer
If CheckBox1.value = vbChecked Then
With GridInstallments
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Selct")) = 1
Next I
End With
Else
With GridInstallments
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Selct")) = 0
Next I
End With
End If

End Sub

Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 3
RemoveGridRow
Case 4
RemoveGridAllRow
End Select
End If
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DcbEmployee1_Change()
DcbEmployee1_Click (0)
End Sub

Private Sub DcbEmployee1_Click(Area As Integer)
If val(DcbEmployee1.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , Me.DcbEmployee1.BoundText, EmpCode
    Me.Text1.Text = EmpCode
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    With GridInstallments
     If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("TypeMofrd")) = "#1;”‰ÊÌ  |#2;‰’› ”‰ÊÌ "
             ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("TypeMofrd")) = "#1;Yearly |#2;Half Yearly"
            End If
    End With

    conection = "select * from  TblApproveCompoYear  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetEmployees Me.DcbEmployee1
   Dcombos.GetBranches Me.DcbBranch1
   Dcombos.GetEmpDepartments Me.DcpDept1
   Dcombos.GetProjects Me.DcbProject1
   Dcombos.GetBranches Me.DcbBrnch2
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
    Dim Sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                  StrSQL = "Delete From TblApproveCompoYearDet Where CoYerID=" & val(Me.TxtSerial1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
          Cn.Execute StrSQL, , adExecuteNoRecords
           End If
   RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
   RsSavRec.Fields("FromDate").value = FromDate.value
   RsSavRec.Fields("ToDate").value = ToDate.value
   RsSavRec.Fields("BrnchID").value = val(Me.DcbBrnch2.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("DeptID").value = val(Me.DcpDept1.BoundText)
   RsSavRec.Fields("ProjectID").value = val(Me.DcbProject1.BoundText)
   RsSavRec.Fields("BrnchID1").value = val(Me.DcbBranch1.BoundText)
   RsSavRec.Fields("EmpID").value = val(Me.DcbEmployee1.BoundText)
   RsSavRec.Fields("Remarks").value = txtRemarks.Text
   RsSavRec.Fields("TotalValue").value = val(Lbl(9).Caption)
   If Me.RdAll.value = True Then
   RsSavRec.Fields("AllEmp").value = 1
   ElseIf RdEmp.value = True Then
   RsSavRec.Fields("AllEmp").value = 2
   End If
   If Me.SelectBranch.value = vbChecked Then
   RsSavRec.Fields("SelBrnch").value = 1
   End If
   If Me.SelectDept.value = vbChecked Then
   RsSavRec.Fields("SelDept").value = 1
   End If
   If Me.SelectProject.value = vbChecked Then
   RsSavRec.Fields("SelProj").value = 1
   End If
    RsSavRec.update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblApproveCompoYearDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim I As Integer
    With Me.GridInstallments
       For I = .FixedRows To .Rows - 1
       If .Cell(flexcpChecked, I, .ColIndex("Selct")) = flexChecked Then
       RsDevsub.AddNew
                RsDevsub("CoYerID").value = val(Me.TxtSerial1.Text)
                RsDevsub("MofrdID").value = IIf((.TextMatrix(I, .ColIndex("MofrdID"))) = "", Null, val(.TextMatrix(I, .ColIndex("MofrdID"))))
                RsDevsub("EmpID").value = IIf((.TextMatrix(I, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(I, .ColIndex("EmpID"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(I, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(I, .ColIndex("DeptID"))))
                RsDevsub("ProjectID").value = IIf((.TextMatrix(I, .ColIndex("ProjID"))) = "", Null, val(.TextMatrix(I, .ColIndex("ProjID"))))
                RsDevsub("BrnchID1").value = IIf((.TextMatrix(I, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(I, .ColIndex("BranchID"))))
                RsDevsub("RecDate1").value = IIf((.TextMatrix(I, .ColIndex("RecDate1"))) = "", Null, (.TextMatrix(I, .ColIndex("RecDate1"))))
                RsDevsub("CompYerID").value = IIf((.TextMatrix(I, .ColIndex("CompYerID"))) = "", Null, val(.TextMatrix(I, .ColIndex("CompYerID"))))
                RsDevsub("TypeMofrd").value = IIf((.TextMatrix(I, .ColIndex("TypeMofrd"))) = "", Null, val(.TextMatrix(I, .ColIndex("TypeMofrd"))))
                RsDevsub("MordValue").value = IIf((.TextMatrix(I, .ColIndex("MordValue"))) = "", Null, val(.TextMatrix(I, .ColIndex("MordValue"))))
                RsDevsub("StFunction").value = IIf((.TextMatrix(I, .ColIndex("StFunction"))) = "", Null, (.TextMatrix(I, .ColIndex("StFunction"))))
                RsDevsub("flg").value = IIf((.TextMatrix(I, .ColIndex("Flg"))) = "", Null, val(.TextMatrix(I, .ColIndex("Flg"))))
       RsDevsub.update
       If val(.TextMatrix(I, .ColIndex("Flg"))) = 2 Then
       Cn.Execute "Update TblComponentYearDet set FlgSel2 =1 where id=" & val(.TextMatrix(I, .ColIndex("CompYerID"))) & ""
       ElseIf val(.TextMatrix(I, .ColIndex("Flg"))) = 1 Then
       Cn.Execute "Update TblComponentYearDet set FlgSel =1 where id=" & val(.TextMatrix(I, .ColIndex("CompYerID"))) & ""
       End If
       Else
         If val(.TextMatrix(I, .ColIndex("Flg"))) = 2 Then
       Cn.Execute "Update TblComponentYearDet set FlgSel2 =null where id=" & val(.TextMatrix(I, .ColIndex("CompYerID"))) & ""
       ElseIf val(.TextMatrix(I, .ColIndex("Flg"))) = 1 Then
       Cn.Execute "Update TblComponentYearDet set FlgSel =null where id=" & val(.TextMatrix(I, .ColIndex("CompYerID"))) & ""
       End If
         End If
     Next I
    End With
    createVoucher
    Dim Msg As String
      Select Case Me.TxtModFlg.Text
        Case "N"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & Chr(13)
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

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim I As Integer
     SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
    Dim Shifttime As Date
        Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
       Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
     Lbl(9).Caption = IIf(IsNull(RsSavRec.Fields("TotalValue").value), 0, RsSavRec.Fields("TotalValue").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    FromDate.value = IIf(IsNull(RsSavRec.Fields("FromDate").value), Date, RsSavRec.Fields("FromDate").value)
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value)
    Me.DcbBrnch2.BoundText = IIf(IsNull(RsSavRec.Fields("BrnchID").value), "", RsSavRec.Fields("BrnchID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcpDept1.BoundText = IIf(IsNull(RsSavRec.Fields("DeptID").value), "", RsSavRec.Fields("DeptID").value)
    Me.DcbProject1.BoundText = IIf(IsNull(RsSavRec.Fields("ProjectID").value), "", RsSavRec.Fields("ProjectID").value)
    Me.DcbBranch1.BoundText = IIf(IsNull(RsSavRec.Fields("BrnchID1").value), "", RsSavRec.Fields("BrnchID1").value)
    Me.DcbEmployee1.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    Me.txtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    If Not (IsNull(RsSavRec.Fields("AllEmp").value)) Then
    If RsSavRec.Fields("AllEmp").value = 1 Then
    RdAll.value = True
    ElseIf RsSavRec.Fields("AllEmp").value = 2 Then
    RdEmp.value = True
    End If
    Else
    RdAll.value = False
    End If

    If Not (IsNull(RsSavRec.Fields("SelBrnch").value)) Then
    If RsSavRec.Fields("SelBrnch").value = 1 Then
    Me.SelectBranch.value = vbChecked
    Else
    SelectBranch.value = vbUnchecked
    End If
    Else
    SelectBranch.value = vbUnchecked
    End If
     If Not (IsNull(RsSavRec.Fields("SelDept").value)) Then
    If RsSavRec.Fields("SelDept").value = 1 Then
    Me.SelectDept.value = vbChecked
    Else
    SelectDept.value = vbUnchecked
    End If
    Else
    SelectDept.value = vbUnchecked
    End If
    If Not (IsNull(RsSavRec.Fields("SelProj").value)) Then
    If RsSavRec.Fields("SelProj").value = 1 Then
    Me.SelectProject.value = vbChecked
    Else
    SelectProject.value = vbUnchecked
    End If
    Else
    SelectProject.value = vbUnchecked
    End If
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData

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
    Dim Sm As Double
    Dim Account_Code_dynamic As String
    Dim I As Integer
If val(Me.DcbBrnch2.BoundText) = 0 Or Me.DcbBrnch2.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
Else
MsgBox "Please Select Branch"
End If
DcbBrnch2.SetFocus
Exit Sub
End If
Account_Code_dynamic = get_account_code_branch(64, my_branch)

    If Account_Code_dynamic = "NO branch" Then
    If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
        MsgBox "Please Create Branch"
        End If
                Exit Sub
            Else

                If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·»œ·«  «·„ﬁœ„…", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If
            

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
    StrRecID = new_id("TblApproveCompoYear", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim Sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
Sql = " SELECT     dbo.TblApproveCompoYearDet.CoYerID, dbo.TblApproveCompoYearDet.ID, dbo.TblApproveCompoYearDet.MofrdID, dbo.mofrdat.mofrad_name, "
Sql = Sql & "                      dbo.mofrdat.mofrad_namee, dbo.TblApproveCompoYearDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
Sql = Sql & "                      dbo.TblApproveCompoYearDet.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
Sql = Sql & "                      dbo.TblApproveCompoYearDet.BrnchID1, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblApproveCompoYearDet.ProjectID,"
Sql = Sql & "                      dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblApproveCompoYearDet.TypeMofrd, dbo.TblApproveCompoYearDet.StFunction,"
Sql = Sql & "                      dbo.TblApproveCompoYearDet.RecDate1, dbo.TblApproveCompoYearDet.MordValue, dbo.TblApproveCompoYearDet.CompYerID,"
Sql = Sql & "                      dbo.TblApproveCompoYearDet.flg"
Sql = Sql & " FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
Sql = Sql & "                      dbo.projects RIGHT OUTER JOIN"
Sql = Sql & "                      dbo.TblApproveCompoYearDet ON dbo.projects.id = dbo.TblApproveCompoYearDet.ProjectID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblBranchesData ON dbo.TblApproveCompoYearDet.BrnchID1 = dbo.TblBranchesData.branch_id ON"
Sql = Sql & "                      dbo.TblEmpDepartments.DeparmentID = dbo.TblApproveCompoYearDet.DeptID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblEmployee ON dbo.TblApproveCompoYearDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.mofrdat ON dbo.TblApproveCompoYearDet.MofrdID = dbo.mofrdat.mofrad_code"
Sql = Sql & " Where (dbo.TblApproveCompoYearDet.CoYerID = " & val(TxtSerial1.Text) & ")"
  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim I As Integer
     
     With Me.GridInstallments
                    For I = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(I, .ColIndex("Ser")) = I
                   .TextMatrix(I, .ColIndex("Selct")) = 1
                   .TextMatrix(I, .ColIndex("TypeMofrd")) = IIf(IsNull(Rs1("TypeMofrd").value), "", Rs1("TypeMofrd").value)
                   .TextMatrix(I, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
                   .TextMatrix(I, .ColIndex("DeptID")) = IIf(IsNull(Rs1("DeptID").value), 0, Rs1("DeptID").value)
                   .TextMatrix(I, .ColIndex("ProjID")) = IIf(IsNull(Rs1("ProjectID").value), 0, Rs1("ProjectID").value)
                   .TextMatrix(I, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BrnchID1").value), 0, Rs1("BrnchID1").value)
                   .TextMatrix(I, .ColIndex("MofrdID")) = IIf(IsNull(Rs1("MofrdID").value), 0, Rs1("MofrdID").value)
                   .TextMatrix(I, .ColIndex("MordValue")) = IIf(IsNull(Rs1("MordValue").value), 0, Rs1("MordValue").value)
                   .TextMatrix(I, .ColIndex("StFunction")) = IIf(IsNull(Rs1("StFunction").value), "", Rs1("StFunction").value)
                   .TextMatrix(I, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(I, .ColIndex("RecDate1")) = IIf(IsNull(Rs1("RecDate1").value), "", Rs1("RecDate1").value)
                   .TextMatrix(I, .ColIndex("CompYerID")) = IIf(IsNull(Rs1("CompYerID").value), 0, Rs1("CompYerID").value)
                   .TextMatrix(I, .ColIndex("Flg")) = IIf(IsNull(Rs1("Flg").value), 0, Rs1("Flg").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(I, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentName").value), "", Rs1("DepartmentName").value)
                   .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_name").value), "", Rs1("Project_name").value)
                   .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                   .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs1("mofrad_name").value), "", Rs1("mofrad_name").value)
                   .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   Else
                   .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs1("mofrad_namee").value), "", Rs1("mofrad_namee").value)
                   .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_nameE").value), "", Rs1("Project_nameE").value)
                   .TextMatrix(I, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentNamee").value), "", Rs1("DepartmentNamee").value)
                   .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   End If
      
                   Rs1.MoveNext
             Next I
End With
        
        Exit Sub
ErrTrap:
    End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Relin
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GridInstallments
If .ColKey(Col) = "Selct" Then
.ComboList = ""
Else
Cancel = True
End If
End With
End Sub
Sub Relin()
Dim I As Integer
Dim SumVal As Double
SumVal = 0
With GridInstallments
For I = 1 To .Rows - 1
If .Cell(flexcpChecked, I, .ColIndex("Selct")) = flexChecked Then
SumVal = SumVal + val(.TextMatrix(I, .ColIndex("MordValue")))
End If
Next I
End With
Lbl(9).Caption = SumVal
End Sub
Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
If RdEmp.value = True Then
If val(Me.DcbEmployee1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÊŸ›"
Else
MsgBox "Please Select Employee"
End If
DcbEmployee1.SetFocus
Exit Sub
End If
End If

If SelectDept.value = vbChecked Then
If val(Me.DcpDept1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·«œ«—…"
Else
MsgBox "Please Select Management"
End If
DcpDept1.SetFocus
Exit Sub
End If
End If
''//////////////
If Me.SelectProject.value = vbChecked Then
If val(Me.DcbProject1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
Else
MsgBox "Please Select Project"
End If
DcbProject1.SetFocus
Exit Sub
End If
End If

If SelectBranch.value = vbChecked Then
If val(Me.DcbBranch1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
Else
MsgBox "Please Select Branch"
End If
DcbBranch1.SetFocus
Exit Sub
End If
End If
filgrid1
filgrid3
End If
End Sub
Sub filgrid3()
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim I As Integer
Dim k As Integer
Dim Sql As String
Sql = "SELECT     dbo.TblComponentYearDet.ID, dbo.TblComponentYearDet.CoYerID, dbo.TblComponentYearDet.MofrdID, dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, "
Sql = Sql & "                      dbo.TblComponentYearDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblComponentYearDet.RecDate2,"
Sql = Sql & "                      dbo.TblComponentYearDet.RecDate1, dbo.TblComponentYearDet.MordValue, dbo.TblComponentYearDet.FlgSel, dbo.TblComponentYearDet.StFunction,"
Sql = Sql & "                      dbo.TblComponentYearDet.TypeMofrd, dbo.TblComponentYearDet.BrnchID1, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
Sql = Sql & "                      dbo.TblComponentYearDet.ProjectID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblComponentYearDet.DeptID,"
Sql = Sql & "                      dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee"
Sql = Sql & " FROM         dbo.TblComponentYearDet LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblEmpDepartments ON dbo.TblComponentYearDet.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.projects ON dbo.TblComponentYearDet.ProjectID = dbo.projects.id LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblBranchesData ON dbo.TblComponentYearDet.BrnchID1 = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblEmployee ON dbo.TblComponentYearDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.mofrdat ON dbo.TblComponentYearDet.MofrdID = dbo.mofrdat.mofrad_code"
Sql = Sql & " WHERE      (dbo.TblComponentYearDet.FlgSel2 IS NULL) "

If val(Me.DcbProject1.BoundText) <> 0 And Me.SelectProject.value = vbChecked Then
Sql = Sql & " and dbo.TblComponentYearDet.ProjectID  =" & val(DcbProject1.BoundText) & " "
End If
If val(DcbBranch1.BoundText) <> 0 And Me.SelectBranch.value = vbChecked Then
Sql = Sql & " and dbo.TblComponentYearDet.BrnchID1  =" & val(DcbBranch1.BoundText) & " "
End If
If val(DcpDept1.BoundText) <> 0 And Me.SelectDept.value = vbChecked Then
Sql = Sql & " and dbo.TblComponentYearDet.DeptID  =" & val(DcpDept1.BoundText) & " "
End If
If val(DcbEmployee1.BoundText) <> 0 And RdEmp.value = True Then
Sql = Sql & " and dbo.TblComponentYearDet.EmpID =" & val(DcbEmployee1.BoundText) & " "
End If
If Not IsNull(Me.FromDate.value) Then
Sql = Sql & " and dbo.TblComponentYearDet.RecDate2 >=" & SQLDate(FromDate.value, True) & " "
End If
If Not IsNull(Me.ToDate.value) Then
Sql = Sql & " and dbo.TblComponentYearDet.RecDate2 <=" & SQLDate(ToDate.value, True) & " "
End If
Sql = Sql & " order by dbo.TblComponentYearDet.RecDate2 ,dbo.TblComponentYearDet.EmpID"
 Rs8.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs8.RecordCount > 0 Then
With GridInstallments
k = .Rows

Rs8.MoveFirst
.Rows = .Rows + Rs8.RecordCount
For I = k To .Rows - 1
.TextMatrix(I, .ColIndex("Flg")) = 2
.TextMatrix(I, .ColIndex("CompYerID")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
.TextMatrix(I, .ColIndex("EmpID")) = IIf(IsNull(Rs8("EmpID").value), 0, Rs8("EmpID").value)
.TextMatrix(I, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DeptID").value), 0, Rs8("DeptID").value)
.TextMatrix(I, .ColIndex("ProjID")) = IIf(IsNull(Rs8("ProjectID").value), 0, Rs8("ProjectID").value)
.TextMatrix(I, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BrnchID1").value), 0, Rs8("BrnchID1").value)
.TextMatrix(I, .ColIndex("FullCode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(I, .ColIndex("StFunction")) = IIf(IsNull(Rs8("StFunction").value), "", Rs8("StFunction").value)
.TextMatrix(I, .ColIndex("MordValue")) = IIf(IsNull(Rs8("MordValue").value), "", Rs8("MordValue").value)
.TextMatrix(I, .ColIndex("MordValue")) = val(.TextMatrix(I, .ColIndex("MordValue")) / 2)
.TextMatrix(I, .ColIndex("MofrdID")) = IIf(IsNull(Rs8("MofrdID").value), 0, Rs8("MofrdID").value)
.TextMatrix(I, .ColIndex("RecDate1")) = IIf(IsNull(Rs8("RecDate2").value), "", Rs8("RecDate2").value)
.TextMatrix(I, .ColIndex("TypeMofrd")) = IIf(IsNull(Rs8("TypeMofrd").value), "", Rs8("TypeMofrd").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs8("mofrad_name").value), "", Rs8("mofrad_name").value)
.TextMatrix(I, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentName").value), "", Rs8("DepartmentName").value)
.TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_name").value), "", Rs8("Project_name").value)
.TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs8("mofrad_namee").value), "", Rs8("mofrad_namee").value)
.TextMatrix(I, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentNamee").value), "", Rs8("DepartmentNamee").value)
.TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_nameE").value), "", Rs8("Project_nameE").value)
.TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)

End If
Rs8.MoveNext
Next I
'.AutoSize 0, .Cols - 1, False
End With
End If
End Sub
Sub filgrid1()
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim I As Integer
Dim k As Integer
Dim Sql As String
Sql = "SELECT     dbo.TblComponentYearDet.ID, dbo.TblComponentYearDet.CoYerID, dbo.TblComponentYearDet.MofrdID, dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, "
Sql = Sql & "                      dbo.TblComponentYearDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblComponentYearDet.RecDate2,"
Sql = Sql & "                      dbo.TblComponentYearDet.RecDate1, dbo.TblComponentYearDet.MordValue, dbo.TblComponentYearDet.FlgSel, dbo.TblComponentYearDet.StFunction,"
Sql = Sql & "                      dbo.TblComponentYearDet.TypeMofrd, dbo.TblComponentYearDet.BrnchID1, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
Sql = Sql & "                      dbo.TblComponentYearDet.ProjectID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblComponentYearDet.DeptID,"
Sql = Sql & "                      dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee"
Sql = Sql & " FROM         dbo.TblComponentYearDet LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblEmpDepartments ON dbo.TblComponentYearDet.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.projects ON dbo.TblComponentYearDet.ProjectID = dbo.projects.id LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblBranchesData ON dbo.TblComponentYearDet.BrnchID1 = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblEmployee ON dbo.TblComponentYearDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.mofrdat ON dbo.TblComponentYearDet.MofrdID = dbo.mofrdat.mofrad_code"
Sql = Sql & " WHERE      (dbo.TblComponentYearDet.FlgSel IS NULL) "

If val(Me.DcbProject1.BoundText) <> 0 And Me.SelectProject.value = vbChecked Then
Sql = Sql & " and dbo.TblComponentYearDet.ProjectID  =" & val(DcbProject1.BoundText) & " "
End If
If val(DcbBranch1.BoundText) <> 0 And Me.SelectBranch.value = vbChecked Then
Sql = Sql & " and dbo.TblComponentYearDet.BrnchID1  =" & val(DcbBranch1.BoundText) & " "
End If
If val(DcpDept1.BoundText) <> 0 And Me.SelectDept.value = vbChecked Then
Sql = Sql & " and dbo.TblComponentYearDet.DeptID  =" & val(DcpDept1.BoundText) & " "
End If
If val(DcbEmployee1.BoundText) <> 0 And RdEmp.value = True Then
Sql = Sql & " and dbo.TblComponentYearDet.EmpID =" & val(DcbEmployee1.BoundText) & " "
End If
If Not IsNull(Me.FromDate.value) Then
Sql = Sql & " and dbo.TblComponentYearDet.RecDate1 >=" & SQLDate(FromDate.value, True) & " "
End If
If Not IsNull(Me.ToDate.value) Then
Sql = Sql & " and dbo.TblComponentYearDet.RecDate1 <=" & SQLDate(ToDate.value, True) & " "
End If
Sql = Sql & " order by dbo.TblComponentYearDet.RecDate1 ,dbo.TblComponentYearDet.EmpID"
 Rs8.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs8.RecordCount > 0 Then
With GridInstallments
k = .Rows

Rs8.MoveFirst
.Rows = .Rows + Rs8.RecordCount
For I = k To .Rows - 1
.TextMatrix(I, .ColIndex("Flg")) = 1
.TextMatrix(I, .ColIndex("CompYerID")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
.TextMatrix(I, .ColIndex("EmpID")) = IIf(IsNull(Rs8("EmpID").value), 0, Rs8("EmpID").value)
.TextMatrix(I, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DeptID").value), 0, Rs8("DeptID").value)
.TextMatrix(I, .ColIndex("ProjID")) = IIf(IsNull(Rs8("ProjectID").value), 0, Rs8("ProjectID").value)
.TextMatrix(I, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BrnchID1").value), 0, Rs8("BrnchID1").value)
.TextMatrix(I, .ColIndex("FullCode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(I, .ColIndex("StFunction")) = IIf(IsNull(Rs8("StFunction").value), "", Rs8("StFunction").value)
.TextMatrix(I, .ColIndex("MordValue")) = IIf(IsNull(Rs8("MordValue").value), "", Rs8("MordValue").value)
If Not IsNull(Rs8("RecDate2").value) Then
.TextMatrix(I, .ColIndex("MordValue")) = val(.TextMatrix(I, .ColIndex("MordValue")) / 2)
End If
.TextMatrix(I, .ColIndex("MofrdID")) = IIf(IsNull(Rs8("MofrdID").value), 0, Rs8("MofrdID").value)
.TextMatrix(I, .ColIndex("RecDate1")) = IIf(IsNull(Rs8("RecDate1").value), "", Rs8("RecDate1").value)
.TextMatrix(I, .ColIndex("TypeMofrd")) = IIf(IsNull(Rs8("TypeMofrd").value), "", Rs8("TypeMofrd").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs8("mofrad_name").value), "", Rs8("mofrad_name").value)
.TextMatrix(I, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentName").value), "", Rs8("DepartmentName").value)
.TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_name").value), "", Rs8("Project_name").value)
.TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs8("mofrad_namee").value), "", Rs8("mofrad_namee").value)
.TextMatrix(I, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentNamee").value), "", Rs8("DepartmentNamee").value)
.TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_nameE").value), "", Rs8("Project_nameE").value)
.TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)

End If
Rs8.MoveNext
Next I
'.AutoSize 0, .Cols - 1, False
End With
End If
End Sub
Private Sub RdEmp_Click()
If Me.RdEmp.value = False Then
Me.DcbEmployee1.BoundText = ""
End If
End Sub

Private Sub SelectBranch_Click()
If Me.SelectBranch.value = vbUnchecked Then
Me.DcbBranch1.BoundText = ""
End If
End Sub

Private Sub SelectDept_Click()
If Me.SelectDept.value = vbUnchecked Then
DcpDept1.BoundText = ""
End If
End Sub

Private Sub SelectProject_Click()
If Me.SelectProject.value = vbUnchecked Then
Me.DcbProject1.BoundText = ""
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text1.Text, EmpID
        Me.DcbEmployee1.BoundText = EmpID
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
Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecID, , adSearchForward, 1
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
    Dim Sql As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim I As Integer
    Dim ID As Double
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
               RemoveGridAllRow
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Delete From TblApproveCompoYearDet Where CoYerID=" & val(Me.TxtSerial1.Text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords

                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
              SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
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
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
                    Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                        End If
                    Case "E"
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
                    Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
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
                   RecID As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    XPDtbTrans.Enabled = True
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    Dim I As Integer
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        GridInstallments.Rows = GridInstallments.Rows + 1
        Me.DCboUserName.BoundText = user_id
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & Chr(13)
            Msg = Msg & " You can not edit this the record now" & Chr(13)
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
    SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
    Me.DCboUserName.BoundText = user_id
 Me.DcbBrnch2.BoundText = Current_branch
 RdAll.value = True
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
    Wrap = Chr(13) + Chr(10)
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
   ' form name
Lbl(7).Caption = "Total"
Lbl(4).Caption = "No"
Lbl(1).Caption = "Date"
Lbl(2).Caption = "Branch"
Lbl(5).Caption = "Remarks"
Lbl(0).Caption = "From Date"
Lbl(6).Caption = "To Date"
CheckBox1.RightToLeft = False
CheckBox1.Caption = "Select All"
Label1(2).Caption = "Approve Allowance"
Cmd(3).Caption = "Delete"
Cmd(4).Caption = "Delete All"

    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"
   Label1(35).Caption = "No.GL"
Command9.Caption = "Print GL"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    Me.Lbl(8).Caption = "by"
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
Lbl(1).Caption = "Date"

SelectDept.RightToLeft = False
SelectDept.Caption = "Management"
SelectProject.RightToLeft = False
SelectProject.Caption = "Project"
ISButton2.Caption = "Add"
SelectBranch.RightToLeft = False
SelectBranch.Caption = "Branch"
RdAll.RightToLeft = False
RdAll.Caption = "All"
RdEmp.RightToLeft = False
RdEmp.Caption = "Select Employee"
'lbl(0).Caption = "Data"
C1Tab1.Caption = "Data"

  With Me.GridInstallments
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("FullCode")) = "Employee Code"
  .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
  .TextMatrix(0, .ColIndex("Name")) = "Component"
  .TextMatrix(0, .ColIndex("StFunction")) = "Proceure"
  .TextMatrix(0, .ColIndex("MordValue")) = "Value"
  .TextMatrix(0, .ColIndex("RecDate1")) = "Payment Date"
  .TextMatrix(0, .ColIndex("Selct")) = "Select"
  .TextMatrix(0, .ColIndex("TypeMofrd")) = "Type Payment"

  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblApproveCompoYear"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub RemoveGridAllRow()
Dim I As Integer
With GridInstallments
For I = 1 To .Rows - 1
        If val(.TextMatrix(I, .ColIndex("Flg"))) = 2 Then
       Cn.Execute "Update TblComponentYearDet set FlgSel2 =null where id=" & val(.TextMatrix(I, .ColIndex("CompYerID"))) & ""
       ElseIf val(.TextMatrix(I, .ColIndex("Flg"))) = 1 Then
       Cn.Execute "Update TblComponentYearDet set FlgSel =null where id=" & val(.TextMatrix(I, .ColIndex("CompYerID"))) & ""
       End If
 Next I
 .Clear flexClearScrollable, flexClearEverything
  .Rows = 1
 End With
End Sub
Private Sub RemoveGridRow()
Dim I As Integer
    With Me.GridInstallments
        If .Row <= 0 Then Exit Sub
        I = .Row
           If val(.TextMatrix(I, .ColIndex("Flg"))) = 2 Then
       Cn.Execute "Update TblComponentYearDet set FlgSel2 =null where id=" & val(.TextMatrix(I, .ColIndex("CompYerID"))) & ""
       ElseIf val(.TextMatrix(I, .ColIndex("Flg"))) = 1 Then
       Cn.Execute "Update TblComponentYearDet set FlgSel =null where id=" & val(.TextMatrix(I, .ColIndex("CompYerID"))) & ""
       End If
        .RemoveItem .Row
    End With
End Sub
