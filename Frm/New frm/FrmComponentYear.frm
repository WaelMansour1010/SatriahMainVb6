VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmComponentYear 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmComponentYear.frx":0000
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
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmComponentYear.frx":6852
      Left            =   15480
      List            =   "FrmComponentYear.frx":6862
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
      Left            =   15600
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
      Left            =   15600
      TabIndex        =   1
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
      Left            =   15480
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
            Picture         =   "FrmComponentYear.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComponentYear.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComponentYear.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComponentYear.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComponentYear.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComponentYear.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComponentYear.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComponentYear.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
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
      ButtonImage     =   "FrmComponentYear.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   9
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
      ButtonImage     =   "FrmComponentYear.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   10
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
      ButtonImage     =   "FrmComponentYear.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8385
      Left            =   0
      TabIndex        =   11
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
         TabIndex        =   12
         Top             =   0
         Width           =   14535
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
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
            ButtonImage     =   "FrmComponentYear.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
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
            ButtonImage     =   "FrmComponentYear.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
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
            ButtonImage     =   "FrmComponentYear.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
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
            ButtonImage     =   "FrmComponentYear.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Œÿ… «·»œ·«  «·„ﬁœ„…"
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
            TabIndex        =   19
            Top             =   240
            Width           =   4080
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "FrmComponentYear.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   870
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   7515
         Width           =   14550
         _cx             =   25665
         _cy             =   1535
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
            Height          =   240
            Left            =   13095
            TabIndex        =   21
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   423
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
            ButtonImage     =   "FrmComponentYear.frx":17E16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   240
            Left            =   11280
            TabIndex        =   22
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   423
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
            ButtonImage     =   "FrmComponentYear.frx":1E678
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   240
            Left            =   9660
            TabIndex        =   23
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   480
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   423
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
            ButtonImage     =   "FrmComponentYear.frx":24EDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   240
            Left            =   7950
            TabIndex        =   24
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   480
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   423
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
            ButtonImage     =   "FrmComponentYear.frx":25274
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   240
            Left            =   6210
            TabIndex        =   25
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   480
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   423
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
            ButtonImage     =   "FrmComponentYear.frx":2560E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   300
            Left            =   5205
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   480
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
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
            ButtonImage     =   "FrmComponentYear.frx":25BA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   240
            Left            =   1665
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   480
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   423
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
            ButtonImage     =   "FrmComponentYear.frx":2C40A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   240
            Left            =   3435
            TabIndex        =   28
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   423
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
            ButtonImage     =   "FrmComponentYear.frx":2C7A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10605
            TabIndex        =   32
            Top             =   90
            Width           =   2670
            _ExtentX        =   4710
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
            Height          =   165
            Left            =   240
            TabIndex        =   37
            Top             =   180
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   1785
            TabIndex        =   36
            Top             =   195
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   165
            Index           =   1
            Left            =   810
            TabIndex        =   35
            Top             =   180
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   165
            Index           =   0
            Left            =   2505
            TabIndex        =   34
            Top             =   180
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   210
            Index           =   8
            Left            =   13575
            TabIndex        =   33
            Top             =   90
            Width           =   900
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6075
         Left            =   0
         TabIndex        =   29
         Top             =   1395
         Width           =   14535
         _cx             =   25638
         _cy             =   10716
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
         Caption         =   "»Ì«‰«  «”«”Ì…|ÌœÊÌ|«·„—›ﬁ« |New Tab"
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
         Flags(3)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   5655
            Left            =   45
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9975
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   4260
               Left            =   0
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   1425
               Width           =   14445
               _cx             =   25479
               _cy             =   7514
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   225
                  Index           =   3
                  Left            =   13125
                  TabIndex        =   38
                  Top             =   3960
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmComponentYear.frx":2CB3E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   225
                  Index           =   4
                  Left            =   11790
                  TabIndex        =   39
                  Top             =   3960
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmComponentYear.frx":2D0D8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                  Height          =   3870
                  Left            =   0
                  TabIndex        =   71
                  Top             =   0
                  Width           =   14445
                  _cx             =   25479
                  _cy             =   6826
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
                  FormatString    =   $"FrmComponentYear.frx":2D672
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
               Height          =   315
               Left            =   6915
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   69
               Top             =   105
               Width           =   6390
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   825
               Left            =   0
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   525
               Width           =   14445
               _cx             =   25479
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
                  Height          =   315
                  Left            =   10770
                  MaxLength       =   50
                  TabIndex        =   57
                  Top             =   480
                  Width           =   720
               End
               Begin XtremeSuiteControls.CheckBox SelectBranch 
                  Height          =   240
                  Left            =   11610
                  TabIndex        =   56
                  Top             =   120
                  Width           =   1035
                  _Version        =   786432
                  _ExtentX        =   1826
                  _ExtentY        =   423
                  _StockProps     =   79
                  Caption         =   "›—⁄ „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdAll 
                  Height          =   285
                  Left            =   12765
                  TabIndex        =   58
                  Top             =   105
                  Width           =   1380
                  _Version        =   786432
                  _ExtentX        =   2434
                  _ExtentY        =   503
                  _StockProps     =   79
                  Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdEmp 
                  Height          =   240
                  Left            =   12045
                  TabIndex        =   59
                  Top             =   480
                  Width           =   2100
                  _Version        =   786432
                  _ExtentX        =   3704
                  _ExtentY        =   423
                  _StockProps     =   79
                  Caption         =   "„ÊŸ› „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbEmployee1 
                  Height          =   315
                  Left            =   6915
                  TabIndex        =   60
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   480
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbBranch1 
                  Height          =   315
                  Left            =   6915
                  TabIndex        =   61
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
                  TabIndex        =   62
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   120
                  Width           =   4395
                  _ExtentX        =   7752
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   765
                  Left            =   120
                  TabIndex        =   63
                  ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                  Top             =   15
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   1349
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
                  ButtonImage     =   "FrmComponentYear.frx":2D97F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin XtremeSuiteControls.CheckBox SelectDept 
                  Height          =   240
                  Left            =   5535
                  TabIndex        =   64
                  Top             =   120
                  Width           =   1260
                  _Version        =   786432
                  _ExtentX        =   2222
                  _ExtentY        =   423
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
                  TabIndex        =   65
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   480
                  Width           =   4395
                  _ExtentX        =   7752
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox SelectProject 
                  Height          =   240
                  Left            =   5535
                  TabIndex        =   66
                  Top             =   480
                  Width           =   1260
                  _Version        =   786432
                  _ExtentX        =   2222
                  _ExtentY        =   423
                  _StockProps     =   79
                  Caption         =   "„‘—Ê⁄ „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin MSDataListLib.DataCombo DcbMofrd 
               Height          =   315
               Left            =   1080
               TabIndex        =   67
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   105
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   210
               Index           =   5
               Left            =   12825
               TabIndex        =   70
               Top             =   105
               Width           =   1260
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ ⁄«„"
               Height          =   210
               Index           =   0
               Left            =   5535
               TabIndex        =   68
               Top             =   120
               Width           =   1260
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   5655
            Left            =   15480
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9975
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
               Height          =   1740
               Left            =   10710
               MaxLength       =   50
               TabIndex        =   42
               Top             =   3570
               Width           =   780
            End
            Begin XtremeSuiteControls.CheckBox BranchSelect 
               Height          =   1380
               Left            =   11610
               TabIndex        =   41
               Top             =   1380
               Width           =   975
               _Version        =   786432
               _ExtentX        =   1720
               _ExtentY        =   2434
               _StockProps     =   79
               Caption         =   "›—⁄ „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton SelectAll 
               Height          =   1725
               Left            =   12765
               TabIndex        =   43
               Top             =   1275
               Width           =   1380
               _Version        =   786432
               _ExtentX        =   2434
               _ExtentY        =   3043
               _StockProps     =   79
               Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton EmpSelect 
               Height          =   1380
               Left            =   12765
               TabIndex        =   44
               Top             =   3570
               Width           =   1380
               _Version        =   786432
               _ExtentX        =   2434
               _ExtentY        =   2434
               _StockProps     =   79
               Caption         =   "„ÊŸ› „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbEmployee 
               Height          =   315
               Left            =   6915
               TabIndex        =   45
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   3570
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBranch 
               Height          =   315
               Left            =   6915
               TabIndex        =   46
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1380
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDepatment 
               Height          =   315
               Left            =   1080
               TabIndex        =   47
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1380
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   4545
               Left            =   120
               TabIndex        =   48
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   765
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   8017
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
               ButtonImage     =   "FrmComponentYear.frx":341E1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin XtremeSuiteControls.CheckBox DeptSelect 
               Height          =   1380
               Left            =   5535
               TabIndex        =   49
               Top             =   1380
               Width           =   1200
               _Version        =   786432
               _ExtentX        =   2117
               _ExtentY        =   2434
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
               TabIndex        =   50
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   3570
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ProjSelect 
               Height          =   1380
               Left            =   5535
               TabIndex        =   51
               Top             =   3570
               Width           =   1200
               _Version        =   786432
               _ExtentX        =   2117
               _ExtentY        =   2434
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
               Height          =   1590
               Index           =   3
               Left            =   12885
               TabIndex        =   52
               Top             =   0
               Width           =   1500
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   5655
            Left            =   15180
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9975
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
               Height          =   6000
               Left            =   0
               TabIndex        =   54
               Top             =   0
               Width           =   14505
               _cx             =   25585
               _cy             =   10583
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   5550
                  Left            =   45
                  TabIndex        =   75
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   14415
                  _cx             =   25426
                  _cy             =   9790
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
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic100 
            Height          =   5655
            Left            =   15780
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9975
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
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   5655
               Left            =   0
               TabIndex        =   73
               Top             =   0
               Width           =   14445
               _cx             =   25479
               _cy             =   9975
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
               BackColorAlternate=   16777088
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
               Rows            =   50
               Cols            =   26
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmComponentYear.frx":3AA43
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
               Begin MSComctlLib.ProgressBar ProgressBar1 
                  Height          =   615
                  Left            =   1200
                  TabIndex        =   74
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   11295
                  _ExtentX        =   19923
                  _ExtentY        =   1085
                  _Version        =   393216
                  Appearance      =   0
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   810
         Left            =   0
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   720
         Width           =   14550
         _cx             =   25665
         _cy             =   1429
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
            Height          =   360
            Left            =   11550
            TabIndex        =   77
            Top             =   285
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   360
            Left            =   9105
            TabIndex        =   78
            Top             =   285
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
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
            Format          =   94044161
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBrnch2 
            Height          =   315
            Left            =   120
            TabIndex        =   79
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
            Top             =   285
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   360
            Left            =   -360
            TabIndex        =   80
            Top             =   285
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
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
            Format          =   94044161
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal XPDtbTransH 
            Height          =   360
            Left            =   7545
            TabIndex        =   81
            Top             =   285
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   330
            Index           =   4
            Left            =   13425
            TabIndex        =   84
            Top             =   285
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   330
            Index           =   1
            Left            =   10185
            TabIndex        =   83
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   330
            Index           =   2
            Left            =   5280
            TabIndex        =   82
            Top             =   300
            Width           =   1005
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
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmComponentYear"
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
 Public LngRow As Long
 Public LngCol As Long
 Dim II As Long
 Sub maxx(Optional ByRef ID As Double = 0)
     
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
    If ID <> 0 Then
   StrSQL = " select max(CompID) as mx from ExpensesSearial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   ID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "ExpensesSearial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("CompID").value = ID
RsDev.update
End If
End Sub
 Function Checked(Optional ID As Double = 0) As Boolean
     Checked = False
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
      Set RsDev = New ADODB.Recordset
    If ID <> 0 Then
   StrSQL = " select * from ExpensesSearial where CompID=" & ID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function
 Public Function cal_value(src As String, Optional EmpID As Double) As Double
    On Error GoTo errortrap
    Dim new_pos As Integer
    Dim i As Integer
    Dim last_pos As Integer
    Dim cuttent_operand As String
    Dim new_str As String
    Dim objScript As Object
    last_pos = 1
    new_str = ""

    For i = 1 To Len(src)

        If Mid(src, i, 1) = "+" Or Mid(src, i, 1) = "-" Or Mid(src, i, 1) = "*" Or Mid(src, i, 1) = "/" Or Mid(src, i, 1) = "=" Then
            new_pos = i
            cuttent_operand = Mid(src, last_pos, new_pos - last_pos)

            If InStr(cuttent_operand, "A") > 0 Then
                cuttent_operand = GetValue(cuttent_operand, EmpID)
                
            End If

            new_str = new_str & cuttent_operand & Mid(src, i, 1)

            If i < Len(src) Then
                last_pos = new_pos + 1
            Else
                GoTo ll
            End If
        End If
 
    Next i

ll:
    new_str = Replace$(new_str, "=", "")

    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"
 
    cal_value = objScript.Eval(new_str)
    cal_value = Round(cal_value, 2)
    Exit Function
errortrap:
    cal_value = 0

End Function
Function GetValue(Optional operand As String, Optional EmpID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
operand = Replace$(operand, "A", "")
sql = "SELECT     [Value]"
sql = sql & " From dbo.EmpSalaryComponent"
sql = sql & " Where (AccountCode = " & val(operand) & ") And (Emp_id = " & EmpID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetValue = IIf(IsNull(Rs3("Value").value), 0, Rs3("Value").value)
Else
GetValue = 0
End If
End Function
Sub RetMofrd(Optional MofrdID As Integer, Optional ByRef Eq_Sys As String, Optional ByRef Eq_Text As String)
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     eq_sys, eq_text"
sql = sql & " From dbo.mofrdat"
sql = sql & " Where (Monthly = 0) And (mofrad_type = " & MofrdID & ")"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Eq_Text = IIf(IsNull(Rs4("eq_text").value), "", Rs4("eq_text").value)
Eq_Sys = IIf(IsNull(Rs4("eq_sys").value), "", Rs4("eq_sys").value)
Else
Eq_Text = ""
Eq_Sys = ""
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
       If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = " SELECT     id, name From mofrad"
    Else
        My_SQL = " SELECT     id, nameE From mofrad"
    End If
      My_SQL = My_SQL & " Where (AllowIntrod = 1)"
      My_SQL = My_SQL & " and id in(  SELECT     mofrad_type From dbo.mofrdat"
      My_SQL = My_SQL & " Where (Monthly = 0))"
    fill_combo Me.DcbMofrd, My_SQL

    conection = "select * from  TblComponentYear  order by  ID "
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

Sub filgrid2(Optional Name As String, Optional account As String, Optional EmpID As Double, Optional Vlaue As Double, Optional FrmDate As Date, Optional todate As Date)
Dim k As Integer
Dim i As Integer
Dim Ind As Integer
Ind = 1
With Grid
k = .Rows - 1
.Rows = .Rows + 1
Do While k < (.Rows - 1)
.TextMatrix(k, .ColIndex("Ser")) = k
'PreCalculteDate k, Ind
.TextMatrix(k, .ColIndex("BranchID")) = val(DcbBrnch2.BoundText)
.TextMatrix(k, .ColIndex("HistoryDate")) = XPDtbTrans.value
.TextMatrix(k, .ColIndex("name")) = Name
.TextMatrix(k, .ColIndex("Messier")) = 1
.TextMatrix(k, .ColIndex("nameE")) = Name
.TextMatrix(k, .ColIndex("TypeExpens")) = 2
.TextMatrix(k, .ColIndex("Account_Code1")) = account
.TextMatrix(k, .ColIndex("EmpID")) = EmpID
.TextMatrix(k, .ColIndex("Account_Code")) = GetAccountEmployee(EmpID)
.TextMatrix(k, .ColIndex("Valu")) = Vlaue
.TextMatrix(k, .ColIndex("FromDate")) = FrmDate
.TextMatrix(k, .ColIndex("ToDate")) = todate
.TextMatrix(k, .ColIndex("Remark2")) = TxtRemarks.Text
.TextMatrix(k, .ColIndex("Distribution")) = 2
k = k + 1
Loop
.AutoSize 0, .Cols - 1, False
End With
End Sub
Function GetAccountEmployee(Optional EmID As Double = 0) As String
Dim Rs7 As ADODB.Recordset
Dim sql As String
If EmID <> 0 Then
sql = "Select Account_Code3 from TblEmployee where Emp_ID =" & EmID & " "
Set Rs7 = New ADODB.Recordset
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetAccountEmployee = IIf(IsNull(Rs7("Account_Code3").value), " ", Rs7("Account_Code3").value)
Else
GetAccountEmployee = ""
End If
End If
End Function
Sub maxx2(Optional ByRef ID As Double = 0, Optional ByRef IDDet As Double = 0)
     
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
    If ID <> 0 Then
   StrSQL = " select max(ID) as mx from ExpensesSearial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   ID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "ExpensesSearial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("ID").value = ID
RsDev.update
End If
    If IDDet <> 0 Then
   StrSQL = " select max(IDDet) as mx from ExpensesSearial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   IDDet = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "ExpensesSearial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("IDDet").value = IDDet
RsDev.update
End If
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRecPeripaid(Optional Name As String, Optional EmpID As Double, Optional Vlue As Double, Optional Fromdate As Date, Optional todate As Date)
  '  On Error GoTo ErrTrap
                     If Me.TxtModFlg.Text = "E" Then
                          StrSQL = "Delete From TblPripaidExpChiled Where AllowID =" & val(TxtSerial1.Text) & ""
                          Cn.Execute StrSQL, , adExecuteNoRecords
                          StrSQL = "Delete From TblPripaidExpensesDet Where AllowID =" & val(TxtSerial1.Text) & ""
                          Cn.Execute StrSQL, , adExecuteNoRecords
                          StrSQL = "Delete From TblPripaidExpenses Where AllowID =" & val(TxtSerial1.Text) & ""
                          Cn.Execute StrSQL, , adExecuteNoRecords
                  End If
    Dim sql As String
    Dim ID As Double
    Dim i As Integer
    Dim Msg As String
    Dim account As String
    Dim Rs2 As ADODB.Recordset
    Dim Rs2Det As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
Dim StrRecID As String
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "„‰ ‘«‘… «·»œ·«  «·„ﬁœ„…"
Else
Msg = "Allowances screen"
End If
account = GetMofrdAccount(val(DcbMofrd.BoundText))
sql = "select * from TblPripaidExpenses where 1=-1 "
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    StrRecID = new_id("TblPripaidExpenses", "ID", "")
    Rs2.AddNew
   Rs2.Fields("ID").value = StrRecID
    Rs2.Fields("AllowID").value = val(TxtSerial1.Text)
    Rs2.Fields("RecordM").value = XPDtbTrans.value
    If Me.TxtRemarks.Text <> "" Then
    Rs2.Fields("Remark").value = Me.TxtRemarks.Text
    Else
    Rs2.Fields("Remark").value = Msg
    End If
    Rs2.Fields("Name").value = DcbMofrd.Text
    Rs2.Fields("NameE").value = DcbMofrd.Text
    Rs2.Fields("BranchID").value = val(Me.DcbBrnch2.BoundText)
    Rs2.Fields("EmpID").value = 0
    Rs2.Fields("Account_Code1").value = account
    Rs2.Fields("Valu").value = Vlue
    Rs2.Fields("HistoryDate").value = XPDtbTrans.value
    Rs2.Fields("FromDate").value = Fromdate
    Rs2.Fields("ToDate").value = todate
    Rs2.Fields("Remark2").value = Me.TxtRemarks.Text
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Rs2.Fields("Messier").value = 1
    Rs2.Fields("TypeExpens").value = 1
    ''/////
    Rs2.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    Rs2.update
   
        With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
      If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 Then
      If val(.TextMatrix(i, .ColIndex("MordValue"))) <> 0 Then
     '  Cn.Execute "Delete from TblComponentYearDet2 where CoYerID2=" & val(.TextMatrix(i, .ColIndex("ID"))) & "  "
       If .TextMatrix(i, .ColIndex("RecDate1")) <> "" Then
       Name = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", "", (.TextMatrix(i, .ColIndex("Name"))))
       EmpID = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", 0, val(.TextMatrix(i, .ColIndex("EmpID"))))
       Vlue = IIf((.TextMatrix(i, .ColIndex("MordValue"))) = "", 0, val(.TextMatrix(i, .ColIndex("MordValue"))))
       account = GetMofrdAccount(IIf((.TextMatrix(i, .ColIndex("MofrdID"))) = "", 0, val(.TextMatrix(i, .ColIndex("MofrdID")))))
       If val(.TextMatrix(i, .ColIndex("TypeMofrd"))) = 1 Then
     '  FiLLRecPeripaid Name, EmpID, Account, Vlue, .TextMatrix(i, .ColIndex("RecDate1")), DateAdd("M", 12, .TextMatrix(i, .ColIndex("RecDate1")))
       filgrid2 Name, account, EmpID, Vlue, .TextMatrix(i, .ColIndex("RecDate1")), DateAdd("M", 12, .TextMatrix(i, .ColIndex("RecDate1")))
       Else
       'FiLLRecPeripaid Name, EmpID, Account, Vlue, .TextMatrix(i, .ColIndex("RecDate1")), DateAdd("M", 6, .TextMatrix(i, .ColIndex("RecDate1")))
       filgrid2 Name, account, EmpID, Vlue / 2, .TextMatrix(i, .ColIndex("RecDate1")), DateAdd("M", 6, .TextMatrix(i, .ColIndex("RecDate1")))
       
       If .TextMatrix(i, .ColIndex("RecDate2")) <> "" Then
      ' FiLLRecPeripaid Name, EmpID, Account, Vlue, .TextMatrix(i, .ColIndex("RecDate2")), DateAdd("M", 6, .TextMatrix(i, .ColIndex("RecDate2")))
       filgrid2 Name, account, EmpID, Vlue / 2, .TextMatrix(i, .ColIndex("RecDate2")), DateAdd("M", 6, .TextMatrix(i, .ColIndex("RecDate2")))
       End If
       End If
       End If

        End If
        End If
     Next i
    End With
    
    
    Dim ID1 As Double
    
    Set Rs2Det = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblPripaidExpensesDet Where (1 = -1)"
    Rs2Det.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim str2 As String
    With Grid
       For i = .FixedRows To .Rows - 1
     If val(.TextMatrix(i, .ColIndex("BranchID"))) <> 0 Then
        ID1 = 0
        ID1 = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("id"))), 0, .TextMatrix(i, .ColIndex("id")))
        
     If Me.Checked2(ID1, 0) = True Then
        Else
       ID1 = 1
        maxx2 ID1, 0
    End If
            .TextMatrix(i, .ColIndex("id")) = ID1
     If ChekExpens(ID1) = False Then
       Rs2Det.AddNew
                Rs2Det("ID").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("id"))), 0, .TextMatrix(i, .ColIndex("id")))
                Rs2Det("PaidExID").value = StrRecID
                Rs2Det("Name").value = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", Null, .TextMatrix(i, .ColIndex("Name")))
                Rs2Det("Messier").value = 1
                Rs2Det("AllowID").value = val(TxtSerial1.Text)
                Rs2Det("NameE").value = IIf((.TextMatrix(i, .ColIndex("NameE"))) = "", Null, .TextMatrix(i, .ColIndex("NameE")))
                Rs2Det("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                Rs2Det("TypeExpens").value = IIf((.TextMatrix(i, .ColIndex("TypeExpens"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeExpens"))))
                Rs2Det("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                Rs2Det("Account_Code").value = IIf((.TextMatrix(i, .ColIndex("Account_Code"))) = "", Null, .TextMatrix(i, .ColIndex("Account_Code")))
                Rs2Det("Account_Code1").value = IIf((.TextMatrix(i, .ColIndex("Account_Code1"))) = "", Null, .TextMatrix(i, .ColIndex("Account_Code1")))
                Rs2Det("HistoryDate").value = IIf((.TextMatrix(i, .ColIndex("HistoryDate"))) = "", Null, .TextMatrix(i, .ColIndex("HistoryDate")))
                Rs2Det("FromDate").value = IIf((.TextMatrix(i, .ColIndex("FromDate"))) = "", Null, .TextMatrix(i, .ColIndex("FromDate")))
                Rs2Det("ToDate").value = IIf((.TextMatrix(i, .ColIndex("ToDate"))) = "", Null, .TextMatrix(i, .ColIndex("ToDate")))
                Rs2Det("Valu").value = IIf((.TextMatrix(i, .ColIndex("Valu"))) = "", Null, .TextMatrix(i, .ColIndex("Valu")))
                Rs2Det("Remark2").value = IIf((.TextMatrix(i, .ColIndex("Remark2"))) = "", Null, .TextMatrix(i, .ColIndex("Remark2")))
                Rs2Det("Distribution").value = IIf((.TextMatrix(i, .ColIndex("Distribution"))) = "", Null, val(.TextMatrix(i, .ColIndex("Distribution"))))
                 If Grid.TextMatrix(i, Grid.ColIndex("StrDistribution")) = "" Then
                   RetrStrEstam str2, i
                    .TextMatrix(i, .ColIndex("StrDistribution")) = str2
                   End If
                Rs2Det("StrDistribution").value = IIf((.TextMatrix(i, .ColIndex("StrDistribution"))) = "", Null, .TextMatrix(i, .ColIndex("StrDistribution")))
                 
                   

      
      Rs2Det.update
      saveDetails2 StrRecID, i, Rs2Det("id").value
      End If
      End If
     Next i
    End With

   End Sub
   Sub saveDetails2(Optional StrID As String, Optional i As Integer = 0, Optional PaidExIDDet As Integer = 0)
Dim RsDetails11 As ADODB.Recordset
 Dim IDDet As Double
Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim j As Integer
Dim st As String
Dim nElements As Integer
Dim k, m As Integer
Dim Diff As Integer
If PaidExIDDet <> 0 Then
Set RsDetails11 = New ADODB.Recordset
If Me.TxtModFlg.Text = "R" Then
StrSQL = "delete From TblPripaidExpChiled  where  PaidExID =" & PaidExIDDet
                   Cn.Execute StrSQL, , adExecuteNoRecords
End If
    StrSQL = "SELECT  *  from dbo.TblPripaidExpChiled Where (1 = -1)"
   RsDetails11.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

k = 0
     If Grid.TextMatrix(i, Grid.ColIndex("StrDistribution")) <> "" Then
          st = Grid.TextMatrix(i, Grid.ColIndex("StrDistribution"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
        
         For j = 0 To nElements - 1
         With Grid
     '   Diff = DateDiff("M", .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate")))
         End With
          RsDetails11.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         Diff = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
         Diff = Diff / 3
         m = 0
        For k = 0 To Diff - 1
        RsDetails11("PaidExID").value = StrID
        RsDetails11("AllowID").value = val(TxtSerial1.Text)
         RsDetails11("PaidExIDDet").value = PaidExIDDet
         RsDetails11("RecDate").value = astrSplit2tems2(m)
         m = m + 1
         RsDetails11("Valu").value = val(astrSplit2tems2(m))
         m = m + 1
         RsDetails11("Remark").value = astrSplit2tems2(m)
         m = m + 1
    
       RsDetails11("ID").value = val(astrSplit2tems2(m))
        m = m + 1
         RsDetails11.update
      Next k
       Next j
          End If
End If

End Sub
      Sub RetrStrEstam(Optional ByRef str1 As String, Optional Row As Integer)
Dim str As String
Dim Diff As Integer
Dim StrtDate As Date
Dim cunt As Integer
Dim IDDet As Double
Dim SumVal As Double
Dim LastQst As Double
cunt = 1
SumVal = 0
  With Grid
If .TextMatrix(Row, .ColIndex("FromDate")) <> "" And .TextMatrix(Row, .ColIndex("ToDate")) <> "" Then
Diff = DateDiff("m", .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate"))) + 1
StrtDate = .TextMatrix(Row, .ColIndex("FromDate"))
Do While cunt <= Diff
  str = str & StrtDate & "#"
  If Diff <> cunt Then
  str = str & Round(val(.TextMatrix(Row, .ColIndex("Valu"))) / Diff, 2) & "#"
  End If
  SumVal = SumVal + Round(val(.TextMatrix(Row, .ColIndex("Valu"))) / Diff, 2)
  If Diff = cunt Then
  LastQst = SumVal - val(.TextMatrix(Row, .ColIndex("Valu")))
  str = str & Round(val(.TextMatrix(Row, .ColIndex("Valu"))) / Diff, 2) - LastQst & "#"
  End If
    IDDet = 1
       maxx2 0, IDDet
  str = str & " " & "#"
  str = str & IDDet & "#"
  
  StrtDate = DateAdd("m", 1, StrtDate)
   str = str & Trim("@")
  str = str & CHR(13)
  cunt = cunt + 1
Loop

  str1 = Trim(str)
End If
  End With
End Sub
   Function ChekExpens(Optional ID As Double) As Boolean
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "SELECT     ID, Paye"
sql = sql & " FROM         dbo.TblPripaidExpensesDet"
sql = sql & " where paye=1 and id=" & ID & ""
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekExpens = True
Else
ChekExpens = False
End If
End Function
   Function Checked2(Optional ID As Double = 0, Optional IDDet As Double = 0) As Boolean
     Checked2 = False
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
      Set RsDev = New ADODB.Recordset
    If ID <> 0 Then
   StrSQL = " select * from ExpensesSearial where ID=" & ID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked2 = True
Else
Checked2 = False
End If
End If
    If IDDet <> 0 Then
  StrSQL = " select * from ExpensesSearial where IDDet=" & IDDet & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
Checked2 = True
Else
Checked2 = False
End If
End If
End Function
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
  Dim sql As String
  Dim account As String
  Dim Name As String
  Dim Vlue As Double
  Dim EmpID As Double
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                  StrSQL = "Delete From TblComponentYearDet Where CoYerID=" & val(Me.TxtSerial1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
           End If
   RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
   RsSavRec.Fields("BrnchID").value = val(Me.DcbBrnch2.BoundText)
   RsSavRec.Fields("MofrdID").value = val(Me.DcbMofrd.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("DeptID").value = val(Me.DcpDept1.BoundText)
   RsSavRec.Fields("ProjectID").value = val(Me.DcbProject1.BoundText)
   RsSavRec.Fields("BrnchID1").value = val(Me.DcbBranch1.BoundText)
   RsSavRec.Fields("EmpID").value = val(Me.DcbEmployee1.BoundText)
   RsSavRec.Fields("Remarks").value = TxtRemarks.Text
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
    StrSQL = "SELECT  *  from TblComponentYearDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
      If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 Then
      If val(.TextMatrix(i, .ColIndex("MordValue"))) <> 0 Then
              ID = 0
        ID = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("ID"))), 0, .TextMatrix(i, .ColIndex("ID")))
        
          If Me.Checked(ID) = True Then
        Else
       ID = 1
        maxx ID
        End If
              .TextMatrix(i, .ColIndex("ID")) = ID
       RsDevsub.AddNew
                RsDevsub("ID").value = val(.TextMatrix(i, .ColIndex("ID")))
                RsDevsub("CoYerID").value = val(Me.TxtSerial1.Text)
                RsDevsub("MofrdID").value = IIf((.TextMatrix(i, .ColIndex("MofrdID"))) = "", Null, val(.TextMatrix(i, .ColIndex("MofrdID"))))
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(i, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DeptID"))))
                RsDevsub("ProjectID").value = IIf((.TextMatrix(i, .ColIndex("ProjID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ProjID"))))
                RsDevsub("BrnchID1").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                RsDevsub("RecDate1").value = IIf((.TextMatrix(i, .ColIndex("RecDate1"))) = "", Null, (.TextMatrix(i, .ColIndex("RecDate1"))))
                RsDevsub("RecDate2").value = IIf((.TextMatrix(i, .ColIndex("RecDate2"))) = "", Null, Trim(.TextMatrix(i, .ColIndex("RecDate2"))))
                RsDevsub("TypeMofrd").value = IIf((.TextMatrix(i, .ColIndex("TypeMofrd"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeMofrd"))))
                RsDevsub("MordValue").value = IIf((.TextMatrix(i, .ColIndex("MordValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("MordValue"))))
                RsDevsub("StFunction").value = IIf((.TextMatrix(i, .ColIndex("StFunction"))) = "", Null, (.TextMatrix(i, .ColIndex("StFunction"))))
                RsDevsub("RecDate1H").value = IIf((.TextMatrix(i, .ColIndex("RecDate1H"))) = "", Null, (.TextMatrix(i, .ColIndex("RecDate1H"))))
                RsDevsub("RecDate2H").value = IIf((.TextMatrix(i, .ColIndex("RecDate2H"))) = "", Null, (.TextMatrix(i, .ColIndex("RecDate2H"))))
       RsDevsub.update
       Cn.Execute "Delete from TblComponentYearDet2 where CoYerID2=" & val(.TextMatrix(i, .ColIndex("ID"))) & "  "

       If .TextMatrix(i, .ColIndex("RecDate1")) <> "" Then
       If .TextMatrix(i, .ColIndex("RecDate2")) = "" Then
                      SaveDetals val(.TextMatrix(i, .ColIndex("ID"))), 12, val(.TextMatrix(i, .ColIndex("MordValue"))), .TextMatrix(i, .ColIndex("RecDate1"))
        Else
        SaveDetals val(.TextMatrix(i, .ColIndex("ID"))), 6, val(.TextMatrix(i, .ColIndex("MordValue"))) / 2, .TextMatrix(i, .ColIndex("RecDate1"))
        SaveDetals val(.TextMatrix(i, .ColIndex("ID"))), 6, val(.TextMatrix(i, .ColIndex("MordValue"))) / 2, .TextMatrix(i, .ColIndex("RecDate2"))
        End If
        End If
        End If
        End If
     Next i
    End With
     ' If .TextMatrix(i, .ColIndex("RecDate1")) <> "" Then
       'Name = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", "", (.TextMatrix(i, .ColIndex("Name"))))
       'EmpID = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", 0, val(.TextMatrix(i, .ColIndex("EmpID"))))
       'Vlue = IIf((.TextMatrix(i, .ColIndex("MordValue"))) = "", 0, val(.TextMatrix(i, .ColIndex("MordValue"))))
       'account = GetMofrdAccount(IIf((.TextMatrix(i, .ColIndex("MofrdID"))) = "", 0, val(.TextMatrix(i, .ColIndex("MofrdID")))))
       'If val(.TextMatrix(i, .ColIndex("TypeMofrd"))) = 1 Then
       FiLLRecPeripaid
       'Else
       'FiLLRecPeripaid Name, EmpID, account, Vlue, .TextMatrix(i, .ColIndex("RecDate1")), DateAdd("M", 6, .TextMatrix(i, .ColIndex("RecDate1")))
       'If .TextMatrix(i, .ColIndex("RecDate2")) <> "" Then
       'FiLLRecPeripaid Name, EmpID, account, Vlue, .TextMatrix(i, .ColIndex("RecDate2")), DateAdd("M", 6, .TextMatrix(i, .ColIndex("RecDate2")))
      ' End If
      ' End If
     '  End If
       
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
Function GetMofrdAccount(Optional ID As Double) As String
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = "SELECT     Account_Code "
sql = sql & " From dbo.MOFRAD"
sql = sql & " WHERE     (id = " & ID & ")"
Rs2.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
GetMofrdAccount = IIf(IsNull(Rs2("Account_Code").value), "", Rs2("Account_Code").value)
Else
GetMofrdAccount = ""
End If
End Function
Sub SaveDetals(Optional CoYerID2 As Double, Optional NoMont As Integer, Optional MordValuen As Double, Optional RecDate As Date)
Dim Rs1 As ADODB.Recordset
Dim StrSQL As String
Dim ValNo As Double
Dim dif As Double
ValNo = Round(MordValuen / NoMont, 2)
dif = (ValNo * NoMont)
dif = MordValuen - dif
    Set Rs1 = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblComponentYearDet2 Where (1 = -1)"
    Rs1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    DTPicker1.value = RecDate
    For i = 1 To NoMont
    Rs1.AddNew
    Rs1("CoYerID2").value = CoYerID2
    Rs1("CoYerID").value = val(TxtSerial1.Text)
    Rs1("RecDate1").value = DTPicker1.value
    Rs1("RecDate1H").value = ToHijriDate(DTPicker1.value)
    If i = NoMont Then
     Rs1("MordValue").value = ValNo + dif
    Else
    Rs1("MordValue").value = ValNo
   End If
    DTPicker1.value = DateAdd("M", 1, DTPicker1.value)
    Rs1.update
    Next i
End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
     SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
    Dim Shifttime As Date
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Me.DcbBrnch2.BoundText = IIf(IsNull(RsSavRec.Fields("BrnchID").value), "", RsSavRec.Fields("BrnchID").value)
    Me.DcbMofrd.BoundText = IIf(IsNull(RsSavRec.Fields("MofrdID").value), "", RsSavRec.Fields("MofrdID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcpDept1.BoundText = IIf(IsNull(RsSavRec.Fields("DeptID").value), "", RsSavRec.Fields("DeptID").value)
    Me.DcbProject1.BoundText = IIf(IsNull(RsSavRec.Fields("ProjectID").value), "", RsSavRec.Fields("ProjectID").value)
    Me.DcbBranch1.BoundText = IIf(IsNull(RsSavRec.Fields("BrnchID1").value), "", RsSavRec.Fields("BrnchID1").value)
    Me.DcbEmployee1.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    Me.TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
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
    Dim i As Integer

    '---------------------- check if data Vaclete -----------------------
    '------------------------------ check if Empcode exist ----------------------
'   StrVacName = IsRecExist("TblEmploymentModel", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtSerial1.text) & "'")
  ' If StrVacName <> "" Then
 '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·«”„ „‰ ﬁ»·"
  '     MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
  '    TxtVacName.SetFocus
 '     Exit Sub
'   End If
    ' -------------------------------------- txtmodflg type -------------------
If val(Me.DcbBrnch2.BoundText) = 0 Or Me.DcbBrnch2.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
Else
MsgBox "Please Select Branch"
End If
DcbBrnch2.SetFocus
Exit Sub
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
    StrRecID = new_id("TblComponentYear", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.Text = StrRecID
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
sql = " SELECT     dbo.TblComponentYearDet.ID, dbo.TblComponentYearDet.CoYerID, dbo.TblComponentYearDet.MofrdID, dbo.TblComponentYearDet.TypeMofrd, "
sql = sql & "                       dbo.TblComponentYearDet.StFunction, dbo.TblComponentYearDet.RecDate1, dbo.TblComponentYearDet.RecDate2, dbo.TblComponentYearDet.MordValue,"
sql = sql & "                       dbo.TblComponentYearDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblComponentYearDet.DeptID,"
sql = sql & "                       dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblComponentYearDet.BrnchID1, dbo.TblBranchesData.branch_name,"
sql = sql & "                       dbo.TblBranchesData.branch_namee, dbo.TblComponentYearDet.ProjectID, dbo.projects.Project_name, dbo.projects.Project_nameE,"
sql = sql & "                       dbo.TblComponentYearDet.RecDate1H , dbo.TblComponentYearDet.RecDate2H, dbo.MOFRAD.Name, dbo.MOFRAD.NameE"
sql = sql & "  FROM         dbo.TblComponentYearDet INNER JOIN"
sql = sql & "                       dbo.mofrad ON dbo.TblComponentYearDet.MofrdID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & "                       dbo.projects ON dbo.TblComponentYearDet.ProjectID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.TblComponentYearDet.BrnchID1 = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmpDepartments ON dbo.TblComponentYearDet.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee ON dbo.TblComponentYearDet.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & " Where (dbo.TblComponentYearDet.CoYerID = " & val(TxtSerial1.Text) & ")"
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs1("ID").value), 0, Rs1("ID").value)
                   .TextMatrix(i, .ColIndex("TypeMofrd")) = IIf(IsNull(Rs1("TypeMofrd").value), 1, Rs1("TypeMofrd").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs1("DeptID").value), 0, Rs1("DeptID").value)
                   .TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(Rs1("ProjectID").value), 0, Rs1("ProjectID").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BrnchID1").value), 0, Rs1("BrnchID1").value)
                   .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(Rs1("MofrdID").value), 0, Rs1("MofrdID").value)
                   .TextMatrix(i, .ColIndex("MordValue")) = IIf(IsNull(Rs1("MordValue").value), 0, Rs1("MordValue").value)
                   .TextMatrix(i, .ColIndex("StFunction")) = IIf(IsNull(Rs1("StFunction").value), "", Rs1("StFunction").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("RecDate1")) = IIf(IsNull(Rs1("RecDate1").value), "", Rs1("RecDate1").value)
                   .TextMatrix(i, .ColIndex("RecDate2")) = IIf(IsNull(Rs1("RecDate2").value), "", Rs1("RecDate2").value)
                   .TextMatrix(i, .ColIndex("RecDate1H")) = IIf(IsNull(Rs1("RecDate1H").value), "", Rs1("RecDate1H").value)
                   .TextMatrix(i, .ColIndex("RecDate2H")) = IIf(IsNull(Rs1("RecDate2H").value), "", Rs1("RecDate2H").value)
                   
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentName").value), "", Rs1("DepartmentName").value)
                   .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_name").value), "", Rs1("Project_name").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_nameE").value), "", Rs1("Project_nameE").value)
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentNamee").value), "", Rs1("DepartmentNamee").value)
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   End If
      
                   Rs1.MoveNext
             Next i
End With
        
        Exit Sub
ErrTrap:
    End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 ' On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Eq_Sys As String
    Dim Eq_Text As String
    Dim LngRow As Long
    With GridInstallments
        Select Case .ColKey(Col)
            Case "Name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MofrdID"), False, True)
                .TextMatrix(Row, .ColIndex("MofrdID")) = StrAccountCode
                RetMofrd val(.TextMatrix(Row, .ColIndex("MofrdID"))), Eq_Sys, Eq_Text
                .TextMatrix(Row, .ColIndex("StFunction")) = Eq_Sys
                .TextMatrix(Row, .ColIndex("MordValue")) = cal_value(Eq_Text, val(.TextMatrix(Row, .ColIndex("EmpID"))))
            Case "TypeMofrd"
            If val(.TextMatrix(Row, .ColIndex("TypeMofrd"))) = 1 Then
            .TextMatrix(Row, .ColIndex("RecDate2")) = ""
            End If
     End Select
    End With
End Sub
Function ChekApprove(Optional ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblComponentYearDet where (FlgSel=1 or FlgSel2=1) and ID=" & ID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ChekApprove = True
Else
ChekApprove = False
End If
End Function
Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GridInstallments
Select Case .ColKey(Col)
Case "Emp_Name"
Cancel = True
Case "FullCode"
Cancel = True
Case "StFunction"
Cancel = True
Case "MordValue"
Cancel = True
Case "TypeMofrd"
.ComboList = ""
Case "RecDate1"
.ComboList = ""
Case "RecDate1H"
.ComboList = ""
Case "RecDate2"
If val(.TextMatrix(Row, .ColIndex("TypeMofrd"))) = 2 Then
.ComboList = ""
Else
Cancel = True
End If
Case "RecDate2H"
If val(.TextMatrix(Row, .ColIndex("TypeMofrd"))) = 2 Then
.ComboList = ""
Else
Cancel = True
End If
End Select
End With
End Sub

Private Sub GridInstallments_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Me.TxtModFlg.Text <> "R" Then
With GridInstallments
Select Case .ColKey(Col)
      Case "RecDate1"
        LngRow = Row
        LngCol = Col
         FrmDateOpProject.Index = 22
        Load FrmDateOpProject
        FrmDateOpProject.Index = 22
        FrmDateOpProject.show vbModal
      Case "RecDate2"
        LngRow = Row
        LngCol = Col
         FrmDateOpProject.Index = 23
        Load FrmDateOpProject
        FrmDateOpProject.Index = 23
        FrmDateOpProject.show vbModal
           Case "RecDate1H"
        LngRow = Row
        LngCol = Col
         FrmDateOpProject.Index = 24
        Load FrmDateOpProject
        FrmDateOpProject.Index = 24
        FrmDateOpProject.show vbModal
        Case "RecDate2H"
        LngRow = Row
        LngCol = Col
         FrmDateOpProject.Index = 25
        Load FrmDateOpProject
        FrmDateOpProject.Index = 25
        FrmDateOpProject.show vbModal
End Select
End With
End If

End Sub

Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With GridInstallments

        Select Case .ColKey(Col)
            Case "Name"
                StrSQL = " select * from mofrdat "
                StrSQL = StrSQL & " where Monthly=0  "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "mofrad_name", "mofrad_code")
                Else
                    StrComboList = .BuildComboList(rs, "mofrad_namee", "mofrad_code")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
           Case "RecDate1"
                .ColComboList(.ColIndex("RecDate1")) = "..."
           Case "RecDate2"
              .ColComboList(.ColIndex("RecDate2")) = "..."
          Case "RecDate1H"
                .ColComboList(.ColIndex("RecDate1H")) = "..."
           Case "RecDate2H"
              .ColComboList(.ColIndex("RecDate2H")) = "..."
        End Select

    End With
End Sub

Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then


If val(Me.DcbMofrd.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„›—œ"
Else
MsgBox "Please Select Component"
End If
DcbMofrd.SetFocus
Exit Sub
End If


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
End If
End Sub
Sub filgrid1()
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim i As Integer
Dim k As Integer
Dim sql As String
Dim Eq_Sys As String
Dim Eq_Text As String
Dim EmpID As Double
sql = "SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, "
sql = sql & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName,"
sql = sql & "                      dbo.TblEmpDepartments.DepartmentNamee , dbo.TblEmployee.project_id, dbo.Projects.Project_name, dbo.Projects.Project_nameE"
sql = sql & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & " WHERE     (1<>-1) "

If val(Me.DcbProject1.BoundText) <> 0 And Me.SelectProject.value = vbChecked Then
sql = sql & " and dbo.TblEmployee.project_id  =" & val(DcbProject1.BoundText) & " "
End If
If val(DcbBranch1.BoundText) <> 0 And Me.SelectBranch.value = vbChecked Then
sql = sql & " and dbo.TblEmployee.BranchId  =" & val(DcbBranch1.BoundText) & " "
End If
If val(DcpDept1.BoundText) <> 0 And Me.SelectDept.value = vbChecked Then
sql = sql & " and dbo.TblEmployee.DepartmentID  =" & val(DcpDept1.BoundText) & " "
End If
If val(DcbEmployee1.BoundText) <> 0 And RdEmp.value = True Then
sql = sql & " and dbo.TblEmployee.Emp_ID =" & val(DcbEmployee1.BoundText) & " "
End If
sql = sql & " ORDER BY dbo.TblEmployee.Emp_ID"
 Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs8.RecordCount > 0 Then
With GridInstallments
k = .Rows

Rs8.MoveFirst
.Rows = .Rows + Rs8.RecordCount
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs8("Emp_ID").value), 0, Rs8("Emp_ID").value)
.TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DepartmentID").value), 0, Rs8("DepartmentID").value)
.TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(Rs8("project_id").value), 0, Rs8("project_id").value)
.TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
.TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
RetMofrd val(DcbMofrd.BoundText), Eq_Sys, Eq_Text
'.TextMatrix(i, .ColIndex("Eq_Text")) = Eq_Text
.TextMatrix(i, .ColIndex("StFunction")) = Eq_Sys
.TextMatrix(i, .ColIndex("MordValue")) = cal_value(Eq_Text, val(.TextMatrix(i, .ColIndex("EmpID"))))
.TextMatrix(i, .ColIndex("MofrdID")) = val(DcbMofrd.BoundText)
.TextMatrix(i, .ColIndex("Name")) = DcbMofrd.Text
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentName").value), "", Rs8("DepartmentName").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_name").value), "", Rs8("Project_name").value)
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
Else
.TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentNamee").value), "", Rs8("DepartmentNamee").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_nameE").value), "", Rs8("Project_nameE").value)
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)

End If
Rs8.MoveNext
Next i
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
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
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
    Dim x As Integer
    Dim i As Integer
    Dim ID As Double
   With Me.GridInstallments
    For i = 1 To .Rows - 1
   If ChekApprove(val(.TextMatrix(i, .ColIndex("ID")))) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ–›  „ «⁄ „«œ »⁄÷ «Ê ﬂ· «·»œ·« "
   Else
   MsgBox "Can not be delete .Linked by approve"
   End If
   Exit Sub
   End If
    Next i
    End With
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
    StrSQL = "Delete From TblComponentYearDet Where CoYerID=" & val(Me.TxtSerial1.Text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
''/////
                         StrSQL = "Delete From TblPripaidExpChiled Where AllowID =" & val(TxtSerial1.Text) & ""
                          Cn.Execute StrSQL, , adExecuteNoRecords
                          StrSQL = "Delete From TblPripaidExpensesDet Where AllowID =" & val(TxtSerial1.Text) & ""
                          Cn.Execute StrSQL, , adExecuteNoRecords
                          StrSQL = "Delete From TblPripaidExpenses Where AllowID =" & val(TxtSerial1.Text) & ""
                          Cn.Execute StrSQL, , adExecuteNoRecords
      ''/////
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
                x = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
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
    Dim i As Integer
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        With Me.GridInstallments
    For i = 1 To .Rows - 1
   If ChekApprove(val(.TextMatrix(i, .ColIndex("ID")))) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·  „ «⁄ „«œ »⁄÷ «Ê ﬂ· «·»œ·« "
   Else
   MsgBox "Can not be edited .Linked by approve"
   End If
   Exit Sub
   End If
    Next i
    End With
        TxtModFlg = "E"
       ' GridInstallments.Rows = GridInstallments.Rows + 1
        Me.Grid.Clear flexClearScrollable, flexClearEverything
        Grid.Rows = 2
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
    SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
         Me.Grid.Clear flexClearScrollable, flexClearEverything
     Grid.Rows = 2
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
    
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    
    lbl(4).Caption = "No"
    lbl(1).Caption = "Date"
    lbl(2).Caption = "Branch"
    lbl(5).Caption = "Remarks"
    lbl(0).Caption = "Component"
    Label1(2).Caption = "Advanced Allowance Plan"
    Cmd(3).Caption = "Delete"
    Cmd(4).Caption = "Delete All"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
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
    C1Tab1.Caption = "Basic Data"

    With Me.GridInstallments
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("FullCode")) = "Employee Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
        .TextMatrix(0, .ColIndex("Name")) = "Component"
        .TextMatrix(0, .ColIndex("StFunction")) = "Proceure"
        .TextMatrix(0, .ColIndex("MordValue")) = "Value"
        .TextMatrix(0, .ColIndex("RecDate1")) = "First Payment Date"
        .TextMatrix(0, .ColIndex("RecDate2")) = "Second Payment Date"
        .TextMatrix(0, .ColIndex("TypeMofrd")) = "Type Payment"
        .TextMatrix(0, .ColIndex("RecDate1H")) = "First Payment Date H"
        .TextMatrix(0, .ColIndex("RecDate2H")) = "Second Payment Date H"
  End With
  
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblComponentYear"
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
Dim i As Integer
   With Me.GridInstallments
    For i = 1 To .Rows - 1
   If ChekApprove(val(.TextMatrix(i, .ColIndex("ID")))) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ–›  „ «⁄ „«œ »⁄÷ «Ê ﬂ· «·»œ·« "
   Else
   MsgBox "Can not be delete .Linked by approve"
   End If
   Exit Sub
   End If
    Next i
    End With
 GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
End Sub
Private Sub RemoveGridRow()
    With Me.GridInstallments
        If .Row <= 0 Then Exit Sub
   If ChekApprove(val(.TextMatrix(.Row, .ColIndex("ID")))) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ–›  „ «⁄ „«œ «·»œ·"
   Else
   MsgBox "Can not be delete .Linked by approve"
   End If
   Exit Sub
   Else
  .RemoveItem .Row
  End If
    End With
End Sub

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
         XPDtbTransH.value = ToHijriDate(XPDtbTrans.value)
End If
End Sub

Private Sub XPDtbTransH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 XPDtbTrans.value = ToGregorianDate(XPDtbTransH.value)
End If
End Sub
