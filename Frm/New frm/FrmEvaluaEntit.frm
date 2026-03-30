VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmEvaluaEntit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmEvaluaEntit.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   14550
   WindowState     =   2  'Maximized
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
      ItemData        =   "FrmEvaluaEntit.frx":6852
      Left            =   15480
      List            =   "FrmEvaluaEntit.frx":6862
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
            Picture         =   "FrmEvaluaEntit.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluaEntit.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluaEntit.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluaEntit.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluaEntit.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluaEntit.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluaEntit.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEvaluaEntit.frx":83B1
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
      ButtonImage     =   "FrmEvaluaEntit.frx":874B
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
      ButtonImage     =   "FrmEvaluaEntit.frx":EFAD
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
      ButtonImage     =   "FrmEvaluaEntit.frx":1580F
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   825
         Left            =   0
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   6720
         Width           =   14610
         _cx             =   25770
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
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Height          =   555
            Left            =   405
            TabIndex        =   70
            Top             =   60
            Width           =   6105
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   3150
               TabIndex        =   77
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   4740
               TabIndex        =   76
               Top             =   240
               Width           =   1140
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   1575
               TabIndex        =   75
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   210
               Left            =   240
               TabIndex        =   74
               Top             =   240
               Width           =   1155
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   9840
            TabIndex        =   72
            Top             =   135
            Width           =   3060
            _ExtentX        =   5398
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
            Height          =   255
            Index           =   8
            Left            =   13080
            TabIndex        =   73
            Top             =   135
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   150
            Index           =   12
            Left            =   14850
            TabIndex        =   71
            Top             =   120
            Width           =   930
         End
      End
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
            ButtonImage     =   "FrmEvaluaEntit.frx":15BA9
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
            ButtonImage     =   "FrmEvaluaEntit.frx":15F43
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
            ButtonImage     =   "FrmEvaluaEntit.frx":162DD
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
            ButtonImage     =   "FrmEvaluaEntit.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«” Õﬁ«ﬁ «· ﬁÌÌ„"
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
            Picture         =   "FrmEvaluaEntit.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5580
         Left            =   0
         TabIndex        =   21
         Top             =   1515
         Width           =   14535
         _cx             =   25638
         _cy             =   9842
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
            Height          =   5160
            Left            =   45
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9102
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
               Height          =   270
               Left            =   210
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   59
               Top             =   105
               Width           =   6615
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   3930
               Left            =   0
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1185
               Width           =   14445
               _cx             =   25479
               _cy             =   6932
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
                  Height          =   255
                  Index           =   3
                  Left            =   12735
                  TabIndex        =   26
                  Top             =   3540
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   450
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
                  ButtonImage     =   "FrmEvaluaEntit.frx":17E16
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   255
                  Index           =   4
                  Left            =   11295
                  TabIndex        =   27
                  Top             =   3540
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   450
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
                  ButtonImage     =   "FrmEvaluaEntit.frx":183B0
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                  Height          =   2970
                  Left            =   0
                  TabIndex        =   61
                  Top             =   390
                  Width           =   14340
                  _cx             =   25294
                  _cy             =   5239
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
                  Cols            =   18
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEvaluaEntit.frx":1894A
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
                  TabIndex        =   67
                  Top             =   3570
                  Width           =   3885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   180
                  Index           =   7
                  Left            =   4440
                  TabIndex        =   66
                  Top             =   3555
                  Width           =   1005
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   765
               Left            =   0
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   465
               Width           =   14445
               _cx             =   25479
               _cy             =   1349
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
                  TabIndex        =   47
                  Top             =   435
                  Width           =   750
               End
               Begin XtremeSuiteControls.CheckBox SelectBranch 
                  Height          =   225
                  Left            =   11595
                  TabIndex        =   46
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
                  Height          =   285
                  Left            =   12735
                  TabIndex        =   48
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
                  Height          =   225
                  Left            =   11895
                  TabIndex        =   49
                  Top             =   435
                  Width           =   2220
                  _Version        =   786432
                  _ExtentX        =   3916
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
                  TabIndex        =   50
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   435
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
                  TabIndex        =   51
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
                  TabIndex        =   52
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
                  Height          =   705
                  Left            =   120
                  TabIndex        =   53
                  ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                  Top             =   15
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   1244
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
                  ButtonImage     =   "FrmEvaluaEntit.frx":18BEA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin XtremeSuiteControls.CheckBox SelectDept 
                  Height          =   225
                  Left            =   5550
                  TabIndex        =   54
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
                  TabIndex        =   55
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   435
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
                  TabIndex        =   56
                  Top             =   435
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
               Height          =   270
               Left            =   11520
               TabIndex        =   62
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   476
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
               Format          =   93913089
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker ToDate 
               Height          =   270
               Left            =   8040
               TabIndex        =   64
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   476
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
               Format          =   93913089
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï  «—ÌŒ"
               Height          =   240
               Index           =   6
               Left            =   9960
               TabIndex        =   65
               Top             =   135
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰  «—ÌŒ"
               Height          =   240
               Index           =   0
               Left            =   13200
               TabIndex        =   63
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
               TabIndex        =   60
               Top             =   105
               Width           =   1230
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   5160
            Left            =   15480
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9102
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
               Height          =   1590
               Left            =   10740
               MaxLength       =   50
               TabIndex        =   32
               Top             =   3255
               Width           =   765
            End
            Begin XtremeSuiteControls.CheckBox BranchSelect 
               Height          =   1260
               Left            =   11595
               TabIndex        =   31
               Top             =   1260
               Width           =   1005
               _Version        =   786432
               _ExtentX        =   1773
               _ExtentY        =   2222
               _StockProps     =   79
               Caption         =   "›—⁄ „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton SelectAll 
               Height          =   1560
               Left            =   12735
               TabIndex        =   33
               Top             =   1170
               Width           =   1395
               _Version        =   786432
               _ExtentX        =   2461
               _ExtentY        =   2752
               _StockProps     =   79
               Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton EmpSelect 
               Height          =   1260
               Left            =   12735
               TabIndex        =   34
               Top             =   3255
               Width           =   1395
               _Version        =   786432
               _ExtentX        =   2461
               _ExtentY        =   2222
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
               TabIndex        =   35
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   3255
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
               TabIndex        =   36
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1260
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
               TabIndex        =   37
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1260
               Width           =   4380
               _ExtentX        =   7726
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   4140
               Left            =   135
               TabIndex        =   38
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   705
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   7303
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
               ButtonImage     =   "FrmEvaluaEntit.frx":1F44C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin XtremeSuiteControls.CheckBox DeptSelect 
               Height          =   1260
               Left            =   5550
               TabIndex        =   39
               Top             =   1260
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   2222
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
               TabIndex        =   40
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   3255
               Width           =   4380
               _ExtentX        =   7726
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ProjSelect 
               Height          =   1260
               Left            =   5550
               TabIndex        =   41
               Top             =   3255
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   2222
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
               Height          =   1440
               Index           =   3
               Left            =   12870
               TabIndex        =   42
               Top             =   0
               Width           =   1530
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   5160
            Left            =   15180
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   9102
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
               Height          =   5460
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   14535
               _cx             =   25638
               _cy             =   9631
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
         TabIndex        =   23
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
            TabIndex        =   28
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
            Format          =   93913089
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBrnch2 
            Height          =   315
            Left            =   1080
            TabIndex        =   57
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
            TabIndex        =   58
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
            TabIndex        =   29
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
            TabIndex        =   24
            Top             =   240
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   1110
         Left            =   0
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4800
         Width           =   14580
         _cx             =   25718
         _cy             =   1958
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
      Begin ImpulseButton.ISButton btnNew 
         Height          =   315
         Left            =   12735
         TabIndex        =   78
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   7830
         Width           =   1140
         _ExtentX        =   2011
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
         ButtonImage     =   "FrmEvaluaEntit.frx":25CAE
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   315
         Left            =   11055
         TabIndex        =   79
         ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
         Top             =   7800
         Width           =   1140
         _ExtentX        =   2011
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
         ButtonImage     =   "FrmEvaluaEntit.frx":2C510
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   315
         Left            =   9375
         TabIndex        =   80
         ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   7800
         Width           =   1140
         _ExtentX        =   2011
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
         ButtonImage     =   "FrmEvaluaEntit.frx":32D72
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   315
         Left            =   7815
         TabIndex        =   81
         ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
         Top             =   7800
         Width           =   1140
         _ExtentX        =   2011
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
         ButtonImage     =   "FrmEvaluaEntit.frx":3310C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   315
         Left            =   6015
         TabIndex        =   82
         ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
         Top             =   7800
         Width           =   1140
         _ExtentX        =   2011
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
         ButtonImage     =   "FrmEvaluaEntit.frx":334A6
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton ISButton5 
         Height          =   315
         Left            =   4245
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
         Top             =   7845
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
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
         ButtonImage     =   "FrmEvaluaEntit.frx":33A40
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton8 
         Height          =   315
         Left            =   2490
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   7845
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
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
         ButtonImage     =   "FrmEvaluaEntit.frx":3A2A2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   315
         Left            =   720
         TabIndex        =   85
         ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
         Top             =   7800
         Width           =   1155
         _ExtentX        =   2037
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
         ButtonImage     =   "FrmEvaluaEntit.frx":3A63C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
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
Attribute VB_Name = "FrmEvaluaEntit"
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
Private Sub Cmd_Click(Index As Integer)
    If Me.TxtModFlg.Text <> "R" Then
        Select Case Index
            Case 3
                RemoveGridRow
            Case 4
                RemoveGridAllRow
        End Select
    End If
Relin
'
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
    Dim conection As String
    Dim My_SQL As String
    
    On Error GoTo ErrTrap
    
    With GridInstallments
        If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("TypeEnti")) = "#1;“Ì«œ… |#2; —ﬁÌ… "
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("TypeEnti")) = "#1;Increase |#2;Upgrade "
        End If
    End With

    conection = "select * from  TblEvaluaEntit order by  ID "
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
Public Sub FiLLRec()
    Dim Sql As String
    Dim ID As Double
    
    On Error GoTo ErrTrap
    
    If Me.TxtModFlg.Text = "E" Then
        StrSQL = "Delete From TblEvaluaEntitDet Where EvlaID=" & val(Me.TxtSerial1.Text)
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
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("TotalValue").value = val(lbl(9).Caption)
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
    
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblEvaluaEntitDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim I As Integer
    With Me.GridInstallments
        For I = .FixedRows To .Rows - 1
            If val(.TextMatrix(I, .ColIndex("EmpID"))) <> 0 Then
                RsDevsub.AddNew
                RsDevsub("EvlaID").value = val(Me.TxtSerial1.Text)
                RsDevsub("EmpID").value = IIf((.TextMatrix(I, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(I, .ColIndex("EmpID"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(I, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(I, .ColIndex("DeptID"))))
                RsDevsub("ProjectID").value = IIf((.TextMatrix(I, .ColIndex("ProjID"))) = "", Null, val(.TextMatrix(I, .ColIndex("ProjID"))))
                RsDevsub("BrnchID1").value = IIf((.TextMatrix(I, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(I, .ColIndex("BranchID"))))
                RsDevsub("TotDigree").value = IIf((.TextMatrix(I, .ColIndex("TotDigree"))) = "", Null, val(.TextMatrix(I, .ColIndex("TotDigree"))))
                RsDevsub("NewValue").value = IIf((.TextMatrix(I, .ColIndex("NewValue"))) = "", Null, val(.TextMatrix(I, .ColIndex("NewValue"))))
                RsDevsub("JobID").value = IIf((.TextMatrix(I, .ColIndex("JobID"))) = "", Null, val(.TextMatrix(I, .ColIndex("JobID"))))
                RsDevsub("ToJob").value = IIf((.TextMatrix(I, .ColIndex("ToJob"))) = "", Null, val(.TextMatrix(I, .ColIndex("ToJob"))))
                RsDevsub("TypeEnti").value = IIf((.TextMatrix(I, .ColIndex("TypeEnti"))) = "", Null, val(.TextMatrix(I, .ColIndex("TypeEnti"))))
                    If val(.TextMatrix(I, .ColIndex("ToJob"))) <> 0 And val(.TextMatrix(I, .ColIndex("TypeEnti"))) = 2 Then
                        Cn.Execute "Update TblEmployee  set JobTypeID =" & val(.TextMatrix(I, .ColIndex("ToJob"))) & "  where Emp_ID=" & val(.TextMatrix(I, .ColIndex("EmpID"))) & ""
                    End If
                RsDevsub.update
            End If
        Next I
    End With
    
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
Public Sub FiLLTXT()
    Dim I As Integer
    Dim Shifttime As Date
    
    On Error GoTo ErrTrap
  
    SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
    
    lbl(9).Caption = IIf(IsNull(RsSavRec.Fields("TotalValue").value), 0, RsSavRec.Fields("TotalValue").value)
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
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    FullGridData

ErrTrap:
End Sub
Private Sub btnSave_Click()
    Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double
    Dim Account_Code_dynamic As String
    Dim I As Integer
    
    On Error GoTo ErrTrap

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
        Case "N"
            AddNewRecored
            AddNewRec
        Case "E"
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
Public Sub AddNewRec()
    Dim StrRecID As String
    
    On Error GoTo ErrTrap
    
    StrRecID = new_id("TblEvaluaEntit", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
 Sub FullGridData()
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    Dim Sql As String
    
    On Error GoTo ErrTrap

    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
    
    Sql = " SELECT     dbo.TblEvaluaEntitDet.EvlaID, dbo.TblEvaluaEntitDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, "
    Sql = Sql & "                      dbo.TblEvaluaEntitDet.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEvaluaEntitDet.ToJob,"
    Sql = Sql & "                      TblEmpJobsTypes_1.JobTypeName AS ToJobTypeName, TblEmpJobsTypes_1.JobTypeNamee AS ToJobTypeNameE, dbo.TblEvaluaEntitDet.TotDigree,"
    Sql = Sql & "                      dbo.TblEvaluaEntitDet.NewValue, dbo.TblEvaluaEntitDet.TypeEnti, dbo.TblEvaluaEntitDet.BrnchID1, dbo.TblBranchesData.branch_name,"
    Sql = Sql & "                      dbo.TblBranchesData.branch_namee, dbo.TblEvaluaEntitDet.ProjectID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblEvaluaEntitDet.DeptID,"
    Sql = Sql & "                      dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee"
    Sql = Sql & " FROM         dbo.TblEvaluaEntitDet LEFT OUTER JOIN"
    Sql = Sql & "                      dbo.TblEmpDepartments ON dbo.TblEvaluaEntitDet.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    Sql = Sql & "                      dbo.projects ON dbo.TblEvaluaEntitDet.ProjectID = dbo.projects.id LEFT OUTER JOIN"
    Sql = Sql & "                      dbo.TblBranchesData ON dbo.TblEvaluaEntitDet.BrnchID1 = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    Sql = Sql & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_1 ON dbo.TblEvaluaEntitDet.ToJob = TblEmpJobsTypes_1.JobTypeID LEFT OUTER JOIN"
    Sql = Sql & "                      dbo.TblEmpJobsTypes ON dbo.TblEvaluaEntitDet.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
    Sql = Sql & "                      dbo.TblEmployee ON dbo.TblEvaluaEntitDet.EmpID = dbo.TblEmployee.Emp_ID"
    Sql = Sql & " Where (dbo.TblEvaluaEntitDet.EvlaID = " & val(TxtSerial1.Text) & ")"
    Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
    End If
    Dim I As Integer
    With Me.GridInstallments
        For I = .FixedRows To Rs1.RecordCount
            .Rows = .FixedRows + Rs1.RecordCount
            .TextMatrix(I, .ColIndex("Ser")) = I
            .TextMatrix(I, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
            .TextMatrix(I, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
            .TextMatrix(I, .ColIndex("JobID")) = IIf(IsNull(Rs1("JobID").value), "", Rs1("JobID").value)
            .TextMatrix(I, .ColIndex("ToJob")) = IIf(IsNull(Rs1("ToJob").value), "", Rs1("ToJob").value)
            .TextMatrix(I, .ColIndex("DeptID")) = IIf(IsNull(Rs1("DeptID").value), 0, Rs1("DeptID").value)
            .TextMatrix(I, .ColIndex("ProjID")) = IIf(IsNull(Rs1("ProjectID").value), 0, Rs1("ProjectID").value)
            .TextMatrix(I, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BrnchID1").value), 0, Rs1("BrnchID1").value)
            .TextMatrix(I, .ColIndex("TotDigree")) = IIf(IsNull(Rs1("TotDigree").value), "", Rs1("TotDigree").value)
            .TextMatrix(I, .ColIndex("NewValue")) = IIf(IsNull(Rs1("NewValue").value), "", Rs1("NewValue").value)
            .TextMatrix(I, .ColIndex("TypeEnti")) = IIf(IsNull(Rs1("TypeEnti").value), "", Rs1("TypeEnti").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("ToJobName")) = IIf(IsNull(Rs1("ToJobTypeName").value), "", Rs1("ToJobTypeName").value)
                .TextMatrix(I, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentName").value), "", Rs1("DepartmentName").value)
                .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_name").value), "", Rs1("Project_name").value)
                .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                .TextMatrix(I, .ColIndex("JobName")) = IIf(IsNull(Rs1("JobTypeName").value), "", Rs1("JobTypeName").value)
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
            Else
                .TextMatrix(I, .ColIndex("ToJobName")) = IIf(IsNull(Rs1("ToJobTypeNameE").value), "", Rs1("ToJobTypeNameE").value)
                .TextMatrix(I, .ColIndex("JobName")) = IIf(IsNull(Rs1("JobTypeNamee").value), "", Rs1("JobTypeNamee").value)
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
    Dim LngRow As Long
    Dim StrAccountCode As String
    With Me.GridInstallments
        Select Case .ColKey(Col)
            Case "ToJobName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ToJob"), False, True)
                .TextMatrix(Row, .ColIndex("ToJob")) = StrAccountCode
            Case "TypeEnti"
                If val(.TextMatrix(Row, .ColIndex("TypeEnti"))) = 1 Then
                    .TextMatrix(Row, .ColIndex("ToJob")) = ""
                    .TextMatrix(Row, .ColIndex("ToJobName")) = ""
                ElseIf val(.TextMatrix(Row, .ColIndex("TypeEnti"))) = 2 Then
                    .TextMatrix(Row, .ColIndex("NewValue")) = ""
                End If
        End Select
    End With
    Relin
End Sub
Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With GridInstallments
        Select Case .ColKey(Col)
            Case "FullCode"
                Cancel = True
            Case "Emp_Name"
                Cancel = True
            Case "JobName"
                Cancel = True
            Case "TotDigree"
                .ComboList = ""
            Case "NewValue"
                If val(.TextMatrix(Row, .ColIndex("TypeEnti"))) = 1 Then
                    .ComboList = ""
                Else
                    Cancel = True
                End If
            Case "ToJobName"
                If val(.TextMatrix(Row, .ColIndex("TypeEnti"))) = 2 Then
                    Cancel = False
                Else
                    Cancel = True
                End If
        End Select
    End With
End Sub
Sub Relin()
    Dim I As Integer
    Dim SumVal As Double
    SumVal = 0
    With GridInstallments
        For I = 1 To .Rows - 1
            If val(.TextMatrix(I, .ColIndex("EmpID"))) <> 0 Then
                SumVal = SumVal + val(.TextMatrix(I, .ColIndex("NewValue")))
            End If
        Next I
    End With
    lbl(9).Caption = SumVal
End Sub
Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrComboList As String
    
    With GridInstallments
        Select Case .ColKey(Col)
            Case "ToJobName"
                StrSQL = "select * from TblEmpJobsTypes"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "JobTypeName", "JobTypeID")
                Else
                    StrComboList = .BuildComboList(rs, "JobTypeNamee", "JobTypeID")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
        End Select
  End With
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
    Dim I As Integer
    Dim k As Integer
    Dim Sql As String
    
    Sql = " SELECT     dbo.TblEvaluation_Details.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.JobTypeID, "
    Sql = Sql & "                       dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, SUM(dbo.TblEvaluation_Details.sum_Degrees) AS Sumsum_Degrees,"
    Sql = Sql & "                       dbo.TblEmployee.BranchID , dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID"
    Sql = Sql & "  FROM         dbo.TblEmpEvaluation LEFT OUTER JOIN"
    Sql = Sql & "                       dbo.TblEvaluation_Details ON dbo.TblEmpEvaluation.ID = dbo.TblEvaluation_Details.HID LEFT OUTER JOIN"
    Sql = Sql & "                       dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
    Sql = Sql & "                       dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID ON dbo.TblEvaluation_Details.Emp_ID = dbo.TblEmployee.Emp_ID"
    Sql = Sql & "  Where (1 = 1) "
    
    If val(Me.DcbProject1.BoundText) <> 0 And Me.SelectProject.value = vbChecked Then
        Sql = Sql & " and dbo.TblEmployee.project_id  =" & val(DcbProject1.BoundText) & " "
    End If
    If val(DcbBranch1.BoundText) <> 0 And Me.SelectBranch.value = vbChecked Then
        Sql = Sql & " and dbo.TblEmployee.BranchID  =" & val(DcbBranch1.BoundText) & " "
    End If
    If val(DcpDept1.BoundText) <> 0 And Me.SelectDept.value = vbChecked Then
        Sql = Sql & " and dbo.TblEmployee.DepartmentID  =" & val(DcpDept1.BoundText) & " "
    End If
    If val(DcbEmployee1.BoundText) <> 0 And RdEmp.value = True Then
        Sql = Sql & " and dbo.TblEvaluation_Details.Emp_id=" & val(DcbEmployee1.BoundText) & " "
    End If
    If Not IsNull(Me.FromDate.value) Then
        Sql = Sql & " and dbo.TblEmpEvaluation.YearID >=" & year(FromDate.value) - 2006 & " "
    End If
    If Not IsNull(Me.ToDate.value) Then
        Sql = Sql & " and dbo.TblEmpEvaluation.YearID <=" & year(ToDate.value) - 2006 & " "
    End If
    If Not IsNull(Me.FromDate.value) Then
        Sql = Sql & " and dbo.TblEmpEvaluation.MonthID >=" & Month(FromDate.value) - 1 & " "
    End If
    If Not IsNull(Me.ToDate.value) Then
        Sql = Sql & " and dbo.TblEmpEvaluation.MonthID <=" & Month(ToDate.value) - 1 & " "
    End If
    
    Sql = Sql & " GROUP BY dbo.TblEvaluation_Details.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.JobTypeID,"
    Sql = Sql & "                      dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id,"
    Sql = Sql & "                      dbo.TblEmployee.DepartmentID"
    
    Rs8.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Rs8.RecordCount > 0 Then
        With GridInstallments
            k = .Rows
            Rs8.MoveFirst
            .Rows = .Rows + Rs8.RecordCount
            For I = k To .Rows - 1
                .TextMatrix(I, .ColIndex("EmpID")) = IIf(IsNull(Rs8("Emp_ID").value), 0, Rs8("Emp_ID").value)
                .TextMatrix(I, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DepartmentID").value), 0, Rs8("DepartmentID").value)
                .TextMatrix(I, .ColIndex("ProjID")) = IIf(IsNull(Rs8("project_id").value), 0, Rs8("project_id").value)
                .TextMatrix(I, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
                .TextMatrix(I, .ColIndex("FullCode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
                .TextMatrix(I, .ColIndex("JobID")) = IIf(IsNull(Rs8("JobTypeID").value), "", Rs8("JobTypeID").value)
                .TextMatrix(I, .ColIndex("TotDigree")) = IIf(IsNull(Rs8("Sumsum_Degrees").value), "", Rs8("Sumsum_Degrees").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(I, .ColIndex("JobName")) = IIf(IsNull(Rs8("JobTypeName").value), "", Rs8("JobTypeName").value)
                    .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
                Else
                    .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)
                    .TextMatrix(I, .ColIndex("JobName")) = IIf(IsNull(Rs8("JobTypeNamee").value), "", Rs8("JobTypeNamee").value)
                End If
                Rs8.MoveNext
            Next I
        End With
    End If
End Sub

Private Sub ISButton8_Click()
    Unload FrmInsurancesSearch
    FrmInsurancesSearch.SendForm = 4
    FrmInsurancesSearch.show
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
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
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
Private Sub BtnCancel_Click()
    Unload Me
End Sub
Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
    BtnLast_Click
End Sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim Sql As String
    Dim X As Integer
    Dim I As Integer
    Dim ID As Double
    
    On Error GoTo ErrTrap
    
    If DoPremis(Do_Delete, Me.Name, True) = False Then
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
            With Me.GridInstallments
                For I = 1 To .Rows - 1
                    If val(.TextMatrix(I, .ColIndex("JobID"))) <> 0 Then
                        Cn.Execute "Update TblEmployee  set JobTypeID =" & val(.TextMatrix(I, .ColIndex("JobID"))) & "  where Emp_ID=" & val(.TextMatrix(I, .ColIndex("EmpID"))) & ""
                    End If
                Next I
            End With
            StrSQL = "Delete From TblEvaluaEntitDet Where EvlaID=" & val(Me.TxtSerial1.Text) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
            RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
            RsSavRec.delete
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            lbl(9).Caption = 0
            SelectDept.value = vbUnchecked
            SelectProject.value = vbUnchecked
            SelectBranch.value = vbUnchecked
            If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
            Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
            End If
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
'        End If
        Me.Refresh
        BtnNext_Click
        Exit Sub
    End If
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
Public Sub EditRec(StrTable As String, RecId As String)
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
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub
Private Sub BtnFirst_Click()
    Dim Msg As String
    
    On Error GoTo ErrTrap
    
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
    Dim Msg As String
    
    On Error GoTo ErrTrap
    
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
    Me.DCboUserName.BoundText = user_id
    Me.DcbBrnch2.BoundText = Current_branch
    lbl(9).Caption = 0
    RdAll.value = True
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    Dim Msg As String
    On Error GoTo ErrTrap
    
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
    Dim Msg As String
    
    On Error GoTo ErrTrap
    
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
    ISButton8.Caption = "Search"
    lbl(7).Caption = "Total"
    lbl(4).Caption = "No"
    lbl(1).Caption = "Date"
    lbl(2).Caption = "Branch"
    lbl(5).Caption = "Remarks"
    lbl(0).Caption = "From Date"
    lbl(6).Caption = "To Date"
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
    'lbl(0).Caption = "Data"
    C1Tab1.Caption = "Data"
    Cmd(3).Caption = "Delete Row"
    Cmd(4).Caption = "Delete all"
    Label1(2).Caption = "Rating Entitlement"
    With Me.GridInstallments
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("FullCode")) = "Employee Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
        .TextMatrix(0, .ColIndex("TotDigree")) = "Total Marks"
        .TextMatrix(0, .ColIndex("TypeEnti")) = "Bonus Type"
        .TextMatrix(0, .ColIndex("NewValue")) = "Value"
        .TextMatrix(0, .ColIndex("ToJobName")) = "Job"
        .TextMatrix(0, .ColIndex("Selct")) = "Select"
        .TextMatrix(0, .ColIndex("TypeMofrd")) = "Type Payment"
  End With
  ISButton8.Caption = "Search"
ErrTrap:
End Sub
Private Sub AddNewRecored()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrTrap
    
    Set rs = New ADODB.Recordset
    My_SQL = "TblEvaluaEntit"
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
             If val(.TextMatrix(I, .ColIndex("JobID"))) <> 0 Then
                Cn.Execute "Update TblEmployee  set JobTypeID =" & val(.TextMatrix(I, .ColIndex("JobID"))) & "  where Emp_ID=" & val(.TextMatrix(I, .ColIndex("EmpID"))) & ""
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
            If val(.TextMatrix(.Row, .ColIndex("JobID"))) <> 0 Then
                Cn.Execute "Update TblEmployee  set JobTypeID =" & val(.TextMatrix(.Row, .ColIndex("JobID"))) & "  where Emp_ID=" & val(.TextMatrix(.Row, .ColIndex("EmpID"))) & ""
            End If
            .RemoveItem .Row
    End With
End Sub
