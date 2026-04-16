VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmGeneralFundReceipt 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”‰œ ﬁ»÷ «·’‰œÊﬁ «·⁄«„  "
   ClientHeight    =   8985
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   16260
   HelpContextID   =   580
   Icon            =   "FrmBankDeposite2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   16260
   Visible         =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8985
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16260
      _cx             =   28681
      _cy             =   15849
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
      BorderWidth     =   2
      ChildSpacing    =   1
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7890
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   16230
         _cx             =   28628
         _cy             =   13917
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
         Caption         =   "."
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7470
            Index           =   2
            Left            =   45
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   45
            Width           =   16140
            _cx             =   28469
            _cy             =   13176
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   735
               Index           =   5
               Left            =   0
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   0
               Width           =   16140
               _cx             =   28469
               _cy             =   1296
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial (Arabic)"
                  Size            =   21.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   4
               MousePointer    =   0
               Version         =   801
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Picture         =   "FrmBankDeposite2.frx":6852
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   0
               ChildSpacing    =   0
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   6
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   2
               PicturePos      =   0
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
               Begin VB.TextBox oldtxtNoteSerial1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3570
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1440
               End
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3930
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   126
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1620
                  _ExtentX        =   2858
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
               Begin ImpulseButton.ISButton btnLast 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   21
                  Top             =   120
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
                  ButtonImage     =   "FrmBankDeposite2.frx":752C
                  ColorButton     =   16777215
                  AcclimateGrayTones=   -1  'True
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnNext 
                  Height          =   315
                  Left            =   1785
                  TabIndex        =   20
                  Top             =   120
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
                  ButtonImage     =   "FrmBankDeposite2.frx":78C6
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnPrevious 
                  Height          =   315
                  Left            =   2385
                  TabIndex        =   19
                  Top             =   120
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
                  ButtonImage     =   "FrmBankDeposite2.frx":7C60
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnFirst 
                  Height          =   315
                  Left            =   2910
                  TabIndex        =   18
                  Top             =   120
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
                  ButtonImage     =   "FrmBankDeposite2.frx":7FFA
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰œ ﬁ»÷ «·’‰œÊﬁ «·⁄«„"
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
                  Left            =   11040
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   120
                  Width           =   3960
               End
               Begin VB.Image Image1 
                  Height          =   375
                  Left            =   15270
                  Picture         =   "FrmBankDeposite2.frx":8394
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   540
               End
            End
            Begin C1SizerLibCtl.C1Elastic Frm2 
               Height          =   7575
               Index           =   1
               Left            =   0
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   -120
               Width           =   16140
               _cx             =   28469
               _cy             =   13361
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
                  Height          =   345
                  Left            =   12690
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   1080
                  Width           =   1500
               End
               Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
                  Height          =   3555
                  Left            =   0
                  TabIndex        =   128
                  Top             =   3510
                  Width           =   16035
                  _cx             =   28284
                  _cy             =   6271
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
                  Rows            =   2
                  Cols            =   19
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBankDeposite2.frx":A113
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
                     Left            =   2400
                     TabIndex        =   129
                     Top             =   1680
                     Visible         =   0   'False
                     Width           =   11295
                     _ExtentX        =   19923
                     _ExtentY        =   1085
                     _Version        =   393216
                     Appearance      =   0
                  End
               End
               Begin VB.TextBox txtManualNo 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   8955
                  RightToLeft     =   -1  'True
                  TabIndex        =   1
                  Top             =   1095
                  Width           =   1500
               End
               Begin VB.CheckBox Check17 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  Height          =   345
                  Left            =   24795
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   2385
                  Width           =   1125
               End
               Begin VB.CheckBox chkDue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ «·‘Ìﬂ«  «·„” Õﬁ…  ›ﬁÿ"
                  Height          =   315
                  Left            =   6015
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   5490
                  Width           =   2880
               End
               Begin VB.TextBox TxtTotalChequesView 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   17790
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   0
                  Width           =   1860
               End
               Begin VB.TextBox TxtTotalCashView 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   585
                  Left            =   1920
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   5160
                  Width           =   1845
               End
               Begin VB.TextBox TxtBankName 
                  Alignment       =   1  'Right Justify
                  Height          =   420
                  Left            =   17610
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   5340
                  Visible         =   0   'False
                  Width           =   2430
               End
               Begin VB.TextBox XXXX 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   19620
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   0
                  Width           =   1905
               End
               Begin VB.TextBox TxtTotalCheques 
                  Alignment       =   1  'Right Justify
                  Height          =   390
                  Left            =   17280
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   4755
                  Width           =   1710
               End
               Begin VB.TextBox TxtTotalCash 
                  Alignment       =   1  'Right Justify
                  Height          =   420
                  Left            =   960
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   5340
                  Width           =   1695
               End
               Begin VB.TextBox txtchequeno 
                  Alignment       =   1  'Right Justify
                  Height          =   420
                  Left            =   17760
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   5340
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.TextBox TxtValue1 
                  Alignment       =   1  'Right Justify
                  Height          =   420
                  Left            =   3435
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   5340
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.TextBox TxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   4425
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   -105
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.Frame Frame1 
                  Caption         =   "„⁄·Ê„« "
                  Height          =   2745
                  Left            =   20490
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1395
                  Width           =   4800
                  Begin MSDataListLib.DataCombo xxx 
                     Height          =   315
                     Index           =   0
                     Left            =   120
                     TabIndex        =   67
                     Top             =   120
                     Width           =   2565
                     _ExtentX        =   4524
                     _ExtentY        =   556
                     _Version        =   393216
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
                  Begin MSDataListLib.DataCombo DCGroup 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   68
                     Top             =   480
                     Width           =   2565
                     _ExtentX        =   4524
                     _ExtentY        =   556
                     _Version        =   393216
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
                  Begin VB.Label lblTotalLate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   1440
                     Width           =   1200
                  End
                  Begin VB.Label lblTotalRevenue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   1155
                     Width           =   1200
                  End
                  Begin VB.Label lblTotlSales 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     ForeColor       =   &H00FF0000&
                     Height          =   315
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   840
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ì »⁄ „Ã„Ê⁄Â"
                     Height          =   315
                     Index           =   11
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Top             =   480
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ì »⁄ ›—⁄"
                     Height          =   315
                     Index           =   10
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   240
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·„ √Œ—« "
                     Height          =   195
                     Index           =   9
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   1440
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «· Õ’Ì·« "
                     Height          =   195
                     Index           =   6
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   1150
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·„»Ì⁄« "
                     Height          =   315
                     Index           =   4
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   840
                     Width           =   1200
                  End
               End
               Begin VB.TextBox txtRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   510
                  Left            =   2520
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   5
                  Top             =   2070
                  Width           =   4785
               End
               Begin VB.CheckBox ChkLocked 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  Height          =   645
                  Left            =   16440
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   2910
                  Width           =   2580
               End
               Begin VB.OptionButton Option2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ Ì«— ’‰›"
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   17235
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   3660
                  Value           =   -1  'True
                  Width           =   1260
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ ﬂ«›Â «·«’‰«›"
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   17190
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   3660
                  Width           =   1800
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   16335
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   5955
                  Width           =   1755
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   705
                  Left            =   17595
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Text            =   "0"
                  Top             =   2910
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox TxtlBanksDepositeId 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   18150
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   900
                  Visible         =   0   'False
                  Width           =   1410
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   225
                  Left            =   16335
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   6165
                  Width           =   2595
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   0
                  Left            =   -4170
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   15885
                  Width           =   2280
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   18150
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   900
                  Visible         =   0   'False
                  Width           =   2310
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2010
                  Left            =   17130
                  TabIndex        =   28
                  Top             =   -630
                  Visible         =   0   'False
                  Width           =   10260
                  _cx             =   18098
                  _cy             =   3545
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBankDeposite2.frx":A3DC
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
               Begin MSComCtl2.DTPicker dbRecordDate 
                  Height          =   360
                  Left            =   6240
                  TabIndex        =   33
                  Top             =   1020
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   635
                  _Version        =   393216
                  Format          =   104792065
                  CurrentDate     =   41640
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   17415
                  TabIndex        =   34
                  Top             =   3525
                  Width           =   4740
                  _ExtentX        =   8361
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo dcproject 
                  Height          =   315
                  Left            =   17280
                  TabIndex        =   36
                  Top             =   2655
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo Dcterm 
                  Height          =   315
                  Left            =   18135
                  TabIndex        =   44
                  Top             =   1395
                  Width           =   3435
                  _ExtentX        =   6059
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSComCtl2.DTPicker dbTodate 
                  Height          =   645
                  Left            =   18555
                  TabIndex        =   52
                  Top             =   2775
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   1138
                  _Version        =   393216
                  Format          =   104792065
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   495
                  Index           =   20
                  Left            =   18015
                  TabIndex        =   57
                  Top             =   3660
                  Width           =   795
                  _ExtentX        =   1402
                  _ExtentY        =   873
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmBankDeposite2.frx":A6B7
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   495
                  Index           =   21
                  Left            =   17130
                  TabIndex        =   58
                  Top             =   3660
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   873
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
                  ButtonImage     =   "FrmBankDeposite2.frx":AA51
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DCEmp 
                  Height          =   315
                  Left            =   17625
                  TabIndex        =   59
                  Top             =   2340
                  Width           =   2805
                  _ExtentX        =   4948
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo dcitems 
                  Height          =   315
                  Left            =   17895
                  TabIndex        =   60
                  Top             =   3660
                  Width           =   4680
                  _ExtentX        =   8255
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin VSFlex8Ctl.VSFlexGrid Grid1 
                  Height          =   2760
                  Left            =   17985
                  TabIndex        =   73
                  Top             =   6630
                  Width           =   10470
                  _cx             =   18468
                  _cy             =   4868
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   24
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBankDeposite2.frx":AFEB
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
               Begin MSDataListLib.DataCombo Dcbank 
                  Height          =   315
                  Left            =   17415
                  TabIndex        =   75
                  Top             =   1815
                  Visible         =   0   'False
                  Width           =   4470
                  _ExtentX        =   7885
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   9000
                  TabIndex        =   3
                  Top             =   2010
                  Width           =   5190
                  _ExtentX        =   9155
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo Dcbank1 
                  Height          =   315
                  Left            =   18870
                  TabIndex        =   80
                  Top             =   1890
                  Visible         =   0   'False
                  Width           =   2475
                  _ExtentX        =   4366
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   525
                  Index           =   9
                  Left            =   17760
                  TabIndex        =   90
                  Top             =   2010
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   926
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmBankDeposite2.frx":B352
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   525
                  Index           =   10
                  Left            =   16995
                  TabIndex        =   91
                  Top             =   2010
                  Visible         =   0   'False
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   926
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› ”ÿ—"
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
                  ButtonImage     =   "FrmBankDeposite2.frx":B6EC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo Dcbranch 
                  Height          =   315
                  Left            =   2520
                  TabIndex        =   4
                  Top             =   1695
                  Width           =   4815
                  _ExtentX        =   8493
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCChequeBox 
                  Height          =   315
                  Left            =   16260
                  TabIndex        =   95
                  Top             =   4920
                  Width           =   6135
                  _ExtentX        =   10821
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker FromDate 
                  Height          =   330
                  Left            =   12930
                  TabIndex        =   112
                  Top             =   2760
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   104792065
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DTPicker2 
                  Height          =   405
                  Left            =   0
                  TabIndex        =   113
                  Top             =   0
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   714
                  _Version        =   393216
                  Format          =   104792065
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   330
                  Left            =   9075
                  TabIndex        =   114
                  Top             =   2760
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   104792065
                  CurrentDate     =   41640
               End
               Begin MSDataListLib.DataCombo DcGeneralBox 
                  Height          =   315
                  Left            =   9000
                  TabIndex        =   2
                  Top             =   1710
                  Width           =   5175
                  _ExtentX        =   9128
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton ISButton3 
                  Height          =   330
                  Left            =   1980
                  TabIndex        =   122
                  ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
                  Top             =   7080
                  Visible         =   0   'False
                  Width           =   1710
                  _ExtentX        =   3016
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
                  ButtonImage     =   "FrmBankDeposite2.frx":BC86
                  ButtonImageDisabled=   "FrmBankDeposite2.frx":124E8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   123
                  ToolTipText     =   "Õ–› «·ﬂ·"
                  Top             =   7080
                  Visible         =   0   'False
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ· "
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
                  ButtonImage     =   "FrmBankDeposite2.frx":316D2
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   450
                  Left            =   2535
                  TabIndex        =   6
                  ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                  Top             =   2760
                  Width           =   4800
                  _ExtentX        =   8467
                  _ExtentY        =   794
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
                  ButtonImage     =   "FrmBankDeposite2.frx":37F34
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
                  Height          =   255
                  Left            =   2655
                  TabIndex        =   124
                  Top             =   1080
                  Width           =   1785
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin ImpulseButton.ISButton ISButton7 
                  Height          =   330
                  Left            =   6000
                  TabIndex        =   130
                  ToolTipText     =   " ÕœÌœ «·ﬂ·"
                  Top             =   7080
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " ÕœÌœ «·ﬂ·"
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
                  ButtonImage     =   "FrmBankDeposite2.frx":3E796
                  ButtonImageDisabled=   "FrmBankDeposite2.frx":44FF8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton ISButton8 
                  Height          =   330
                  Left            =   4080
                  TabIndex        =   131
                  ToolTipText     =   " ÕœÌœ «·ﬂ·"
                  Top             =   7080
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "«·€«¡ «· ÕœÌœ"
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
                  ButtonImage     =   "FrmBankDeposite2.frx":641E2
                  ButtonImageDisabled=   "FrmBankDeposite2.frx":6AA44
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcBoxMan 
                  Height          =   315
                  Left            =   9000
                  TabIndex        =   135
                  Top             =   2355
                  Width           =   5190
                  _ExtentX        =   9155
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboDebitSide 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   137
                  Top             =   0
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’‰œÊﬁ «·„‰œÊ»"
                  Height          =   285
                  Index           =   17
                  Left            =   14535
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   2340
                  Width           =   1335
               End
               Begin VB.Label TotalTXT 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   360
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   7080
                  Width           =   2055
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ã„Ê⁄"
                  Height          =   240
                  Index           =   3
                  Left            =   11160
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   7200
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   285
                  Index           =   30
                  Left            =   4545
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   1080
                  Width           =   1395
               End
               Begin VB.Shape Shape3 
                  BorderColor     =   &H0000C0C0&
                  BorderWidth     =   2
                  Height          =   2415
                  Left            =   2415
                  Top             =   960
                  Width           =   13605
               End
               Begin VB.Image Image2 
                  Height          =   255
                  Left            =   645
                  Picture         =   "FrmBankDeposite2.frx":89C2E
                  Stretch         =   -1  'True
                  Top             =   1080
                  Width           =   255
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ… :"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   225
                  Index           =   28
                  Left            =   645
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   1125
                  Width           =   1545
               End
               Begin VB.Shape Shape1 
                  BorderWidth     =   2
                  Height          =   1470
                  Left            =   135
                  Top             =   1575
                  Width           =   2160
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "„‰ Œ·«· Â–… «·‘«‘…   „ﬂ‰ „‰ ⁄„· ﬁ»÷ „‰ ﬂ· ”‰œ«  «· Õ’Ì· ··„‰«œÌ»"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   1395
                  Index           =   26
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   1575
                  Width           =   2040
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ ÌœÊÌ"
                  Height          =   345
                  Index           =   25
                  Left            =   11205
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   1095
                  Width           =   930
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï  «—ÌŒ"
                  Height          =   330
                  Index           =   23
                  Left            =   11130
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   2760
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   330
                  Index           =   22
                  Left            =   14610
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   2865
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·‘Ìﬂ«  «·„Õœœ…"
                  Height          =   405
                  Index           =   21
                  Left            =   18015
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   0
                  Width           =   1470
               End
               Begin VB.Label TxtPaymentCounts 
                  Alignment       =   1  'Right Justify
                  Caption         =   "0"
                  Height          =   510
                  Left            =   16200
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   0
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   405
                  Index           =   19
                  Left            =   22110
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   0
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ Õ«›Ÿ… «·‘Ìﬂ« "
                  Height          =   405
                  Index           =   18
                  Left            =   8745
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   5880
                  Width           =   1605
               End
               Begin VB.Label lblBranch 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄ „‰›– «·⁄„·Ì…"
                  Height          =   360
                  Left            =   7545
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   1695
                  Width           =   1410
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·‘Ìﬂ« "
                  ForeColor       =   &H000000FF&
                  Height          =   375
                  Left            =   17640
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   0
                  Width           =   1350
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·‰ﬁœ"
                  ForeColor       =   &H000000FF&
                  Height          =   360
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   5340
                  Width           =   1080
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·‘Ìﬂ"
                  Height          =   285
                  Left            =   17220
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁÌ„Â"
                  Height          =   360
                  Index           =   0
                  Left            =   11685
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»‰ﬂ"
                  Height          =   420
                  Index           =   16
                  Left            =   24960
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   1935
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—’Ìœ"
                  Height          =   345
                  Index           =   0
                  Left            =   5820
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   -105
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’‰œÊﬁ «·—∆Ì”Ì"
                  Height          =   300
                  Index           =   15
                  Left            =   14370
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   1710
                  Width           =   1560
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰œÊ» "
                  Height          =   285
                  Index           =   14
                  Left            =   14550
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   1995
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìœ«⁄«  ‘Ìﬂ« "
                  ForeColor       =   &H00FF0000&
                  Height          =   270
                  Index           =   13
                  Left            =   8340
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   5340
                  Width           =   1860
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìœ«⁄«  ‰ﬁœÌ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Index           =   12
                  Left            =   17970
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   2355
                  Width           =   2220
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   300
                  Index           =   3
                  Left            =   7545
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   2190
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   645
                  Index           =   2
                  Left            =   17145
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   2775
                  Width           =   330
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰œÊ»"
                  Height          =   375
                  Index           =   0
                  Left            =   16995
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   2340
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   240
                  Index           =   5
                  Left            =   7455
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   1095
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  Height          =   375
                  Index           =   8
                  Left            =   17280
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   4590
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”‰œ"
                  Height          =   315
                  Index           =   7
                  Left            =   14490
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   1095
                  Width           =   1320
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   825
                  Left            =   14580
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   1530
                  Width           =   1020
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   300
               Index           =   1
               Left            =   7275
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   90
               Width           =   795
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   900
         Left            =   0
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   7920
         Width           =   16050
         _cx             =   28310
         _cy             =   1588
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   1
         ChildSpacing    =   1
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
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3300
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   585
            Width           =   1170
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   12825
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   1395
            Visible         =   0   'False
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14737632
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
            ButtonImage     =   "FrmBankDeposite2.frx":90480
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   315
            Left            =   12510
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   1545
            Visible         =   0   'False
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌÀ"
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
            ButtonImage     =   "FrmBankDeposite2.frx":9081A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13650
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1830
            Visible         =   0   'False
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   14.25
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBankDeposite2.frx":90BB4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   14340
            TabIndex        =   48
            Tag             =   "Delete Row"
            Top             =   1200
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmBankDeposite2.frx":90F4E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   11
            Left            =   1650
            TabIndex        =   16
            Top             =   510
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   661
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
            ButtonImage     =   "FrmBankDeposite2.frx":90F6A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   375
            Left            =   90
            TabIndex        =   17
            Top             =   510
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   661
            Caption         =   "«·„—›ﬁ« "
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
            ButtonImage     =   "FrmBankDeposite2.frx":977CC
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   11970
            TabIndex        =   120
            Top             =   0
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   14880
            TabIndex        =   7
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   480
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBankDeposite2.frx":9E02E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   12000
            TabIndex        =   9
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBankDeposite2.frx":A4890
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   13560
            TabIndex        =   8
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   480
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBankDeposite2.frx":A4C2A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   10680
            TabIndex        =   10
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBankDeposite2.frx":AB48C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   6960
            TabIndex        =   13
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBankDeposite2.frx":AB826
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   5520
            TabIndex        =   14
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBankDeposite2.frx":ABDC0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   9480
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmBankDeposite2.frx":AC15A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   330
            Left            =   8400
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBankDeposite2.frx":B29BC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   29
            Left            =   14835
            TabIndex        =   121
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌœ —ﬁ„"
            Height          =   315
            Index           =   24
            Left            =   4515
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   585
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   20
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   105
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   37
            Left            =   1515
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   105
            Width           =   1455
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   105
            Width           =   615
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   8415
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   105
            Width           =   870
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   105
            Width           =   1110
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   3330
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   120
            Width           =   915
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   25
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "⁄—÷"
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
      ButtonImage     =   "FrmBankDeposite2.frx":B2D56
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   1470
      Left            =   -7320
      Top             =   -3960
      Width           =   2145
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   27
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   106
      Top             =   9360
      Width           =   7155
   End
End
Attribute VB_Name = "FrmGeneralFundReceipt"
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

Private Sub Cmd_Click(Index As Integer)
If Index = 11 Then
    ShowGL_cc TxtNoteSerial.Text, , 1063, val(Me.TxtNoteID.Text)
End If
End Sub

Private Sub DcboBox_Change()
       If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
          Dim Current_case As Integer, s As String, mBoxID As Long
          Dim rsOut As New ADODB.Recordset
            Set rsOut = New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & val(Me.DcboBox.BoundText)
            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                mBoxID = val(rsOut!BoxID & "")
                DcBoxMan.BoundText = mBoxID
            End If
            
            
        End If

End Sub

Private Sub DcBoxMan_Click(Area As Integer)
Dim mBoxID As Long
        mBoxID = val(DcBoxMan.BoundText)
        Dim StrTempAccountCode     As String
        StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)
        DcboDebitSide.BoundText = StrTempAccountCode
        
End Sub

 Private Sub ISButton6_Click()
 Load FrmGeneralFundReceiptSerch
 FrmGeneralFundReceiptSerch.show vbModal
 End Sub
   Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TBLGeneralFundReceipt  "
    
    
     '       If SystemOptions.usertype <> UserAdmin Then
        conection = conection & " WHERE     BranchID in(" & Current_branchSql & ")"
     
        conection = conection & " order by  IDGFR "
    'End If
    
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName '' user
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetBoxes Me.DcGeneralBox
    Dcombos.GetBoxes Me.DcBoxMan
    Dcombos.GetSalesRepData Me.DcboBox
    Dcombos.GetAccountingCodes Me.DcboDebitSide
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
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
   On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
     StrSQL = "Delete From TBLGeneralFRJoin Where IDGFR='" & val(TxtSerial1.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    RsSavRec.Fields("ManualNo").value = IIf(TxtManualNo.Text <> "", Trim(TxtManualNo.Text), Null)
    RsSavRec.Fields("DateM").value = dbRecordDate.value
    RsSavRec.Fields("DateH").value = Me.Txt_DateHigri.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("GeneralBoxID").value = val(Me.DcGeneralBox.BoundText)
    RsSavRec.Fields("DelegateID").value = val(Me.DcboBox.BoundText)
    RsSavRec.Fields("FromDate").value = Fromdate.value
    RsSavRec.Fields("ToDate").value = ToDate.value
    RsSavRec.Fields("Explan").value = IIf(TxtRemarks.Text <> "", Trim(TxtRemarks.Text), Null)
    RsSavRec.Fields("TotallVal").value = val(Me.TotalTXT.Caption)
    
    RsSavRec.Fields("BoxManID").value = val(Me.DcBoxMan.BoundText)
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TBLGeneralFRJoin Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    
    
'DB_updateField "TBLGeneralFRJoin", "Account_Name", "nvarchar(4000)"
    With FlexGrid
       For i = .FixedRows To .Rows - 1
      If .ValueMatrix(i, .ColIndex("Select")) <> False Then
                RsDevsub.AddNew
                RsDevsub("IDGFR").value = Me.TxtSerial1.Text
                RsDevsub("NoteSerial").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial"))) = "", Null, .TextMatrix(i, .ColIndex("NoteSerial")))
                RsDevsub("ManualNoJoin").value = IIf((.TextMatrix(i, .ColIndex("ManoulNO"))) = "", Null, .TextMatrix(i, .ColIndex("ManoulNO")))
                RsDevsub("ReceiptDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, .TextMatrix(i, .ColIndex("NoteDate")))
                RsDevsub("ReceiptValue").value = IIf((.TextMatrix(i, .ColIndex("Valuue"))) = "", Null, .TextMatrix(i, .ColIndex("Valuue")))
                RsDevsub("AcountCode").value = IIf((.TextMatrix(i, .ColIndex("Account_Serial"))) = "", Null, .TextMatrix(i, .ColIndex("Account_Serial")))
                RsDevsub("AcountId").value = IIf((.TextMatrix(i, .ColIndex("AcountId"))) = "", Null, .TextMatrix(i, .ColIndex("AcountId")))
                RsDevsub("AcountName").value = IIf((.TextMatrix(i, .ColIndex("Account_Name"))) = "", Null, .TextMatrix(i, .ColIndex("Account_Name")))
                RsDevsub("ExplanJoin").value = IIf((.TextMatrix(i, .ColIndex("Remark"))) = "", Null, .TextMatrix(i, .ColIndex("Remark")))
                RsDevsub("BoxID").value = IIf((.TextMatrix(i, .ColIndex("BoxID"))) = "", Null, .TextMatrix(i, .ColIndex("BoxID")))
                RsDevsub("BankID").value = IIf((.TextMatrix(i, .ColIndex("BankID"))) = "", Null, .TextMatrix(i, .ColIndex("BankID")))
                RsDevsub("BankName").value = IIf((.TextMatrix(i, .ColIndex("BankName"))) = "", Null, .TextMatrix(i, .ColIndex("BankName")))
                RsDevsub("BranchJoinID").value = IIf((.TextMatrix(i, .ColIndex("BranchId"))) = "", Null, .TextMatrix(i, .ColIndex("BranchId")))
                RsDevsub("Remark").value = IIf((.TextMatrix(i, .ColIndex("pymentacount"))) = "", Null, .TextMatrix(i, .ColIndex("pymentacount")))
            RsDevsub.update
      End If
     Next i
     createVoucher
     End With
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
               Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
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
    Dim i As Integer
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("IDGFR").value), "", RsSavRec.Fields("IDGFR").value): ProgressBar1.value = 10
    TxtManualNo.Text = IIf(IsNull(RsSavRec.Fields("ManualNo").value), "", RsSavRec.Fields("ManualNo").value): ProgressBar1.value = 10
    dbRecordDate.value = IIf(IsNull(RsSavRec.Fields("DateM").value), Date, RsSavRec.Fields("DateM").value): ProgressBar1.value = 30
    Txt_DateHigri.value = IIf(IsNull(RsSavRec.Fields("DateH").value), "", RsSavRec.Fields("DateH").value): ProgressBar1.value = 40
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 50
    DcGeneralBox.BoundText = IIf(IsNull(RsSavRec.Fields("GeneralBoxID").value), "", RsSavRec.Fields("GeneralBoxID").value): ProgressBar1.value = 60
    DcboBox.BoundText = IIf(IsNull(RsSavRec.Fields("DelegateID").value), "", RsSavRec.Fields("DelegateID").value): ProgressBar1.value = 70
    Fromdate.value = IIf(IsNull(RsSavRec.Fields("FromDate").value), Date, RsSavRec.Fields("FromDate").value): ProgressBar1.value = 80
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value): ProgressBar1.value = 90
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Explan").value), "", RsSavRec.Fields("Explan").value): ProgressBar1.value = 10
    TotalTXT.Caption = IIf(IsNull(RsSavRec.Fields("TotallVal").value), "", RsSavRec.Fields("TotallVal").value): ProgressBar1.value = 10
        TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value): ProgressBar1.value = 10
            TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value): ProgressBar1.value = 10
            

    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value): ProgressBar1.value = 20
    DcBoxMan.BoundText = IIf(IsNull(RsSavRec.Fields("BoxManID").value), "", RsSavRec.Fields("BoxManID").value): ProgressBar1.value = 20
    
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 80
    FullGridData
    ProgressBar1.Visible = False
    ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
  Sub FullGridData()
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
  sql = "SELECT       dbo.TBLGeneralFRJoin.ID, dbo.TBLGeneralFRJoin.IDGFR, dbo.TBLGeneralFRJoin.NoteSerial, dbo.TBLGeneralFRJoin.ManualNoJoin, dbo.TBLGeneralFRJoin.ReceiptDate,"
  sql = sql + "       dbo.TBLGeneralFRJoin.ReceiptValue, dbo.TBLGeneralFRJoin.AcountCode, dbo.TBLGeneralFRJoin.AcountName, dbo.TBLGeneralFRJoin.ExplanJoin, dbo.TBLGeneralFRJoin.BoxID,"
  sql = sql + "       dbo.TBLGeneralFRJoin.BankID, dbo.TBLGeneralFRJoin.BankName, dbo.TBLGeneralFRJoin.BranchJoinID, dbo.TBLGeneralFRJoin.Remark, dbo.ACCOUNTS.Account_Name,"
  sql = sql + "        dbo.ACCOUNTS.Account_NameEng, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE, dbo.BanksData.BankName AS BankNameBankData, dbo.BanksData.BankNamee,"
  sql = sql + "        dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE, dbo.ACCOUNTS.Account_Code, dbo.TblBranchesData.branch_id"
  sql = sql + "         FROM         dbo.TBLGeneralFRJoin LEFT OUTER JOIN"
  sql = sql + "       dbo.TblBranchesData ON dbo.TBLGeneralFRJoin.BranchJoinID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  sql = sql + "       dbo.BanksData ON dbo.TBLGeneralFRJoin.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
  sql = sql + "       dbo.TblBoxesData ON dbo.TBLGeneralFRJoin.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
  sql = sql + "       dbo.ACCOUNTS ON dbo.TBLGeneralFRJoin.AcountId = dbo.ACCOUNTS.Account_Code"
       
  sql = sql + "  Where (dbo.TBLGeneralFRJoin.IDGFR = " & val(TxtSerial1.Text) & ") "
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.FlexGrid
            For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs1("IDGFR").value), "", Rs1("IDGFR").value)
                   .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(Rs1("NoteSerial").value), "", Rs1("NoteSerial").value)
                   .TextMatrix(i, .ColIndex("ManoulNO")) = IIf(IsNull(Rs1("ManualNoJoin").value), "", Rs1("ManualNoJoin").value)
                   .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs1("ReceiptDate").value), "", Rs1("ReceiptDate").value)
                   .TextMatrix(i, .ColIndex("Valuue")) = IIf(IsNull(Rs1("ReceiptValue").value), "", Rs1("ReceiptValue").value)
                   .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(Rs1("AcountCode").value), "", Rs1("AcountCode").value)
                   .TextMatrix(i, .ColIndex("AcountId")) = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
                   .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(Rs1("ExplanJoin").value), "", Rs1("ExplanJoin").value)
                   .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(Rs1("BoxID").value), "", Rs1("BoxID").value)
                   .TextMatrix(i, .ColIndex("IBan")) = IIf(IsNull(Rs1("BankID").value), "", Rs1("BankID").value)
                   .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(Rs1("BankID").value), "", Rs1("BankID").value)
                   .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(Rs1("branch_id").value), "", Rs1("branch_id").value)
                   .TextMatrix(i, .ColIndex("pymentacount")) = IIf(IsNull(Rs1("Remark").value), "", Rs1("Remark").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    
                   .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
                   .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(Rs1("BoxName").value), "", Rs1("BoxName").value)
                   .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(Rs1("BankName").value), "", Rs1("BankName").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                    Else
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs1("Account_NameEng").value), "", Rs1("Account_NameEng").value)
                    .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(Rs1("BoxNameE").value), "", Rs1("BoxNameE").value)
                    .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(Rs1("BankNamee").value), "", Rs1("BankNamee").value)
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_nameE").value), "", Rs1("branch_nameE").value)
                    End If
                   Rs1.MoveNext
             Next i
              .AutoSize 0, .Cols - 1, False
        End With
        
        Exit Sub
     End Sub
   Function addrow1()
   On Error GoTo ErrTrap
     Dim Msg As String
     Dim LngRow As Long
     Dim LngFindRow As Long
     Dim des As String
     On Error Resume Next
     Dim rs As New ADODB.Recordset
     Dim i As Integer
     
    StrSQL = "SELECT     dbo.Notes.NoteSerial, dbo.Notes.NoteDate, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,"
    StrSQL = StrSQL & "  dbo.ACCOUNTS.Account_Code AS AcountCodee, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.Notes.NoteCashingType,"
    StrSQL = StrSQL & "  dbo.BanksData.BankID, dbo.BanksData.BankName, dbo.Notes.BankName AS BanknameNotes, dbo.BanksData.BankNamee, dbo.BanksData.IBan, dbo.BanksData.account_no,"
    StrSQL = StrSQL & "  dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE, dbo.Notes.Remark, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "  dbo.TblBranchesData.branch_namee, dbo.DOUBLE_ENTREY_VOUCHERS.Collected, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.Notes.ManulaNO, dbo.TblEmployee.Emp_Code,"
    StrSQL = StrSQL & "  dbo.TblEmployee.emp_name , dbo.TblEmployee.fullcode, dbo.TblEmployee.Emp_Namee, dbo.Notes.EmpID"
    StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
    StrSQL = StrSQL & "  dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblEmployee ON dbo.Notes.EmpId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
     
     StrSQL = StrSQL & " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1) And (dbo.Notes.NoteType = 4) And (dbo.DOUBLE_ENTREY_VOUCHERS.Collected Is Null) And (dbo.Notes.EmpId = " & Me.DcboBox.BoundText & ")"
     

     StrSQL = StrSQL + " and (RecordDate >=" & SQLDate(Fromdate.value, True) & ")"
     StrSQL = StrSQL + " and (RecordDate <=" & SQLDate(ToDate.value, True) & ")"
     StrSQL = StrSQL + " and (Notes.NoteSerial Not In (Select NoteSerial from TBLGeneralFRJoin ))"
     
     StrSQL = StrSQL & "  ORDER BY dbo.Notes.NoteSerial"
     rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          
     FlexGrid.Clear flexClearScrollable, flexClearEverything
     FlexGrid.Rows = 1

  For i = 1 To rs.RecordCount
        Me.FlexGrid.Rows = Me.FlexGrid.Rows + 1
        LngRow = Me.FlexGrid.Rows - 1
        Dim NoteCashingType As Integer
  
  With Me.FlexGrid
           .TextMatrix(i, .ColIndex("Ser")) = i
           .TextMatrix(LngRow, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
           .TextMatrix(LngRow, .ColIndex("ManoulNO")) = IIf(IsNull(rs("ManulaNO").value), "", rs("ManulaNO").value)
           .TextMatrix(LngRow, .ColIndex("NoteDate")) = IIf(IsNull(rs("NoteDate").value), "", rs("NoteDate").value)
           .TextMatrix(LngRow, .ColIndex("Valuue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
           .TextMatrix(LngRow, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
           .TextMatrix(LngRow, .ColIndex("AcountId")) = IIf(IsNull(rs("AcountCodee").value), "", rs("AcountCodee").value)
           .TextMatrix(LngRow, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
           .TextMatrix(LngRow, .ColIndex("BankID")) = IIf(IsNull(rs("BankID").value), "", rs("BankID").value)
           .TextMatrix(LngRow, .ColIndex("BranchId")) = IIf(IsNull(rs("branch_id").value), "", rs("branch_id").value)
     If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(LngRow, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            .TextMatrix(LngRow, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
            .TextMatrix(LngRow, .ColIndex("BankName")) = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
            .TextMatrix(LngRow, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
     Else
         .TextMatrix(LngRow, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
         .TextMatrix(LngRow, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxNamee").value), "", rs("BoxNamee").value)
         .TextMatrix(LngRow, .ColIndex("BankName")) = IIf(IsNull(rs("BankNamee").value), "", rs("BankNamee").value)
         .TextMatrix(LngRow, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
    End If
    
     If .TextMatrix(LngRow, .ColIndex("BankName")) = "" Then
               .TextMatrix(LngRow, .ColIndex("BankName")) = IIf(IsNull(rs("BanknameNotes").value), "", rs("BanknameNotes").value)
     End If
     
          .TextMatrix(LngRow, .ColIndex("IBan")) = IIf(IsNull(rs("IBan").value), "", rs("IBan").value)
          .TextMatrix(LngRow, .ColIndex("Remark")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
          .AutoSize 0, .Cols - 1, False
    End With
        rs.MoveNext
   Next i
   TotalGrid
ErrTrap:
End Function
 Private Sub FlexGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 On Error GoTo ErrTrap
      TotalGrid
ErrTrap:
  End Sub
 Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
  On Error GoTo ErrTrap
     TotalGrid
ErrTrap:
 End Sub
  Sub TotalGrid()
  On Error GoTo ErrTrap
    With Me.FlexGrid
      Me.TotalTXT.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Valuue"), .Rows - 1, .ColIndex("Valuue"))
      Me.TotalTXT.Caption = 0
      Dim i As Long
      For i = 1 To FlexGrid.Rows - 1
        If FlexGrid.ValueMatrix(i, FlexGrid.ColIndex("Select")) Then
            Me.TotalTXT.Caption = val(Me.TotalTXT.Caption) + val(FlexGrid.TextMatrix(i, FlexGrid.ColIndex("Valuue")))
        End If
      Next
    End With
ErrTrap:
  End Sub
Private Sub ISButton2_Click()
 ' If txtManualNo.Text = "" Then
 '       If SystemOptions.UserInterface = ArabicInterface Then
 '           MsgBox "⁄›Ê« ...«·—Ã«¡ ﬂ «»… «·—ﬁ„ «·ÌœÊÌ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
 '           txtManualNo.SetFocus
 '            Exit Sub
 '    Else
 '           MsgBox "Write Expenses ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 '          txtManualNo.SetFocus
 '           Exit Sub
 '       End If
 '    End If
 '    ''''''''''''''''''''''''''''''''''''''''''''''''''''
        If DcGeneralBox.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·’‰œÊﬁ «·—∆Ì”Ì", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            DcGeneralBox.SetFocus
           Exit Sub
            Else
            MsgBox "Write Box Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            DcGeneralBox.SetFocus
         End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
       If DcboBox.Text = "" Then
       If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„ «·„‰œÊ»", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            DcboBox.SetFocus
             Exit Sub
             Else
            MsgBox "Write Delegate Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboBox.SetFocus
            Exit Sub
            End If
     End If
      '+++++++++++++++++++++++++++++++++++++++++++++++
        If Dcbranch.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Dcbranch.SetFocus
         End If
      End If
     '+++++++++++++++++++++++++++++++++++++++++++++++
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     addrow1
        End Sub
Private Sub ISButton3_Click()
  On Error Resume Next
    With Me.FlexGrid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub ISButton4_Click()
On Error Resume Next
Me.FlexGrid.Clear flexClearScrollable, flexClearEverything
cleargriid
End Sub

Private Sub ISButton7_Click()
   On Error GoTo ErrTrap
    Dim Selrow As Integer
    Dim DelRow As Integer
    With Me.FlexGrid
                  Selrow = True
                  For DelRow = .FixedRows To .Rows - 1
                  If val(.TextMatrix(DelRow, .ColIndex("Select"))) = vbUnchecked Then
                 .TextMatrix(DelRow, .ColIndex("Select")) = Selrow
                  Else
                  End If
           Next DelRow
    End With
 Exit Sub
ErrTrap:
End Sub
Private Sub ISButton8_Click()
  On Error GoTo ErrTrap
    Dim Selrow As Integer
    Dim DelRow As Integer
    With Me.FlexGrid
                  Selrow = False
                  For DelRow = .FixedRows To .Rows - 1
                  If val(.TextMatrix(DelRow, .ColIndex("Select"))) = True Then
                 .TextMatrix(DelRow, .ColIndex("Select")) = Selrow
                  Else
                  End If
           Next DelRow
    End With
 Exit Sub
ErrTrap:
End Sub
Private Sub Txt_DateHigri_LostFocus()
  VBA.Calendar = vbCalGreg
           ' XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
End Sub

' change date to hj
  Private Sub XPDtbTrans_Change()
  If Me.TxtModFlg.Text <> "R" Then
              'Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
   End If
   End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
       '+++++++++++++++++++++++++++++++++++++++++++++++
        If DcGeneralBox.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·’‰œÊﬁ «·—∆Ì”Ì", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            DcGeneralBox.SetFocus
            Exit Sub
            Else
            MsgBox "Write Box Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            DcGeneralBox.SetFocus
         End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
       If DcboBox.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„ «·„‰œÊ»", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            DcboBox.SetFocus
             Exit Sub
             Else
            MsgBox "Write Delegate Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboBox.SetFocus
            Exit Sub
            End If
     End If
            '+++++++++++++++++++++++++++++++++++++++++++++++
      If Dcbranch.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Dcbranch.SetFocus
         End If
      End If
     '+++++++++++++++++++++++++++++++++++++++++++++++
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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TBLGeneralFundReceipt", "IDGFR", "")
    RsSavRec.AddNew
    RsSavRec.Fields("IDGFR").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
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
    RsSavRec.find "IDGFR=" & RecId, , adSearchForward, 1
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
If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
    
    
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
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
                RsSavRec.find "IDGFR=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                 StrSQL = "Delete From TBLGeneralFRJoin Where IDGFR='" & val(TxtSerial1.Text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
               cleargriid
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
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
        'Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.ISButton6.Enabled = False
        ISButton5.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
              
    ElseIf TxtModFlg.Text = "R" Then
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.ISButton6.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton5.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
    
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.ISButton6.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
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
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
    
    
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
         Me.TxtManualNo.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
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
    cleargriid
    TxtModFlg.Text = "N"
    Me.DCboUserName.BoundText = user_id
    Dcbranch.BoundText = Current_branch
    Me.TxtManualNo.SetFocus
    TxtSerial1.Enabled = False
    Me.FlexGrid.Clear flexClearScrollable, flexClearEverything
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
      cleargriid
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
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
     cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
Private Sub ISButton5_Click()
'On Error GoTo ErrTrap
   If val(Me.TxtSerial1.Text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
Function print_report(Optional TxtSerial1 As String)
'On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    sql = "SELECT    dbo.TBLGeneralFRJoin.ID, dbo.TBLGeneralFRJoin.IDGFR, dbo.TBLGeneralFRJoin.NoteSerial, dbo.TBLGeneralFRJoin.ManualNoJoin, dbo.TBLGeneralFRJoin.ReceiptDate,"
    sql = sql & "     dbo.TBLGeneralFRJoin.ReceiptValue, dbo.TBLGeneralFRJoin.AcountCode, dbo.TBLGeneralFRJoin.AcountName, dbo.TBLGeneralFRJoin.ExplanJoin, dbo.TBLGeneralFRJoin.BoxID,"
    sql = sql & "     dbo.TBLGeneralFRJoin.BankID, dbo.TBLGeneralFRJoin.BankName, dbo.TBLGeneralFRJoin.BranchJoinID, dbo.TBLGeneralFRJoin.Remark, dbo.ACCOUNTS.Account_Name,"
    sql = sql & "     dbo.ACCOUNTS.Account_NameEng, TblBoxesData_1.BoxName, TblBoxesData_1.BoxNameE, dbo.BanksData.BankName AS BankNameBankData, dbo.BanksData.BankNamee,"
    sql = sql & "     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.ACCOUNTS.Account_Code, dbo.TblBranchesData.branch_id, dbo.TBLGeneralFundReceipt.IDGFR AS IDGFRMM,"
    sql = sql & "     dbo.TBLGeneralFundReceipt.ManualNo, dbo.TBLGeneralFundReceipt.DateM, dbo.TBLGeneralFundReceipt.DateH, dbo.TBLGeneralFundReceipt.BranchID,"
    sql = sql & "     dbo.TBLGeneralFundReceipt.GeneralBoxID, dbo.TBLGeneralFundReceipt.DelegateID, dbo.TBLGeneralFundReceipt.FromDate, dbo.TBLGeneralFundReceipt.ToDate,"
    sql = sql & "     dbo.TBLGeneralFundReceipt.Explan, dbo.TBLGeneralFundReceipt.TotallVal, dbo.TBLGeneralFundReceipt.UserID, TblBranchesData_1.branch_id AS branch_idMM,"
    sql = sql & "     TblBranchesData_1.branch_name AS branch_nameMM, TblBranchesData_1.branch_namee AS branch_nameeMM, TblBoxesData_1.BoxID AS BoxIDMM,"
    sql = sql & "      TblBoxesData_1.BoxName AS BoxNameMM, TblBoxesData_1.BoxNameE AS BoxNameEMM, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee"
    sql = sql & "      FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    sql = sql & "     dbo.TBLGeneralFundReceipt ON dbo.TblEmployee.Emp_ID = dbo.TBLGeneralFundReceipt.DelegateID LEFT OUTER JOIN"
    sql = sql & "     dbo.TblBoxesData TblBoxesData_1 ON dbo.TBLGeneralFundReceipt.GeneralBoxID = TblBoxesData_1.BoxID LEFT OUTER JOIN"
    sql = sql & "     dbo.TBLGeneralFRJoin ON dbo.TBLGeneralFundReceipt.IDGFR = dbo.TBLGeneralFRJoin.IDGFR LEFT OUTER JOIN"
    sql = sql & "     dbo.TblBranchesData ON dbo.TBLGeneralFRJoin.BranchJoinID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & "     dbo.BanksData ON dbo.TBLGeneralFRJoin.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
    sql = sql & "     dbo.TblBoxesData TblBoxesData_2 ON dbo.TBLGeneralFRJoin.BoxID = TblBoxesData_2.BoxID LEFT OUTER JOIN"
    sql = sql & "     dbo.ACCOUNTS ON dbo.TBLGeneralFRJoin.AcountId = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    sql = sql & "      dbo.TblBranchesData TblBranchesData_1 ON dbo.TBLGeneralFundReceipt.BranchID = TblBranchesData_1.branch_id"
    sql = sql & "     Where (dbo.TBLGeneralFundReceipt.IDGFR = " & val(Me.TxtSerial1.Text) & ")"
   ' sql = sql & "   WHERE  (dbo.TBLGeneralFundReceipt.IDGFR = 3 )"
    'Where (dbo.TBLGeneralFundReceipt.IDGFR = " & val(TxtSerial1.text) & ")"
   
   If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "GeneralFundReceiptRPT.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "GeneralFundReceiptRPTEE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If RsData.BOF Or RsData.EOF Then
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
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
          xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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
ErrTrap:
  End Function
Private Sub ChangeLang()
    On Error GoTo ErrTrap
   ' form name
    Me.Caption = "General Fund Receipt"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(7).Caption = "Receipt NO."
    Me.lbl(25).Caption = "Manual NO."
    Me.lbl(5).Caption = "Receipt Date"
    Me.lbl(30).Caption = "HJ Date"
    Me.lblBranch.Caption = "Branch"
    Me.lbl(15).Caption = "Main Fund"
    Me.lbl(14).Caption = "Representative"
    Me.lbl(3).Caption = "Remarks"
    Me.lbl(22).Caption = "From Date"
    Me.lbl(23).Caption = "To Date"
    Me.lbl(28).Caption = "Notice"
    Me.lbl(26).Caption = "From this screen you can receive all of the collection Receipts from the Representatives "
    Me.lbl(24).Caption = "Receipt NO."
    Me.Cmd(11).Caption = "Print Receipt"
    Me.CmdAttach.Caption = "Attachments"
    Me.Label2(3).Caption = "Total"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.lbl(20).Caption = "Current Record"
    Me.lbl(37).Caption = "NO. Recordes"
    Me.lbl(29).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    ISButton2.Caption = "ADD"
    ISButton7.Caption = "Select All"
    ISButton8.Caption = "Un Select"
    ISButton3.Caption = "Remove Select"
    ISButton4.Caption = "Remove All"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton5.Caption = "Print"
    ISButton6.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    With Me.FlexGrid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "Receipt NO."
        .TextMatrix(0, .ColIndex("ManoulNO")) = "Manual NO."
        .TextMatrix(0, .ColIndex("NoteDate")) = "Receipt Date"
        .TextMatrix(0, .ColIndex("Valuue")) = "Value"
        .TextMatrix(0, .ColIndex("Account_Serial")) = "Acount Code"
        .TextMatrix(0, .ColIndex("Account_Name")) = "Acount Name"
        .TextMatrix(0, .ColIndex("Remark")) = "Explan"
        .TextMatrix(0, .ColIndex("BoxName")) = "Box"
        .TextMatrix(0, .ColIndex("IBan")) = "Bank Acount"
        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
        .TextMatrix(0, .ColIndex("pymentacount")) = "Remark"
      End With
ErrTrap:
End Sub
Private Sub txtManualNo_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  DcGeneralBox.SetFocus
  End If
ErrTrap:
End Sub
Private Sub DcGeneralBox_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  DcboBox.SetFocus
  End If
ErrTrap:
End Sub
Private Sub DcboBox_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  TxtRemarks.SetFocus
  End If
ErrTrap:
End Sub
Private Sub TxtRemarks_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Call btnSave_Click
  End If
ErrTrap:
End Sub
Private Sub cleargriid()
Me.FlexGrid.Rows = 1
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TBLGeneralFundReceipt"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end







Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = "”‰œ ﬁ»÷ «·⁄«„ " & TxtSerial1 & " · " & DcGeneralBox.Text
Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "TBLGeneralFundReceipt"
Filedname = "IDGFR"
ContNo = TxtSerial1

Notevalue = val(TotalTXT)


                     If Me.TxtModFlg = "N" Then
                                 CreateNotes NoteID, (dbRecordDate.value), val(Dcbranch.BoundText), 1063, Notevalue, NoteSerial, TxtSerial1, tablename, Filedname, ContNo, des, Txt_DateHigri.value
                                     TxtNoteID.Text = NoteID
                                    TxtNoteSerial.Text = NoteSerial
                    Else
                                      If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                    CreateNotes NoteID, (dbRecordDate.value), val(Dcbranch.BoundText), 1063, Notevalue, NoteSerial, TxtSerial1, tablename, Filedname, ContNo, des, Txt_DateHigri.value
                                                       TxtNoteID.Text = NoteID
                                                  TxtNoteSerial.Text = NoteSerial
                                    Else
                                                  sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                  sql = sql & ",NoteSerial1='" & TxtSerial1 & "'"
                                                    sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                     Cn.Execute sql
                                                     
                                       End If
                         
                    End If
'ReLineGrid
CREATE_VOUCHER_GE val(TxtNoteID.Text), val(Dcbranch.BoundText), user_id, dbRecordDate.value
RsSavRec.Resync adAffectCurrent


End Function



Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        

 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
     
    my_branch = BranchID

 
'        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
'GoTo ll
            
  
            StrTempDes = "”‰œ ﬁ»÷ ⁄«„    " & TxtSerial1 & "  ··’‰œÊﬁ «·⁄«„   " & DcGeneralBox.Text
            LngDevNO = LngDevNO + 1
'Notevalue = val(TxtTotalContract) + val(TxtPayAmini) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtEnternet)
Notevalue = 0
 
 Dim Account_Code_dynamic80 As String
      

            
'll:
   LngDevNO = 0
  
 Dim mBoxID  As Long
 If val(TotalTXT.Caption) > 0 Then  '«Ì«„ ‰«ﬁ
       '«·⁄„Ì· œ«∆‰
       Notevalue = Abs(val(TotalTXT.Caption))
   LngDevNO = LngDevNO + 1
        
        mBoxID = val(DcGeneralBox.BoundText)
        
        StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)
       
        
   
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·’‰œÊﬁ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
        LngDevNO = LngDevNO + 1


            
            
  End If
  ''*************

   

  
   
   
'*************************************************
If val(TotalTXT.Caption) > 0 Then
       '«·⁄„Ì· „œÌ‰
       Notevalue = Abs(val(TotalTXT.Caption))
   LngDevNO = LngDevNO + 1
          
        mBoxID = val(DcBoxMan.BoundText)
        
        StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)
       
        
   
        
      
      If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  


            
            
  End If
  
  
   
   
'**************************************************
     
     
     

     
'**************************************************************************’Ì«‰…
     
   
ErrTrap:
End Function





