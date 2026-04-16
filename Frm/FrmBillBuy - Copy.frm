VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBillBuy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›« Ê—… «·„‘ —Ì«   "
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15870
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   100
   Icon            =   "FrmBillBuy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmBillBuy.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   15870
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      TabIndex        =   69
      Top             =   9285
      Width           =   15870
      _ExtentX        =   27993
      _ExtentY        =   26
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   9285
      Left            =   0
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   15870
      _cx             =   27993
      _cy             =   16378
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      AutoSizeChildren=   8
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
      GridRows        =   5
      GridCols        =   6
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmBillBuy.frx":2B2C
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5400
         Left            =   15
         TabIndex        =   21
         Top             =   2895
         Width           =   15825
         _cx             =   27914
         _cy             =   9525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         BackColor       =   14871017
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   $"FrmBillBuy.frx":2BC6
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
         DogEars         =   0   'False
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   1
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "FrmBillBuy.frx":2C50
         Picture(1)      =   "FrmBillBuy.frx":2FEA
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   4935
            Left            =   18270
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AutoSizeChildren=   0
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
               Height          =   4800
               Index           =   19
               Left            =   0
               TabIndex        =   198
               TabStop         =   0   'False
               Top             =   0
               Width           =   15750
               _cx             =   27781
               _cy             =   8467
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               AutoSizeChildren=   0
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
               Begin VB.TextBox TxtManualNo1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   480
                  TabIndex        =   210
                  Top             =   840
                  Width           =   1410
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   206
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.OptionButton BillBasedOn 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ ”‰œ«  «·«” ·«„"
                  Height          =   195
                  Index           =   1
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   205
                  Top             =   600
                  Width           =   4215
               End
               Begin VB.OptionButton BillBasedOn 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "√Ê«„— «·‘—«¡"
                  Height          =   195
                  Index           =   2
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   204
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   4095
               End
               Begin VB.OptionButton BillBasedOn 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "›« Ê—… „‘ —Ì« "
                  Height          =   195
                  Index           =   0
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   203
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   4335
               End
               Begin VB.Frame Frame3 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  ﬁÌœ «·›« Ê—Â"
                  Height          =   1575
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   199
                  Top             =   1800
                  Width           =   4695
                  Begin VB.TextBox TxtNoteSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   1560
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   200
                     Top             =   600
                     Width           =   2625
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     CausesValidation=   0   'False
                     Height          =   375
                     Index           =   10
                     Left            =   120
                     TabIndex        =   201
                     Top             =   600
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                     BackColor       =   14871017
                     FontName        =   "Arial"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ColorButton     =   14871017
                     ColorHighlight  =   16777215
                     ColorHoverText  =   16711680
                     ColorShadow     =   4210752
                     ColorOutline    =   0
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16711680
                     ColorTextShadow =   4210752
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·ﬁÌœ ··›« Ê—Â"
                     Height          =   195
                     Index           =   62
                     Left            =   1920
                     TabIndex        =   202
                     Top             =   240
                     Width           =   2175
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid GRID1 
                  Height          =   2085
                  Left            =   5160
                  TabIndex        =   207
                  Tag             =   "1"
                  Top             =   840
                  Width           =   9255
                  _cx             =   16325
                  _cy             =   3678
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBillBuy.frx":3384
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
               End
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   1725
                  Left            =   5160
                  TabIndex        =   208
                  Tag             =   "1"
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   9255
                  _cx             =   16325
                  _cy             =   3043
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBillBuy.frx":351E
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
               Begin ImpulseButton.ISButton CmdAttach 
                  Height          =   375
                  Left            =   3720
                  TabIndex        =   217
                  Top             =   3360
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   661
                  ButtonPositionImage=   1
                  Caption         =   "«·„—›ﬁ« "
                  BackColor       =   14871017
                  FontName        =   "Arial"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «–‰ «·«” ·«„ «·ÌœÊÌ"
                  Height          =   405
                  Index           =   69
                  Left            =   3060
                  TabIndex        =   211
                  Top             =   900
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›« Ê—Â »‰«¡ ⁄·Ï"
                  Height          =   300
                  Index           =   61
                  Left            =   12240
                  TabIndex        =   209
                  Top             =   120
                  Width           =   2160
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   4935
            Left            =   17970
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AutoSizeChildren=   0
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
               Height          =   4800
               Index           =   9
               Left            =   0
               TabIndex        =   188
               TabStop         =   0   'False
               Top             =   0
               Width           =   15750
               _cx             =   27781
               _cy             =   8467
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               AutoSizeChildren=   0
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
               Begin VB.TextBox TXTFactoryExpensesVat 
                  Alignment       =   2  'Center
                  Height          =   405
                  Left            =   7920
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   335
                  Top             =   3390
                  Width           =   1215
               End
               Begin VB.CheckBox ChSameCurrncy 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰›” ⁄„·… «·›« Ê—…"
                  Height          =   240
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   314
                  Top             =   3120
                  Visible         =   0   'False
                  Width           =   2040
               End
               Begin VB.CheckBox ChAddToTotal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«÷«›… «·Ï «·«Ã„«·Ì"
                  Height          =   240
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   258
                  Top             =   2880
                  Width           =   2040
               End
               Begin VB.TextBox TXTFactoryExpenses 
                  Alignment       =   2  'Center
                  Height          =   405
                  Left            =   7920
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   189
                  Top             =   2910
                  Width           =   1215
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
                  Height          =   2340
                  Left            =   240
                  TabIndex        =   190
                  Top             =   480
                  Width           =   15360
                  _cx             =   27093
                  _cy             =   4128
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Rows            =   1
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBillBuy.frx":3611
                  ScrollTrack     =   0   'False
                  ScrollBars      =   2
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
                  Begin VB.PictureBox PicDes 
                     BorderStyle     =   0  'None
                     Height          =   1635
                     Left            =   240
                     RightToLeft     =   -1  'True
                     ScaleHeight     =   1635
                     ScaleWidth      =   2925
                     TabIndex        =   191
                     Top             =   960
                     Visible         =   0   'False
                     Width           =   2925
                     Begin VB.TextBox TxtDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000018&
                        BorderStyle     =   0  'None
                        Height          =   1125
                        Left            =   30
                        MultiLine       =   -1  'True
                        RightToLeft     =   -1  'True
                        ScrollBars      =   3  'Both
                        TabIndex        =   192
                        Top             =   360
                        Visible         =   0   'False
                        Width           =   2115
                     End
                     Begin VB.Label LblDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H8000000C&
                        Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                        ForeColor       =   &H0000C8FF&
                        Height          =   315
                        Left            =   0
                        RightToLeft     =   -1  'True
                        TabIndex        =   193
                        Top             =   0
                        Width           =   2445
                     End
                  End
                  Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   194
                     ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   2955
                     _cx             =   1973752924
                     _cy             =   1973748268
                     Alignment       =   0
                     Appearance      =   3
                     AutoSearch      =   0   'False
                     BackColor       =   -2147483624
                     BackgroundColor =   -2147483633
                     BorderColor     =   0
                     BorderVisible   =   -1  'True
                     Caption         =   "SmartCombo1"
                     CaptionAlignment=   4
                     CaptionBackColor=   -2147483633
                     BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     CaptionForeColor=   -2147483630
                     CaptionHeight   =   15
                     CaptionOnTop    =   0   'False
                     CaptionMultiLine=   0
                     Checkbox3D      =   0   'False
                     CheckboxAlignment=   5
                     CheckboxBackColor=   16777215
                     CheckboxSize    =   13
                     CheckboxValue   =   0
                     BrowsePictureAlignment=   5
                     BrowsePictureStretchH=   0
                     BrowsePictureStretchV=   0
                     Enabled         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   0
                     Gap             =   0
                     HideSelection   =   -1  'True
                     Locked          =   0   'False
                     MaxLength       =   0
                     MultiLine       =   0
                     OnFocus         =   3
                     PasswordChar    =   ""
                     Picture         =   "FrmBillBuy.frx":3912
                     PictureAlignment=   5
                     PictureBackColor=   -2147483624
                     PictureStretchH =   0
                     PictureStretchV =   0
                     Redraw          =   -1  'True
                     ScrollBar       =   0
                     Style           =   0
                     Text            =   ""
                     UnderLine       =   0   'False
                     Enabled0        =   -1  'True
                     Position0       =   0
                     Tip0            =   "Caption"
                     Visible0        =   0   'False
                     Width0          =   90
                     Enabled1        =   -1  'True
                     Position1       =   1
                     Tip1            =   ""
                     Visible1        =   -1  'True
                     Width1          =   32
                     Enabled2        =   -1  'True
                     Position2       =   2
                     Tip2            =   "Check Box (Space, Ctrl + Space)"
                     Visible2        =   0   'False
                     Width2          =   16
                     Enabled3        =   -1  'True
                     Position3       =   3
                     Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                     Visible3        =   -1  'True
                     Width3          =   145
                     Enabled4        =   -1  'True
                     Position4       =   4
                     Tip4            =   "Left Spinner (Alt + Left)"
                     Visible4        =   0   'False
                     Width4          =   16
                     Enabled5        =   -1  'True
                     Position5       =   5
                     Tip5            =   "Right Spinner (Alt + Right)"
                     Visible5        =   0   'False
                     Width5          =   16
                     Enabled6        =   -1  'True
                     Position6       =   6
                     Tip6            =   "Up Spinner (Ctrl + Up)"
                     Visible6        =   0   'False
                     Width6          =   16
                     Enabled7        =   -1  'True
                     Position7       =   7
                     Tip7            =   "Down Spinner (Ctrl + Down)"
                     Visible7        =   0   'False
                     Width7          =   16
                     Enabled8        =   -1  'True
                     Position8       =   8
                     Tip8            =   "Browse (Alt + Enter)"
                     Visible8        =   0   'False
                     Width8          =   16
                     Enabled9        =   -1  'True
                     Position9       =   9
                     Tip9            =   " (Alt + Down, F4)"
                     Visible9        =   -1  'True
                     Width9          =   16
                     Enabled10       =   -1  'True
                     Position10      =   10
                     Tip10           =   "Right Arrow (Alt + >)"
                     Visible10       =   0   'False
                     Width10         =   16
                  End
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   9
                  Left            =   13200
                  TabIndex        =   195
                  Top             =   2880
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› ”ÿ—"
                  BackColor       =   14871017
                  FontName        =   "Arial"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmBillBuy.frx":3EAC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì  «·„’—Ê›«  «· ﬁœÌ—ÌÂ"
                  Height          =   375
                  Left            =   9120
                  RightToLeft     =   -1  'True
                  TabIndex        =   197
                  Top             =   3000
                  Width           =   2055
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„’—Ê›«   ﬁœÌ—ÌÂ"
                  Height          =   255
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   196
                  Top             =   120
                  Width           =   3855
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4935
            Index           =   0
            Left            =   45
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AutoSizeChildren=   8
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
            GridRows        =   6
            GridCols        =   6
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmBillBuy.frx":4446
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   3570
               Left            =   30
               TabIndex        =   23
               Top             =   1080
               Width           =   15645
               _cx             =   27596
               _cy             =   6297
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Rows            =   2
               Cols            =   33
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmBillBuy.frx":44EA
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
            Begin MSComctlLib.Toolbar TBar 
               Height          =   630
               Left            =   480
               TabIndex        =   24
               Top             =   4665
               Width           =   11535
               _ExtentX        =   20346
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   930
               Index           =   4
               Left            =   30
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   30
               Width           =   15675
               _cx             =   27649
               _cy             =   1640
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Begin VB.CommandButton Command8 
                  Caption         =   "«œ—«Ã"
                  Height          =   285
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   340
                  Top             =   0
                  Width           =   705
               End
               Begin VB.TextBox txtAccountCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4380
                  RightToLeft     =   -1  'True
                  TabIndex        =   338
                  Top             =   0
                  Width           =   1305
               End
               Begin VB.TextBox txtItemCodeSearch 
                  BackColor       =   &H0000FFFF&
                  Height          =   270
                  Left            =   14340
                  TabIndex        =   11
                  Top             =   645
                  Width           =   1275
               End
               Begin VB.TextBox TxtItemsIDes 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   13560
                  TabIndex        =   322
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1695
               End
               Begin VB.TextBox TxtItemCodeB1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   12285
                  TabIndex        =   299
                  Top             =   0
                  Width           =   1695
               End
               Begin VB.TextBox TxtShortName 
                  Height          =   300
                  Left            =   7155
                  TabIndex        =   254
                  Top             =   0
                  Width           =   4110
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0000FFFF&
                  Height          =   330
                  Left            =   705
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   630
                  Width           =   1410
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0000FFFF&
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   3375
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   630
                  Width           =   2820
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0000FFFF&
                  Height          =   330
                  Left            =   2115
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   630
                  Width           =   1260
               End
               Begin VB.ComboBox CboItemCase 
                  BackColor       =   &H0000FFFF&
                  Height          =   315
                  Left            =   6195
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Top             =   630
                  Width           =   2535
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   8745
                  TabIndex        =   13
                  Top             =   630
                  Width           =   3540
                  _ExtentX        =   6244
                  _ExtentY        =   582
                  _Version        =   393216
                  BackColor       =   65535
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   12270
                  TabIndex        =   12
                  Top             =   630
                  Width           =   1980
                  _ExtentX        =   3493
                  _ExtentY        =   582
                  _Version        =   393216
                  BackColor       =   65535
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   450
                  Left            =   0
                  TabIndex        =   18
                  Top             =   630
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   794
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
                  BackColor       =   14871017
                  FontName        =   "Arial"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackStyle       =   0
                  ButtonImage     =   "FrmBillBuy.frx":4A45
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin MSDataListLib.DataCombo cmbAccounts 
                  Height          =   315
                  Left            =   810
                  TabIndex        =   337
                  Top             =   0
                  Width           =   3555
                  _ExtentX        =   6271
                  _ExtentY        =   582
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Õ”«»"
                  Height          =   195
                  Index           =   82
                  Left            =   5310
                  RightToLeft     =   -1  'True
                  TabIndex        =   339
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»«—ﬂÊœ"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   270
                  Index           =   95
                  Left            =   13845
                  TabIndex        =   300
                  Top             =   0
                  Width           =   1410
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»ÕÀ «·”—Ì⁄"
                  Height          =   315
                  Index           =   97
                  Left            =   11145
                  TabIndex        =   255
                  Top             =   30
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   330
                  Index           =   26
                  Left            =   705
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   405
                  Width           =   1410
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   330
                  Index           =   27
                  Left            =   2115
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   405
                  Width           =   1845
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   330
                  Index           =   28
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   405
                  Width           =   3525
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   330
                  Index           =   29
                  Left            =   7485
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   405
                  Width           =   2535
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   330
                  Index           =   30
                  Left            =   9105
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   405
                  Width           =   3510
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   360
                  Index           =   31
                  Left            =   12645
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   345
                  Width           =   1980
               End
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   345
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   244
               Top             =   4305
               Width           =   435
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4935
            Index           =   2
            Left            =   16470
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            BackColor       =   255
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   0
            ChildSpacing    =   0
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
            GridRows        =   3
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmBillBuy.frx":4DDF
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2235
               Index           =   10
               Left            =   0
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   2700
               Width           =   15735
               _cx             =   27755
               _cy             =   3942
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               AutoSizeChildren=   8
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
               GridRows        =   10
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmBillBuy.frx":4E50
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   435
                  Index           =   14
                  Left            =   15
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   767
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Ìﬂ« "
                     Height          =   345
                     Index           =   2
                     Left            =   12645
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     Top             =   90
                     Width           =   1110
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   345
                     Left            =   6390
                     TabIndex        =   29
                     Top             =   90
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   609
                     Caption         =   " ”ÃÌ· «·‘Ìﬂ« "
                     BackColor       =   14871017
                     FontName        =   "Arial"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin MSDataListLib.DataCombo dcbanks 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   82
                     Top             =   0
                     Width           =   2370
                     _ExtentX        =   4180
                     _ExtentY        =   582
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·»‰ﬂ"
                     Height          =   330
                     Left            =   2370
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   150
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   345
                     Index           =   18
                     Left            =   7095
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   90
                     Width           =   825
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈Ã„«·Ï ﬁÌ„… «·‘Ìﬂ« "
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
                     Height          =   345
                     Index           =   16
                     Left            =   7920
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   90
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·‘Ìﬂ« "
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
                     Height          =   345
                     Index           =   17
                     Left            =   10980
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   90
                     Visible         =   0   'False
                     Width           =   1110
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   345
                     Index           =   19
                     Left            =   10005
                     RightToLeft     =   -1  'True
                     TabIndex        =   30
                     Top             =   90
                     Width           =   975
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   1095
                  Left            =   15
                  TabIndex        =   144
                  Top             =   465
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   1931
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBillBuy.frx":4EEE
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
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2415
               Index           =   7
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   285
               Width           =   15735
               _cx             =   27755
               _cy             =   4260
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               AutoSizeChildren=   8
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
               GridRows        =   10
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmBillBuy.frx":5022
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   465
                  Index           =   12
                  Left            =   15
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   820
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   780
                     Index           =   1
                     Left            =   12645
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   -225
                     Width           =   975
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   690
                     Index           =   1
                     Left            =   10290
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   45
                     Width           =   1245
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   840
                     Index           =   1
                     Left            =   7365
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   45
                     Width           =   1815
                  End
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   225
                     Left            =   2775
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   90
                     Width           =   1260
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   345
                     Left            =   135
                     TabIndex        =   40
                     Top             =   90
                     Width           =   1665
                     _ExtentX        =   2937
                     _ExtentY        =   609
                     ButtonPositionImage=   1
                     Caption         =   "Õ”«» «·√ﬁ”«ÿ"
                     BackColor       =   14871017
                     Enabled         =   0   'False
                     FontName        =   "Arial"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmBillBuy.frx":50C0
                     ColorButton     =   14871017
                     ColorHighlight  =   16777215
                     ColorHoverText  =   16711680
                     ColorShadow     =   4210752
                     ColorOutline    =   0
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16711680
                     ColorTextShadow =   4210752
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                     Height          =   1065
                     Index           =   21
                     Left            =   4440
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   1155
                     Width           =   975
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   705
                     Index           =   14
                     Left            =   9315
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   90
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   750
                     Index           =   15
                     Left            =   11820
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   90
                     Width           =   405
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   705
                  Left            =   15
                  TabIndex        =   119
                  Top             =   495
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   1244
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBillBuy.frx":545A
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   225
                  Index           =   13
                  Left            =   15
                  TabIndex        =   128
                  TabStop         =   0   'False
                  Top             =   1935
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   397
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„… «·„»œ∆Ì…"
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
                     Height          =   165
                     Index           =   37
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   143
                     Top             =   30
                     Width           =   1665
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   135
                     RightToLeft     =   -1  'True
                     TabIndex        =   142
                     Top             =   30
                     Width           =   285
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   165
                     Left            =   3195
                     RightToLeft     =   -1  'True
                     TabIndex        =   141
                     Top             =   30
                     Width           =   420
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   12225
                     RightToLeft     =   -1  'True
                     TabIndex        =   140
                     Top             =   30
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»… «·›«∆œ…"
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
                     Height          =   165
                     Index           =   35
                     Left            =   12645
                     RightToLeft     =   -1  'True
                     TabIndex        =   139
                     Top             =   30
                     Width           =   840
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·›«∆œ…"
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
                     Height          =   165
                     Index           =   34
                     Left            =   14460
                     RightToLeft     =   -1  'True
                     TabIndex        =   138
                     Top             =   30
                     Width           =   1110
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   13335
                     RightToLeft     =   -1  'True
                     TabIndex        =   137
                     Top             =   30
                     Width           =   1125
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„»·€ «·ﬂ·Ï"
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
                     Height          =   165
                     Index           =   36
                     Left            =   10695
                     RightToLeft     =   -1  'True
                     TabIndex        =   136
                     Top             =   30
                     Width           =   1395
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   9735
                     RightToLeft     =   -1  'True
                     TabIndex        =   135
                     Top             =   30
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·√ﬁ”«ÿ"
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
                     Height          =   165
                     Index           =   38
                     Left            =   8205
                     RightToLeft     =   -1  'True
                     TabIndex        =   134
                     Top             =   30
                     Width           =   1530
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   7650
                     RightToLeft     =   -1  'True
                     TabIndex        =   133
                     Top             =   30
                     Width           =   555
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ê· ﬁ”ÿ"
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
                     Height          =   165
                     Index           =   40
                     Left            =   6525
                     RightToLeft     =   -1  'True
                     TabIndex        =   132
                     Top             =   30
                     Width           =   1125
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   5145
                     RightToLeft     =   -1  'True
                     TabIndex        =   131
                     Top             =   30
                     Width           =   1380
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "› —… «· ﬁ”Ìÿ"
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
                     Height          =   165
                     Index           =   42
                     Left            =   3615
                     RightToLeft     =   -1  'True
                     TabIndex        =   130
                     Top             =   30
                     Width           =   1530
                  End
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   165
                     Left            =   2085
                     RightToLeft     =   -1  'True
                     TabIndex        =   129
                     Top             =   30
                     Width           =   1110
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   285
               Index           =   11
               Left            =   0
               TabIndex        =   120
               TabStop         =   0   'False
               Top             =   0
               Width           =   15735
               _cx             =   27755
               _cy             =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   120
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   0
                  Width           =   15
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   90
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   60
                  Width           =   15
               End
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   345
                  Index           =   0
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   0
                  Width           =   15
               End
               Begin MSDataListLib.DataCombo DcboCurrency 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   121
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   30
                  _ExtentX        =   53
                  _ExtentY        =   582
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„·…"
                  Height          =   225
                  Index           =   20
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   345
                  Index           =   13
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   90
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   345
                  Index           =   12
                  Left            =   105
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   90
                  Width           =   15
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4935
            Index           =   15
            Left            =   16770
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   12
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
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   1
            ChildSpacing    =   1
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
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   7
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmBillBuy.frx":552B
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   210
               Index           =   18
               Left            =   15
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   1095
               Width           =   15705
               _cx             =   27702
               _cy             =   370
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   270
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   420
                  Left            =   105
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   49
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   390
                  Index           =   43
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   47
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   210
               Index           =   17
               Left            =   15
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   1095
               Visible         =   0   'False
               Width           =   15705
               _cx             =   27702
               _cy             =   370
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   180
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   45
                  Width           =   15
               End
               Begin VB.TextBox TxtTaxStampValue 
                  Alignment       =   1  'Right Justify
                  Height          =   225
                  Left            =   105
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   33
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   195
                  Index           =   41
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   45
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "$"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   48
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   45
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   795
               Index           =   16
               Left            =   15
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   285
               Visible         =   0   'False
               Width           =   15705
               _cx             =   27702
               _cy             =   1402
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   585
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   30
               End
               Begin VB.TextBox TxtTaxAddValue 
                  Alignment       =   1  'Right Justify
                  Height          =   480
                  Left            =   105
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   75
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   32
                  Left            =   15
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   120
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   390
                  Index           =   39
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   46
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   255
               Index           =   8
               Left            =   15
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   15
               Visible         =   0   'False
               Width           =   15705
               _cx             =   27702
               _cy             =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   5
               MousePointer    =   0
               Version         =   801
               BackColor       =   14871017
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   0
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
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   240
                  Left            =   105
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   45
                  Width           =   15
               End
               Begin VB.CheckBox XPChkTAX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   150
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   30
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   150
                  Index           =   25
                  Left            =   15
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   90
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   165
                  Index           =   22
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Index           =   45
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   15
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈÷«›… √Ì… „·«ÕŸ«  ⁄·Ï «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   44
               Left            =   15
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   1095
               Width           =   15705
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   4935
            Left            =   17070
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AutoSizeChildren=   0
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
            Begin VB.Frame Frame1 
               Height          =   4800
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   0
               Width           =   15750
               Begin VB.TextBox TxtVATCustoms1 
                  Height          =   405
                  Left            =   1200
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   318
                  Top             =   3000
                  Width           =   1935
               End
               Begin VB.CommandButton Command6 
                  Caption         =   "Command6"
                  Height          =   375
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  TabIndex        =   171
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.TextBox TXTToTAlELSHahn 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   1200
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   170
                  Text            =   "0"
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.TextBox Txt_EXport 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   9600
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   169
                  Top             =   2880
                  Width           =   1890
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "⁄—÷ «·„’—Ê›« "
                  Height          =   480
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   168
                  Top             =   3240
                  Width           =   2220
               End
               Begin VSFlex8UCtl.VSFlexGrid Grid 
                  Height          =   2325
                  Left            =   240
                  TabIndex        =   172
                  Tag             =   "1"
                  Top             =   480
                  Width           =   15255
                  _cx             =   26908
                  _cy             =   4101
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Rows            =   50
                  Cols            =   13
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBillBuy.frx":55A2
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  Caption         =   "ﬁÌ„… «· VAT ··Ã„«—ﬂ"
                  Height          =   255
                  Index           =   81
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   319
                  Top             =   3000
                  Width           =   1455
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·„’—Ê›« "
                  Height          =   285
                  Index           =   60
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   175
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1800
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰œ«  «·’—›"
                  Height          =   285
                  Index           =   54
                  Left            =   12600
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   240
                  Width           =   2640
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì  ”‰œ«  «·„’—Ê›« "
                  Height          =   285
                  Index           =   51
                  Left            =   11670
                  RightToLeft     =   -1  'True
                  TabIndex        =   173
                  Top             =   3000
                  Width           =   1920
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   4935
            Left            =   17370
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AutoSizeChildren=   0
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
               Height          =   4800
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   0
               Width           =   15750
               Begin VB.CommandButton Command7 
                  Caption         =   "Command7"
                  Height          =   195
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   178
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "⁄—÷ ÿ·»«  «·‘—«¡"
                  Height          =   480
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   177
                  Top             =   3000
                  Width           =   2010
               End
               Begin VSFlex8UCtl.VSFlexGrid GRID2 
                  Height          =   2205
                  Left            =   5040
                  TabIndex        =   179
                  Tag             =   "1"
                  Top             =   600
                  Width           =   7695
                  _cx             =   13573
                  _cy             =   3889
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBillBuy.frx":57BD
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÿ·»«  «·‘—«¡  Ê «·›Ê« Ì— «·„»œ∆ÌÂ"
                  Height          =   285
                  Index           =   57
                  Left            =   8280
                  RightToLeft     =   -1  'True
                  TabIndex        =   180
                  Top             =   240
                  Width           =   4440
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   4935
            Left            =   17670
            TabIndex        =   164
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AutoSizeChildren=   0
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
               Height          =   6120
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   181
               Top             =   0
               Width           =   15750
               Begin VB.TextBox TxtVATCustoms 
                  Height          =   405
                  Left            =   240
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   316
                  Top             =   3000
                  Width           =   1770
               End
               Begin VB.CommandButton Command5 
                  Caption         =   " Œ’Ì’"
                  Height          =   480
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   184
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   2220
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "⁄—÷ «·›Ê« Ì— «·„«·Ì…"
                  Height          =   480
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   183
                  Top             =   2880
                  Width           =   2220
               End
               Begin VB.TextBox txt_total_bill 
                  Height          =   405
                  Left            =   10200
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   182
                  Top             =   2880
                  Width           =   1770
               End
               Begin VSFlex8UCtl.VSFlexGrid grid4 
                  Height          =   2325
                  Left            =   240
                  TabIndex        =   185
                  Tag             =   "1"
                  Top             =   480
                  Width           =   14055
                  _cx             =   24791
                  _cy             =   4101
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Rows            =   50
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBillBuy.frx":58C5
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  Caption         =   "ﬁÌ„… «· VAT ··Ã„«—ﬂ"
                  Height          =   255
                  Index           =   80
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   317
                  Top             =   3120
                  Width           =   1455
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›Ê« Ì— «·„«·ÌÂ"
                  Height          =   285
                  Index           =   64
                  Left            =   10800
                  RightToLeft     =   -1  'True
                  TabIndex        =   187
                  Top             =   120
                  Width           =   3120
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·›Ê« Ì—"
                  Height          =   285
                  Index           =   59
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   3000
                  Width           =   2040
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   4935
            Left            =   18570
            TabIndex        =   301
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Begin VB.CheckBox VstReverse 
               Alignment       =   1  'Right Justify
               Caption         =   "«·«Õ ”«» «·⁄ﬂ”Ì"
               Height          =   255
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   331
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.TextBox txtManulaVat 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   240
               TabIndex        =   324
               Top             =   0
               Width           =   1215
            End
            Begin VB.CheckBox ChecVAT 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   270
               Left            =   12315
               RightToLeft     =   -1  'True
               TabIndex        =   308
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox TxtValueAdded 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   306
               Top             =   4560
               Width           =   2610
            End
            Begin VSFlex8UCtl.VSFlexGrid VatGrid 
               Height          =   3945
               Left            =   135
               TabIndex        =   302
               Tag             =   "1"
               Top             =   480
               Width           =   15600
               _cx             =   27517
               _cy             =   6959
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmBillBuy.frx":5A89
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
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«œŒ«· «·‰”»… «·ÌœÊÌ…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   148
               Left            =   1560
               TabIndex        =   325
               Top             =   120
               Width           =   1800
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·«Ã„«·Ì"
               Height          =   300
               Index           =   76
               Left            =   12045
               TabIndex        =   307
               Top             =   4560
               Width           =   540
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "«’‰«› «·ﬁÌ„… «·„÷«›…"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   12180
               TabIndex        =   303
               Top             =   120
               Width           =   3015
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   4935
            Index           =   1
            Left            =   18870
            TabIndex        =   348
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   8705
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   180
               Left            =   14430
               RightToLeft     =   -1  'True
               TabIndex        =   349
               Top             =   270
               Width           =   975
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
               Height          =   3525
               Left            =   180
               TabIndex        =   350
               Top             =   540
               Width           =   17145
               _cx             =   30242
               _cy             =   6218
               Appearance      =   2
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmBillBuy.frx":5BBB
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
               ExplorerBar     =   1
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
               Caption         =   "«·«Ã„«·Ì »œÊ‰ ›« "
               Height          =   240
               Index           =   111
               Left            =   10575
               TabIndex        =   356
               Top             =   4245
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   240
               Index           =   112
               Left            =   8025
               TabIndex        =   355
               Top             =   4245
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì «·›« "
               Height          =   240
               Index           =   113
               Left            =   6015
               TabIndex        =   354
               Top             =   4245
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   240
               Index           =   114
               Left            =   4005
               TabIndex        =   353
               Top             =   4245
               Width           =   1650
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   240
               Index           =   115
               Left            =   210
               TabIndex        =   352
               Top             =   4245
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’«›Ì"
               Height          =   240
               Index           =   116
               Left            =   2370
               TabIndex        =   351
               Top             =   4245
               Width           =   1455
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   585
         Index           =   6
         Left            =   15
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   15
         Width           =   15840
         _cx             =   27940
         _cy             =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   24
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
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "›« Ê—… «·„‘ —Ì«   "
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   0
         ChildSpacing    =   0
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   7
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
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»«·„Ê—œ"
            Height          =   225
            Index           =   2
            Left            =   7050
            TabIndex        =   343
            Top             =   330
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»œÊ‰ ﬁÌœ"
            Height          =   225
            Index           =   1
            Left            =   6000
            TabIndex        =   333
            Top             =   240
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»«·›—⁄"
            Height          =   225
            Index           =   0
            Left            =   7050
            TabIndex        =   332
            Top             =   60
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4920
            PasswordChar    =   "*"
            TabIndex        =   327
            Top             =   240
            Width           =   945
         End
         Begin VB.CommandButton cmdReSave 
            Caption         =   "÷»ÿ «·Õ—ﬂ« "
            Height          =   285
            Left            =   11280
            TabIndex        =   326
            Top             =   240
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2955
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   315
            Left            =   8565
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   105
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   345
            Left            =   9345
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   90
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   10575
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   150
            Visible         =   0   'False
            Width           =   630
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   390
            Left            =   5040
            TabIndex        =   75
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":5ED4
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2700
            TabIndex        =   76
            Top             =   105
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":626E
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   3
            Left            =   1440
            TabIndex        =   77
            Top             =   105
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":6608
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   1
            Left            =   3810
            TabIndex        =   78
            Top             =   105
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":69A2
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   2
            Left            =   165
            TabIndex        =   79
            Top             =   105
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":6D3C
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   390
            Left            =   5955
            TabIndex        =   80
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":70D6
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   480
            Left            =   11295
            TabIndex        =   81
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":7670
            ButtonImageHover=   "FrmBillBuy.frx":834A
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker txtFromDateReSave 
            Height          =   315
            Left            =   9750
            TabIndex        =   328
            Top             =   240
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            Format          =   238944257
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtToDateReSave 
            Height          =   315
            Left            =   8010
            TabIndex        =   329
            Top             =   240
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            Format          =   238944257
            CurrentDate     =   38784
         End
         Begin VB.Label LBLGross 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   315
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   11880
            Picture         =   "FrmBillBuy.frx":9024
            Stretch         =   -1  'True
            Top             =   0
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
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
            Index           =   67
            Left            =   4935
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   120
            Width           =   6855
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2265
         Index           =   5
         Left            =   15
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   615
         Width           =   15810
         _cx             =   27887
         _cy             =   3995
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VB.TextBox txtAdvPay 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2655
            TabIndex        =   344
            Top             =   1290
            Width           =   1170
         End
         Begin VB.TextBox txtContainerNo 
            BackColor       =   &H0000FFFF&
            Height          =   345
            Left            =   7080
            TabIndex        =   341
            Top             =   360
            Width           =   1140
         End
         Begin VB.CheckBox chkTaxExempt 
            Alignment       =   1  'Right Justify
            Caption         =   "„⁄›Ì"
            Height          =   315
            Left            =   4320
            TabIndex        =   336
            Top             =   1140
            Width           =   975
         End
         Begin VB.CommandButton cmdInsertItems 
            Caption         =   "...."
            Height          =   285
            Left            =   9675
            TabIndex        =   334
            Top             =   780
            Width           =   615
         End
         Begin VB.TextBox poTransaction_ID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   330
            Text            =   "Text3"
            Top             =   1680
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox TxtVATNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8025
            MaxLength       =   55
            RightToLeft     =   -1  'True
            TabIndex        =   312
            Top             =   780
            Width           =   990
         End
         Begin VB.TextBox TXtResonVAT 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   310
            Top             =   1800
            Width           =   3360
         End
         Begin VB.ComboBox Dcbtyp 
            BackColor       =   &H0000FFFF&
            Height          =   330
            ItemData        =   "FrmBillBuy.frx":CC8C
            Left            =   6795
            List            =   "FrmBillBuy.frx":CC8E
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Text            =   "Dcbtyp"
            Top             =   1875
            Width           =   1410
         End
         Begin VB.CheckBox ChkCompsBill 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›« Ê—… „Ã„⁄Â"
            Height          =   240
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   1545
            Width           =   2040
         End
         Begin VB.TextBox TxtPhone 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   10470
            TabIndex        =   219
            Top             =   1110
            Width           =   1650
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   13260
            TabIndex        =   215
            Top             =   1410
            Width           =   1305
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   825
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   300
            Width           =   2775
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   540
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   214
               Top             =   240
               Width           =   2520
            End
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   12630
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   1110
            Width           =   1935
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   315
            Left            =   13260
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   780
            Width           =   1305
         End
         Begin VB.TextBox TxtLCNO 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   720
            MaxLength       =   55
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   375
            Width           =   1125
         End
         Begin VB.TextBox txtManualNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   300
            Left            =   11130
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   0
            Width           =   1140
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   330
            ItemData        =   "FrmBillBuy.frx":CC90
            Left            =   11100
            List            =   "FrmBillBuy.frx":CC92
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   5595
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   -180
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   300
            Left            =   13275
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   0
            Width           =   1305
         End
         Begin VB.TextBox TXT_order_no 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   330
            Left            =   8835
            MaxLength       =   55
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   360
            Width           =   1230
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            ItemData        =   "FrmBillBuy.frx":CC94
            Left            =   5040
            List            =   "FrmBillBuy.frx":CC96
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   375
            Width           =   1230
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   330
            Left            =   7860
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5880
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1110
            Width           =   1065
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   15780
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1710
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   10320
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   -240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   315
            Left            =   13260
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   1920
            Width           =   1305
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   210
            Left            =   13365
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   855
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
            Caption         =   " ÕÊÌ· «·Ï  «–‰ «÷«›… "
            Height          =   240
            Left            =   -345
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1230
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   195
            Left            =   675
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   1230
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox txt_Currency_rate 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   285
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Text            =   "1"
            Top             =   15
            Width           =   960
         End
         Begin MSDataListLib.DataCombo DcCurrency 
            Height          =   315
            Left            =   1245
            TabIndex        =   87
            Top             =   0
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   10305
            TabIndex        =   7
            Top             =   780
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   65535
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   9690
            TabIndex        =   9
            Top             =   1875
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   65535
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   330
            Left            =   13260
            TabIndex        =   4
            Top             =   375
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            _Version        =   393216
            Format          =   238944257
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   270
            Left            =   14475
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   825
            Visible         =   0   'False
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":CC98
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCproject 
            Height          =   330
            Left            =   4155
            TabIndex        =   107
            Top             =   1500
            Width           =   4050
            _ExtentX        =   7144
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTArrivalDate 
            Height          =   300
            Left            =   3000
            TabIndex        =   109
            Top             =   780
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            Format          =   238944257
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   330
            Left            =   7530
            TabIndex        =   2
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   65535
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   2520
            TabIndex        =   148
            Top             =   360
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   3030
            TabIndex        =   3
            Top             =   30
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   65535
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   330
            Left            =   315
            TabIndex        =   156
            Top             =   375
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷"
            BackColor       =   12632256
            ForeColor       =   16711680
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   12632256
            ColorHighlight  =   16777215
            ColorHoverText  =   255
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   255
            ColorTextShadow =   -2147483637
         End
         Begin MSComCtl2.DTPicker DtpDelayDate 
            Height          =   300
            Left            =   5520
            TabIndex        =   157
            Top             =   780
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            Format          =   238944257
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   9690
            TabIndex        =   216
            Top             =   1410
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   360
            Left            =   10020
            TabIndex        =   218
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   1065
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ButtonImage     =   "FrmBillBuy.frx":D032
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker BLDate 
            Height          =   300
            Left            =   4155
            TabIndex        =   320
            Top             =   1875
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            Format          =   238944257
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   2640
            TabIndex        =   323
            Top             =   360
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ›⁄Â „"
            Height          =   285
            Index           =   83
            Left            =   3510
            TabIndex        =   347
            Top             =   1395
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   285
            Index           =   96
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   346
            Top             =   1395
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Index           =   110
            Left            =   0
            TabIndex        =   345
            Top             =   1305
            Width           =   1665
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«—«„ﬂÊ"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   7740
            TabIndex        =   342
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·»Ê·Ì’…"
            Height          =   285
            Left            =   5445
            RightToLeft     =   -1  'True
            TabIndex        =   321
            Top             =   1875
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ VAT"
            Height          =   285
            Index           =   79
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   313
            Top             =   780
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·”»»"
            Height          =   285
            Index           =   78
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   311
            Top             =   1920
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·ﬁÌ„… «·„÷«›…"
            Height          =   285
            Index           =   77
            Left            =   8310
            RightToLeft     =   -1  'True
            TabIndex        =   309
            Top             =   1875
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ·Ì›Ê‰"
            Height          =   300
            Index           =   84
            Left            =   11940
            TabIndex        =   220
            Top             =   1170
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„‰œÊ»"
            Height          =   225
            Index           =   72
            Left            =   14610
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   1500
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   180
            Index           =   71
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   780
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ «·‰ﬁœÌ"
            Height          =   195
            Index           =   70
            Left            =   14610
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            Height          =   285
            Left            =   6765
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   780
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«⁄ „«œ"
            Height          =   210
            Index           =   68
            Left            =   1830
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   390
            Width           =   660
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·›« Ê—…"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5700
            TabIndex        =   153
            Top             =   60
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›« Ê—… «·„Ê—œ"
            Height          =   195
            Index           =   53
            Left            =   12210
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰œÊﬁ"
            Height          =   240
            Index           =   2
            Left            =   4365
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   375
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   285
            Index           =   66
            Left            =   10110
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   375
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   285
            Index           =   65
            Left            =   12375
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   375
            Width           =   795
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10530
            TabIndex        =   111
            Top             =   75
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  Ê’Ê· «·‘Õ‰Â"
            Height          =   300
            Index           =   56
            Left            =   4275
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   780
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‘—Ê⁄"
            Height          =   285
            Index           =   58
            Left            =   8310
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·Œ’„"
            Height          =   300
            Index           =   11
            Left            =   6930
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   1170
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·”œ«œ"
            Height          =   285
            Index           =   10
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   375
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·›« Ê—…"
            Height          =   285
            Index           =   7
            Left            =   14550
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   375
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "««·⁄„·…"
            Height          =   195
            Index           =   9
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   75
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   195
            Index           =   8
            Left            =   14625
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   15
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„Ê—œ"
            Height          =   195
            Index           =   6
            Left            =   14625
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   750
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   285
            Index           =   5
            Left            =   9150
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   1110
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„Œ“‰"
            Height          =   300
            Index           =   4
            Left            =   14505
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   1875
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   55
            Left            =   5685
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1170
            Width           =   210
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «–‰ «·«÷«›…"
            Height          =   225
            Index           =   52
            Left            =   -1065
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   1050
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   8310
         Width           =   15840
         _cx             =   27940
         _cy             =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtSum 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Height          =   375
            Left            =   5565
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   223
            TabStop         =   0   'False
            Top             =   -90
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3105
            TabIndex        =   224
            Top             =   90
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LblValueAdded 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   7560
            TabIndex        =   305
            Top             =   30
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„Â „÷«›…"
            Height          =   255
            Index           =   75
            Left            =   8355
            RightToLeft     =   -1  'True
            TabIndex        =   304
            Top             =   90
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«÷«›« "
            Height          =   255
            Index           =   74
            Left            =   10155
            RightToLeft     =   -1  'True
            TabIndex        =   257
            Top             =   90
            Width           =   495
         End
         Begin VB.Label TxtAddValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   9360
            TabIndex        =   256
            Top             =   0
            Width           =   750
         End
         Begin VB.Label LblTotalview 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5460
            TabIndex        =   243
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label LblDiscountsTotalview 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   12030
            TabIndex        =   242
            Top             =   0
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label LblTotalAllview 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   525
            Left            =   13380
            TabIndex        =   241
            Top             =   -120
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   195
            Index           =   3
            Left            =   15180
            RightToLeft     =   -1  'True
            TabIndex        =   240
            Top             =   150
            Width           =   570
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   1260
            RightToLeft     =   -1  'True
            TabIndex        =   239
            Top             =   90
            Width           =   735
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   238
            Top             =   90
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   255
            Index           =   1
            Left            =   4455
            RightToLeft     =   -1  'True
            TabIndex        =   237
            Top             =   90
            Width           =   690
         End
         Begin VB.Label LblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   5460
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   255
            Index           =   23
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   235
            Top             =   90
            Width           =   150
         End
         Begin VB.Label LblTotalAll 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   13605
            RightToLeft     =   -1  'True
            TabIndex        =   234
            Top             =   30
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   255
            Index           =   50
            Left            =   12795
            RightToLeft     =   -1  'True
            TabIndex        =   233
            Top             =   90
            Width           =   600
         End
         Begin VB.Label LblDiscountsTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   12030
            RightToLeft     =   -1  'True
            TabIndex        =   232
            Top             =   0
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   255
            Index           =   24
            Left            =   6900
            RightToLeft     =   -1  'True
            TabIndex        =   231
            Top             =   90
            Width           =   600
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   6180
            TabIndex        =   230
            Top             =   0
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   660
            Index           =   63
            Left            =   7935
            TabIndex        =   229
            Top             =   615
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   495
            Index           =   0
            Left            =   2205
            RightToLeft     =   -1  'True
            TabIndex        =   228
            Top             =   90
            Width           =   750
         End
         Begin VB.Label LblCommision 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   10695
            TabIndex        =   227
            Top             =   0
            Width           =   750
         End
         Begin VB.Label LblCommisionV 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   226
            Top             =   -255
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„Ê·« "
            Height          =   255
            Index           =   73
            Left            =   11370
            RightToLeft     =   -1  'True
            TabIndex        =   225
            Top             =   90
            Width           =   615
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   510
         Index           =   0
         Left            =   15
         TabIndex        =   245
         TabStop         =   0   'False
         Top             =   8760
         Width           =   15840
         _cx             =   27940
         _cy             =   900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   1
            Left            =   12255
            TabIndex        =   246
            Top             =   90
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   2
            Left            =   10470
            TabIndex        =   19
            Top             =   90
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   3
            Left            =   8760
            TabIndex        =   247
            Top             =   90
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   4
            Left            =   7020
            TabIndex        =   248
            Top             =   90
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   5
            Left            =   5265
            TabIndex        =   249
            Top             =   90
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   6
            Left            =   90
            TabIndex        =   250
            Top             =   90
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   7
            Left            =   3600
            TabIndex        =   251
            Top             =   90
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   330
            Left            =   1560
            TabIndex        =   252
            Top             =   90
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”ŒÂ „„«À·Â"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   0
            Left            =   14010
            TabIndex        =   253
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   330
            Left            =   2640
            TabIndex        =   259
            Top             =   90
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
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
   Begin VB.Frame FramePay 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„»·€ «·„œ›Ê⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   510
      RightToLeft     =   -1  'True
      TabIndex        =   260
      Top             =   420
      Visible         =   0   'False
      Width           =   14415
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Height          =   5055
         Left            =   90
         TabIndex        =   278
         Top             =   180
         Width           =   5535
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   0
            Left            =   4320
            TabIndex        =   279
            Top             =   3970
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":D42F
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   1
            Left            =   2160
            TabIndex        =   280
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":DBEF
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   2
            Left            =   3240
            TabIndex        =   281
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":E1F1
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   3
            Left            =   4320
            TabIndex        =   282
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":E9D8
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   4
            Left            =   2160
            TabIndex        =   283
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":F1ED
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   5
            Left            =   3240
            TabIndex        =   284
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":F978
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   6
            Left            =   4320
            TabIndex        =   285
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":10137
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   7
            Left            =   2160
            TabIndex        =   286
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":108D1
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   8
            Left            =   3240
            TabIndex        =   287
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":10FD4
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   9
            Left            =   4320
            TabIndex        =   288
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":117EF
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   10
            Left            =   3240
            TabIndex        =   289
            Top             =   3970
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":11F7E
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   11
            Left            =   2160
            TabIndex        =   290
            Top             =   3970
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":12AC5
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   12
            Left            =   120
            TabIndex        =   291
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":12FB7
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   13
            Left            =   1200
            TabIndex        =   292
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":1381E
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   2895
            Index           =   14
            Left            =   120
            TabIndex        =   293
            Top             =   2040
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   5106
            Caption         =   ""
            BackColor       =   16777215
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuy.frx":13F2F
            ButtonImageDisabled=   "FrmBillBuy.frx":152DD
            ColorButton     =   16777215
         End
         Begin VB.Label LBLPayVal 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   960
            TabIndex        =   298
            Top             =   360
            Width           =   3375
         End
         Begin VB.Image Image13 
            Height          =   1035
            Left            =   120
            Picture         =   "FrmBillBuy.frx":15678
            Stretch         =   -1  'True
            Top             =   120
            Width           =   5295
         End
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "1500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   3000
         TabIndex        =   277
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   4200
         TabIndex        =   276
         Top             =   7320
         Width           =   1335
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   1935
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   269
         Top             =   4440
         Width           =   3840
         Begin VB.TextBox TxtNetValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   600
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   272
            Top             =   240
            Width           =   2460
         End
         Begin VB.TextBox TxtPayedValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   555
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   271
            Top             =   840
            Width           =   2445
         End
         Begin VB.TextBox TxtRemainValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   555
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   270
            Top             =   1320
            Width           =   2445
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   101
            Left            =   2640
            TabIndex        =   275
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„œ›Ê⁄"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   100
            Left            =   2640
            TabIndex        =   274
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ »ﬁÌ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   99
            Left            =   2640
            TabIndex        =   273
            Top             =   1440
            Width           =   855
         End
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   1560
         TabIndex        =   268
         Top             =   7320
         Width           =   1455
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   240
         TabIndex        =   267
         Top             =   7320
         Width           =   1335
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   4200
         TabIndex        =   266
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   3000
         TabIndex        =   265
         Top             =   6720
         Width           =   1215
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   1560
         TabIndex        =   264
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   263
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   8640
         TabIndex        =   262
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   5760
         TabIndex        =   261
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin ImpulseButton.ISButton CMDPAy 
         Height          =   1215
         Left            =   240
         TabIndex        =   294
         Top             =   5450
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2143
         Caption         =   "”œ«œ"
         ForeColor       =   16777215
         FontSize        =   24
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBillBuy.frx":15A2E
         ColorHoverText  =   16777215
         ColorToggledText=   16777215
         ColorToggledHoverText=   16777215
         AlignmentIgnoreImage=   -1  'True
      End
      Begin VSFlex8UCtl.VSFlexGrid Grid22 
         Height          =   3885
         Left            =   5760
         TabIndex        =   295
         Top             =   600
         Width           =   8325
         _cx             =   14684
         _cy             =   6853
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483640
         ForeColor       =   65280
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483641
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483640
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
         Rows            =   6
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   650
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmBillBuy.frx":15FA8
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
      Begin VB.Label lblexit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   90
         Left            =   9120
         TabIndex        =   297
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   10800
         TabIndex        =   296
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmBillBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim NewGrid As ClsGrid
Dim TTP As clstooltip
Dim BuyReport As ClsBuyReport
Dim cSearchDcbo(3) As clsDCboSearch
Dim OtherInformation As New ClsGLOther
Public BolPrint As Boolean
Dim WithEvents m_MnuShowNewItemsPrices As Menu
Attribute m_MnuShowNewItemsPrices.VB_VarHelpID = -1
Dim WithEvents m_MenuViewList As Menu
Attribute m_MenuViewList.VB_VarHelpID = -1
Dim WithEvents m_MenuShowItemCostEffect As Menu
Attribute m_MenuShowItemCostEffect.VB_VarHelpID = -1
Dim WithEvents m_FrmSearch As Form
Attribute m_FrmSearch.VB_VarHelpID = -1
Dim IsClicKCommand4 As Boolean
Dim bank_account As String
Public invoiceSerach As Boolean
Dim IsVouc         As Boolean
 Dim MintDone As Integer
Dim general_noteid As Long
Dim CurrentVoucherNo As String
Dim CurrentVoucherSerialNo As String
Dim DateChanged As Boolean
Dim StroreChanged As Boolean
Dim TxtNoteSerial1V As String
Dim Account_Code_dynamic101 As String
Dim Account_Code_dynamic102 As String
Dim mValue As Double
Dim mIsFinishSave As Boolean
Dim IsSaveWithOutMsg As Boolean
Dim mIsStart As Boolean

Public Sub chkTaxExempt_Click()
 Dim i As Integer
If Me.TxtModFlg.Text <> "R" Then
If chkTaxExempt.value = vbChecked Then
    ChecVAT.value = vbUnchecked
Else
    ChecVAT.value = vbChecked
End If
    If ChecVAT.value = vbChecked Then

        With Me.VatGrid
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = True
            Next i

        End With

    Else

        With Me.VatGrid

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = False
            Next i

        End With

    End If
    RelinVatGrid
    End If
End Sub

Private Sub cmbAccount_Click(Area As Integer)

End Sub

Private Sub cmbAccounts_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        Account_search.show
        'Account_search.mIndex = Index
        Account_search.case_id = 7897289
    End If
End Sub

Private Sub Command8_Click()
  Dim Num  As Long
  Dim ClsAcc As ClsAccounts
     Set ClsAcc = New ClsAccounts

                   ' .TextMatrix(row, .ColIndex("Account_Code")) = StrAccountCode
 
                    
  For Num = 1 To FG.rows - 1 'RsDetails.RecordCount
    
        
            FG.TextMatrix(Num, FG.ColIndex("Account_Name")) = Trim(Me.cmbAccounts.Text)
            FG.TextMatrix(Num, FG.ColIndex("Account_Code")) = Trim(Me.cmbAccounts.BoundText)
            FG.TextMatrix(Num, FG.ColIndex("Account_Serial2")) = ClsAcc.Get_Account_Serial(Me.cmbAccounts.BoundText)
       
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num
End Sub

Private Sub txtContainerNo_Change()
 If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    s = "Select NoteSerial1 from Transactions where ContainerNo = '" & Trim(txtContainerNo) & "' and Transaction_Type = 29"
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        CBoBasedON.ListIndex = 1
        TXT_order_no = rsDummy!NoteSerial1 & ""
        
        
        
    End If
 End If
End Sub

Private Sub txtItemCodeSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DCboItemsCode.SetFocus
End If
End Sub

Private Sub cmdInsertItems_Click()
    
        '*******************
        
        Dim StrSQL  As String
        Dim rs2  As ADODB.Recordset
        Dim mUnitPurPrice As Double
         Dim mUnitId As Long
         Dim LngItemID As Long
        Dim mName As String
        Dim mUnitName As String
        StrSQL = " SELECT TblItemsUnits.UnitPurPrice,TblItemsUnits.UnitSalesPrice,dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID,TblItems.* "
        StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
        StrSQL = StrSQL & "                   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID INNER JOIN"
        StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
        StrSQL = StrSQL & "  Where DefaultSupplier = " & val(DBCboClientName.BoundText) & " And TblItemsUnits.DefaultUnit = 1"

        Set rs2 = New ADODB.Recordset
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        Do While Not rs2.EOF
            LngItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
            mUnitId = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
            mUnitPurPrice = IIf(IsNull(rs2("UnitPurPrice").value), 0, rs2("UnitPurPrice").value)
            mName = IIf(IsNull(rs2("ItemName").value), 0, rs2("ItemName").value)
            mUnitName = IIf(IsNull(rs2("UnitName").value), 0, rs2("UnitName").value)
        

        '         ParrtNoCode = "" 'IIf(IsNull(RsTemp("ParrtNoCode").value), "", RsTemp("ParrtNoCode").value)
        '        ItemDetailedCode = "" 'IIf(IsNull(RsTemp("ItemDetailedCode").value), "", RsTemp("ItemDetailedCode").value)
        
        If LngItemID = 0 Then GoTo NextRow
        ' If mCode = "" and  Then GoTo NextRow
         
        With Me.FG

            If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
                .rows = .rows + 1
            End If
               ' NewGrid.FillGrid

            .TextMatrix(.rows - 1, FG.ColIndex("Code")) = LngItemID
            .TextMatrix(.rows - 1, FG.ColIndex("Name")) = LngItemID
            .TextMatrix(.rows - 1, FG.ColIndex("Name")) = mName
            .TextMatrix(.rows - 1, FG.ColIndex("Count")) = 1
            .TextMatrix(.rows - 1, FG.ColIndex("DiscountType")) = 1
        
            .TextMatrix(.rows - 1, FG.ColIndex("Serial")) = "" ' IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
            .TextMatrix(.rows - 1, FG.ColIndex("HaveSerial")) = "" ' IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemCase")) = "" ' IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ColorID")) = 1 ' IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemSize")) = 1 ' IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ClassID")) = 1 ' IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.cell(flexcpData, .rows - 1, FG.ColIndex("UnitID")) = mUnitId ' IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            'FG.TextMatrix(.Rows - 1, FG.ColIndex("ParrtNoCode")) = ParrtNoCode
            'FG.TextMatrix(.Rows - 1, FG.ColIndex("ItemDetailedCode")) = ItemDetailedCode
            .TextMatrix(.rows - 1, FG.ColIndex("Price")) = mUnitPurPrice
            ' .TextMatrix(.Rows - 1, FG.ColIndex("ShowPrice")) = mUnitPurPrice
            .TextMatrix(.rows - 1, FG.ColIndex("Valu")) = val(.TextMatrix(.rows - 1, .ColIndex("Price"))) * val(.TextMatrix(.rows - 1, .ColIndex("Count")))

            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = mUnitName
            Else
                FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = mUnitName
            End If

        End With

NextRow:
        rs2.MoveNext
    Loop


End Sub

Private Sub txtPassword_Change()
If Trim(txtPassword) = "Alex2025" Then
    cmdReSave.Visible = True
    txtFromDateReSave.Visible = True
    txtToDateReSave.Visible = True
    chkIsBranch(0).Visible = True
    chkIsBranch(2).Visible = True
    chkIsBranch(1).Visible = True
Else
    cmdReSave.Visible = False
    txtFromDateReSave.Visible = False
    chkIsBranch(2).Visible = False
    txtToDateReSave.Visible = False
   chkIsBranch(0).Visible = False
   chkIsBranch(1).Visible = False
End If
txtFromDateReSave.value = Date
txtToDateReSave.value = Date
End Sub

Function FIllSotreIfEmpty()
Dim i As Integer
With FG
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("StoreID2"))) = 0 Then
.TextMatrix(i, .ColIndex("StoreID2")) = val(DCboStoreName.BoundText)
End If
Next i
End With
End Function
Function CheckCompositeAccount() As Boolean
Dim Account_Code_dynamic  As String
If ChkCompsBill.value = vbUnchecked Then CheckCompositeAccount = True: Exit Function
CheckCompositeAccount = True
    Account_Code_dynamic = get_account_code_branch(4, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox " ÌÊÃœ Œÿ√ ›Ì Õ”«» «·„‘ —Ì« ", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox " ÌÊÃœ Œÿ√ ›Ì Õ”«» «·„‘ —Ì« ", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If
                
                
    Account_Code_dynamic = get_account_code_branch(96, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox " ÌÊÃœ Œÿ√ ›Ì Õ”«» «·⁄„Ê·« ", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox " ÌÊÃœ Œÿ√ ›Ì Õ”«» «·⁄„Ê·« ", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If
                
Exit Function
ErrTrap:
CheckCompositeAccount = False
End Function

Sub RelinVatGrid()
Dim k As Integer
If FG.ColIndex("Vat") = -1 Then Exit Sub
If val(DcbTyp.ListIndex) <> -1 Then
For k = FG.FixedRows To FG.rows - 1
FG.TextMatrix(k, FG.ColIndex("Vat")) = 0
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = 0
FG.TextMatrix(k, FG.ColIndex("TypeVAT")) = 0
Next k
VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.rows = 2
End If

Dim i As Integer
Dim SmValu As Double
SmValu = 0
With VatGrid
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
For k = FG.FixedRows To FG.rows - 1
If k = i And val(FG.TextMatrix(k, FG.ColIndex("Code"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) And val(FG.TextMatrix(k, FG.ColIndex("Valu"))) = val(.TextMatrix(i, .ColIndex("Valu"))) Then
FG.TextMatrix(k, FG.ColIndex("Vat")) = val(.TextMatrix(i, .ColIndex("Vat")))
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = val(.TextMatrix(i, .ColIndex("Vatyo")))
End If
Next k

SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
.TextMatrix(i, .ColIndex("Typ")) = ""
Else
For k = FG.FixedRows To FG.rows - 1
If k = i And val(FG.TextMatrix(k, FG.ColIndex("Code"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) And val(FG.TextMatrix(k, FG.ColIndex("Valu"))) = val(.TextMatrix(i, .ColIndex("Valu"))) Then
FG.TextMatrix(k, FG.ColIndex("Vat")) = 0
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = 0
End If
Next k
If val(Me.DcbTyp.ListIndex) > -1 Then
.TextMatrix(i, .ColIndex("Typ")) = val(Me.DcbTyp.ListIndex) + 1
End If
End If
Next i
End With
TxtValueAdded.Text = Format(SmValu, ".##")
LblValueAdded.Caption = Format(SmValu, ".##")
Me.LblTotal.Caption = val(XPTxtSum.Text) + val(LblValueAdded.Caption)
End Sub
Function SaveItemsData(Optional Transaction_ID As String = 0, Optional StoreID3 As Integer)
If SystemOptions.WorkWithItemsDetails = False Then Exit Function
    Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDes As String
    Dim RowNum As Integer
    Dim StrSQL As String
    strFilterText = ","
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete ItemsDetails   where Transaction_ID= " & val(Me.XPTxtBillID.Text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  ItemsDetails Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim UnitName As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim Name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.rows - 1
If StoreID3 <> 0 Then
        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("StoreID2"))) = StoreID3 Then
            
           If FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) <> "" Then
                AllDes = FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea"))
                astrSplitItems = Split(AllDes, strFilterText)
         NooFRows = UBound(astrSplitItems) + 1
                For intX = 0 To NooFRows - 2
             
                
                          RsgGrantee.AddNew
                         astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
                         RsgGrantee("ItemDetailedCode").value = (astrSplitItems1(0))
                         RsgGrantee("ParrtNoCode").value = (astrSplitItems1(1))
                         RsgGrantee("count").value = val(astrSplitItems1(2))
                         RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 1, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))  ' val(astrSplitItems1(3))
                         RsgGrantee("ColorID").value = val(astrSplitItems1(4))
                         RsgGrantee("sizeid").value = val(astrSplitItems1(5))
                         RsgGrantee("ClassId").value = val(astrSplitItems1(6))
                         RsgGrantee("ProductionDate").value = IIf(IsDate((astrSplitItems1(7))), astrSplitItems1(7), Null)
                         RsgGrantee("ExpireDate").value = IIf(IsDate((astrSplitItems1(8))), astrSplitItems1(8), Null)
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
                        RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                       RsgGrantee("EffectN").value = 1
                    RsgGrantee.update
                                    Next intX
                Else
                If FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) <> "" Then
                RsgGrantee.AddNew
                RsgGrantee("ParrtNoCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))
                RsgGrantee("count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
                RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RsgGrantee("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RsgGrantee("sizeid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RsgGrantee("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
                RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
                RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                RsgGrantee("ItemDetailedCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))
                RsgGrantee("EffectN").value = 1
                RsgGrantee.update
                  
               End If
         
                   
               End If
            End If

  Else
   If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            
           If FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) <> "" Then
                AllDes = FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea"))
                astrSplitItems = Split(AllDes, strFilterText)
         NooFRows = UBound(astrSplitItems) + 1
                For intX = 0 To NooFRows - 2
             
                
                          RsgGrantee.AddNew
                         astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
                         RsgGrantee("ItemDetailedCode").value = (astrSplitItems1(0))
                         RsgGrantee("ParrtNoCode").value = (astrSplitItems1(1))
                         RsgGrantee("count").value = val(astrSplitItems1(2))
                         RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 1, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))  ' val(astrSplitItems1(3))
                         RsgGrantee("ColorID").value = val(astrSplitItems1(4))
                         RsgGrantee("sizeid").value = val(astrSplitItems1(5))
                         RsgGrantee("ClassId").value = val(astrSplitItems1(6))
                         RsgGrantee("ProductionDate").value = IIf(IsDate((astrSplitItems1(7))), astrSplitItems1(7), Null)
                         RsgGrantee("ExpireDate").value = IIf(IsDate((astrSplitItems1(8))), astrSplitItems1(8), Null)
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
                        RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                       RsgGrantee("EffectN").value = 1
                    RsgGrantee.update
                                    Next intX
                Else
                If FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) <> "" Then
                RsgGrantee.AddNew
                RsgGrantee("ParrtNoCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))
                RsgGrantee("count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
                RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RsgGrantee("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RsgGrantee("sizeid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RsgGrantee("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
                RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
                RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                RsgGrantee("ItemDetailedCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))
                RsgGrantee("EffectN").value = 1
                RsgGrantee.update
                  
               End If
         
                   
               End If
            End If
  End If

    Next RowNum

End Function

 
Function SaveGoldData()
    Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDes As String
    Dim RowNum As Integer
    Dim StrSQL As String
    strFilterText = ","
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete TblGoldDetail   where Transaction_ID= " & val(Me.XPTxtBillID.Text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  TblGoldDetail Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim UnitName As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim Name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            If FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails")) <> "" Then
                AllDes = FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails"))
                astrSplitItems = Split(AllDes, strFilterText)
         NooFRows = UBound(astrSplitItems) + 1
                For intX = 0 To NooFRows - 2
                        
                  
                        RsgGrantee.AddNew
 
'                        RsgGrantee("itemid").value = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))
'    astrSplitItems = Split(AllIDS, strFilterText)
    

 
 
    
    astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
            RsgGrantee("TTypeId").value = val(astrSplitItems1(0))
            RsgGrantee("typeid").value = val(astrSplitItems1(1))
            RsgGrantee("uniteid").value = val(astrSplitItems1(2))
            RsgGrantee("type").value = val(astrSplitItems1(3))
            RsgGrantee("price").value = val(astrSplitItems1(4))
            RsgGrantee("weight").value = val(astrSplitItems1(5))
            RsgGrantee("Count").value = val(astrSplitItems1(6))
            RsgGrantee("InstallPrice").value = val(astrSplitItems1(7))
         RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
         
      
  
               
                         
                         RsgGrantee.update
                  
                       
                Next intX
                    
            End If

        End If

    Next RowNum

End Function

Public Sub RetriveSerials(ItemID As String, _
                          ItemName As String, _
                          seriallist As String, _
                          currentrow As Long, Optional Price As Double, Optional UnitID As Double = 1, Optional UnitName As String)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(strInputString, strFilterText)
    Dim i As Integer
 
   
    Num = currentrow

    '  For Num = currentrow To UBound(astrSplitItems)+currentrow
    
    Dim CurrentSerial As String
 
   
    '*****************************************************
    For intX = 0 To UBound(astrSplitItems)
   FG.cell(flexcpData, Num, FG.ColIndex("Code")) = ItemID
   FG.TextMatrix(Num, FG.ColIndex("Code")) = ItemID
   
        FG.TextMatrix(Num, FG.ColIndex("Name")) = ItemID
        
        
         FG.TextMatrix(Num, FG.ColIndex("UnitID")) = ItemID
        FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = 1
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = 0
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = 0
    
         
         
        FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
        FG.TextMatrix(Num, FG.ColIndex("Serial")) = astrSplitItems(intX)
        
        FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = UnitID
        FG.TextMatrix(Num, FG.ColIndex("UnitID")) = UnitName
                     FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True

        
If val(Price) > 0 Then
            FG.TextMatrix(Num, FG.ColIndex("price")) = Price
        End If
        
        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.rows = FG.rows + 1
 
        Num = Num + 1
    If intX = UBound(astrSplitItems) Then
    NewGrid.Calculate Num
    NewGrid.DtpBillDate_Change
        End If
    Next
     
     
    TxtFillData.Text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Function CheckMyData() As Boolean
    CheckMyData = True
my_branch = val(Me.dcBranch.BoundText)
    If TxtNoteSerial.Text = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": GoTo ErrTrap
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": GoTo ErrTrap
            Else
                TxtNoteSerial.Text = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
        
   '                 If TxtNoteSerial1.Text = "" Then
   '                                     If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22) = "error" Then
   '                                         MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ›« Ê—… „‘ —Ì«  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": GoTo ErrTrap
   '                                     Else
   '
   ''                                                         If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22) = "" Then
                                                       '         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… „‘ —Ì«   ÌœÊÌ« ﬂ„« Õœœ   ": GoTo ErrTrap
    '                                                        Else
    '                                                            TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22)
                                              '              End If
                                        'End If
             '       End If
 
 Dim NoteSerial1str  As String
     If TxtNoteSerial1.Text = "" Then
    
    NoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22, , val(DCboStoreName.BoundText))
                    If NoteSerial1str = "error" Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ›« Ê—… „‘ —Ì«  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": GoTo ErrTrap
                    Else
                                   
                        If NoteSerial1str = "" Then
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… „‘ —Ì«   ÌœÊÌ« ﬂ„« Õœœ   ": GoTo ErrTrap
                        Else
                            TxtNoteSerial1.Text = NoteSerial1str
                        End If
                    End If
    End If
    
    If BillBasedOn(0).value = True Then

        If Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , val(DCboStoreName.BoundText)) = "" Then
                                
            If Trim$(TxtManualNo1) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·«” ·«„ ÌœÊÌ« Õ Ï Ì „ «‰‘«¡ «·”‰œ  ":  GoTo ErrTrap
            
            Else
                TxtNoteSerial1V = TxtManualNo1
            End If
            
        End If
                       
    End If

    Exit Function
ErrTrap:
    CheckMyData = False
End Function
Function CheckAcconts() As Boolean
CheckAcconts = False


            Account_Code_dynamic101 = get_account_code_branch(101, my_branch)
            Account_Code_dynamic102 = get_account_code_branch(102, my_branch)
             If Account_Code_dynamic101 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·„œÌ‰ ·«Ê«„— «·‘—«¡  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
              
              
                  If Account_Code_dynamic102 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·œ«∆‰ ·«Ê«„— «·‘—«¡  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
              
  
   CheckAcconts = True
   Exit Function
ErrTrap:
      CheckAcconts = False
End Function
Public Sub Convert()
    Cmd_Click (0)
End Sub

Public Sub Cala()
    NewGrid.Calculate 1, , , True
End Sub

Private Sub BillBasedOn_Click(Index As Integer)

    Select Case Index

        Case 0

            If BillBasedOn(0).value = True Then
                
                FillVoucherGrid (0)
                GRID1.Enabled = True
            End If

        Case 1

            If BillBasedOn(1).value = True Then
                
                FillVoucherGrid (1)
                GRID1.Enabled = True
            End If

        Case 2

            If BillBasedOn(2).value = True Then
                
                '            FillOrderGrid
                '            GRID2.Enabled = True
            End If

    End Select

End Sub

Function FillVoucherGrid(Optional OPtype As Integer = 0)
    ' ⁄»∆…  ”‰œ«   «·’—›
    On Error Resume Next

    With Me.GRID1
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
    'My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=20   and   dbo.TblCustemers.CusID=" & Val(DBCboClientName.BoundText)
    If OPtype = 0 Then
        My_SQL = "SELECT dbo.Transactions.NoteID,dbo.Transactions.ManualNO,dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_ID= " & val(Text1.Text) & " or Transaction_ID in ( " & GetTrnasectionID(val(XPTxtBillID.Text), 22) & " )"
    Else
        'My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_ID= " & Val(Text1.text)
        My_SQL = "SELECT dbo.Transactions.NoteID, dbo.Transactions.ManualNO, dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where   ( (nots='" & Me.XPTxtBillID.Text & "' and  Transaction_Type=20) or(Transaction_ID in ( " & GetTrnasectionID(val(XPTxtBillID.Text), 22) & " ))or ( Transaction_Type=20   and  closed =0 and (nots='' or nots is null) ) and  (dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText) & ")) "
    End If

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID1
        .rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .rows - 1
             
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
                .TextMatrix(i, .ColIndex("ManualNO")) = IIf(IsNull(RsExp.Fields("ManualNO").value), "", RsExp.Fields("ManualNO").value)
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), 0, RsExp.Fields("NoteID").value)
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("P1")) = "⁄—÷ «·”‰œ"
                    .TextMatrix(i, .ColIndex("P2")) = "ÿ»«⁄Â  «·ﬁÌœ"
                Else
                    .TextMatrix(i, .ColIndex("P1")) = "View VCHR"
                    .TextMatrix(i, .ColIndex("P2")) = "Print GE"

                End If

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    GRID1.Visible = True

End Function

Function CloseIssueVoucher()
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
  
    If BillBasedOn(1).value = False Then Exit Function
DeleteTransactiomsVoucher val(Text1.Text)

    With GRID1

        For i = 1 To .rows - 1
     
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.Text) & ",nots2='" & Me.TxtNoteSerial1.Text & "' where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
            Else
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) ' & "nots=" & "" & "nots2=" & ""
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function
Function deletelinktoVoucher()
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
    If BillBasedOn(1).value = False Then Exit Function

    With GRID1

        For i = 1 To .rows - 1
     
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                'sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2='" & Me.TxtNoteSerial1.text & "' where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
           sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) ' & "nots=" & "" & "nots2=" & ""
             
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function

Private Sub BLDate_Change()
DBCboClientName_Change
End Sub

Private Sub CBoBasedON_Change()
If Me.TxtModFlg.Text <> "R" Then

TXT_order_no.Text = ""
End If

    If Me.CBoBasedON.ListIndex = 0 Then

    ElseIf Me.CBoBasedON.ListIndex = 1 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(66).Caption = "—ﬁ„ «·«„—  "
        Else
            lbl(66).Caption = "Order NO"
        End If

    ElseIf Me.CBoBasedON.ListIndex = 2 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(66).Caption = "—ﬁ„Â«"
        Else
            lbl(66).Caption = "NO:"
        End If
    End If

    If TXT_order_no.Text <> "" Then
        Txt_order_no_Change
    End If

End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub



Private Sub ChAddToTotal_Click()
If ChAddToTotal.value = vbChecked Then
txtAddValue.Caption = val(TXTFactoryExpenses.Text)
TXTFactoryExpensesVat = ""
If val(Fg_Journal.rows) > 1 Then
    TXTFactoryExpensesVat.Text = Fg_Journal.Aggregate(flexSTSum, Fg_Journal.FixedRows, Fg_Journal.ColIndex("Vat"), Fg_Journal.rows - 1, Fg_Journal.ColIndex("Vat"))
End If
LblValueAdded.Caption = val(LblValueAdded.Caption) + val(TXTFactoryExpensesVat.Text)
Else
LblValueAdded.Caption = val(LblValueAdded.Caption) - val(TXTFactoryExpensesVat.Text)
txtAddValue.Caption = 0
End If
XPTxtSum_Change
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            LblPrecenType.Caption = ""
            LblPrecenValue.Caption = ""
            LblInstallTotal.Caption = ""
            LblInstallCount.Caption = ""
            LblFirstInstallDate.Caption = ""
            LblInstallmentType.Caption = ""
        End With

    End If

End Sub

Private Sub ChkTaxAdd_Click()

    If ChkTaxAdd.value = Checked Then
        TxtTaxAddValue.Enabled = True
        lbl(39).Enabled = True
        lbl(46).Enabled = True
    Else
        TxtTaxAddValue.Text = ""
        TxtTaxAddValue.Enabled = False
        lbl(39).Enabled = False
        lbl(46).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxSerivce_Click()
    On Error GoTo ErrTrap

    If ChkTaxSerivce.value = Checked Then
        TxtTaxServiceValue.Enabled = True
        lbl(43).Enabled = True
        lbl(47).Enabled = True
    Else
        TxtTaxServiceValue.Text = ""
        TxtTaxServiceValue.Enabled = False
        lbl(43).Enabled = False
        lbl(47).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxStamp_Click()

    If ChkTaxStamp.value = Checked Then
        TxtTaxStampValue.Enabled = True
        lbl(41).Enabled = True
        lbl(48).Enabled = True
    Else
        TxtTaxStampValue.Text = ""
        TxtTaxStampValue.Enabled = False
        lbl(41).Enabled = False
        lbl(48).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Function RemoveFactoryExpenses()

    With Me.Fg_Journal
  
        If .Row <= 0 Then Exit Function
        .RemoveItem .Row
    End With

      ReLineGrid

End Function

Function CheckExpens(Optional ByRef Row As Integer) As Boolean
Dim i As Integer
With Fg_Journal
CheckExpens = True
For i = 1 To .rows - 1
If .TextMatrix(i, .ColIndex("AccountName")) <> "" Then
If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
Row = i
CheckExpens = False
Exit Function
End If
End If
Next i
End With
End Function
 Function Retrive_Items_data1()
    Dim StrSQL  As String
    Dim row_count As Long
    Dim Num As Long
    Dim i As Long
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    StrSQL = "select * from TblItems where ItemID in(" & TxtItemsIDes.Text & ")"
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If rs2.RecordCount > 0 Then
        
        If FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code")) = "" Then
      FG.rows = FG.rows - 1
        End If
     With FG
     row_count = FG.rows
       rs2.MoveFirst
       .rows = rs2.RecordCount + .rows
        For Num = row_count To .rows - 1 'RsDetails.RecordCount
        .TextMatrix(Num, .ColIndex("Code")) = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
      
        rs2.MoveNext
        Next Num
        For i = row_count To .rows - 1 'RsDetails.RecordCount
          NewGrid.Grid_AfterEdit i, .ColIndex("Code")
        Next i
        NewGrid.Grid_AfterEdit row_count, .ColIndex("Code")
    End With
    End If


End Function


Private Function CheckPOItems() As Boolean
    Dim StrSQL  As String
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    If CBoBasedON.ListIndex <> 1 Then
    Exit Function
    
    End If
    '----------------------------
    StrSQL = "Select * From Transaction_Details Where Transaction_ID=" & val(Me.poTransaction_ID.Text) & ""
    StrSQL = StrSQL + " Order  By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    CheckPOItems = False

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FG

            For i = .FixedRows To .rows - 1

                If .TextMatrix(i, .ColIndex("Name")) <> "" Then
                    If rs.filter <> adFilterNone Then
                        rs.filter = adFilterNone
                    End If

                    rs.MoveFirst
                    rs.filter = "Item_ID=" & val(.TextMatrix(i, .ColIndex("Name")))

                    If rs.BOF Or rs.EOF Then
                        Msg = "«·’‰› : " & .cell(flexcpTextDisplay, i, .ColIndex("Name"))
                        Msg = Msg & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & i
                        Msg = Msg & CHR(13) & "·„ Ìﬂ‰ „ÊÃÊœ ›Ï «„— «·‘—«¡ —ﬁ„ : " & Me.TXT_order_no.Text
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        CheckPOItems = False
                        rs.Close
                        Set rs = Nothing
                        Exit Function
                    ElseIf FG.cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexChecked Then
                        rs.Find "ItemSerial='" & Trim(.TextMatrix(i, .ColIndex("Serial"))) & "'", , adSearchForward, 1

                        If rs.BOF Or rs.EOF Then
                            Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«·:  " & Trim(.TextMatrix(i, .ColIndex("Serial")))
                            Msg = Msg & CHR(13) & "„‰ «·’‰› : " & .cell(flexcpTextDisplay, i, .ColIndex("Name"))
                            Msg = Msg & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  : " & i
                            Msg = Msg & CHR(13) & "·„ Ìﬂ‰ „ÊÃÊœ ›Ï «„— «·‘—«¡ —ﬁ„  : " & Me.TXT_order_no.Text
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            CheckPOItems = False
                            rs.Close
                            Set rs = Nothing
                            Exit Function
                        End If
                    End If
                End If

            Next i

        End With

    End If

    '----------------------------

    '----------------------------
    CheckPOItems = True
End Function

Private Sub Cmd_Click(Index As Integer)
    '       On Error GoTo ErrTrap
    Dim AskOption As Boolean
    Dim intDef    As Integer
    Dim Msg       As String

    BolPrint = True
 
    Select Case Index
    
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            With Me.Grid4
                .rows = .FixedRows
   
            End With
            VatGrid.Clear flexClearScrollable, flexClearEverything
            VatGrid.rows = 1
            Command2.Enabled = True
            Txt_EXport.Enabled = True
            '  Grid.Visible = True
            clear_all Me
            TxtModFlg.Text = "N"
            ' Me.TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
            BillBasedOn(0).value = True
            BLDate.value = Date
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))

            If BillType = 22 Then
                TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=22"))
            End If
        
            If BillType = 1 Then
                TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=1"))
            End If

            '      TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "",  True  )
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", 1)
            '  DCboStoreName.BoundText = intDef
            XPTab301.CurrTab = 0
            '        FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.rows - 1
            Command2_Click
            
            If SystemOptions.DefaultIsCreditPurchase = False Then
                CboPayMentType.ListIndex = 0
            Else
                CboPayMentType.ListIndex = 1
            End If
            
            Dim dstore       As Integer
            Dim dBox         As Integer
            Dim usertype     As Integer
            Dim EmpID        As Integer
            Dim userbranchid As Integer
            Dim cust1        As Integer
            'GetBranchData branch_id, dstore, dBox
            Dim boxid1       As Integer
            GetUserData user_id, usertype, userbranchid, , dBox, , EmpID, , , dstore, cust1, boxid1
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = True
                '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
                DBCboClientName.BoundText = cust1
                Me.DcboBox.BoundText = boxid1
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
                '                TxtStoreID.Enabled = True
                Me.DcboBox.BoundText = ""
            End If

            If SystemOptions.usertype <> UserAdminAll Then
                If checkmanyBranches = False Then
                    Me.dcBranch.Enabled = True
                Else
                    Me.dcBranch.Enabled = True
                End If
                    
                If checkmanyStores = False Then
                    Me.DCboStoreName.Enabled = True
                                    
                Else
                    Me.DCboStoreName.Enabled = True
 
                End If
                                  
            End If
            
            Me.dcBranch.BoundText = Current_branch
            Me.CBoBasedON.ListIndex = 0
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.rows = 2
            Fg_Journal.Enabled = True
            GRID1.rows = 1
            GRID1.Enabled = True
          
            Dccurrency.BoundText = MainCurrency()
            TxtNoteSerial1V = ""
            If Voucher_coding(val(Me.dcBranch.BoundText), XPDtbBill.value, 6, 150, 22) = "" Then
                TxtNoteSerial1.locked = False
            Else
                TxtNoteSerial1.locked = True
 
            End If
            XPTxtBillID.Text = 0
            
            If SystemOptions.IsHiddenTransportInv Then
                CBoBasedON.ListIndex = 1
            End If
        Case 1

            If ChekClodePeriod(XPDtbBill.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
            If ChekPaymet() = True And cmdReSave.Visible = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ   "
                Else
                    Msg = "Can not be allowed to edite this process"
                    Msg = Msg & CHR(13) & "There repayment process   "
                End If
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
                  
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            ' If SystemOptions.usertype = UserNormal Then
            '     If AvailableDeal = False Then
            '         Exit Sub
            '     End If
            ' End If

            TxtModFlg.Text = "E"
            If Trim(txtPassword) <> "Alex2025" Then
                Me.DCboUserName.BoundText = user_id
            End If
            Fg_Journal.rows = Fg_Journal.rows + 1
            Fg_Journal.Enabled = True
            DateChanged = False
            CuurentLogdata
            StroreChanged = False
            NewGrid.CboDiscount_Type_Change
            '    Command4_Click
    
        Case 2
         
            If SystemOptions.POMustentryAndBillMustEntry = True And (TXT_order_no.Text = "" Or CBoBasedON.ListIndex = 0) Then
                MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ·⁄œ„ «Œ Ì«— »‰«¡ ⁄·Ì Ê ÕœÌœ «·—ﬁ„", vbCritical
                Exit Sub
            End If
            If CheckPOItems = False And CBoBasedON.ListIndex = 1 Then
                Exit Sub
            End If
         
            If CboPayMentType.ListIndex = 1 Then
                XPChkPayType(1).value = 1
                '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
                XPTxtValue(1).Text = val(LblTotal.Caption)

            Else
                XPChkPayType(0).value = 1
                '  XPTxtValue(0).text = Val(LblTotalAll.Caption)
                XPTxtValue(0).Text = val(LblTotal.Caption)

            End If
            If SystemOptions.PoCreateVoucher = True Then

                If CheckAcconts = False Then Exit Sub

            End If

            If ChkCompsBill.value = vbChecked Then
                CboPayMentType.ListIndex = 1

                If DBCboClientName.BoundText = 1 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "„‰ ›÷·ﬂ √œŒ· «”„  «·„Ê—œ   «·«Ã·"
                    Else
                        Msg = "Select Customer Name"
    
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DBCboClientName.SetFocus
                    Sendkeys "{F4}"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
    
            End If
        
            If Dccurrency.BoundText = "" Then
    
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "«Œ — «·⁄„·… «Ê·« "
                Else
                    Msg = "Select Currency First"
    
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Dccurrency.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
    
            End If
        
            Dim NoteSerialT As String
            If SystemOptions.DontDuplicateManulaNoInPurchase = True Then
                If ChekInvoiceNoPurchasemanualExist(val(Me.XPTxtBillID.Text), val(DBCboClientName.BoundText), Me.TxtManualNO, NoteSerialT) = True Then
        
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ  €ÌÌ—  «·—ﬁ„ «·ÌœÊÌ ·«‰ „ﬂ—— „‰ ﬁ»· ›Ì ›« Ê—… —ﬁ„  " & NoteSerialT
                    Else
                        MsgBox "Manual No Already Exist   " & NoteSerialT
                    End If
                          
                    Exit Sub
                          
                End If
        
            End If
        
            If ChekClodePeriod(XPDtbBill.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
                  
            If CheckCompositeAccount = False Then
             
            End If
            Dim i As Integer
            If CheckExpens(i) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox i & " " & "·«ÌÊÃœ Õ”«» ··„’—Ê› ›Ì «·”ÿ— —ﬁ„"
                Else
                    MsgBox "There is no expense account in line " & " " & i
                End If
                Exit Sub
            End If

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«  "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
 
            If val(DBCboClientName.BoundText) = 1 And CboPayMentType.ListIndex = 1 Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = " Cash Vendor can't be credit  "
                Else
                    Msg = "«·„Ê—œ «·‰ﬁœÌ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ «Ã·"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                ' Dcbranch.SetFocus
                '  SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            my_branch = val(Me.dcBranch.BoundText)

            '   If Me.TxtModFlg.text = "N" Then
             
            ' End If
            If val(CboPayMentType.ListIndex) = 2 And SystemOptions.AllowPurchasesMultyPayed = True Then
                If val(TxtRemainValue2.Text) <> 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ…"
                    Else
                        MsgBox "The  value is incorrect"
                    End If
                    Exit Sub
                End If
                If val(TxtPayedValue2.Text) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ «œŒ«· «·ﬁÌ„… "
                    Else
                        MsgBox "The  value is incorrect"
                    End If
                    Exit Sub
                End If
                If val(TxtPayedValue2.Text) <> val(LblTotal.Caption) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ «œŒ«· «·ﬁÌ„… «·’ÕÌÕ… "
                    Else
                        MsgBox "The  value is incorrect"
                    End If
                    Exit Sub
                End If
                FramePay.Visible = False
            End If
            If val(Me.TxtValueAdded.Text) > 0 Then
                If GetValueAddedAccount(XPDtbBill.value, , , 1, 22) = False Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
                    Else
                        MsgBox "Value added account not specified"
                    End If
                    Exit Sub
                End If
            End If
            If TxtModFlg.Text = "E" Then
                FIllSotreIfEmpty
            End If
            If NewGrid.CaseStoredLin() = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ Ì—ÃÏ  ⁄œÌ· «·„Œ“‰ «·—∆Ì”Ì"
                Else
                    MsgBox "Please change the main store"
                End If
                DCboStoreName.SetFocus
                Exit Sub
            End If
            If CheckGeidExpensss() = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ «œŒ«· «·Õ”«»«  ›Ì «·„’—Ê›«  «· ﬁœÌ—Ì…"
                Else
                    MsgBox "Please select Account in Expensess"
                End If
                Exit Sub
            End If
            If IsSaveWithOutMsg Then GoTo SaveDirect
                  Dim StrSQL As String
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=22 "
            If val(XPTxtBillID) <> 0 Then
                StrSQL = StrSQL & "  and Transaction_Id= " & val(XPTxtBillID)
            End If
            
            StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
                   If SystemOptions.usertype <> UserAdminAll Then
                  '     StrSQL = StrSQL & " AND   BranchId=" & Current_branch
                   End If
            
            
            If SystemOptions.usertype <> UserAdminAll Then
            
                 If SystemOptions.FixedCustomer = 1 Then
                   StrSQL = StrSQL & " and  UserID = " & user_id
                    End If
            
               Me.dcBranch.Enabled = True
             
             
            End If
            
                   StrSQL = StrSQL & " Order by Transaction_ID"
 
            
            '    StrSQL = StrSQL & "  and Transaction_ID in (Select Transaction_Details.Transaction_ID  from Transaction_Details where Item_Id = 89)"
            
            
            'StrSQL = StrSQL & "  and Transaction_ID =14895"
                   Set rs = New ADODB.Recordset
                   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
            
                   If Not (rs.EOF Or rs.BOF) Then
                       rs.MoveLast
                   End If
            
SaveDirect:
            SaveData

        Case 3
            Undo

        Case 4
            If ChekClodePeriod(XPDtbBill.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
                  
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            '   If SystemOptions.usertype = UserNormal Then
            '       Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
            '       MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
            '       Exit Sub
            '   End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
            If m_FrmSearch Is Nothing Then
                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = PurchaseTransaction
                If SystemOptions.UserInterface = ArabicInterface Then
                    m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… „‘ —Ì« "
                Else
                    m_FrmSearch.Caption = "Search About Purchase Invoice"
                End If

                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal
            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’À… »‘«‘… ›« Ê—… «·‘—«¡ «·Õ«·Ì…"
                Msg = Msg & CHR(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘… ›« Ê—…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                m_FrmSearch.Visible = True
                m_FrmSearch.ZOrder 0
                m_FrmSearch.SetFocus
            End If

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
              
                FrmPrintOptions.show vbModal
            End If

            If BolPrint = False Then
                Exit Sub
            End If
        
            printing
        
        Case 9
            RemoveFactoryExpenses

        Case 10
            ShowGL_cc TxtNoteSerial.Text, , 200
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
     On Error Resume Next
           If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtNoteSerial1, "0812201403"

End Sub

Private Sub SumChecks()

    With Me.FgCheques

        If .rows > 1 Then
            Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .rows - 1, .ColIndex("CheckValue"))
        Else
            Me.lbl(19).Caption = 0
            Me.lbl(18).Caption = 0
        End If

    End With

End Sub

Private Sub CmdHelp_Click()
'    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
'    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd


   TxtModFlg.Text = "N"
            Me.XPTxtBillID.Text = ""
 
            Me.DCboUserName.BoundText = user_id
              'Me.DcBranch.BoundText = Current_branch
     TxtNoteSerial.Text = ""
     TxtNoteSerial1.Text = ""
 
     
End Sub

Private Sub CmdInfo_Click()
    Me.PopupMenu mdifrmmain.MnuInvPurchase
End Sub
Private Sub cmdReSave_Click()
    Dim Ids() As Long
    Dim cnt As Long, i As Long
    
    IsSaveWithOutMsg = True
    
    'Õ„¯· rs »«·›·« — » «⁄ ﬂ
    XPBtnMove_Click 2
    If rs Is Nothing Then GoTo Endsub
    If rs.EOF And rs.BOF Then GoTo Endsub
    
    '«Ã„⁄ «·‹ IDs ›Ì „’›Ê›…
    rs.MoveFirst
    Do While Not rs.EOF
        cnt = cnt + 1
        ReDim Preserve Ids(1 To cnt)
        Ids(cnt) = CLng(val(rs!Transaction_ID & ""))
        rs.MoveNext
    Loop
    rs.MoveFirst
    
    '·› ⁄·Ï «·‹ IDs („‰ €Ì— „«  ⁄ „œ ⁄·Ï rs ‰Â«∆Ì)
    For i = 1 To cnt
        ResaveOneTransaction Ids(i)
        DoEvents
    Next i

Endsub:
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"
    Cmd(2).Enabled = True
End Sub

Private Sub ResaveOneTransaction(ByVal TransID As Long)
    On Error GoTo eh
    
    '???? ????????
    Me.TxtModFlg.Text = "R"
    Me.Retrive TransID
    DoEvents
    
    '???? Edit
    Cmd_Click 1
    DoEvents
    
    '?? ????? ????? ????? ??????/?????
    NewGrid.DtpBillDate_Change
    DoEvents
    
    'Save
    IsSaveWithOutMsg = True
    Cmd_Click 2
    DoEvents
    
    Exit Sub
eh:
    '???????: ??? ??? ???????? ???? ????
    Debug.Print "Resave error TransID=" & TransID & " Err=" & Err.Number & " " & Err.Description
    Err.Clear
End Sub

'
Private Sub cmdReSave_ClickOld()
   Dim s As String
   Dim i As Double
   IsSaveWithOutMsg = True
     XPBtnMove_Click (2)
    DoEvents
For i = 1 To rs.RecordCount
  Cmd_Click (1)
  DoEvents
  NewGrid.DtpBillDate_Change
  DoEvents
  DoEvents
  IsSaveWithOutMsg = True
  DoEvents
  
  Cmd_Click (2)
  
  
  XPBtnMove_Click (0)
  If i > rs.RecordCount Then GoTo Endsub
  
Next i

Endsub:
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"
    Cmd(2).Enabled = True
  '  Dim s As String
  '  Dim rsDummy As ADODB.Recordset
  '  XPBtnMove_Click (2)
  '  DoEvents
  '
  '  XPBtnMove_Click (1)
  '  DoEvents
  '  Set rsDummy = New ADODB.Recordset
  '      s = " SELECT * FROM Transactions WHERE Transaction_Type = 22 "
  '      s = s & "   and ( Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & " and "
  '      s = s & "   Transaction_Date <=   " & SQLDate(txtToDateReSave.value, True) & " )"
  '      s = s & " ORDER BY  Transaction_Date, BranchId, Transaction_ID"
  '  rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
  '
  '  Do While Not rsDummy.EOF
  '      On Error GoTo NextRow
  '
  '      mIsFinishSave = False
  '      mIsStart = True
  '      Me.TxtModFlg.Text = "R"
  '      Me.Retrive val(rsDummy!Transaction_ID & "")
  '
  '
  '      DoEvents
'11 :
  '      DoEvents
  '      If mIsFinishSave And mIsStart Then
'            IsSaveWithOutMsg = True
'            Me.TxtModFlg.Text = "E"
'            Cmd_Click (1)
'            DoEvents
'            DoEvents
'            DoEvents
'
            
'
'            SaveData True
'            mIsStart = False
'        Else
'            GoTo 11
'        End If
'        DoEvents
'        DoEvents
'        DoEvents
'        DoEvents
'
        
'        DoEvents
                 
                 
                 
'NextRow:
'        rsDummy.MoveNext
        
        
'    Loop
'    IsSaveWithOutMsg = False
'    MsgBox " „ «·Õ›Ÿ"

End Sub

Private Sub DcboEmp_Change()
DcboEmp_Click 0
End Sub

Private Sub DcboEmp_Click(Area As Integer)
Dim StoreID As Double
 If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
  StoreID = get_StoreBYPurchasePerson(val(Me.DcboEmp.BoundText))
 If StoreID <> 0 Then
 DCboStoreName.BoundText = StoreID
 End If
 End If
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcbTyp_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTyp.ListIndex) = -1 Then
NewGrid.DtpBillDate_Change
Else
RelinVatGrid
End If
End If
End Sub

Private Sub DcbTyp_Click()
DcbTyp_Change
End Sub

Private Sub dcproject_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
       If KeyCode = vbKeyF3 Then
           FrmProjectSearch.lblSearchtype.Caption = 29
               FrmProjectSearch.show vbModal
        End If
End If
End Sub

Private Sub Grid22_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid2
End Sub

Private Sub Grid22_Click()
If TxtPayedValue2.Text = "" Or val(TxtPayedValue2.Text) = 0 Then
With Me.Grid22
.TextMatrix(.Row, .ColIndex("Value")) = LBLPayVal.Caption
ReLineGrid2
End With
End If
End Sub
Private Sub CmdINSTALLMENT_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim i As Integer

    If XPTxtValue(1).Text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

        If XPTxtValue(1).Enabled = True Then
            XPTxtValue(1).SetFocus
        End If

        Exit Sub
    End If

    Load FrmInstallMent
    Set FrmInstallMent.Frm = Me

    With FrmInstallMent

        If Me.TxtModFlg.Text = "R" Then
            .Tag = "R"
            .Retrive val(XPTxtValue(1).Tag)
        Else
            .Tag = "N"
            .Txt(1).Text = XPTxtValue(1).Text
            .LblNoteID.Caption = XPTxtSerial(1).Text
            .CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            .Txt(3).Text = val(LblPrecenValue.Caption)
            .Txt(5).Text = val(LblInstallCount.Caption)

            If IsDate(Me.LblFirstInstallDate.Caption) Then
                .Dtp_First.value = Me.LblFirstInstallDate.Caption
            End If

            .Txt(7).Text = val(LblInstallSeprator.Caption)

            If val(LblInstallmentType.Tag) = 0 Then
                .OptInt(0).value = True
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                .OptInt(1).value = True
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                .OptInt(2).value = True
            End If

            With .FG
                .rows = Me.FgInstallments.rows

                For i = 1 To Me.FgInstallments.rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Value")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Value"))
                    .TextMatrix(i, .ColIndex("Due_Date")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Due_Date"))
                Next i

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdNotes_Click()
    ShowRelatedNotes val(Me.XPTxtBillID.Text), 1
End Sub

Private Sub CmdNotes_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Dim StrTemp As String

    If val(Me.CmdNotes.Tag) = 0 Then
        Me.CmdNotes.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… ⁄„·Ì«  „«·Ì… „ﬁœ«—Â« : " & val(Me.CmdNotes.Tag)
        Me.CmdNotes.ToolTipText = StrTemp
    End If

End Sub

Private Sub CMDPAy_Click()
'
If val(CboPayMentType.ListIndex) = 2 And SystemOptions.AllowPurchasesMultyPayed = True Then
If val(TxtRemainValue2.Text) <> 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ…"
Else
MsgBox "The  value is incorrect"
End If
Exit Sub
End If
FramePay.Visible = False
End If
End Sub

Private Sub CmdRetruns_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim StrTemp As String

    If val(Me.CmdRetruns.Tag) = 0 Then
        Me.CmdRetruns.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… Õ—ﬂ«   Ã«—Ì… √Œ—Ï ·Â« ⁄·«ﬁ… »Â« ≈Ã„«·ÌÂ«: " & val(Me.CmdRetruns.Tag)
        Me.CmdRetruns.ToolTipText = StrTemp
    End If

End Sub

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer, Optional StoreID3 As Integer)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim SngTemp2  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim total_shahn As Single
    Dim usedaccount As Integer
    Dim StoredID5 As Double
    Dim StoredID6 As Integer
    Dim ExpenssValue As Double
Dim NoteID2 As Double
NoteID2 = general_noteid
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    SngTemp = ((NewGrid.GetItemsTotal(ItemsGoodType) + val(TXTFactoryExpenses.Text) - val(LblDiscountsTotal.Caption) + val(LblCommision.Caption)) * val(txt_Currency_rate.Text) + val(TXTToTAlELSHahn.Text) - val(TXTFactoryExpenses.Text) * val(txt_Currency_rate.Text))
    SngTemp2 = ((NewGrid.GetItemsTotal(ItemsGoodType) + val(TXTFactoryExpenses.Text) - val(LblDiscountsTotal.Caption) + val(LblCommision.Caption)) + val(TXTToTAlELSHahn.Text) - val(TXTFactoryExpenses.Text))
    StoredID6 = val(DCboStoreName.BoundText)
    
If StoreID3 <> 0 Then

ExpenssValue = val(TXTFactoryExpenses.Text) * GetItemsTotalExpensessByStore(val(XPTxtBillID.Text), StoreID3)
ExpenssValue = Round(ExpenssValue, SystemOptions.SysDefCurrencyForamt)
SngTemp2 = GetItemsTotalByStore(val(XPTxtBillID.Text), StoreID3) + ExpenssValue
SngTemp = (GetItemsTotalByStore(val(XPTxtBillID.Text), StoreID3) + ExpenssValue) * val(txt_Currency_rate.Text) + ((val(txt_total_bill.Text) + val(Txt_EXport.Text)) * GetItemsTotalExpensessByStore(val(XPTxtBillID.Text), StoreID3))

StoredID6 = StoreID3
End If

StoredID5 = StoredID6
    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·”‰œ «·«” ·«„", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 Then
        
                    StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            End If

            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
                End If
            
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V
            End If

            LngDevNO = LngDevNO + 1

If SystemOptions.SupplierReciveGE = True And SystemOptions.autoReseiveVoucher = False Then
SngTemp = 0
End If


            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·”‰œ «·«” ·«„", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                    Account_Code_dynamic = StrTempAccountCode
                ElseIf usedaccount = 0 Then
        
                    Account_Code_dynamic = get_store_Account(StoredID6, "Account_Code")
                End If

            Else
                Account_Code_dynamic = get_store_Account(StoredID6, "Account_Code")
            End If

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
            If val(Me.DBCboClientName.BoundText) > 2 And val(CboPayMentType.ListIndex) <> 0 Then
                OtherInformation.NextAccount_Code = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
             Else
                OtherInformation.NextAccount_Code = get_account_code_branch(4, my_branch)
             End If
             
         '  OtherInformation.NextAccount_Code = get_account_code_branch(4, my_branch)
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.rows - 1
If StoreID3 <> 0 Then
                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And FG.TextMatrix(i, FG.ColIndex("StoreID")) = StoreID3 Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), StoredID6, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = 0

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.Text)
    
                        total_shahn = Round((((line_value) / (val(LblTotal.Caption) * val(txt_Currency_rate.Text))) * val(TXTToTAlELSHahn.Text)), SystemOptions.SysDefCurrencyForamt)  'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                        line_value = line_value + total_shahn + val(FG.TextMatrix(i, FG.ColIndex("LineShahn")))
                        line_value = Round(line_value, SystemOptions.SysDefCurrencyForamt)
                        
                        
                       SngTemp2 = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        total_shahn = Round((((line_value) / (val(LblTotal.Caption))) * val(TXTToTAlELSHahn.Text)), SystemOptions.SysDefCurrencyForamt)   'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                        SngTemp2 = SngTemp2 + total_shahn + val(FG.TextMatrix(i, FG.ColIndex("LineShahn")))
                        SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)
                        
     
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                        Else
                            StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
                        End If
   
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If
               Else
                 If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), StoredID6, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = 0

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.Text)
    
                        total_shahn = Round((((line_value) / (val(LblTotal.Caption) * val(txt_Currency_rate.Text))) * val(TXTToTAlELSHahn.Text)), SystemOptions.SysDefCurrencyForamt)  'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                        line_value = line_value + total_shahn + val(FG.TextMatrix(i, FG.ColIndex("LineShahn")))
                        line_value = Round(line_value, SystemOptions.SysDefCurrencyForamt)
                        
                        
                       SngTemp2 = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        total_shahn = Round((((line_value) / (val(LblTotal.Caption))) * val(TXTToTAlELSHahn.Text)), SystemOptions.SysDefCurrencyForamt)   'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                        SngTemp2 = SngTemp2 + total_shahn + val(FG.TextMatrix(i, FG.ColIndex("LineShahn")))
                        SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)
                        
     
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                        Else
                            StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
                        End If
   
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If
                   End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(Me.LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text) '+ Val(TXTToTAlELSHahn.text)
        SngTemp2 = (NewGrid.GetItemsTotal(ItemsGoodType) - val(Me.LblDiscountsTotal.Caption))
        If ChAddToTotal.value = vbChecked Then
       SngTemp = SngTemp + val(TXTFactoryExpenses.Text) * val(txt_Currency_rate.Text)
       SngTemp2 = SngTemp2 + val(TXTFactoryExpenses.Text)
       End If
       If StoreID3 <> 0 Then
       If ChAddToTotal.value = vbChecked Then
       SngTemp2 = GetItemsTotalByStore(val(XPTxtBillID.Text), StoreID3) + ExpenssValue
        SngTemp = (GetItemsTotalByStore(val(XPTxtBillID.Text), StoreID3) + ExpenssValue) * val(txt_Currency_rate.Text)
      
      Else
        SngTemp = GetItemsTotalByStore(val(XPTxtBillID.Text), StoreID3) * val(txt_Currency_rate.Text)
      SngTemp2 = SngTemp
      End If
     End If
        If SngTemp > 0 Then
            
            If SystemOptions.PoCreateVoucher = True And CboPayMentType.ListIndex = 1 Then
            If TXT_order_no.Text <> "" Then
               StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            GoTo NewGl3
            End If
            
         
            End If
            
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then
          '  If val(Me.DBCboClientName.BoundText) > 2 And val(CboPayMentType.ListIndex) <> 0 Then
          '      Account_Code_dynamic = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
          '   Else
          '      Account_Code_dynamic = get_account_code_branch(4, my_branch)
          '   End If
              
                 Account_Code_dynamic = get_account_code_branch(4, my_branch)
                 
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„‘ —Ì«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                If val(DCDocTypes.BoundText) > 0 Then
                    getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

                    If StrTempAccountCode = "" And usedaccount = 1 Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ «·«” ·«„", vbCritical
                        GoTo ErrTrap
                    ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                    ElseIf usedaccount = 0 Then
        
                        StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
                    End If

                Else
                    StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
                End If
NewGl3:
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
                End If
            
                LngDevNO = LngDevNO + 1
                
                
OtherInformation.NextAccount_Code = get_store_Account(StoredID6, "Account_Code")


If SystemOptions.SupplierReciveGE = True And SystemOptions.autoReseiveVoucher = False Then
SngTemp = 0
End If





                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                    GoTo ErrTrap
                End If
         
         
         
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.rows - 1
If StoreID3 <> 0 Then
                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("StoreID2"))) = StoreID3 Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 4)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»   «·„‘ —Ì«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = 0
                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.Text)
                            SngTemp2 = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
                            '  total_shahn = Round((((line_value) / (Val(LblTotal.Caption) * Val(txt_Currency_rate.text))) * Val(TXTToTAlELSHahn.text)), 2)  'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                            '  line_value = line_value + total_shahn + Val(FG.TextMatrix(I, FG.ColIndex("LineShahn")))
                            line_value = Round(line_value, SystemOptions.SysDefCurrencyForamt)
     
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                            Else
                                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If
Else
         If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 4)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»   «·„‘ —Ì«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = 0
                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.Text)
                            SngTemp2 = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
                            '  total_shahn = Round((((line_value) / (Val(LblTotal.Caption) * Val(txt_Currency_rate.text))) * Val(TXTToTAlELSHahn.text)), 2)  'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                            '  line_value = line_value + total_shahn + Val(FG.TextMatrix(I, FG.ColIndex("LineShahn")))
                            line_value = Round(line_value, 2)
     
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                            Else
                                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If
End If
                    Next i

                End With

            End If
        End If

        'ﬁÌœ «·„’—Ê›« 
        Dim Account_code As String
        Dim Note_Value As Double
        Dim Note_Value2 As Double
        
        With Grid

            For i = 1 To Grid.rows - 1

                If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                    Else
                        StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_code = Grid.TextMatrix(i, Grid.ColIndex("Account_code"))
                    If StoreID3 <> 0 Then
                    Note_Value = Grid.TextMatrix(i, Grid.ColIndex("Note_value")) * GetItemsTotalExpensessByStore(val(XPTxtBillID.Text), StoreID3)
                    Else
                    Note_Value = Grid.TextMatrix(i, Grid.ColIndex("Note_value"))
                    End If

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
        
            Next

        End With

        'ﬁÌœ «·›Ê« Ì—
     
        With Grid4

            For i = 1 To Grid4.rows - 1

                If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                                            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                    Else
                        StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text
                    End If
                                                        
                    LngDevNO = LngDevNO + 1
                    Account_code = Grid4.TextMatrix(i, Grid4.ColIndex("Account_code"))
                    If StoreID3 <> 0 Then
                    Note_Value = Grid4.TextMatrix(i, Grid4.ColIndex("Note_value")) * GetItemsTotalExpensessByStore(val(XPTxtBillID.Text), StoreID3)
                    Else
                    Note_Value = Grid4.TextMatrix(i, Grid4.ColIndex("Note_value"))
                    End If

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
         
            Next
   
        End With

        '«·„’—Ê›«  «·„»«‘—…
   
          If ChAddToTotal.value = vbUnchecked Then
            Dim mDisc As String
        With Fg_Journal

            For i = 1 To .rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And val(.TextMatrix(i, .ColIndex("value"))) <> 0 Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text & CHR(13) & Trim(.TextMatrix(i, .ColIndex("des")))
                    Else
                        StrTempDes = "Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text & CHR(13) & Trim(.TextMatrix(i, .ColIndex("des")))
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_code = .TextMatrix(i, .ColIndex("AccountCode"))
                    If StoreID3 <> 0 Then
                    Note_Value = val(.TextMatrix(i, .ColIndex("value"))) * GetItemsTotalExpensessByStore(val(XPTxtBillID.Text), StoreID3) * val(txt_Currency_rate.Text)
                    Else
                    Note_Value = val(.TextMatrix(i, .ColIndex("value"))) * val(txt_Currency_rate.Text)
                    End If
                    Note_Value2 = val(.TextMatrix(i, .ColIndex("value")))
                     
                    mDisc = Trim(.TextMatrix(i, .ColIndex("des")))
                  '  mDisc = ""
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , mDisc, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
        
            Next

        End With
            '  Else
          '    Note_Value = val(TXTFactoryExpenses.Text) * val(txt_Currency_rate.Text)
 '   If Note_Value > 0 Then
 '          If SystemOptions.UserInterface = ArabicInterface Then
 '                       StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.txtNoteSerial1.Text
 '                   Else
 '                       StrTempDes = "Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.txtNoteSerial1.Text
 '                   End If
 '          Account_Code_dynamic = get_account_code_branch(4, my_branch)
 ''                 LngDevNO = LngDevNO + 1
  '
  '               Note_Value2 = val(TXTFactoryExpenses.Text)
  '                  If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
  '                      GoTo ErrTrap
  '                  End If
  ' End If
End If

    End If



''
Dim CommissionAccount As String
        LngDevNO = LngDevNO + 1
                  CommissionAccount = get_account_code_branch(96, my_branch)
  
                    
                    Note_Value = val(LblCommision.Caption) * val(txt_Currency_rate.Text)
If Note_Value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, CommissionAccount, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
   End If
   updateNotesValueAndNobytext NoteID2
ErrTrap:
End Function

Function CreateRecieveVouchers2()

    If BillBasedOn(1).value = True Then Exit Function
   '  On Error GoTo errortrap
    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double
    Dim rs As ADODB.Recordset
    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '>>>>>>>>>>>>>>>>>>>>>>>>>
    CurrentVoucherNo = ""
    CurrentVoucherSerialNo = ""
    CurrentVoucherNo = Trim(GetVoucherGLNO(val(Text1.Text), CurrentVoucherSerialNo))
          
TxtNoteSerial1V = ""
   ' DeleteTransactiomsVoucher val(Text1.Text)

    ' rs.Close
 
    '        rs.Close


Dim sql As String
Dim rs2 As ADODB.Recordset
Dim StoreID3 As Integer
Set rs2 = New ADODB.Recordset
sql = " SELECT     StoreID2"
sql = sql & " From dbo.Transaction_Details"
sql = sql & " Where (Transaction_ID = " & val(XPTxtBillID.Text) & ")"
sql = sql & " GROUP BY StoreID2"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst

For i = 1 To rs2.RecordCount
StoreID3 = IIf(IsNull(rs2("StoreID2").value), 0, rs2("StoreID2").value)

     Set rs = New ADODB.Recordset
 '   StrSQL = "select * from Transactions where Transaction_ID = " & val(XPTxtBillID.Text)   ' & TxtTransSerial.text & " and Transaction_type = 22"
   ' rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=20"))
  
    Dim Transaction_ID As Long
   
    my_branch = val(Me.dcBranch.BoundText)
    Dim general_noteid As Long
    Dim RsNotesGeneral As ADODB.Recordset
    Dim TxtNoteSerialV As String
 'Dim txtNoteSerial1V As String
            
    my_branch = val(Me.dcBranch.BoundText)

    If TxtNoteSerialV = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Function
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Function
            Else
                TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
        
        Dim TxtNoteSerial1Vstr As String
        
        
    If TxtNoteSerial1V = "" Then
    TxtNoteSerial1Vstr = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , StoreID3)
        If TxtNoteSerial1Vstr = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «÷«›… ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Function
        Else
                       
            If TxtNoteSerial1Vstr = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «÷«›…  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Function
            Else
                TxtNoteSerial1V = TxtNoteSerial1Vstr
            End If
        End If
    End If
 
    If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
        TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
                    TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
                If StroreChanged <> True Then
                    TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
                 
                
                End If
    
    Else
            TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , StoreID3)
            CurrentVoucherNo = ""
    
    
    End If
    
    
 
        Dim sql22 As String
        
     Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
     sql22 = "INSERT INTO  Transactions (CBoBasedON,Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,BranchId,nots2,TransactionComment)SELECT 5," & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 20,CusID,StoreID,UserID,Emp_ID,nots='" & TxtNoteSerial1.Text & "',NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId ," & TxtNoteSerial1 & "  ,TransactionComment  From Transactions Where Transaction_ID =" & XPTxtBillID.Text + " And Transaction_Type = 22"
     
     Cn.Execute sql22
     SaveTrnasectionID val(XPTxtBillID), Transaction_ID, 22
    
   ' rs!nots = Transaction_ID
   ' rs.update
    

    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
    '
    
  
       sql = "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO,OldQty,OldCost,NewQty,NewCost,StoreID2,length,OUTR,INR,Height,Width ,NoCount) " & "SELECT   ( ( (round(Commisionvalue," & SystemOptions.SysDefCurrencyForamt & ")+showPrice-( round(discountvalue," & SystemOptions.SysDefCurrencyForamt & ")+TotalDiscountPerLine)*QtyBySmalltUnit)*" & val(txt_Currency_rate.Text) & ")+(ToTAlELSHahn+LineShahn-(" & val(TXTFactoryExpenses.Text) & "  * LineExpenses))*QtyBySmalltUnit)+(" & val(TXTFactoryExpenses.Text) & "  * LineExpenses * " & val(txt_Currency_rate.Text) & ") "
       sql = sql & ",guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, (( ( round(Commisionvalue," & SystemOptions.SysDefCurrencyForamt & ")+ Price-(round(discountvalue," & SystemOptions.SysDefCurrencyForamt & " )+TotalDiscountPerLine))*" & val(txt_Currency_rate.Text) & ")+(ToTAlELSHahn+LineShahn) ), ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID," & Me.XPDtbBill.value & ",ExpiryDate,LotNO ,OldQty,OldCost,NewQty,NewCost ,StoreID2,length,OUTR,INR,Height,Width ,NoCount From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.Text & "  And StoreID2 = " & StoreID3 & ""
         Cn.Execute sql
       Cn.Execute "Update Transactions set StoreID =" & StoreID3 & " where Transaction_ID=" & Transaction_ID & " "
       Dim StrTransaction_ID As String
       StrTransaction_ID = Transaction_ID
    SaveItemsData StrTransaction_ID, StoreID3
    UpdateTransactionsCost CStr(Transaction_ID)
    Dim NoteID As Long
    Dim NoteDate As Date
    Dim NoteSerial As String
    Dim Notevalue As Double
    Dim des As String
If CurrentVoucherNo <> "" Then
NoteSerial = CurrentVoucherNo
End If
'TxtNoteSerialV
Dim SngTemp As Double
    SngTemp = ((NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption) + val(LblCommision.Caption)) * val(txt_Currency_rate.Text) + val(TXTToTAlELSHahn.Text))
     CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 160, SngTemp, NoteSerial, TxtNoteSerial1V, "Transactions", "Transaction_ID", Transaction_ID, TxtNoteSerial1V, ToHijriDate(XPDtbBill.value)
          ' TxtNoteID.text = NoteID
     general_noteid = NoteID
     
     CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText), StoreID3
     rs2.MoveNext
     TxtNoteSerial1V = ""
     TxtNoteSerialV = ""
     NoteSerial = ""
    
     Next i
   End If
ErrTrap:
End Function
Function CreateRecieveVouchers() As Boolean
    On Error GoTo ErrTrap
CreateRecieveVouchers = False

    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double
    Dim rs As ADODB.Recordset
    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '>>>>>>>>>>>>>>>>>>>>>>>>>

On Error GoTo ErrTrap
IsVouc = False
If NewGrid.CaseStoredLin() = 2 Then
    CurrentVoucherNo = ""
    CurrentVoucherSerialNo = ""
    CurrentVoucherNo = Trim(GetVoucherGLNO(val(Text1.Text), CurrentVoucherSerialNo))
          
TxtNoteSerial1V = ""
    DeleteTransactiomsVoucher val(Text1.Text)
    
    If SystemOptions.NotCrtResvVouchProjects = True And val(dcproject.BoundText) <> 0 Then
        IsVouc = True
        CreateRecieveVouchers = True
        Exit Function
    Else
        CreateRecieveVouchers2
        CreateRecieveVouchers = True
        IsVouc = True
    Exit Function
End If
Else
    If BillBasedOn(1).value = True Then IsVouc = True: CreateRecieveVouchers = True: Exit Function
        CurrentVoucherNo = ""
    CurrentVoucherSerialNo = ""
    CurrentVoucherNo = Trim(GetVoucherGLNO(val(Text1.Text), CurrentVoucherSerialNo))
          
TxtNoteSerial1V = ""
    DeleteTransactiomsVoucher val(Text1.Text)
        If SystemOptions.NotCrtResvVouchProjects = True And dcproject.BoundText <> "" Then
        IsVouc = True
        CreateRecieveVouchers = True
    Exit Function
    End If
   '  On Error GoTo errortrap

    ' rs.Close
 
    '        rs.Close


'salimhere************************
   ' Set rs = New ADODB.Recordset

   ' StrSQL = "select * from Transactions where Transaction_ID = " & val(XPTxtBillID.Text)   ' & TxtTransSerial.text & " and Transaction_type = 22"
   ' rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 'salimhere************************
 
    MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=20"))
  
    Dim Transaction_ID As Long
   
    my_branch = val(Me.dcBranch.BoundText)
    Dim general_noteid As Long
    Dim RsNotesGeneral As ADODB.Recordset
    Dim TxtNoteSerialV As String
 'Dim txtNoteSerial1V As String
            
    my_branch = val(Me.dcBranch.BoundText)

    If TxtNoteSerialV = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Function
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Function
            Else
                TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
        
        Dim TxtNoteSerial1Vstr As String
        
        
    If TxtNoteSerial1V = "" Then
    TxtNoteSerial1Vstr = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , val(DCboStoreName.BoundText))
        If TxtNoteSerial1Vstr = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «÷«›… ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Function
        Else
                       
            If TxtNoteSerial1Vstr = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «÷«›…  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Function
            Else
                TxtNoteSerial1V = TxtNoteSerial1Vstr
            End If
        End If
    End If
 
    If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
        TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
                    TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
                If StroreChanged <> True Then
                    TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
                 
                
                End If
    
    Else
            TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , val(DCboStoreName.BoundText))
            CurrentVoucherNo = ""
    
    
    End If
    
    
 
        Dim sql22 As String
        
     Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
     sql22 = "INSERT INTO  Transactions (CBoBasedON,Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,BranchId,nots2,TransactionComment)SELECT 5," & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 20,CusID,StoreID,UserID,Emp_ID,nots='" & TxtNoteSerial1.Text & "',NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId ," & TxtNoteSerial1 & "  ,TransactionComment  From Transactions Where Transaction_ID =" & XPTxtBillID.Text + " And Transaction_Type = 22"
     
     Cn.Execute sql22
    
    
       StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.Text)
        Cn.Execute StrSQL
 
 'salimhere************************
    'rs!nots = Transaction_ID
    'rs.update
    'salimhere************************
    
    'Create big notes
'    Set RsNotesGeneral = New ADODB.Recordset
'        StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
'   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   'RsNotesGeneral.AddNew
   ' RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
   ' general_noteid = RsNotesGeneral("NoteID").value
 
   ' RsNotesGeneral("Transaction_ID").value = Transaction_ID
   ' RsNotesGeneral("NoteDate").value = XPDtbBill.value
   ' RsNotesGeneral("NoteType").value = 160
   ' RsNotesGeneral("Note_Value").value = Null
   ' RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
 '
 '    RsNotesGeneral("remark").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
     
 '   RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
 '   RsNotesGeneral("numbering_type1").value = sand_numbering_type(9) '«–‰ «÷«›…
 '   RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
 '   RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
 '   RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
 '   RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
 '   general_noteid = RsNotesGeneral("NoteID").value
 '
 '   RsNotesGeneral.update
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
    '
    Dim sql As String
  
       sql = "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO,OldQty,OldCost,NewQty,NewCost,length,OUTR,INR,Height,Width ,NoCount) " & "SELECT   ( ( (round(Commisionvalue," & SystemOptions.SysDefCurrencyForamt & ")+showPrice-( round(discountvalue," & SystemOptions.SysDefCurrencyForamt & " "
     sql = sql & " )+TotalDiscountPerLine)*QtyBySmalltUnit)*" & val(txt_Currency_rate.Text) & ")+(ToTAlELSHahn+LineShahn)*QtyBySmalltUnit) ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, (( ( round(Commisionvalue,2)+ Price-(round(discountvalue,2)+TotalDiscountPerLine))*" & val(txt_Currency_rate.Text) & ")+(ToTAlELSHahn+LineShahn) ), ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID," & Me.XPDtbBill.value & ",ExpiryDate,LotNO ,OldQty,OldCost,NewQty,NewCost ,length,OUTR,INR,Height,Width ,NoCount From dbo.Transaction_Details Where Transaction_ID =" & val(XPTxtBillID.Text)
    Cn.Execute sql
    SaveItemsData (Transaction_ID) ' ›«’Ì· «·«’‰«›
    UpdateTransactionsCost CStr(Transaction_ID)  '«· ﬂ·›Â »«·‘ﬂ· «·ÃœÌœ
 Dim NoteID As Long
  Dim NoteDate As Date
    Dim NoteSerial As String
    Dim Notevalue As Double
    Dim des As String
If CurrentVoucherNo <> "" Then
    NoteSerial = CurrentVoucherNo
End If
'TxtNoteSerialV
Dim SngTemp As Double
    SngTemp = ((NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption) + val(LblCommision.Caption)) * val(txt_Currency_rate.Text) + val(TXTToTAlELSHahn.Text))
     CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 160, SngTemp, NoteSerial, TxtNoteSerial1V, "Transactions", "Transaction_ID", Transaction_ID, TxtNoteSerial1V, ToHijriDate(XPDtbBill.value)
          ' TxtNoteID.text = NoteID
     general_noteid = NoteID
     CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)
    End If
    IsVouc = True
    
    
        CreateRecieveVouchers = True
 Exit Function
    '
 
ErrTrap:
    IsVouc = True
    
    
    
    
    CreateRecieveVouchers = False
    




End Function

Private Sub Command1_Click()
    CreateRecieveVouchers
End Sub

Private Sub Command2_Click()

    '    ⁄»∆… «·«–Ê‰ «·„’—Ê›« 
    If CBoBasedON.ListIndex = 0 Or CBoBasedON.ListIndex = 1 Or TXT_order_no.Text = "" Then

        With Me.Grid
            .rows = .FixedRows
   
        End With

   '     Exit Sub

    End If

    With Me.Grid
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    'Dim i As Integer
    Dim i As Double
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3   and order_no='" & Me.Txt_order_no.text & "' and(Transaction_ID1 is null or Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ")  )  "
'    My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where  dbo.Notes.NoteType = 3   and order_no='" & Me.TXT_order_no.text & "'"
    'My_SQL = ""
My_SQL = " SELECT     dbo.Notes.NoteID, dbo.Notes.Buy, dbo.Notes.NoteSerial, dbo.Notes.ItemID, dbo.Notes.Note_Value, dbo.ExpensesType.Name, dbo.ExpensesType.Account_Code, "
My_SQL = My_SQL & "   dbo.notes_all.BasedONID"
My_SQL = My_SQL & "  FROM         dbo.Notes INNER JOIN"
My_SQL = My_SQL & "  dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID INNER JOIN"
My_SQL = My_SQL & "    dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"
'My_SQL = My_SQL & "  WHERE     (dbo.Notes.ORDER_NO = '" & txt_ORDER_NO & "') AND (dbo.Notes.NoteType = 3) AND ( dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2 )"


My_SQL = "SELECT   DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,  dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Buy, dbo.Notes.NoteSerial, dbo.Notes.ItemID, dbo.Notes.Note_Value, dbo.ExpensesType.Name, dbo.ExpensesType.Account_Code,"
My_SQL = My_SQL & "                        dbo.notes_all.BasedONID , dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 ,dbo.notes_all.VATCustoms "
My_SQL = My_SQL & "  FROM         dbo.Notes INNER JOIN"
My_SQL = My_SQL & "                        dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID INNER JOIN"
My_SQL = My_SQL & "                        dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID INNER JOIN"
My_SQL = My_SQL & "                        dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
If Me.TxtModFlg.Text = "R" Or Me.TxtModFlg.Text = "" Then
            If CBoBasedON.ListIndex = 0 Then ' »·«
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 3)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.Text)
            Else
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 3) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2) "
            End If

ElseIf Me.TxtModFlg.Text = "E" Then


            If CBoBasedON.ListIndex = 0 Then
            My_SQL = My_SQL + " WHERE   (   dbo.Notes.NoteType = 3   AND  dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0 and     dbo.notes_all.BasedONID = 0   and   ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) ) and  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.Text)
            Else
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 3) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2)"
            End If

ElseIf Me.TxtModFlg.Text = "N" Then


            If CBoBasedON.ListIndex = 0 Then
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 3)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and   ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) "
            Else
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 3) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2) and     ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) "
            End If
            
End If

My_SQL = My_SQL + " and ( dbo.DOUBLE_ENTREY_VOUCHERS.hideline = 0 or dbo.DOUBLE_ENTREY_VOUCHERS.hideline is null)"

    RsExp.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Dim StrSQL  As String

    With Me.Grid
        .rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst
TxtVATCustoms1.Text = IIf(IsNull(RsExp("VATCustoms").value), 0, RsExp("VATCustoms").value)
            For i = 1 To .rows - 1
                   .TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")) = IIf(IsNull(RsExp.Fields("Double_Entry_Vouchers_ID").value), 0, RsExp.Fields("Double_Entry_Vouchers_ID").value)
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsExp.Fields("ItemID").value), "", RsExp.Fields("ItemID").value)
    
                StrSQL = "select * from TblItems where ItemID=" & val(.TextMatrix(i, .ColIndex("ItemID")))
                Dim rs As New ADODB.Recordset
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

3                 If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(i, .ColIndex("ItemName")) = ""
                    .TextMatrix(i, .ColIndex("ItemCode")) = ""
 
                End If
               
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Name").value), "", RsExp.Fields("Name").value)
               
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
            
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), "", RsExp.Fields("NoteID").value)
           
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsExp.Fields("Note_Value").value), "", RsExp.Fields("Note_Value").value)
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsExp.Fields("Account_Code").value), "", RsExp.Fields("Account_Code").value)
            
                If IsNull(RsExp.Fields("buy").value) Then
                    .TextMatrix(i, .ColIndex("Select")) = 0
                Else

                    If RsExp.Fields("buy").value = False Then
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    ElseIf RsExp.Fields("buy").value = True Then
                        .TextMatrix(i, .ColIndex("Select")) = 1
                    Else
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    End If
           
                End If
           
           '     .TextMatrix(i, .ColIndex("Select")) = 1
               
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid.Visible = True

    '    ⁄»∆… «·«–Ê‰ «·„œ›Ê⁄« 

    Expenses_update_total

End Sub

Private Sub Command3_Click()

    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.Grid2
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
    My_SQL = "SELECT dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  (Transaction_Type=6  or Transaction_Type=29)and NOT(ORDER_NO IS NULL) AND CLOSED= 0 and   dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText)

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid2
        .rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("order_no").value), "", RsExp.Fields("order_no").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
0                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid2.Visible = True

End Sub

Private Sub Command4_Click()



'If Me.TxtModFlg.Text = "" And txt_ORDER_NO.Text = "" Then
'        With Me.grid4
'            .Rows = .FixedRows
'
'        End With
'Exit Sub
'End If
       With Me.Grid4
            .rows = .FixedRows
   
        End With
    'If Not Fg.TextMatrix(Fg.Row, Fg.ColIndex("Code")) = "" Then
    '    ⁄»∆… «·«–Ê‰ «·„’—Ê›« 

    'Frame2.Caption = FG.TextMatrix(FG.Row, FG.ColIndex("name"))

    If CBoBasedON.ListIndex = 0 Or CBoBasedON.ListIndex = 1 Or TXT_order_no.Text = "" Then

        With Me.Grid4
            .rows = .FixedRows
   
        End With

    '    Exit Sub

    End If

    With Me.Grid4
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
        '
        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Notes.Item_id,dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3 and order_no='" & Me.TXT_order_no.text & "' " & "AND (ITEM_ID=" & Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code"))) & " or  ITEM_ID is null)  and(Transaction_ID1 is null or Transaction_ID1=" & Val(Me.XPTxtBillID.text) & "))  "
'    My_SQL = "SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
'    My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial,"
'    My_SQL = My_SQL + " dbo.Notes.order_no, dbo.DOUBLE_ENTREY_VOUCHERS.ItemID ,  dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.buy,dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 "
'    My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
'    My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
'    My_SQL = My_SQL + "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
'    'My_SQL = My_SQL + " WHERE      (dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 is null or dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ") and  (dbo.Notes.NoteType = 80) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.ORDER_NO = '" & Me.Txt_order_no.text & "')"
  '  My_SQL = My_SQL + " WHERE       (dbo.Notes.NoteType = 80) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.ORDER_NO = '" & Me.TXT_order_no.text & "')"




My_SQL = " SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
My_SQL = My_SQL + "  dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial,"
My_SQL = My_SQL + "  dbo.Notes.ORDER_NO, dbo.DOUBLE_ENTREY_VOUCHERS.ItemID, dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,"
My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.buy , dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1, dbo.notes_all.BasedONID ,dbo.notes_all.VATCustoms"
My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
My_SQL = My_SQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
My_SQL = My_SQL + " dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"

If Me.TxtModFlg.Text = "R" Or Me.TxtModFlg.Text = "" Then
            If CBoBasedON.ListIndex = 0 Then ' »·«
            My_SQL = My_SQL + " WHERE  ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and    (dbo.Notes.NoteType = 80)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.Text)
            Else
            My_SQL = My_SQL + " WHERE   ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and    (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2) "
            End If

ElseIf Me.TxtModFlg.Text = "E" Then


            If CBoBasedON.ListIndex = 0 Then
            My_SQL = My_SQL + " WHERE   ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and   (   dbo.Notes.NoteType = 80   AND  dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0 and     dbo.notes_all.BasedONID = 0   and   ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) ) or  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.Text)
            Else
            My_SQL = My_SQL + " WHERE     ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and   (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2)"
            End If

ElseIf Me.TxtModFlg.Text = "N" Then


            If CBoBasedON.ListIndex = 0 Then
            My_SQL = My_SQL + " WHERE   ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and     (dbo.Notes.NoteType = 80)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and   ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) "
            Else
            My_SQL = My_SQL + " WHERE    ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and    (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2) and     ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) "
            End If
            
End If

My_SQL = My_SQL & " AND (dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 IS NULL " & _
                 "      OR dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.Text) & ")"

My_SQL = My_SQL + " and ( dbo.DOUBLE_ENTREY_VOUCHERS.hideline = 0 or dbo.DOUBLE_ENTREY_VOUCHERS.hideline is null)"
My_SQL = My_SQL + "  order by dbo.DOUBLE_ENTREY_VOUCHERS.buy desc ,dbo.Notes.NoteSerial1"
    RsExp.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    Dim StrSQL As String
    Dim rs As New ADODB.Recordset

    With Me.Grid4
        .rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst
TxtVATCustoms.Text = IIf(IsNull(RsExp("VATCustoms").value), 0, RsExp("VATCustoms").value)
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")) = IIf(IsNull(RsExp.Fields("Double_Entry_Vouchers_ID").value), 0, RsExp.Fields("Double_Entry_Vouchers_ID").value)
           
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsExp.Fields("ItemID").value), "", RsExp.Fields("ItemID").value)
    
                StrSQL = "select * from TblItems where ItemID=" & val(.TextMatrix(i, .ColIndex("ItemID")))
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(i, .ColIndex("ItemName")) = ""
                    .TextMatrix(i, .ColIndex("ItemCode")) = ""
 
                End If
               
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Account_Name").value), "", RsExp.Fields("Account_Name").value)
 
                Else
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Account_NameEng").value), "", RsExp.Fields("Account_NameEng").value)
                End If
 
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
 
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), "", RsExp.Fields("NoteID").value)
 
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsExp.Fields("Value").value), "", RsExp.Fields("Value").value)
 
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsExp.Fields("Account_Code").value), "", RsExp.Fields("Account_Code").value)
 
'                If IsNull(RsExp.Fields("buy").value) Then
'                    .TextMatrix(i, .ColIndex("Select")) = 0
'                Else
'
'                    If RsExp.Fields("buy").value = False Then
'                        .TextMatrix(i, .ColIndex("Select")) = 0
'                    ElseIf RsExp.Fields("buy").value = True Then
'                        .TextMatrix(i, .ColIndex("Select")) = 1
'                    Else
'                        .TextMatrix(i, .ColIndex("Select")) = 0
'                    End If
'
'                End If
 
 
 If Not IsNull(RsExp("Transaction_ID1").value) Then
    If val(RsExp("Transaction_ID1").value) = val(Me.XPTxtBillID.Text) Then
        .TextMatrix(i, .ColIndex("Select")) = 1
    Else
        .TextMatrix(i, .ColIndex("Select")) = 0
    End If
Else
    .TextMatrix(i, .ColIndex("Select")) = 0
End If

                ' .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("buy").value), _
                  0, RsExp.Fields("buy").value)
                  If CBoBasedON.ListIndex = 1 And Me.TxtModFlg.Text = "R" Then
           '   .TextMatrix(i, .ColIndex("Select")) = 1
              End If
              
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid4.Visible = True

    ' End If
  
    update_finincial_invoice_total
Exit Sub
'If Me.TxtModFlg.Text = "" And txt_ORDER_NO.Text = "" Then
'        With Me.grid4
'            .Rows = .FixedRows
'
'        End With
'Exit Sub
'End If
       With Me.Grid4
            .rows = .FixedRows
   
        End With
    'If Not Fg.TextMatrix(Fg.Row, Fg.ColIndex("Code")) = "" Then
    '    ⁄»∆… «·«–Ê‰ «·„’—Ê›« 

    'Frame2.Caption = FG.TextMatrix(FG.Row, FG.ColIndex("name"))

    If CBoBasedON.ListIndex = 0 Or CBoBasedON.ListIndex = 1 Or TXT_order_no.Text = "" Then

        With Me.Grid4
            .rows = .FixedRows
   
        End With

    '    Exit Sub

    End If

    With Me.Grid4
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
        '
        '    .AutoSize 0, .Cols - 1, False
    End With

'    Dim i As Integer
'    Dim RsExp As ADODB.Recordset
'    Dim My_SQL As String
'
'    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Notes.Item_id,dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3 and order_no='" & Me.TXT_order_no.text & "' " & "AND (ITEM_ID=" & Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code"))) & " or  ITEM_ID is null)  and(Transaction_ID1 is null or Transaction_ID1=" & Val(Me.XPTxtBillID.text) & "))  "
'    My_SQL = "SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
'    My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial,"
'    My_SQL = My_SQL + " dbo.Notes.order_no, dbo.DOUBLE_ENTREY_VOUCHERS.ItemID ,  dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.buy,dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 "
'    My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
'    My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
'    My_SQL = My_SQL + "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
'    'My_SQL = My_SQL + " WHERE      (dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 is null or dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ") and  (dbo.Notes.NoteType = 80) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.ORDER_NO = '" & Me.Txt_order_no.text & "')"
  '  My_SQL = My_SQL + " WHERE       (dbo.Notes.NoteType = 80) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.ORDER_NO = '" & Me.TXT_order_no.text & "')"




My_SQL = " SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
My_SQL = My_SQL + "  dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial,"
My_SQL = My_SQL + "  dbo.Notes.ORDER_NO, dbo.DOUBLE_ENTREY_VOUCHERS.ItemID, dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,"
My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.buy , dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1, dbo.notes_all.BasedONID ,dbo.notes_all.VATCustoms"
My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
My_SQL = My_SQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
My_SQL = My_SQL + " dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"

If Me.TxtModFlg.Text = "R" Or Me.TxtModFlg.Text = "" Then
            If CBoBasedON.ListIndex = 0 Then ' »·«
            My_SQL = My_SQL + " WHERE  ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and    (dbo.Notes.NoteType = 80)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.Text)
            Else
            My_SQL = My_SQL + " WHERE   ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and    (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2) "
            End If

ElseIf Me.TxtModFlg.Text = "E" Then


            If CBoBasedON.ListIndex = 0 Then
         '   My_SQL = My_SQL + " WHERE   ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and   (   dbo.Notes.NoteType = 80   AND  dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0 and     dbo.notes_all.BasedONID = 0   and   ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) ) or  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.text)
            My_SQL = My_SQL + " WHERE  ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and    (dbo.Notes.NoteType = 80)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.Text)
            Else
            My_SQL = My_SQL + " WHERE     ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and   (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2)"
            End If

ElseIf Me.TxtModFlg.Text = "N" Then


            If CBoBasedON.ListIndex = 0 Then
            My_SQL = My_SQL + " WHERE   ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and     (dbo.Notes.NoteType = 80)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and   ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) "
            Else
            My_SQL = My_SQL + " WHERE    ( DOUBLE_ENTREY_VOUCHERS.FlgVat is null ) and    (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2) and     ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) "
            End If
            
End If

My_SQL = My_SQL + " and ( dbo.DOUBLE_ENTREY_VOUCHERS.hideline = 0 or dbo.DOUBLE_ENTREY_VOUCHERS.hideline is null)"
My_SQL = My_SQL + "  order by dbo.DOUBLE_ENTREY_VOUCHERS.buy desc ,dbo.Notes.NoteSerial1"
    RsExp.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

'    Dim StrSQL As String
'    Dim rs As New ADODB.Recordset

    With Me.Grid4
        .rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst
TxtVATCustoms.Text = IIf(IsNull(RsExp("VATCustoms").value), 0, RsExp("VATCustoms").value)
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")) = IIf(IsNull(RsExp.Fields("Double_Entry_Vouchers_ID").value), 0, RsExp.Fields("Double_Entry_Vouchers_ID").value)
           
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsExp.Fields("ItemID").value), "", RsExp.Fields("ItemID").value)
    
                StrSQL = "select * from TblItems where ItemID=" & val(.TextMatrix(i, .ColIndex("ItemID")))
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(i, .ColIndex("ItemName")) = ""
                    .TextMatrix(i, .ColIndex("ItemCode")) = ""
 
                End If
               
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Account_Name").value), "", RsExp.Fields("Account_Name").value)
 
                Else
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Account_NameEng").value), "", RsExp.Fields("Account_NameEng").value)
                End If
 
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
 
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), "", RsExp.Fields("NoteID").value)
 
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsExp.Fields("Value").value), "", RsExp.Fields("Value").value)
 
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsExp.Fields("Account_Code").value), "", RsExp.Fields("Account_Code").value)
 
                If IsNull(RsExp.Fields("buy").value) Then
                    .TextMatrix(i, .ColIndex("Select")) = 0
                Else

                    If RsExp.Fields("buy").value = False Then
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    ElseIf RsExp.Fields("buy").value = True Then
                        .TextMatrix(i, .ColIndex("Select")) = 1
                    Else
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    End If
           
                End If
 
                .TextMatrix(i, .ColIndex("Select")) = 1
                
                ' .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("buy").value), _
                  0, RsExp.Fields("buy").value)
                  If CBoBasedON.ListIndex = 1 And Me.TxtModFlg.Text = "R" Then
             ' .TextMatrix(i, .ColIndex("Select")) = 1
              End If
              
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid4.Visible = True

    ' End If
  
    update_finincial_invoice_total
       
End Sub

Private Sub save_expenses()
   Dim Item_ID As Integer
    Dim i As Double
    Dim sql As String
    ' ﬁÊ„ » ÕÌÀ ﬂ· ”ÿ— ›Ì «·ﬁÌœ »Ê÷⁄ —ﬁ„ «·⁄„·Ì… Ê—ﬁ„ «·’‰› Ê buy - Notes
    'Item_ID = Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code")))
    ' ›—Ì€  ﬂ·›… «·‘Õ‰ ⁄·Ï „” ÊÏ «·”ÿ—

    With Grid

        For i = 1 To Grid.rows - 1
      
         '   Cn.BeginTrans
 
           ' If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           '     check_item_Exist_in_Grid val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("Note_value"))), True
        '
        '        sql = "update notes set Transaction_ID1=" & val(Me.XPTxtBillID.text) & " , buy='1',itemid=" & IIf(val(.TextMatrix(i, .ColIndex("itemid"))) = 0, "Null", val(.TextMatrix(i, .ColIndex("itemid")))) & " where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
        '
        '    Else
        '        sql = "update notes set Transaction_ID1=null ,  buy=Null,itemid=Null  where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
'
'            End If


          If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                check_item_Exist_in_Grid val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("Note_value")))
        
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=" & val(Me.XPTxtBillID.Text) & " , buy='1',itemid=" & IIf(val(.TextMatrix(i, .ColIndex("itemid"))) = 0, "Null", val(.TextMatrix(i, .ColIndex("itemid")))) & " where Double_Entry_Vouchers_ID=" & val(.TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")))
        
            Else
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=null , buy=Null,itemid=Null where Double_Entry_Vouchers_ID=" & val(.TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")))

            End If
            
            Cn.Execute sql

          '  Cn.CommitTrans

        Next

    End With

    Expenses_update_total

End Sub

Function Expenses_update_total()
    Dim i As Integer
    On Error Resume Next
    Txt_EXport.Text = 0

    If Grid.rows = 1 Then Exit Function

    With Grid

        For i = 1 To Grid.rows - 1
        
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked And Grid.TextMatrix(i, Grid.ColIndex("ItemID")) = "" Then
            
                Txt_EXport.Text = val(Txt_EXport.Text) + val(Grid.TextMatrix(i, Grid.ColIndex("note_value")))
            End If
            
            If val(Grid.TextMatrix(i, Grid.ColIndex("select"))) = 0 Then
                Grid.TextMatrix(i, Grid.ColIndex("ItemID")) = ""
                Grid.TextMatrix(i, Grid.ColIndex("ItemCode")) = ""
                Grid.TextMatrix(i, Grid.ColIndex("ItemName")) = ""
            
            End If
            
        Next
 
    End With
       
End Function

Private Sub Save_Financial_invoice()
    'FG.TextMatrix(FG.Row, FG.ColIndex("LineShahn")) = Val(Me.txt_item_expenses.text)
    ' ﬁÊ„ » ÕÌÀ ﬂ· ”ÿ— ›Ì «·ﬁÌœ »Ê÷⁄ —ﬁ„ «·⁄„·Ì… Ê—ﬁ„ «·’‰› Ê buy - Double entry Voucher
    Dim Item_ID As Integer
    Dim i As Integer
    Dim sql As String

    'Item_ID = Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code")))
    ' ›—Ì€  ﬂ·›… «·‘Õ‰ ⁄·Ï „” ÊÏ «·”ÿ—
    With FG

        For i = 1 To FG.rows - 1
        
            .TextMatrix(i, .ColIndex("LineShahn")) = 0
      
        Next i

    End With
    If Not IsClicKCommand4 Then Exit Sub
    With Grid4
 
        For i = 1 To Grid4.rows - 1
      
'            Cn.BeginTrans
 
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                check_item_Exist_in_Grid val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("Note_value")))
        
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=" & val(Me.XPTxtBillID.Text) & " , buy='1',itemid=" & IIf(val(Grid4.TextMatrix(i, Grid4.ColIndex("itemid"))) = 0, "Null", val(Grid4.TextMatrix(i, Grid4.ColIndex("itemid")))) & " where Double_Entry_Vouchers_ID=" & val(Grid4.TextMatrix(i, Grid4.ColIndex("Double_Entry_Vouchers_ID")))
        
            Else
          '      sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=null , buy=Null,itemid=Null where Double_Entry_Vouchers_ID=" & val(Grid4.TextMatrix(i, Grid4.ColIndex("Double_Entry_Vouchers_ID")))
                sql = "update DOUBLE_ENTREY_VOUCHERS " & _
      "set Transaction_ID1=null, buy=null, itemid=null " & _
      "where Double_Entry_Vouchers_ID=" & val(Grid4.TextMatrix(i, Grid4.ColIndex("Double_Entry_Vouchers_ID"))) & _
      "  and Transaction_ID1=" & val(Me.XPTxtBillID.Text)


            End If

            Cn.Execute sql

'            Cn.CommitTrans

        Next

    End With

    update_finincial_invoice_total

    '    DoEvents
    '    Command4_Click
End Sub

Function update_finincial_invoice_total()
    On Error Resume Next
    Dim i As Integer
    txt_total_bill.Text = 0

    If Grid4.rows = 1 Then Exit Function

    With Grid4

        For i = 1 To Grid4.rows - 1
        
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked And Grid4.TextMatrix(i, Grid4.ColIndex("ItemID")) = "" Then
                txt_total_bill.Text = val(txt_total_bill.Text) + val(Grid4.TextMatrix(i, Grid4.ColIndex("note_value")))
  
            End If
            
            If val(Grid4.TextMatrix(i, Grid4.ColIndex("select"))) = 0 Then
                Grid4.TextMatrix(i, Grid4.ColIndex("ItemID")) = ""
                Grid4.TextMatrix(i, Grid4.ColIndex("ItemCode")) = ""
                Grid4.TextMatrix(i, Grid4.ColIndex("ItemName")) = ""
            
            End If

        Next

    End With

End Function

Private Sub Command5_Click()

    Save_Financial_invoice
       
End Sub

Private Sub Command6_Click()
    save_expenses
End Sub

Private Sub DBCboClientName_Change()
    Dim StrSQL As String
    Dim CreditInterval As Integer
    Dim CreditIntervalID As Integer
    Dim RsTemp As ADODB.Recordset

    On Error GoTo ErrTrap
    Dim fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 2
    TxtSearchCode.Text = fullcode

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
                '   CboPayMentType.locked = True
                '   CboPayMentType.ListIndex = 0
            Else
                '   CboPayMentType.locked = False
            End If
        End If
    End If

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        StrSQL = "Select * From TblCustemers Where CusID=" & val(DBCboClientName.BoundText)
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not (IsNull(RsTemp("Trans_DiscountTypePur").value)) Then
                If RsTemp("Trans_DiscountTypePur").value = 0 Then
                    '     mina           Me.XPCboDiscountType.ListIndex = 0
                    '   mina             Me.XPTxtDiscountVal.text = 0
                ElseIf RsTemp("Trans_DiscountTypePur").value = 1 Then
                    Me.XPCboDiscountType.ListIndex = 1
                    Me.XPTxtDiscountVal.Text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                ElseIf RsTemp("Trans_DiscountTypePur").value = 2 Then
                    Me.XPCboDiscountType.ListIndex = 2
                    Me.XPTxtDiscountVal.Text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                End If

            Else
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.Text = 0
            End If

        Else
            Me.XPCboDiscountType.ListIndex = 0
            '     mina   Me.XPTxtDiscountVal.text = 0
        End If
      
      Me.TxtVATNO.Text = IIf(IsNull(RsTemp("VATNO").value), "", RsTemp("VATNO").value)
      CreditInterval = IIf(IsNull(RsTemp("CreditInterval").value), 0, RsTemp("CreditInterval").value)
      CreditIntervalID = IIf(IsNull(RsTemp("CreditIntervalID").value), 0, RsTemp("CreditIntervalID").value)
      If CreditIntervalID = 0 Then
      DtpDelayDate.value = DateAdd("d", CreditInterval, BLDate.value)
      ElseIf CreditIntervalID = 1 Then
      DtpDelayDate.value = DateAdd("m", CreditInterval, BLDate.value)
      ElseIf CreditIntervalID = 2 Then
      DtpDelayDate.value = DateAdd("yyyy", CreditInterval, BLDate.value)
      End If
      
      
      
        RsTemp.Close
        Set RsTemp = Nothing
                
    End If

    If BillBasedOn(1).value = True Then
                
        FillVoucherGrid (1)
                
    End If
    
    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 1
        FrmCompanySearch.show vbModal
    
    End If
    
    
        If KeyCode = vbKeyF5 Then
        ReloadCombos

    End If
    
 

End Sub
Function ReloadCombos()
             Dim Dcombos As New ClsDataCombos
 Dim StrSQL As String
 
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True
    Dcombos.GetBranches dcBranch
    Dcombos.GetDocTypebyid Me.DCDocTypes, 22, val(Me.dcBranch.BoundText)
    Dcombos.GetBanks Me.DcboBankName
    StrSQL = "  select  BankID,BankName  from BanksData   "
    fill_combo Dcbanks, StrSQL
    StrSQL = " select id,code from currency"
    fill_combo Me.Dccurrency, StrSQL
    StrSQL = " select id,Project_name from projects"
    fill_combo Me.dcproject, StrSQL
 Dcombos.GetStores Me.DCboStoreName
End Function

Private Sub dcbanks_KeyUp(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF5 Then
        ReloadCombos

    End If
End Sub

Private Sub DcboBox_KeyUp(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF5 Then
        ReloadCombos

    End If
    
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 3
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 3
        FrmItemSearch.show vbModal
    End If
    
End Sub

Private Sub DCboStoreName_Change()

TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 6) = True Or CheckStoreCoding(val(dcBranch.BoundText), 9) = True Then
     TxtNoteSerial1V = ""
     
StroreChanged = True



  CurrentVoucherNo = ""
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    DateChanged = True
    
    
     End If
     
    End If


End Sub

 

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF5 Then
        ReloadCombos

    End If
    
    
End Sub

Private Sub Dcbranch_Change()
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        Dcombos.GetDocTypebyid Me.DCDocTypes, 22, val(Me.dcBranch.BoundText)
    End If

    If dcBranch.BoundText = "" Then TxtNoteSerial1.locked = True: Exit Sub

    If Voucher_coding(val(Me.dcBranch.BoundText), XPDtbBill.value, 6, 150, 22) = "" Then
        TxtNoteSerial1.locked = False
    Else
        TxtNoteSerial1.locked = True
 
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    Dcbranch_Change

    If Voucher_coding(val(val(Me.dcBranch.BoundText)), XPDtbBill.value, 6, 150, , 22) = "" Then Exit Sub
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    
         TxtNoteSerial1V = ""
     




  CurrentVoucherNo = ""
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    DateChanged = True
    
    
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

 '   If KeyCode = vbKeyF5 Then
 '       Dim Dcombos As ClsDataCombos
 '
 '       Set Dcombos = New ClsDataCombos
 '
 '   End If
    
            If KeyCode = vbKeyF5 Then
        ReloadCombos

    End If
    
    

End Sub

Private Sub DcCurrency_Change()

    If Me.TxtModFlg.Text = "" Or Me.TxtModFlg.Text = "R" Or cmdReSave.Visible = True Then Exit Sub
    If Me.Dccurrency.BoundText <> "" Then
        txt_Currency_rate.Text = get_currency_rate(Me.Dccurrency.BoundText)
    Else
        txt_Currency_rate.Text = 1
    End If
     ReLineGrid
     ChAddToTotal_Click
End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim StrSQL As String
        StrSQL = " select id,code from currency"
 
        fill_combo Me.Dccurrency, StrSQL
    End If

End Sub

Private Sub DCDocTypes_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

   ' If KeyCode = vbKeyF5 Then
   '     Dim Dcombos As ClsDataCombos
        '
        'Set Dcombos = New ClsDataCombos
        'Dcombos.GetDocTypebyid Me.DCDocTypes, 22, val(Me.dcBranch.BoundText)
    'End If


            If KeyCode = vbKeyF5 Then
        ReloadCombos

    End If
    
End Sub

Private Sub Ele_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 6

            If Me.WindowState = vbNormal Then
                Me.WindowState = vbMaximized
            Else
                Me.WindowState = vbNormal
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    fill_bill_items_table

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 150
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 150

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
End Sub

Private Sub fg_Click()

    'Command4_Click
End Sub

Function fill_bill_items_table() ' ﬁÊ„ Â–… «·œ«·Â »Õ›Ÿ «’‰«› «·›« Ê—… ›Ì ÃœÊ· „ƒﬁ  ·«” Œœ«„Â« ›Ì «· Ê“Ì⁄ ⁄·Ï «·„’—Ê›«   Ê«·›Ê« Ì—
    Dim bill_items As ADODB.Recordset
    Set bill_items = New ADODB.Recordset
    Dim StrSqlDel As String
    Dim RowNum As Integer
    bill_items.Open "[temp_bill_items]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    StrSqlDel = "delete From temp_bill_items"
    Cn.Execute StrSqlDel, , adExecuteNoRecords
 
    With FG

        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
20152015 '                bill_items.AddNew
'                bill_items("ItemID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
'                bill_items.update
            End If

        Next RowNum

    End With

End Function

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)
        
            Case "Vatyo"
                If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
                    .TextMatrix(Row, .ColIndex("Vat")) = 0
                    If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) <> 0 Then
                        .TextMatrix(Row, .ColIndex("Price")) = val(.TextMatrix(Row, .ColIndex("PriceTotal")))
                    End If
                    If .rows > Row Then
                        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
                            .RemoveItem Row + 1
                        End If
                    End If
                End If
            Case "PriceTotal"
                AddVAT Row
                Me.TXTFactoryExpenses.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Price"), .rows - 1, .ColIndex("Price"))
                TXTFactoryExpensesVat.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat"))
           Case "Account_Name2"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Accountcode2"), False, True)
                .TextMatrix(Row, .ColIndex("Accountcode2")) = StrAccountCode
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If
AddVAT Row
            Case "value", "Price", "ChSameCurrncy"
                Dim sgl As String
           
                .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("Price")))
                AddVAT Row
                Me.TXTFactoryExpenses.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
                
                
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.TXTFactoryExpenses.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If
          TXTFactoryExpensesVat.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat"))
        ' ReLineGrid
    End With

    ReLineGrid
    ChAddToTotal_Click
End Sub

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
Dim SumValue As Double
SumValue = 0
    With Fg_Journal

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                                Dim sgl As String
 
                .TextMatrix(i, .ColIndex("value")) = val(.TextMatrix(i, .ColIndex("Price")))
                       
            End If
            If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
            SumValue = SumValue + val(.TextMatrix(i, .ColIndex("Value")))
           End If
        Next i
Me.TXTFactoryExpenses.Text = SumValue
TXTFactoryExpensesVat.Text = Fg_Journal.Aggregate(flexSTSum, Fg_Journal.FixedRows, Fg_Journal.ColIndex("Vat"), Fg_Journal.rows - 1, Fg_Journal.ColIndex("Vat"))


    End With

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)
            Case "Vat"
                 Cancel = True
        Case "Vatyo"
              If val(.TextMatrix(Row, .ColIndex("ForcedFlg"))) = 1 Then
                 Cancel = True
              Else
              .ComboList = ""
              End If
            Case "value"
               Cancel = True
              Case "Price"
                .ComboList = ""
              Case "ChSameCurrncy"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.cell(flexcpData, r, c)) <> "String" Then
            TxtDes.Text = ""
        Else
            '
            TxtDes.Text = Fg_Journal.cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                  
                    Order_no_search.show
                     Order_no_search.RetrunType = 4
                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.show
                    FrmExpensesSearch.RetrunType = 9
                End If
            Case "Account_Name2"
                  If KeyCode = vbKeyF3 Then
                      Account_search.show
                     Account_search.case_id = 350055
         End If
 
        End Select

    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
          Case "Account_Name2"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "AccountName"

                '      StrSQL = "select * from Expenses_accounts"
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts"
                Else
                    StrSQL = "select * from Expenses_accounts_eng "
                End If
                 
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                'StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
            
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                StrSQL = "  select fullcode,name from terms_operations "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode", "fullcode")

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub


Sub AddVAT(Optional Row As Long)
If True = True Then
Dim ForcedFlg As Integer
Dim valuee As Double
Dim AccountVATDept As String
Dim i As Integer
Dim k As Integer
Dim ClsAcc  As New ClsAccounts

With Fg_Journal
.TextMatrix(Row, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("Price"))
.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbBill.value, .TextMatrix(Row, .ColIndex("AccountCode2")), val(dcBranch.BoundText), ForcedFlg)

If val(txtManulaVat.Text) > 0 Then
.TextMatrix(Row, .ColIndex("Vatyo")) = val(txtManulaVat.Text)

End If

.TextMatrix(Row, .ColIndex("Rate")) = 1
.TextMatrix(Row, .ColIndex("Rate")) = val(.TextMatrix(Row, .ColIndex("Vatyo"))) / 100 + 1
If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) > 0 And val(.TextMatrix(Row, .ColIndex("Rate"))) > 0 Then
.TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("PriceTotal"))) / val(.TextMatrix(Row, .ColIndex("Rate"))), 2)

End If
valuee = val(.TextMatrix(Row, .ColIndex("value")))

.TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg
'.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)

'new idea
 If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) > 0 Then
.TextMatrix(Row, .ColIndex("Vat")) = Round(val(.TextMatrix(Row, .ColIndex("PriceTotal"))) - valuee, 2)
Else
.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)
End If


GetValueAddedAccount XPDtbBill.value, AccountVATDept
If AccountVATDept = "" And val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «·Õ”«» «·„œÌ‰ ›Ì ‘«‘… «⁄œ«œ  «·›« "
Else
MsgBox "Please Enter Account In VAT Settings"
End If
.TextMatrix(Row, .ColIndex("Vat")) = 0
.TextMatrix(Row, .ColIndex("Vatyo")) = 0
Exit Sub
End If
''/////////////
If val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
   If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
    DeleteGridCurrRow Row
   For i = 1 To 1
         .AddItem " ", .Row + i
  k = .Row + i
.TextMatrix(k, .ColIndex("CurrRow")) = Row
 
If i = 1 Then
'.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(AccountVATDept)
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_Name(, AccountVATDept)
.TextMatrix(k, .ColIndex("AccountCode")) = AccountVATDept
.TextMatrix(k, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("Vat"))
.TextMatrix(k, .ColIndex("Price")) = .TextMatrix(Row, .ColIndex("Vat"))
Else
'.TextMatrix(k, .ColIndex("AccountCode")) = DcboCredi.BoundText
'.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, DcboCreditSide.BoundText)
'.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Vat"))
.TextMatrix(k, .ColIndex("Price")) = .TextMatrix(Row, .ColIndex("Vat"))
End If
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " ﬁÌ„… „÷«›…"
Else
.TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " VAT  "
End If
.TextMatrix(k, .ColIndex("FlgVat")) = 1
.TextMatrix(k, .ColIndex("ExpensesID")) = .TextMatrix(Row, .ColIndex("ExpensesID"))
.TextMatrix(k, .ColIndex("opr_fullcode")) = .TextMatrix(Row, .ColIndex("opr_fullcode"))
'.TextMatrix(k, .ColIndex("CarName")) = .TextMatrix(Row, .ColIndex("CarName"))
.TextMatrix(k, .ColIndex("Order_No")) = .TextMatrix(Row, .ColIndex("Order_No"))
'.TextMatrix(k, .ColIndex("CarId")) = .TextMatrix(Row, .ColIndex("CarId"))
Next i
End If
End If
End With
End If
End Sub

Sub DeleteGridCurrRow(Optional CurrRow As Long)
Dim i As Integer
With Fg_Journal
i = .rows
Do
i = i - 1
If val(.TextMatrix(i, .ColIndex("CurrRow"))) = CurrRow Then
.RemoveItem i
End If
Loop While i > 1
End With
End Sub


Function fillExpensesFactoryGrid()
 
    '  «·’‰«⁄Ì…   ⁄»∆… «·«–Ê‰ «·„’—Ê›« 
    With Me.Fg_Journal
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
  '  My_SQL = "SELECT * from TblProductOrderFactoryexpenses where Transaction_ID=" & val(Me.XPTxtBillID.Text)
My_SQL = " SELECT     dbo.TblProductOrderFactoryExpenses.id, dbo.TblProductOrderFactoryExpenses.Transaction_ID, dbo.TblProductOrderFactoryExpenses.[Value],"
My_SQL = My_SQL & "                       dbo.TblProductOrderFactoryExpenses.des, dbo.TblProductOrderFactoryExpenses.ChSameCurrncy, dbo.TblProductOrderFactoryExpenses.Price,"
My_SQL = My_SQL & "                      dbo.TblProductOrderFactoryExpenses.Accountcode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng,"
My_SQL = My_SQL & "                      dbo.TblProductOrderFactoryExpenses.Accountcode2, ACCOUNTS_1.Account_Name AS Account_Name2, ACCOUNTS_1.Account_Serial AS Account_Serial2,"
My_SQL = My_SQL & "                      ACCOUNTS_1.Account_NameEng AS Account_NameE2,TblProductOrderFactoryExpenses.FlgVat,TblProductOrderFactoryExpenses.Vatyo,"
My_SQL = My_SQL & "                      TblProductOrderFactoryExpenses.Vat ,TblProductOrderFactoryExpenses.ForcedFlg,TblProductOrderFactoryExpenses.PriceTotal,TblProductOrderFactoryExpenses.CurrRow"
My_SQL = My_SQL & " FROM         dbo.TblProductOrderFactoryExpenses LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblProductOrderFactoryExpenses.Accountcode2 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.ACCOUNTS ON dbo.TblProductOrderFactoryExpenses.Accountcode = dbo.ACCOUNTS.Account_Code"
My_SQL = My_SQL & " Where (dbo.TblProductOrderFactoryExpenses.Transaction_ID = " & val(XPTxtBillID.Text) & ")"
    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim StrSQL  As String

    With Me.Fg_Journal
        .rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .rows - 1
                   
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Accountcode2")) = IIf(IsNull(RsExp.Fields("Accountcode2").value), "", RsExp.Fields("Accountcode2").value)
                .TextMatrix(i, .ColIndex("Accountcode")) = IIf(IsNull(RsExp.Fields("Accountcode").value), "", RsExp.Fields("Accountcode").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Account_Name2")) = IIf(IsNull(RsExp.Fields("Account_Name2").value), "", RsExp.Fields("Account_Name2").value)
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsExp.Fields("Account_Name").value), "", RsExp.Fields("Account_Name").value)
            Else
                .TextMatrix(i, .ColIndex("Account_Name2")) = IIf(IsNull(RsExp.Fields("Account_NameE2").value), "", RsExp.Fields("Account_NameE2").value)
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsExp.Fields("Account_NameEng").value), "", RsExp.Fields("Account_NameEng").value)
            End If
               
               .TextMatrix(i, .ColIndex("CurrRow")) = IIf(IsNull(RsExp.Fields("CurrRow").value), "", RsExp.Fields("CurrRow").value)
               .TextMatrix(i, .ColIndex("FlgVat")) = IIf(IsNull(RsExp.Fields("FlgVat").value), "", RsExp.Fields("FlgVat").value)
               .TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(RsExp.Fields("Vatyo").value), "", RsExp.Fields("Vatyo").value)
               .TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(RsExp.Fields("Vat").value), "", RsExp.Fields("Vat").value)
               .TextMatrix(i, .ColIndex("ForcedFlg")) = IIf(IsNull(RsExp.Fields("ForcedFlg").value), "", RsExp.Fields("ForcedFlg").value)
               .TextMatrix(i, .ColIndex("PriceTotal")) = IIf(IsNull(RsExp.Fields("PriceTotal").value), "", RsExp.Fields("PriceTotal").value)
               
               

            
               
                .TextMatrix(i, .ColIndex("value")) = IIf(Not IsNumeric(RsExp.Fields("value").value), 0, RsExp.Fields("value").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(Not IsNumeric(RsExp.Fields("Price").value), 0, RsExp.Fields("Price").value)
                .TextMatrix(i, .ColIndex("ChSameCurrncy")) = IIf(Not IsNumeric(RsExp.Fields("ChSameCurrncy").value), 0, RsExp.Fields("ChSameCurrncy").value)
            
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsExp.Fields("des").value), "", RsExp.Fields("des").value)
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    With Me.Fg_Journal
        Me.TXTFactoryExpenses.Text = .Aggregate(flexSTSum, .FixedRows - 1, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
    End With
 
End Function

Private Sub Form_Activate()
    Set m_MnuShowNewItemsPrices = mdifrmmain.MnuInvPurchaseMnu2
    Set m_MenuViewList = mdifrmmain.MnuInvPurchaseMnu1
    Set m_MenuShowItemCostEffect = mdifrmmain.MnuInvPurchaseMnu4

End Sub

Private Sub CmdRetruns_Click()
    ShowRelatedTransactions val(Me.XPTxtBillID.Text), 1
End Sub

Private Sub Form_Resize()
'  Me.WindowState = 2
End Sub

Public Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
       
    With Grid

        Select Case .ColKey(Col)
   
            Case "ItemID"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))
    
                StrSQL = "select * from QRY_temp_bill_items where ItemID=" & Trim(.TextMatrix(Row, Col))
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
 
                End If
 
                check_item_Exist_in_Grid val(.TextMatrix(Row, .ColIndex("ItemID"))), val(.TextMatrix(Row, .ColIndex("Note_value")))

            Case "ItemCode"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "select * from QRY_temp_bill_items where ItemCode='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                    
                Else
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
 
                End If
 
            Case "ItemName"
                  
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
    
                Set ClsAcc = New ClsAccounts
      
                .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
                 
                StrSQL = "select * from QRY_temp_bill_items where ItemID= " & val(StrAccountCode)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
            
                    .TextMatrix(Row, .ColIndex("ItemCode")) = rs("ItemCode").value
                Else
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                   
                End If

        End Select

        'to Add new row if needed
        If Row = .rows - 1 Then
            '    .Rows = .Rows + 1
        End If

    End With

    Expenses_update_total
     
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "ItemName" Then
            .ComboList = ""
        End If
   
    End With

End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
 
    On Error GoTo ErrTrap

    With Me.Grid

        Select Case .ColKey(Col)

            Case "Items"
                On Error GoTo ErrTrap

         
         
         
         Dim Frm As FrmItemsDetails1
    Set Frm = FrmItemsDetails1
    Dim Item_ID As Integer
    Dim GroupID As Integer
    Dim ExpiryValue As Integer
    Dim ExpiryType As Integer

    With Me.Grid
        Frm.lbl(3).Caption = (Row + 1)
        Frm.lbl(4).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("NoteSerial"))
        Frm.lbl(5).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("name"))
        Frm.TxtValue = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Note_value"))
        
       ' Item_ID = val(Me.Grid.TextMatrix(Row, .ColIndex("Code")))

' Frm.TxtUnitID = .Cell(flexcpData, .Row, .ColIndex("UnitID"))
'Frm.TxtItemID.text = Item_ID
 
    
        Set Frm.FG = Me.Grid
        Frm.LngRow = Row
        Frm.LngCol = Col
    
        'Frm.GranteeStartDate.value = Date
    
        If 1 = 1 Then
        '    If FrmBillBuy.TxtModFlag.text = "R" Then
             '  Frm.TxtComment.locked = True
        '      Frm.Cmd(2).Enabled = False
        '     Frm.Grid.Enabled = False
             '    Frm.txtWages.Enabled = False
                
        '    End If

 
        
            
            If Me.Grid.ColIndex("SelectedItem") <> -1 Then
                Frm.AllIDS = IIf(Grid.TextMatrix(Row, Grid.ColIndex("SelectedItem")) = "", "", Grid.TextMatrix(Row, Grid.ColIndex("SelectedItem")))
            End If
            
            
            
 
            Frm.FillGridWithData
        
        End If

    End With

    Frm.show
   
   
   
              
          End Select
          
          End With
ErrTrap:
End Sub

Private Sub Grid_Click()
    ' Expenses_update_total

End Sub

Public Function close_order2(order_no As String)
   If Trim(TXT_order_no) = "" Then Exit Function
    
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim i As Integer
    Dim Result As Integer
    Result = 1
    StrSQL = "select * from  items_qty_not_recieved_in_order where  order_no='" & order_no & "'"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then

        For i = 1 To RsDetails.RecordCount

            If IsNull(RsDetails("net").value) Then Result = 0: GoTo ll
            If RsDetails("net").value <> 0 Then
                Result = 0
                GoTo ll
            End If

            RsDetails.MoveNext
        Next i
 
    End If

ll:
    Dim sql As String
    sql = "update Transactions Set closed = " & Result & " Where Transaction_Type = 6 and order_no='" & Me.TXT_order_no & "'"
    Cn.Execute sql

End Function

Public Function items_qty_not_recieved_in_order(Item_ID As Integer, _
                                                order_no As String) As Integer
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    StrSQL = "select * from  items_qty_not_recieved_in_order where Item_ID=" & Item_ID & " and order_no='" & order_no & "'"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then

        items_qty_not_recieved_in_order = IIf(IsNull(RsDetails("net").value), IIf(IsNull(RsDetails("sum_qty").value), 0, RsDetails("sum_qty").value), RsDetails("net").value)

    Else
        items_qty_not_recieved_in_order = 0
    End If

End Function

Function Retrive_orders_data(Transaction_ID As Double, Optional str As String)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim row_count As Double
    Dim Num As Double

    StrSQL = "Select * from transactions where Transaction_ID=" & Transaction_ID
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Function
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), 1, rs("Currency_id").value)
        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
        TxtLcNo.Text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))
    TxtBillComment.Text = IIf(IsNull(rs("TransactionComment")), "", (rs("TransactionComment").value))
Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    
 If IsNull(rs("chkTaxExempt").value) Then
        Me.chkTaxExempt.value = vbUnchecked
    Else
        Me.chkTaxExempt.value = IIf(rs("chkTaxExempt").value = 0, vbUnchecked, vbChecked)
    End If
    End If
    
  

    If rs.EOF Or rs.BOF Then
        Exit Function
    End If
If CBoBasedON.ListIndex = 1 Then
    StrSQL = "SELECT TblItems.HaveSerial, dbo.[GetBalanceQtyPO3] (Transaction_Details.Item_ID,N'" & Trim(TXT_order_no) & "'," & val(Me.XPTxtBillID) & ") as Showqty6 , * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
Else
    StrSQL = "SELECT TblItems.HaveSerial,Showqty as ShowQty6, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
End If
    'StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID
 If str = "" And Transaction_ID <> 0 Then
     StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID
 ElseIf str <> "" Then
 StrSQL = StrSQL + " where Transaction_ID in (" & str & ")"
Else
Exit Function
 End If
 If CBoBasedON.ListIndex = 1 Then
    StrSQL = StrSQL + " And dbo.[GetBalanceQtyPO3] (Transaction_Details.Item_ID,N'" & Trim(TXT_order_no) & "'," & val(Me.XPTxtBillID) & ") <> 0"
End If
'str
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
           Dim ClsAcc As ClsAccounts
     Set ClsAcc = New ClsAccounts
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.rows - 1 'RsDetails.RecordCount
    
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = DTArrivalDate.value
         
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
                  '    Fg.TextMatrix(Num, Fg.ColIndex("Count")) = items_qty_not_recieved_in_order(Fg.TextMatrix(Num, Fg.ColIndex("Code")), Fg.TextMatrix(Num, Fg.ColIndex("order_no")))
            
            
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Showqty6")), "", (RsDetails("Showqty6").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), "", (RsDetails("ShowPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("StoreID2")) = val(Me.DCboStoreName.BoundText)
            
            FG.TextMatrix(Num, FG.ColIndex("Account_Name")) = Trim(Me.cmbAccounts.Text)
            FG.TextMatrix(Num, FG.ColIndex("Account_Code")) = Trim(Me.cmbAccounts.BoundText)
           


                   ' .TextMatrix(row, .ColIndex("Account_Code")) = StrAccountCode
 
                    
  
            FG.TextMatrix(Num, FG.ColIndex("Account_Serial2")) = ClsAcc.Get_Account_Serial(Me.cmbAccounts.BoundText)
 
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassId")) = IIf(IsNull(RsDetails("ClassId")), 1, (RsDetails("ClassId").value))
            If val(FG.TextMatrix(Num, FG.ColIndex("DiscountType"))) = 2 Then
            
                FG.TextMatrix(Num, FG.ColIndex("Valu")) = (val(FG.TextMatrix(Num, FG.ColIndex("Count"))) * val(FG.TextMatrix(Num, FG.ColIndex("Price")))) - val(FG.TextMatrix(Num, FG.ColIndex("DiscountVal")))
            Else
                FG.TextMatrix(Num, FG.ColIndex("Valu")) = val(FG.TextMatrix(Num, FG.ColIndex("Count"))) * val(FG.TextMatrix(Num, FG.ColIndex("Price")))
            End If
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
    

 
 
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If
  chkTaxExempt_Click
End Function

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Grid

        Select Case .ColKey(Col)

            Case "ItemName"
       
                StrSQL = "Select * from QRY_temp_bill_items"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid4.BuildComboList(rs, "ItemName", "ItemID")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Grid1_Click()

    With GRID1

        Select Case .Col

            Case 2
 
                '       If .Cell(flexcpChecked, .Row, .ColIndex("select")) = flexChecked Then
                '            Retrive_orders_data (Val(.TextMatrix(.Row, .ColIndex("Transaction_ID"))))
                '
                '
                '        End If

                With FG
                    .Clear flexClearScrollable, flexClearEverything
                    .rows = 1
       
                End With
 
                fillVchr

            Case 8
            FrmInpout.XPBtnMove_Click (2)

                FrmInpout.Retrive val(.TextMatrix(.Row, 1))

            Case 9
                'ShowGL_cc .TextMatrix(.Row, .ColIndex("NoteSerial")), , 200
                ShowGL_cc val(.TextMatrix(.Row, .ColIndex("NoteSerial"))), , 200, val(.TextMatrix(.Row, .ColIndex("NoteID")))

        End Select

    End With

End Sub

Function fillVchr()
Dim str As String
Dim Transaction_ID As Double
str = ""

    Dim i As Integer
        
    With GRID1

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
            
            
            Transaction_ID = val(.TextMatrix(i, .ColIndex("Transaction_ID")))
                
           str = val(.TextMatrix(i, .ColIndex("Transaction_ID"))) & "," & str
            
            End If

        Next i

    End With

If str <> "" Then
str = mId(str, 1, Len(str) - 1)
End If

Retrive_orders_data Transaction_ID, str
                 If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                            NewGrid.DtpBillDate_Change
                            NewGrid.Calculate 1, , , True
                        End If

End Function

Private Sub GRID2_Click()

    With Grid2

        If .cell(flexcpChecked, .Row, .ColIndex("select")) = flexChecked Then
            Retrive_orders_data (val(Grid2.TextMatrix(Grid2.Row, Grid2.ColIndex("Transaction_ID"))))
            
        End If

    End With

End Sub
 
Private Function check_item_Exist_in_Grid(ItemID As Integer, _
                                          value As Single, _
                                          Optional addition As Boolean)
    Dim i As Integer
    On Error Resume Next

    With FG

        For i = 1 To FG.rows - 1

            If .TextMatrix(i, .ColIndex("Code")) = CStr(ItemID) Then
                If addition = False Then
                    .TextMatrix(i, .ColIndex("LineShahn")) = value
                Else
                    .TextMatrix(i, .ColIndex("LineShahn")) = val(.TextMatrix(i, .ColIndex("LineShahn"))) + value
                End If

                Exit Function
    
            End If

        Next i

    End With
 
End Function

Private Sub grid4_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
       
    With Grid4

        Select Case .ColKey(Col)
   
            Case "ItemID"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "select * from QRY_temp_bill_items where ItemID=" & Trim(.TextMatrix(Row, Col))
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
 
                End If
 
                check_item_Exist_in_Grid val(.TextMatrix(Row, .ColIndex("ItemID"))), val(.TextMatrix(Row, .ColIndex("Note_value")))

            Case "ItemCode"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))
         
                StrSQL = "select * from QRY_temp_bill_items where ItemCode='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                    
                Else
            
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
                    
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                End If
 
            Case "ItemName"
                  
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
    
                Set ClsAcc = New ClsAccounts
      
                .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
                 
                StrSQL = "select * from QRY_temp_bill_items where ItemID= " & val(StrAccountCode)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
            
                    .TextMatrix(Row, .ColIndex("ItemCode")) = rs("ItemCode").value
                    .TextMatrix(Row, .ColIndex("ItemID")) = rs("ItemID").value
                Else
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                   
                End If

        End Select

        'to Add new row if needed
        If Row = .rows - 1 Then
            '    .Rows = .Rows + 1
        End If

    End With

    update_finincial_invoice_total
End Sub

Private Sub grid4_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With Grid4

        If .ColKey(Col) <> "ItemName" Then
            .ComboList = ""
        End If
   
    End With

End Sub

Private Sub grid4_Click()
    'update_finincial_invoice_total
       
End Sub

Private Sub grid4_StartEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    IsClicKCommand4 = True
    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Grid4

        Select Case .ColKey(Col)

            Case "ItemName"
       
                StrSQL = "Select * from QRY_temp_bill_items"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid4.BuildComboList(rs, "ItemName", "ItemID")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
    FrmLC.show
    FrmLC.Retrive Trim(Me.TxtLcNo.Text)

End Sub

Private Sub ISButton2_Click()
ShowGL_cc TxtNoteSerial.Text, , 200, val(Me.TXTNoteID.Text)
End Sub

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalView.Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub lblexit_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
If Index = 90 Then
'
If val(CboPayMentType.ListIndex) = 2 And SystemOptions.AllowPurchasesMultyPayed = True Then
If val(TxtRemainValue2.Text) <> 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ…"
Else
MsgBox "The  value is incorrect"
End If
Exit Sub
End If
FramePay.Visible = False
End If
End If
Else
FramePay.Visible = False
End If
End Sub

Private Sub LblTotal_Change()

    If CboPayMentType.ListIndex = 1 Then
        XPTxtValue(1).Text = LblTotal.Caption
    ElseIf CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 2 Then
        XPTxtValue(0).Text = LblTotal.Caption
    End If
         
    LblTotalView.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblTotal_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Private Sub LblTotalAll_Change()
    LblTotalAllView.Caption = Format(val(LblTotalAll.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub m_FrmSearch_Unload(Cancel As Integer)
    Set m_FrmSearch = Nothing
End Sub

Private Sub m_MenuShowItemCostEffect_Click()

    If Me.TxtModFlg.Text = "R" Then
        ShowItemCostEffectForTrans 1, , Trim$(Me.TxtTransSerial.Text)
    End If

End Sub

Private Sub m_MenuViewList_Click()
    Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

    With FG
        .Cols = 9
        .RowHeightMin = 320
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .ColKey(0) = "Transaction_ID"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·„Ê—œ"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
    
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 7) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 8) = "≈Ã„«·Ï «·›« Ê—…"

        ',
        'QryTransactionsTotal.TransSum
        'QryTransactionsTotal.TransNet,
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT TOP 100 PERCENT QryTransactionsTotal.Transaction_ID," & "QryTransactionsTotal.Transaction_Serial, QryTransactionsTotal.Transaction_Date, " & "dbo.TblCustemers.CusName, QryTransactionsTotal.PaymentType, dbo.TblStore.StoreName," & "QryTransactionsTotal.Trans_DiscountType,QryTransactionsTotal.Trans_Discount ," & "QryTransactionsTotal.TotalAfterTax "
            StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal LEFT OUTER JOIN "
            StrSQL = StrSQL + "dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID " & "LEFT OUTER JOIN dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
            StrSQL = StrSQL + " Where (QryTransactionsTotal.Transaction_Type = 1)"
            StrSQL = StrSQL + " ORDER BY QryTransactionsTotal.Transaction_ID "
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT QryTransactionsTotal.Transaction_ID , QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.Transaction_Date,TblCustemers.CusName, QryTransactionsTotal.PaymentType," & "TblStore.StoreName,TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax "
            StrSQL = StrSQL + "FROM (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN QryTransactionsTotal " & "ON TblCustemers.CusID = QryTransactionsTotal.CusID) ON TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID) " & "LEFT JOIN TblStore ON QryTransactionsTotal.StoreID = TblStore.StoreID "
            StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type=1 "
            StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adAsyncExecute + adAsyncFetch
        Set cProgress = New ClsProgress
        BolFrmLoaded = True
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop

        If BolFrmLoaded = True Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .ColKey(0) = "Transaction_ID"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·„Ê—œ"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 7) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 8) = "≈Ã„«·Ï «·›« Ê—…"
        .ColKey(8) = "TotalAfterTax"
        'Rs.Close
        'Set Rs = Nothing
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "TotalAfterTax"
    FrmView.vsfGroup1.update
    FrmView.BolRetrunOnDblClick = True
    FrmView.SetDblClickRetrun Me, "Transaction_ID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·›Ê« Ì— «·„‘ —Ì« "
    FrmView.show
End Sub

Private Sub m_MnuShowNewItemsPrices_Click()

    If Not NewGrid Is Nothing Then
        NewGrid.ShowNewItemsPrice
    End If

End Sub

Private Sub SearchCashCustomer_Click()
frmCashCustomerSearch.RetrunType = 1
frmCashCustomerSearch.show

End Sub

Private Sub Text3_Change()
RelinVatGrid
End Sub

Private Sub Txt_EXport_Change()
    Me.TXTToTAlELSHahn.Text = IIf(Not IsNumeric(Txt_EXport.Text), 0, val(Txt_EXport.Text)) + IIf(Not IsNumeric(txt_total_bill.Text), 0, val(txt_total_bill.Text)) + IIf(Not IsNumeric(TXTFactoryExpenses.Text), 0, val(TXTFactoryExpenses.Text))
    Me.TXTToTAlELSHahn.Text = IIf(Not IsNumeric(Txt_EXport.Text), 0, val(Txt_EXport.Text)) + IIf(Not IsNumeric(txt_total_bill.Text), 0, val(txt_total_bill.Text)) + IIf(Not IsNumeric(TXTFactoryExpenses.Text), 0, val(TXTFactoryExpenses.Text) * val(Me.txt_Currency_rate.Text))
    Me.TXTToTAlELSHahn.Text = Round(Me.TXTToTAlELSHahn.Text, 2)
End Sub

Private Sub Txt_order_no_Change()
Dim StrSQL As String
    With Me.Grid4
        .rows = .FixedRows
 
    End With

    With Me.Grid
        .rows = .FixedRows
 
    End With

    If TXT_order_no.Text = "" Then
        txt_total_bill.Text = ""
        Txt_EXport.Text = ""
    End If
If Trim(Me.TxtModFlg.Text) <> "R" And Trim(Me.TxtModFlg.Text) <> "" Then
    Command4_Click
 End If
 
    Command2_Click
    Command3_Click
    Dim Transaction_ID As String
    Dim Transaction_Type As Integer

    If CBoBasedON.ListIndex = 1 Then
        Transaction_Type = 29
    ElseIf CBoBasedON.ListIndex = 2 Then
        Transaction_Type = 17
    Else
        Transaction_Type = 0
   '     Exit Sub
    End If

If CBoBasedON.ListIndex = 2 Then '›« Ê—… „»œ∆Ì…
    Transaction_ID = get_transactionData("NoteSerial1", TXT_order_no.Text, "Transaction_ID", Transaction_Type)
ElseIf CBoBasedON.ListIndex = 3 Then
       If Me.TxtModFlg = "R" Then Exit Sub
        Dim orderStatus As Integer
     
        MintDone = 0
        Dim rs2 As New ADODB.Recordset
        StrSQL = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXT_order_no.Text) & " "
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If rs2.RecordCount > 0 Then
                orderStatus = IIf(IsNull(rs2("OrderStatus").value), 0, rs2("OrderStatus").value)
                TxtCashCustomerName.Text = IIf(IsNull(rs2("ClientName").value), "", rs2("ClientName").value)
                'DCOPrType =
                
                
                
'
'                DcbyearFactor.Text = val(Rs2!YearFact & "")
'                TxtPlatNo = Trim(Rs2!PlateNo & "")
'                DcbCarType.BoundText = val(Rs2!CarTypeID & "")
'
'                TxtManualNo2(2).Text = Trim(Rs2!Shaseh & "")
'                 TxtManualNo2(1).Text = Trim(Rs2!CarMeter & "")
                
                
                 
                If orderStatus = 2 Or orderStatus = 4 Or orderStatus = 5 Then
                    MintDone = 1
                End If
                If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
                    Dim s As String
                    Dim RsData3 As New ADODB.Recordset
                    
                    
                    s = "Select TblCardAuthorizationReformItems.qty, tblitems.itemid,TblCardAuthorizationReformItems.Price ,TblCardAuthorizationReformItems.TotalWithVat ,tblItems.ItemCode,tblItems.ItemName from TblCardAuthorizationReformItems Left Outer Join tblItems On tblItems.ItemID =TblCardAuthorizationReformItems.ItemID Left Outer join TblCardAuthorizationReform On TblCardAuthorizationReform.Id = TblCardAuthorizationReformItems.id"
                    
                    s = s & "  Where (dbo.TblCardAuthorizationReform.WorkOrder = " & val(TXT_order_no.Text) & ") "
                           
                     RsData3.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     FG.rows = 1
                     Do While Not RsData3.EOF
                        FG.rows = FG.rows + 1
                  FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code")) = RsData3!ItemID & ""
                        FG.TextMatrix(FG.rows - 1, FG.ColIndex("Name")) = RsData3!ItemID & ""
                        FG.TextMatrix(FG.rows - 1, FG.ColIndex("Price")) = RsData3!Price & ""
                        FG.TextMatrix(FG.rows - 1, FG.ColIndex("Count")) = RsData3!Qty & ""
                        
                            FG.TextMatrix(FG.rows - 1, FG.ColIndex("ColorID")) = 1
        
            FG.TextMatrix(FG.rows - 1, FG.ColIndex("ItemSize")) = 1
            FG.TextMatrix(FG.rows - 1, FG.ColIndex("ClassID")) = 1
            Dim UnitID As Long
            Dim UnitName As String
           GetDefaultItemUnit val(FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code"))), UnitID, UnitName
 

     FG.cell(flexcpData, FG.rows - 1, FG.ColIndex("UnitID")) = UnitID
            FG.TextMatrix(FG.rows - 1, FG.ColIndex("UnitID")) = UnitName
        
        FG.TextMatrix(FG.rows - 1, FG.ColIndex("DiscountType")) = 1
            FG.TextMatrix(FG.rows - 1, FG.ColIndex("DiscountVal")) = 0

                        
                        RsData3.MoveNext
                     Loop
              If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                  NewGrid.DtpBillDate_Change

         NewGrid.Calculate 1, , , True
     End If
           
           
                     Exit Sub
                End If
            Else
                TxtCashCustomerName.Text = ""
                MintDone = -1
            End If
           ' LoadCar
           Exit Sub

Else
Transaction_ID = get_transactionData("order_no", TXT_order_no.Text, "Transaction_ID", Transaction_Type)
End If

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        Retrive_orders_data (val(Transaction_ID))
    End If
'NewGrid.DtpBillDate_Change
              If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                  NewGrid.DtpBillDate_Change

         NewGrid.Calculate 1, , , True
     End If
     chkTaxExempt_Click
End Sub

Private Sub TXT_total_payments_Change()
    'Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, Val(Txt_EXport.text)) + IIf(Not IsNumeric(TXT_total_payments.text), 0, Val(TXT_total_payments.text))
End Sub

Private Sub Txt_order_no_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim StrSQL As String
   Dim rs2 As New ADODB.Recordset
   If KeyCode <> vbKeyReturn Then Exit Sub
   If CBoBasedON.ListIndex = 1 Then
        StrSQL = "SELECT NoteSerial1 FROM Transactions where Transaction_Type = 22 and IsNull(order_no,0)  = '" & val(TXT_order_no.Text) & "' and Transaction_ID <> " & val(XPTxtBillID)
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs2.EOF Then
            MsgBox "Â–« «·«„— ·« Ì„ﬂ‰ «œ—«ÃÂ ›ﬁœ «œ—Ã „‰ ﬁ»· ›Ï «·›« Ê—… —ﬁ„" & rs2!NoteSerial1 & ""
            TXT_order_no = ""
         Exit Sub
        End If

   ElseIf CBoBasedON.ListIndex = 3 Then
       
       
        If SystemOptions.MaintOrderCantRepeatBillBuy Then
            
            StrSQL = "SELECT NoteSerial1 FROM Transactions where Transaction_Type = 22 and IsNull(order_no,0)  = '" & val(TXT_order_no.Text) & "' and Transaction_ID <> " & val(XPTxtBillID)
            rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Not rs2.EOF Then
                MsgBox "Â–« «·«„— ·« Ì„ﬂ‰ «œ—«ÃÂ ›ﬁœ «œ—Ã „‰ ﬁ»· ›Ï «·›« Ê—… —ﬁ„" & rs2!NoteSerial1 & ""
                TXT_order_no = ""
                Exit Sub
            End If
        End If

        Dim orderStatus As Integer
     
        MintDone = 0
        Set rs2 = New ADODB.Recordset
        StrSQL = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXT_order_no.Text) & " "
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If rs2.RecordCount > 0 Then
                orderStatus = IIf(IsNull(rs2("OrderStatus").value), 0, rs2("OrderStatus").value)
                TxtCashCustomerName.Text = IIf(IsNull(rs2("ClientName").value), "", rs2("ClientName").value)
                If orderStatus = 2 Or orderStatus = 4 Or orderStatus = 5 Then
                    MintDone = 1
                End If
            Else
                TxtCashCustomerName.Text = ""
                MintDone = -1
            End If
        Exit Sub
    End If
End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then

        If CBoBasedON.ListIndex = 0 Then
            Exit Sub
                
        Else
                
            TXT_order_no.Text = ""
            
           Order_no_search.show
            Order_no_search.RetrunType = 3
            Order_no_search.mTransactionID = CLng(val(XPTxtBillID.Text))
        
           Order_no_search.lblSpecificsearch.Caption = val(CBoBasedON.ListIndex)
        End If

    End If

End Sub

Private Sub txt_total_bill_Change()
    Me.TXTToTAlELSHahn.Text = IIf(Not IsNumeric(Txt_EXport.Text), 0, val(Txt_EXport.Text)) + IIf(Not IsNumeric(txt_total_bill.Text), 0, val(txt_total_bill.Text)) + IIf(Not IsNumeric(TXTFactoryExpenses.Text), 0, val(TXTFactoryExpenses.Text))
Me.TXTToTAlELSHahn.Text = IIf(Not IsNumeric(Txt_EXport.Text), 0, val(Txt_EXport.Text)) + IIf(Not IsNumeric(txt_total_bill.Text), 0, val(txt_total_bill.Text)) + IIf(Not IsNumeric(TXTFactoryExpenses.Text), 0, val(TXTFactoryExpenses.Text) * val(Me.txt_Currency_rate.Text))
Me.TXTToTAlELSHahn.Text = Round(Me.TXTToTAlELSHahn.Text, 2)
End Sub

Private Sub TXTFactoryExpenses_Change()
    Me.TXTToTAlELSHahn.Text = IIf(Not IsNumeric(Txt_EXport.Text), 0, val(Txt_EXport.Text)) + IIf(Not IsNumeric(txt_total_bill.Text), 0, val(txt_total_bill.Text)) + IIf(Not IsNumeric(TXTFactoryExpenses.Text), 0, val(TXTFactoryExpenses.Text))
Me.TXTToTAlELSHahn.Text = IIf(Not IsNumeric(Txt_EXport.Text), 0, val(Txt_EXport.Text)) + IIf(Not IsNumeric(txt_total_bill.Text), 0, val(txt_total_bill.Text)) + IIf(Not IsNumeric(TXTFactoryExpenses.Text), 0, val(TXTFactoryExpenses.Text) * val(Me.txt_Currency_rate.Text))
Me.TXTToTAlELSHahn.Text = Round(Me.TXTToTAlELSHahn.Text, 2)
End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.Text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub TxtLcNo_KeyUp(KeyCode As Integer, _
                          Shift As Integer)
       
    If KeyCode = vbKeyF3 Then
        Order_no_search3.show
        Order_no_search3.RetrunType = 2
         
    End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)

    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 2
        DBCboClientName.BoundText = CUSTID
         On Error Resume Next
         txtItemCodeSearch.SetFocus
    End If

End Sub

Private Sub TxtShortName_KeyDown(KeyCode As Integer, Shift As Integer)
'   LoadSpecificItems
SerchItems (TxtShortName.Text)
DoEvents
DoEvents
DoEvents
DoEvents

        If KeyCode = vbKeyReturn Then
        
        
   DCboItemsName.SetFocus
   DCboItemsName.BoundText = ""
        Sendkeys "{F4}"
        End If
End Sub

Sub SerchItems(Optional str As String)
 
Dim sql As String
Dim SQL1 As String
   
    SerchItemspUBLIC str, sql, SQL1
    fill_combo DCboItemsCode, sql
  fill_combo DCboItemsName, SQL1
        
         
End Sub

Sub SerchItemsxx(Optional str As String)
 
Set DCboItemsCode.RowSource = Nothing
Set DCboItemsName.RowSource = Nothing
If str <> "" Then
Dim sql As String
Dim SQL1 As String
 
Dim StrWhere As String
  Dim astrSplit2tems2() As String
  Dim j As Integer
  Dim nElements As Integer
  Dim SearchString As String
StrWhere = ""
SearchString = ""
sql = " select  ItemID,barCodeNO   from  dbo.TblItems where TblItems.IsArchive=0"
If SystemOptions.UserInterface = ArabicInterface Then
SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If

          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
          If nElements = 0 Then
                               If SystemOptions.UserInterface = ArabicInterface Then
                            StrWhere = " and (ItemName Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%'  or fullcode Like N'%" & Trim(str) & "%') "
                    Else
                            StrWhere = " and (ItemNamee Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%' or fullcode Like N'%" & Trim(str) & "%' ) "
                    End If
                    
          End If
        If nElements > 0 Then
        
     '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        SearchString = ""
        For j = 0 To nElements
        
         SearchString = SearchString & "%" & Trim(astrSplit2tems2(j))
             '     SearchString = "%" & Trim(astrSplit2tems2(j)) & SearchString
                  
        '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
        '   StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         SearchString = SearchString & "%"
                             If SystemOptions.UserInterface = ArabicInterface Then

             StrWhere = StrWhere + " and (ItemName Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             Else
              StrWhere = StrWhere + " and (ItemNamee Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             End If
        '-  StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
      
         End If
        
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemName "
    End If


    SQL1 = SQL1 & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        SQL1 = SQL1 + " Order BY ItemName "
    Else
        SQL1 = SQL1 + " Order BY ItemNamee "
    End If
    
   End If
    fill_combo DCboItemsCode, sql
        fill_combo DCboItemsName, SQL1
        DoEvents
        DoEvents
                  If str = "" Then
                                 sql = " select  ItemID,barCodeNO   from  dbo.TblItems where TblItems.IsArchive=0"
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                 SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
                                     SQL1 = SQL1 + " Order BY ItemName "
                                 Else
                                 SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
                                     SQL1 = SQL1 + " Order BY ItemNameE "
                                 End If
                                 
                                     fill_combo DCboItemsCode, sql
                                         fill_combo DCboItemsName, SQL1
                End If
        
       Exit Sub
       
If str <> "" Then
'Dim Sql As String
'Dim StrWhere As String
'  Dim astrSplit2tems2() As String
'  Dim j As Integer
'  Dim nElements As Integer
StrWhere = ""
If SystemOptions.UserInterface = ArabicInterface Then
sql = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
sql = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If
          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
        If nElements > 0 Then
        StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        For j = 1 To nElements - 1
        
           StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
           StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         End If
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemNamee "
    End If


   End If
   
        fill_combo DCboItemsName, sql
        
End Sub
Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreID
    End If
End Sub

Private Sub TxtValueAdded_Change()
RelinVatGrid
End Sub

Private Sub ChecVAT_Click()
  Dim i As Integer
If Me.TxtModFlg.Text <> "R" Then
    If ChecVAT.value = vbChecked Then

        With Me.VatGrid
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = True
            Next i

        End With

    Else

        With Me.VatGrid

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = False
            Next i

        End With

    End If
    RelinVatGrid
    End If
End Sub

Private Sub VatGrid_Click()
RelinVatGrid
End Sub

Public Sub XPBtnMove_Click(Index As Integer)
invoiceSerach = False
    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2
Me.TxtModFlg.Text = ""

        Dim StrSQL As String
     StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=22 "
     
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
            If SystemOptions.usertype <> UserAdminAll Then
           '     StrSQL = StrSQL & " AND   BranchId=" & Current_branch
            End If


     If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  UserID = " & user_id
             End If
  
        Me.dcBranch.Enabled = True
      
      
    End If
    
            StrSQL = StrSQL & " Order by Transaction_ID"
          If cmdReSave.Visible = True Then
    
    StrSQL = " SELECT * FROM Transactions WHERE Transaction_Type = 22 "
    StrSQL = StrSQL & "   and ( Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & " and "
    StrSQL = StrSQL & "   Transaction_Date <=   " & SQLDate(txtToDateReSave.value, True) & " )"
    
    
   ' StrSQL = StrSQL & "  and NoteSerial1 = '120040023'"
    
    If chkIsBranch(0).value = vbChecked And val(Me.dcBranch.BoundText) > 0 Then
        StrSQL = StrSQL & "  and BranchID =  " & val(Me.dcBranch.BoundText)
    End If
    
    
    If chkIsBranch(2).value = vbChecked And val(Me.DBCboClientName.BoundText) > 0 Then
        StrSQL = StrSQL & "  and CusId =  " & val(Me.DBCboClientName.BoundText)
        StrSQL = StrSQL & "  and isnull(ToTAlELSHahn,0) <> 0"
        ElseIf chkIsBranch(2).value = vbChecked Then
                'StrSQL = StrSQL & "  and isnull(ToTAlELSHahn,0) <> 0"
                StrSQL = StrSQL & "  and  IsNull(BillBasedOn,0) = 0"
                
        
    End If
    
    
    
     If chkIsBranch(1).value = vbChecked Then
        StrSQL = StrSQL & "  and Transaction_ID in "
        StrSQL = StrSQL & "  ( Select Transactions.Transaction_ID from Transactions where Transaction_Type=22 and NoteId not In (SELECT IsNull(notes_id,0) FROM DOUBLE_ENTREY_VOUCHERS where Credit_Or_Debit = 0))"
    End If
    
'    StrSQL = StrSQL & "  and Transaction_ID in (Select Transaction_Details.Transaction_ID  from Transaction_Details where Item_Id = 89)"

End If
'StrSQL = StrSQL & "  and Transaction_ID =14895"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If
   Me.TxtModFlg.Text = "R"

   
        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select
    If IsSaveWithOutMsg Then Exit Sub
     
    Retrive

    Command2_Click
  '  Command4_Click
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            '        Cmd_Click (0)
        Else
            '      SendKeys "{TAB}"
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            'XPBtnRemove_Click
        End If
    End If

  '  If KeyCode = vbKeyF5 Then
  '      If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
  '          XPBtnNewClients_Click
  '      End If
  '  End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
                'XPFillData_Click
            End If
        End If
    End If

    If Shift = 2 Then
        XPTab301.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTab301.CurrTab = 0 Then
                XPTab301.CurrTab = 1

                If XPChkPayType(0).Enabled = True Then
                    XPChkPayType(0).SetFocus
                End If

            Else
                XPTab301.CurrTab = 0
                FG.SetFocus
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    invoiceSerach = True
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.POMustentryAndBillMustEntry = True Then
        Me.TXT_order_no.locked = True
    End If
    
    If mdifrmmain.taxes = True Then
        XPTab301.TabVisible(8) = True
    Else
        XPTab301.TabVisible(8) = False
    End If

    If SystemOptions.AllowEditVaTManulay = True Then
        txtManulaVat.Enabled = True
        txtManulaVat.Visible = True
    Else
        txtManulaVat.Enabled = False
        txtManulaVat.Text = 0
        txtManulaVat.Visible = False
    End If

    ScreenNameArabic = "  ›« Ê—… „‘ —Ì«  "
    ScreenNameEnglish = " Purchase Invoice  "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 150

    With Me.VatGrid

        If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("Typ")) = "#1; ·„ ÌﬁÊ„ «·„Ê—œ »«÷«›… ﬁÌ„…|#2; «·„Ê—œ „⁄›Ï"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .ColComboList(.ColIndex("Typ")) = "#1;Supplier did not add VAT|#2;Supplier is exempt "
        End If

    End With

    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.DcbTyp
            .Clear
            .AddItem "·„ ÌﬁÊ„ «·„Ê—œ »«÷«›… ﬁÌ„…"
            .AddItem "«·„Ê—œ „⁄›Ï"
            .AddItem " «·«Õ ”«» «·⁄ﬂ”Ì "
    
        End With

    Else

        With Me.DcbTyp
            .Clear
            .AddItem "Supplier did not add VAT"
            .AddItem "Supplier is exempt"
        End With

    End If
If SystemOptions.IsHiddenTransportInv Then
        'lbl(66).Caption = "—ﬁ„ ÿ·» «—«„ﬂÊ"
    End If
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL    As String
    Dim Num       As Integer
    Dim StrList   As String
    Dim Dcombos   As ClsDataCombos
    Dim BGround   As New ClsBackGroundPic
    Dim RsNote    As New ADODB.Recordset
    On Error GoTo ErrTrap
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    'fill_combo dcBranch, My_SQL
  
    'dcBranch
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If
    
    SetDtpickerDate XPDtbBill
    Set NewGrid = New ClsGrid
    NewGrid.GridTrans = PurchaseTransaction
    Set NewGrid.TxtValueAdded = TxtValueAdded
    Set NewGrid.VatGrid = Me.VatGrid
    Set NewGrid.Grid = Me.FG
    Set NewGrid.txtManulaVat = Me.txtManulaVat
    Set NewGrid.TxtInvID = Me.Text1
    Set NewGrid.TxtModFlag = Me.TxtModFlg
    Set NewGrid.txtTotal = Me.XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.TxtShortName = Me.TxtShortName
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.TxtItemCodeB1 = TxtItemCodeB1

    Set NewGrid.LBLGross = LBLGross

    
    '-----------------------------------------------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    '-----------------------------------------------------------------------------
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.cmbAccounts = cmbAccounts
    
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.StoreName = DCboStoreName
    Set NewGrid.LblCommision = Me.LblCommision
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
    Set NewGrid.LblTaxSalesValue = Me.lbl(25)
    Set NewGrid.LblTaxAddValue = Me.lbl(32)
    Set NewGrid.LblTaxStampValue = Me.lbl(33)
    Set NewGrid.LblTaxServiceValue = Me.lbl(49)
    Set NewGrid.Customer = Me.DBCboClientName
    
    Set NewGrid.txtItemCodeSearch = txtItemCodeSearch
     NewGrid.frmname = Me.Name
    FG.WallPaper = BGround.Picture

    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date

    With XPCboDiscountType
        .AddItem "·«ÌÊÃœ Œ’„"
        .AddItem "Œ’„ »ﬁÌ„…"
        .AddItem "Œ’„ »‰”»…"
    End With

    With CboPayMentType
        .AddItem "‰ﬁœ«"
        .AddItem "¬Ã·"

        If SystemOptions.AllowPurchasesMultyPayed = True Then
            .AddItem "„ ⁄œœ"
        End If

        .AddItem " ÕÊÌ· »‰ﬂÌ"
      
    End With

    With Me.CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "√„— ‘—«¡"
        .AddItem "›« Ê—… „»œ∆ÌÂ"
        .AddItem "«„— «’·«Õ Ê—‘"
    End With

    NewGrid.FillGrid
    Set Dcombos = New ClsDataCombos
    Dcombos.GetSalesRepDatapurchase Me.DcboEmp
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True
    Dcombos.GetBranches dcBranch
    Dcombos.GetDocTypebyid Me.DCDocTypes, 22, val(Me.dcBranch.BoundText)
    Dcombos.GetBanks Me.DcboBankName
    
    Dcombos.GetAccountingCodes cmbAccounts
    
    StrSQL = "  select  BankID,BankName  from BanksData   "
    fill_combo Dcbanks, StrSQL
 
    StrSQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, StrSQL

    StrSQL = " select id,Project_name from projects"
 
    fill_combo Me.dcproject, StrSQL
      
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    cSearchDcbo(0).SetBuddyText Me.TxtCusID
    
    Dcombos.GetStores Me.DCboStoreName
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName
    ' cSearchDcbo(2).SetBuddyText Me.TxtStoreID
    '-----------------------------------------
    SetDtpickerDate Me.DtpDelayDate
    '≈⁄œ«œ Ã—œ «·√ﬁ”«ÿ
    ChkInstall.value = Unchecked
    ChkInstall.Enabled = False

    With Me.FgInstallments
        .rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCheques
        .rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    Me.XPChkTAX.value = vbUnchecked
    XPChkTAX_Click
    Me.ChkTaxAdd.value = vbUnchecked
    ChkTaxAdd_Click
    Me.ChkTaxStamp.value = vbUnchecked
    ChkTaxStamp_Click
    Me.ChkTaxSerivce.value = vbUnchecked
    ChkTaxSerivce_Click

    If SystemOptions.UserInterface = EnglishInterface Then
     
        SetInterface Me
        ChangeLang
    End If

    '  Resize_Form Me, TransactionSize
  
    With Grid
        .ColComboList(.ColIndex("Items")) = "..."
    End With

    '-----------------------------------------------------------------------------
    Dim rsOut As New ADODB.Recordset
    Dim Msg   As String
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!checkinpo = True Then
            '            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=22"
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=-1"
            StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
            StrSQL = StrSQL & "  Order by Transaction_ID"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
            'Resize_Form Me, TransactionSize
            BillType = 22
    
        Else
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=1"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
            '  Resize_Form Me, TransactionSize
            BillType = 1
            Exit Sub
        End If
    End If

    Me.TxtModFlg.Text = "R"
    '  Command2_Click
    '  Command4_Click

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

ErrTrap:
End Sub

Sub RetriveValueAdded()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.rows = 1
sql = " SELECT     dbo.TransactionValueAdded.Transaction_Type, dbo.TransactionValueAdded.Transaction_ID, dbo.TransactionValueAdded.Vat, dbo.TransactionValueAdded.Vatyo,"
sql = sql & " dbo.TransactionValueAdded.ItemID , dbo.TblItems.itemname, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee ,dbo.TransactionValueAdded.selectd ,dbo.TransactionValueAdded.Typ ,dbo.TransactionValueAdded.Valu "
sql = sql & " FROM         dbo.TransactionValueAdded LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TransactionValueAdded.ItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 22) And (dbo.TransactionValueAdded.Transaction_ID = " & val(XPTxtBillID.Text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.VatGrid
rs2.MoveFirst
.rows = .rows + rs2.RecordCount
For i = 1 To .rows - 1
 .TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
.TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(rs2("Vat").value), "", rs2("Vat").value)
.TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(rs2("Vatyo").value), "", rs2("Vatyo").value)
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("select")) = IIf(IsNull(rs2("selectd").value), 0, rs2("selectd").value)
.TextMatrix(i, .ColIndex("Typ")) = IIf(IsNull(rs2("Typ").value), "", rs2("Typ").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Sub SaveValueAdded()
chkTaxExempt_Click
Cn.Execute "Update tblItems set DefaultSupplier =  " & DBCboClientName.BoundText & " Where ItemId In (SELECT Item_ID FROM Transaction_Details WHERE Transaction_ID  = " & val(val(XPTxtBillID.Text)) & ")"
Dim sss As String
sss = "Update TblItemsUnits set UnitPurPrice =  "
sss = sss & " (SELECT Top 1 Transaction_Details.ShowPrice FROM Transaction_Details WHERE Transaction_ID  = " & val(val(XPTxtBillID.Text)) & " "
sss = sss & " and Transaction_Details.Item_Id =TblItemsUnits.ItemID and TblItemsUnits.UnitId = Transaction_Details.UnitId )"
sss = sss & " Where ItemId In (SELECT Item_ID FROM Transaction_Details WHERE Transaction_ID  = " & val(val(XPTxtBillID.Text)) & ")"
Cn.Execute sss


Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = "Select * from  TransactionValueAdded where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.VatGrid
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
rs2.AddNew
rs2("Transaction_ID").value = val(Me.XPTxtBillID.Text)
rs2("Transaction_Type").value = 22
rs2("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
rs2("Vatyo").value = val(.TextMatrix(i, .ColIndex("Vatyo")))
rs2("Vat").value = val(.TextMatrix(i, .ColIndex("Vat")))
rs2("Valu").value = val(.TextMatrix(i, .ColIndex("Valu")))
rs2("Typ").value = val(.TextMatrix(i, .ColIndex("Typ")))
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
rs2("selectd").value = 1
Else
rs2("selectd").value = 0
End If
rs2.update
End If
Next i
End With
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
CmdAttach.Caption = "Attachments"
lbl(97).Caption = "Smart Search"
lbl(74).Caption = "Add Value"
ChAddToTotal.Caption = "Add To Total"
ISButton2.Caption = "Print GLV"
lbl(95).Caption = "Barcode"
lbl(79).Caption = "VAT"
lbl(80).Caption = "Customs Value"
lbl(81).Caption = "Customs Value"
Label5.Caption = "B/L Date"
'ChSameCurrncy.RightToLeft = False
'ChSameCurrncy.Caption = "Same Currency"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(69).Caption = "I. Manual No"
    lbl(68).Caption = "LC No"
    Label4.Caption = "Doc Type"
    'Label3.Caption = "Shioment No."
    'Label4.Caption = "Order No."
    '''////////
    Me.XPTab301.TabCaption(8) = "VAT"
Label22.Caption = "Data of VAT"
lbl(75).Caption = "VAT"
lbl(76).Caption = "Total"
With VatGrid
.TextMatrix(0, .ColIndex("index")) = "Serial"
.TextMatrix(0, .ColIndex("Code")) = "Item Code"
.TextMatrix(0, .ColIndex("Name")) = "Item Name"
.TextMatrix(0, .ColIndex("Vatyo")) = "Percentage"
.TextMatrix(0, .ColIndex("Vat")) = "Value"
.TextMatrix(0, .ColIndex("select")) = "Select"
.TextMatrix(0, .ColIndex("Valu")) = "Item Value"
.TextMatrix(0, .ColIndex("Typ")) = "Status"
End With
lbl(77).Caption = "Status VAT"
lbl(78).Caption = "Reason"
    ''/////////
    ChecVAT.RightToLeft = False
    ChecVAT.Caption = "Select All"
    Ele(13).Visible = False
    Frame5.Caption = "Notes"
    lbl(72).Caption = "Employee"
    ChkCompsBill.Caption = "Comps Bill"
    lbl(84).Caption = "Tel"
    lbl(73).Caption = "Commision"
    Command4.Caption = "Financial Invoice"
    XPCboDiscountType.Clear
    XPCboDiscountType.AddItem "NO"
    XPCboDiscountType.AddItem "Value"
    XPCboDiscountType.AddItem "Percent"
    CboPayMentType.Clear
    CboPayMentType.AddItem "Cash"
    CboPayMentType.AddItem "Credit"
    
      If SystemOptions.AllowPurchasesMultyPayed = True Then
            CboPayMentType.AddItem "Multi"
      End If
      CboPayMentType.AddItem "Bank transfer"
      
    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(1) = "Securities"
    Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.XPTab301.TabCaption(5) = "Expenses Vouchers"
    Me.XPTab301.TabCaption(4) = "Purchase Orders and Performa Invoices"
    Me.XPTab301.TabCaption(3) = "Fn invoices"
    Me.XPTab301.TabCaption(6) = "Estimated Expenses"
    Me.XPTab301.TabCaption(7) = " Linked voucher"
    '«·ÊﬁÊ› ⁄‰œ «›Ê« Ì— «·„«·Ì… ⁄‘«‰ «··€Â «·«‰Ã·Ì“Ì…
    lbl(57).Caption = "Purcahase order and Performa Invoices"
    lbl(64).Caption = "Financial Invoices"
    Label19.Caption = "Discretionary Expenses"

    lbl(52).Caption = "RCV VCHR No."
    lbl(58).Caption = "Project"
    Label3.Caption = "Branch"
    lbl(65).Caption = " Based On"
    lbl(56).Caption = "O. Arival Date"
    lbl(66).Caption = "NO."
    lbl(63).Caption = "Total Qty "
    lbl(70).Caption = "Cash Supp."
ISButton1.Caption = "View"

    With CBoBasedON
        CBoBasedON.Clear
        CBoBasedON.AddItem "WithOut"
        CBoBasedON.AddItem "Purchase Order"
        CBoBasedON.AddItem "Performa Invoices"
        CBoBasedON.AddItem "Work Order"

    End With

    ' lbl(53).Caption = "Order No:"
    lbl(54).Caption = "Expenses"
    '  lbl(56).Caption = "Payment Voucher"
    '  lbl(57).Caption = "Total Payment"
    lbl(60).Caption = "Total"
 
    lbl(51).Caption = "Total Expenses"
    Command3.Caption = "View P.O. For Vendor"
    Me.Caption = "Purchase Invoice"
    Ele(6).Caption = Me.Caption
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    With FG
        .TextMatrix(0, .ColIndex("NewItem")) = "New ItemID"
 
    End With

    lbl(3).Caption = "Total"
    lbl(50).Caption = "Discount"
    lbl(24).Caption = "Net"
    lbl(1).Caption = "By"
    lbl(0).Caption = "Record#"
    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = "Status"
    lbl(28).Caption = "Serial "
    lbl(27).Caption = "Qty"
    lbl(26).Caption = "Price"
    lbl(4).Caption = "Inventory"
    lbl(6).Caption = "Vendor"
    lbl(7).Caption = "Date"
    lbl(8).Caption = "Bill#"
    lbl(9).Caption = "Currency"

    lbl(5).Caption = "Discount type"
    lbl(11).Caption = "Discount Value"
    lbl(10).Caption = "Pay. Method"
    Command1.Caption = "Convert to Recived VCHR"
    Command2.Caption = "Show Payment VCHR"
    lbl(44).Caption = "Comment"
    XPChkPayType(0).Caption = "Cash"
    lbl(13).Caption = "Value"
    lbl(12).Caption = "ID"
    lbl(2).Caption = "Box"
    lbl(20).Caption = "Currency"
    XPChkPayType(1).Caption = "Credit"
    lbl(15).Caption = "Value"
    lbl(14).Caption = "ID"
    Label1.Caption = "Due Date"
    ChkInstall.Caption = "Installment"
    CmdINSTALLMENT.Caption = "Calc"
    Label2.Caption = "Bank"
    CmdCheque.Caption = "Register"

    With FgInstallments
        .TextMatrix(0, .ColIndex("QestID")) = "ID"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due_Date"
 
    End With

    With FgCheques
 
        .TextMatrix(0, .ColIndex("CheckValue")) = "Value"
        .TextMatrix(0, .ColIndex("CheckNumber")) = "Cheque Number"
        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
        .TextMatrix(0, .ColIndex("ReleaseDate")) = "Release Date"
 
    End With

    With Me.FG
 
        .TextMatrix(0, .ColIndex("order_no")) = "P/O NO."
    End With

    lbl(53).Caption = "Vendor Bill"

    With Me.Grid
 
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("NoteID")) = "NoteID"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "NoteID"

        .TextMatrix(0, .ColIndex("Note_Value")) = "Note_Value"
        .TextMatrix(0, .ColIndex("name")) = "name"

        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
    End With

    With Me.Grid2
 
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Order Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"

    End With

    'With Me.grid
 
    '.TextMatrix(0, .ColIndex("Select")) = "Select"
    '.TextMatrix(0, .ColIndex("NoteID")) = "NoteID"
    '.TextMatrix(0, .ColIndex("NoteSerial")) = "NoteID"
    '
    '.TextMatrix(0, .ColIndex("Note_Value")) = "Note_Value"
    '.TextMatrix(0, .ColIndex("name")) = "Based ON"
    '
 
    'End With

    With Me.Grid4
        '

        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "NoteID"
        .TextMatrix(0, .ColIndex("name")) = "Account Name"

        .TextMatrix(0, .ColIndex("Note_Value")) = "Note_Value"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"

    End With
 
    Cmd(9).Caption = "Delete Row"
    Label18.Caption = "Total"
 
    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Name"
        .TextMatrix(0, .ColIndex("Account_Name2")) = "Account Name"
        .TextMatrix(0, .ColIndex("Price")) = "Value"
        .TextMatrix(0, .ColIndex("value")) = "Total"
        .TextMatrix(0, .ColIndex("ChSameCurrncy")) = "Same Currency"
        .TextMatrix(0, .ColIndex("des")) = "des"
    End With

    lbl(61).Caption = "Bill type"

    BillBasedOn(0).Caption = "Direct Purchase Invoice"
    BillBasedOn(1).Caption = "From Recieve Vouchers"
    BillBasedOn(2).Caption = "Purchase Orders"

    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Name"
        .TextMatrix(0, .ColIndex("value")) = "value"

        .TextMatrix(0, .ColIndex("des")) = "des"
    End With

    With Me.GRID1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("noteserial1")) = "Voucher NO"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "JE Voucher NO"
        .TextMatrix(0, .ColIndex("ManualNO")) = "Manual NO."
    End With

    With Me.GRID1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("noteserial1")) = "Voucher NO"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "JE Voucher NO"
    End With

    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "Order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
    End With

    Frame3.Caption = "JE Voucher NO"
    lbl(62).Caption = "JE Voucher NO"
    Cmd(10).Caption = "Print JE"
 
    lbl(59).Caption = "Total Financial Invoice"
    Command5.Caption = "Save"
    XPChkPayType(2).Caption = "Cheques"

End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·›« Ê—…   " & TxtNoteSerial1.Text & CHR(13) & " —ﬁ„ ›« Ê—… «·„Ê—œ   " & TxtManualNO.Text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & " «·Œ“Ì‰… " & DcboBox.Text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.Text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.Text & CHR(13) & "‰Ê⁄ «·”‰œ " & DCDocTypes & CHR(13) & "»‰«¡ ⁄·Ï " & CBoBasedON & "»—ﬁ„   " & TXT_order_no & CHR(13) & "ÿ—Ìﬁ… «·œ›⁄ " & CboPayMentType & CHR(13) & "‰Ê⁄ «·Œ’„ " & XPCboDiscountType & CHR(13) & "ﬁÌ„… «·Œ’„ " & XPTxtDiscountVal & CHR(13) & "  Ê’Ê· «·‘Õ‰… " & DTArrivalDate & CHR(13) & "  «·«” Õﬁ«ﬁ " & DtpDelayDate & CHR(13) & " «·⁄„·Â " & Dccurrency & CHR(13) & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill No " & TxtNoteSerial1.Text & CHR(13) & "Supplier Bill No " & TxtManualNO.Text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Box " & DcboBox.Text & CHR(13) & " Store  " & DCboStoreName.Text & CHR(13) & " Supplier/Cuxtomer" & DBCboClientName.Text & CHR(13) & "Doc Type" & DCDocTypes & CHR(13) & "Based On" & CBoBasedON & "No :   " & TXT_order_no & CHR(13) & "Payment Type" & CboPayMentType & CHR(13) & "Discount Type  " & XPCboDiscountType & CHR(13) & " Discount Vaalue   " & XPTxtDiscountVal & CHR(13) & " Shipment Arival Date" & DTArrivalDate & CHR(13) & "Due Date " & DtpDelayDate & CHR(13) & " Currency " & Dccurrency & CHR(13) & " GE NO" & TxtNoteSerial
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 150, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 150, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 150

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set BuyReport = Nothing
    Set m_MnuShowNewItemsPrices = Nothing

    If Not m_FrmSearch Is Nothing Then
        Unload m_FrmSearch
        Set m_FrmSearch = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Purcahase Invoice"
        
            Else
                Me.Caption = "›« Ê—… ‘—«¡"
            End If
    
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
     
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPBtnNewClients.Enabled = False
        
            XPCboDiscountType.locked = True
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            Me.XPTxtDiscountVal.locked = True
        
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
        
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
            FG.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
         '       Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            
            End If
        
            CboPayMentType.locked = True
            DtpDelayDate.Enabled = False
            Ele(4).Enabled = False
        
            XPChkTAX.Enabled = False
            ChkTaxAdd.Enabled = False
            ChkTaxSerivce.Enabled = False
            ChkTaxStamp.Enabled = False
        
        Case "N"
            '   Me.Caption = "›« Ê—… ‘—«¡( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            FG.rows = 2
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            FG.Editable = flexEDKbdMouse
            XPDtbBill.value = Date
            '        XPFillData.Enabled = True
            XPCboDiscountType.ListIndex = 0
            CboPayMentType.ListIndex = 0
            CboPayMentType.locked = False
            DtpDelayDate.Enabled = True
            DtpDelayDate.value = Date
            Ele(4).Enabled = True
        
            CboItemCase.ListIndex = 0
        
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True

        Case "E"
            '   Me.Caption = "›« Ê—… ‘—«¡(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
        
            FG.Enabled = True
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            DtpDelayDate.Enabled = True
        
            If XPChkPayType(0).value = Checked Then
                XPChkPayType_Click (0)
            End If

            If XPChkPayType(1).value = Checked Then
                XPChkPayType_Click (1)
            End If

            If XPChkPayType(2).value = Checked Then
                XPChkPayType_Click (2)
            End If

            If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 2 Then
                CboPayMentType_Change
            End If

            FG.Editable = flexEDKbdMouse
        
            CboPayMentType.locked = False
            DBCboClientName_Change
            Ele(4).Enabled = True
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
    End Select

    Exit Sub
ErrTrap:
    Stop
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsTest As ADODB.Recordset
    Dim Num As Long
    Dim Msg As String
    Dim i As Integer
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset

  '    On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting
 IsClicKCommand4 = False

    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""
    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.Text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.Text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.Text = ""
    ChkTaxSerivce.value = vbUnchecked
    Me.TxtTaxServiceValue.Text = ""

    '---------------------------------------------
    '---------------------------------------------
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
      Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.Text = "T"
    Screen.MousePointer = vbArrowHourglass
    BLDate.value = IIf(IsNull(rs("BLDate").value), Date, rs("BLDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    Me.dcproject.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(67).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.Text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    If Not (IsNull(rs("CashCustomerPhone").value)) Then
        Me.TxtPhone.Text = rs("CashCustomerPhone").value
    Else
        Me.TxtPhone.Text = ""
    End If
    TxtVATCustoms.Text = IIf(IsNull(rs("VATCustoms").value), 0, (rs("VATCustoms").value))
    TxtVATCustoms1.Text = IIf(IsNull(rs("VATCustoms1").value), 0, (rs("VATCustoms1").value))
 TxtValueAdded.Text = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))
 LblValueAdded.Caption = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))
 Me.DcbTyp.ListIndex = IIf(IsNull(rs("Typ").value), -1, (rs("Typ").value))
 TXtResonVAT.Text = IIf(IsNull(rs("ResonVAT").value), "", (rs("ResonVAT").value))
 TxtVATNO.Text = IIf(IsNull(rs("VATNO").value), "", (rs("VATNO").value))
 poTransaction_ID.Text = IIf(IsNull(rs("poTransaction_ID").value), "", (rs("poTransaction_ID").value))
     txtContainerNo = IIf(IsNull(rs("ContainerNo").value), "", rs("ContainerNo").value)
    If Not (IsNull(rs("CompsBill").value)) Then
         
                If (rs("CompsBill").value) = 0 Then
                          ChkCompsBill.value = vbUnchecked
                Else
                          ChkCompsBill.value = vbChecked
                End If
         
    Else
      ChkCompsBill.value = vbUnchecked
    End If
    
    
    
    If Not (IsNull(rs("VstReverse").value)) Then
         
                If (rs("VstReverse").value) = 0 Then
                          VstReverse.value = vbUnchecked
                Else
                          VstReverse.value = vbChecked
                End If
         
    Else
      VstReverse.value = vbUnchecked
    End If
    
    
 If IsNull(rs("chkTaxExempt").value) Then
        Me.chkTaxExempt.value = vbUnchecked
    Else
        Me.chkTaxExempt.value = IIf(rs("chkTaxExempt").value = 0, vbUnchecked, vbChecked)
    End If
    
    
    
  '   If Not (IsNull(rs("ChSameCurrncy").value)) Then
  '
  '              If (rs("ChSameCurrncy").value) = 0 Then
  '                        ChSameCurrncy.value = vbUnchecked
  ''              Else
   '                       ChSameCurrncy.value = vbChecked
   '             End If
   '
   ' Else
   '   ChSameCurrncy.value = vbUnchecked
   ' End If
    
    
    TXT_order_no.Text = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))
    TxtManualNO.Text = IIf(IsNull(rs("ManualNO").value), "", (rs("ManualNO").value))

    TxtManualNo1.Text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))
txtManulaVat.Text = IIf(IsNull(rs("txtManulaVat").value), 0, (rs("txtManulaVat").value))
txtManulaVat.Text = val(txtManulaVat.Text)
 
 
    txt_Currency_rate.Text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    DtpDelayDate.value = IIf(IsNull(rs("DueDate").value), Date, (rs("DueDate").value))
    XPTxtBillID.Text = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
    TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    DTArrivalDate.value = IIf(IsNull(rs("ArrivalDate").value), Date, (rs("ArrivalDate").value))
''//
txtAddValue.Caption = IIf(IsNull(rs("AddValue").value), 0, (rs("AddValue").value))
If Not IsNull(rs("AddToTotal").value) Then
If rs("AddToTotal").value = 1 Then
ChAddToTotal.value = vbChecked
Else
ChAddToTotal.value = vbUnchecked
End If
Else
ChAddToTotal.value = vbUnchecked
End If
DTArrivalDate.value = IIf(IsNull(rs("ArrivalDate").value), Date, (rs("ArrivalDate").value))

    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), 0, rs("Trans_DiscountType").value)

    XPTxtDiscountVal.Text = IIf(IsNull(rs("Trans_Discount").value), "", Trim(rs("Trans_Discount").value))
    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.Text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.Text = ""
    End If

    TXTToTAlELSHahn.Text = IIf(Not IsNumeric(rs("ToTAlELSHahn").value), 0, rs("ToTAlELSHahn").value)
    Txt_EXport.Text = IIf(Not IsNumeric(rs("total_expenses").value), 0, rs("total_expenses").value)

If IsSaveWithOutMsg Then
    TXTToTAlELSHahn.Text = 0
End If
If val(TXTToTAlELSHahn.Text) <> 0 Then
    txt_total_bill.Text = 0
End If
    txt_total_bill.Text = 0
    Command4_Click
    'txt_total_bill.text = IIf(Not IsNumeric(rs("total_payments").value), 0, rs("total_payments").value)
    If val(txt_total_bill) <> 0 Then
        
    End If
    
    
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)

    TxtBillComment.Text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
    
    If IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value) = 3 Then
        If SystemOptions.AllowPurchasesMultyPayed Then
            CboPayMentType.ListIndex = 3
        Else
         CboPayMentType.ListIndex = 2
        End If
    Else
        CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    End If
   
   
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    '
    Me.DcboBankName.BoundText = IIf(IsNull(rs("BankID").value), "", rs("BankID").value)
    
    Text1.Text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
    'Text1.text = IIf(IsNull(Rs("nots").Value), "", (Rs("nots").Value))

    'txt_Shipment_no.text = IIf(IsNull(Rs("Shipment_no").value), "", Trim(Rs("Shipment_no").value))
    'Txt_order_no.text = IIf(IsNull(Rs("order_no").value), "", Trim(Rs("order_no").value))
    Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)

    TxtLcNo.Text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))

    '÷—»Ì… «·„»Ì⁄« 
    XPTxtTaxValue.Text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(rs("TaxAddValue").value) Then
        If rs("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.Text = rs("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(rs("TaxStampValue").value) Then
        If rs("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.Text = rs("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(rs("TaxServiceValue").value) Then
        If rs("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.Text = rs("TaxServiceValue").value
        End If
    End If

    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    XPTxtSum.Text = ""
   CBoBasedON.ListIndex = IIf(IsNull(rs("CBoBasedON").value), 0, (rs("CBoBasedON").value))

    If Not IsNull(rs("BillBasedOn").value) Then

        If rs("BillBasedOn").value = 0 Then
            BillBasedOn(0).value = True
            BillBasedOn_Click (0)
        ElseIf rs("BillBasedOn").value = 1 Then
            BillBasedOn(1).value = True
            BillBasedOn_Click (1)
        ElseIf rs("BillBasedOn").value = 2 Then
            BillBasedOn(2).value = True
            BillBasedOn_Click (2)
        End If
    
    Else

        BillBasedOn(0).value = True
        BillBasedOn_Click (0)
    End If

    StrSQL = "SELECT TblItems.HaveSerial, "
    If SystemOptions.IsGeometricProportions Then
        StrSQL = StrSQL & "  TblItemsUnits.ForUnit , TblItemsUnits.MethodCalc,TblItemsUnits.PartItemQty,"
    End If
    
    StrSQL = StrSQL & "  *,ACCOUNTS.Account_Serial,ACCOUNTS.Account_Name,projects.Project_name,projects.Project_nameE FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID  "
    StrSQL = StrSQL & "  INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "  Left outer JOIN dbo.projects ON dbo.Transaction_Details.ProjectID = dbo.projects.Id"
    If SystemOptions.IsGeometricProportions Then
    
        StrSQL = StrSQL & "  LEFT OUTER JOIN TblItemsUnits"
        StrSQL = StrSQL & "              ON  dbo.TblItemsUnits.Unitid = dbo.TblUnites.UnitID"
        StrSQL = StrSQL & "              and dbo.TblItemsUnits.ItemID = TblItems.ItemID"
        
       
    End If
StrSQL = StrSQL & "  LEFT OUTER JOIN ACCOUNTS "
StrSQL = StrSQL & "              ON  dbo.ACCOUNTS.Account_Code = dbo.Transaction_Details.Account_Code"
    StrSQL = StrSQL + "  where Transaction_ID=" & val(rs("Transaction_ID").value) & " order by Transaction_Details.id "
    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim mTaxExemptTotal As Double
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("TypeVAT")) = IIf(IsNull(RsDetails("TypeVAT")), "", (RsDetails("TypeVAT").value))
            FG.TextMatrix(Num, FG.ColIndex("LineShahn")) = IIf(IsNull(RsDetails("LineShahn")), 0, (RsDetails("LineShahn").value))
            If Me.FG.ColIndex("SalesValue") <> -1 Then
                FG.TextMatrix(Num, FG.ColIndex("SalesValue")) = IIf(IsNull(RsDetails("SalesValue")), 0, (RsDetails("SalesValue").value))
            End If
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            
            FG.TextMatrix(Num, FG.ColIndex("Account_Name")) = IIf(IsNull(RsDetails("Account_Name")), "", Trim(RsDetails("Account_Name").value))
            FG.TextMatrix(Num, FG.ColIndex("Account_Code")) = IIf(IsNull(RsDetails("Account_Code")), "", Trim(RsDetails("Account_Code").value))
            FG.TextMatrix(Num, FG.ColIndex("Account_Serial2")) = IIf(IsNull(RsDetails("Account_Serial")), "", Trim(RsDetails("Account_Serial").value))
            FG.TextMatrix(Num, FG.ColIndex("projectid")) = IIf(IsNull(RsDetails("ProjectID")), "", Trim(RsDetails("ProjectID").value))
            
            
            FG.TextMatrix(Num, FG.ColIndex("project")) = IIf(IsNull(RsDetails("Project_name")), "", Trim(RsDetails("Project_name").value))
            
            'FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").Value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))


            If SystemOptions.IsGeometricProportions Then
            
                FG.TextMatrix(Num, FG.ColIndex("ForUnit")) = IIf(IsNull(RsDetails("ForUnit")), "", (RsDetails("ForUnit").value))
                FG.TextMatrix(Num, FG.ColIndex("MethodCalc")) = IIf(IsNull(RsDetails("MethodCalc")), "", (RsDetails("MethodCalc").value))
                FG.TextMatrix(Num, FG.ColIndex("ItemQty")) = IIf(IsNull(RsDetails("PartItemQty")), "", (RsDetails("PartItemQty").value))

            End If
            
             FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
            FG.TextMatrix(Num, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))
            FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
            FG.TextMatrix(Num, FG.ColIndex("OUTR")) = IIf(IsNull(RsDetails("OUTR")), "", (RsDetails("OUTR").value))
            FG.TextMatrix(Num, FG.ColIndex("INR")) = IIf(IsNull(RsDetails("INR")), "", (RsDetails("INR").value))
            FG.TextMatrix(Num, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
            
            If SystemOptions.IsGeometricProportions = True Then


                FG.TextMatrix(Num, FG.ColIndex("UnitPrice")) = IIf(IsNull(RsDetails("UnitPrice")), "", (RsDetails("UnitPrice").value))
            End If

        

        


            FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If

            ' FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").Value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If
             FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
             FG.TextMatrix(Num, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", (RsDetails("Vatyo").value))
             
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 0, (RsDetails("ItemSize").value))
           FG.TextMatrix(Num, FG.ColIndex("ItemsDetailsNewidea")) = IIf(IsNull(RsDetails("ItemsDetailsNewidea")), "", (RsDetails("ItemsDetailsNewidea").value))
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
            
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
                  FG.TextMatrix(Num, FG.ColIndex("GoldDetails")) = IIf(IsNull(RsDetails("GoldDetails")), "", (RsDetails("GoldDetails").value))
        FG.TextMatrix(Num, FG.ColIndex("Wages")) = IIf(IsNull(RsDetails("Wages")), "", (RsDetails("Wages").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            'FG.TextMatrix(Num, FG.ColIndex("OpeningBurcahseQty")) = IIf(IsNull(RsDetails("OpeningBurcahseQty").value), "", RsDetails("OpeningBurcahseQty").value)
            '      FG.TextMatrix(Num, FG.ColIndex("OpeningBurcahseValue")) = IIf(IsNull(RsDetails("OpeningBurcahseValue").value), "", RsDetails("OpeningBurcahseValue").value)
            '       FG.TextMatrix(Num, FG.ColIndex("OpeningSalesQty")) = IIf(IsNull(RsDetails("OpeningSalesQty").value), "", RsDetails("OpeningSalesQty").value)
            '        FG.TextMatrix(Num, FG.ColIndex("OpeningSalesValue")) = IIf(IsNull(RsDetails("OpeningSalesValue").value), "", RsDetails("OpeningSalesValue").value)
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", RsDetails("FoxyNo").value)
            FG.TextMatrix(Num, FG.ColIndex("StoreID2")) = IIf(IsNull(RsDetails("StoreID2")), "", (RsDetails("StoreID2").value))
            FG.TextMatrix(Num, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
        
                   '***********************************************************
        FG.TextMatrix(Num, FG.ColIndex("SBillNO")) = IIf(IsNull(RsDetails("SBillNO").value), "", RsDetails("SBillNO").value)
            FG.TextMatrix(Num, FG.ColIndex("ExtraType")) = IIf(IsNull(RsDetails("ExtraType")), "", (RsDetails("ExtraType").value))
            FG.TextMatrix(Num, FG.ColIndex("ExtraVal")) = IIf(IsNull(RsDetails("ExtraVal")), "", (RsDetails("ExtraVal").value))
        
        'FG.Cell(flexcpData, Num, FG.ColIndex("SupplierID")) = IIf(IsNull(RsDetails("SupplierID")), "", (RsDetails("SupplierID").value))
            FG.TextMatrix(Num, FG.ColIndex("SupplierID")) = IIf(IsNull(RsDetails("SupplierID")), "", (RsDetails("SupplierID").value))
            
               FG.TextMatrix(Num, FG.ColIndex("ScurrencyID")) = IIf(IsNull(RsDetails("ScurrencyID")), 1, (RsDetails("ScurrencyID").value))
            FG.TextMatrix(Num, FG.ColIndex("Scurrenyrate")) = IIf(IsNull(RsDetails("Scurrenyrate")), 1, (RsDetails("Scurrenyrate").value))
      FG.TextMatrix(Num, FG.ColIndex("Scurrenyrate")) = IIf(IsNull(RsDetails("Scurrenyrate")), 1, (RsDetails("Scurrenyrate").value))
         '***********************************************************


                   'Wael
            If FG.ValueMatrix(i, FG.ColIndex("chkTaxExempt")) = True Then
                mTaxExemptTotal = mTaxExemptTotal + ((val(FG.TextMatrix(i, FG.ColIndex("Count"))) * val(FG.TextMatrix(i, FG.ColIndex("Price"))))) - val(FG.TextMatrix(i, FG.ColIndex("DiscountValue")))
            End If

            RsDetails.MoveNext

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

        '  FG.AutoSize 0, FG.Cols - 1, False
    End If
   'wael
    LblValueAdded.Tag = mTaxExemptTotal

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).Text = ""
    XPTxtValue(1).Text = ""

    XPTxtSerial(0).Text = ""
    XPTxtSerial(1).Text = ""
   ' DtpDelayDate.value = Date
    StrSQL = "select * From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                'Me.TxtNoteID(0).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(0).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).Text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                'Me.TxtNoteID(1).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(1).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).Text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 13 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
            End If
        
            RsNotes.MoveNext
        Next Num

    End If

    Set RsNotes = New ADODB.Recordset
    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
    StrSQL = StrSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
    StrSQL = StrSQL + " Where NoteType=13 AND NOTES.Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + " Order BY Notes.NoteID"
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgCheques
        .rows = .FixedRows

        If Not (RsNotes.BOF Or RsNotes.EOF) Then
            .rows = .FixedRows + RsNotes.RecordCount

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("CheckValue")) = IIf(IsNull(RsNotes("Note_Value").value), "", RsNotes("Note_Value").value)
                .TextMatrix(i, .ColIndex("CheckNumber")) = IIf(IsNull(RsNotes("ChqueNum").value), "", RsNotes("ChqueNum").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsNotes("BankName").value), "", RsNotes("BankName").value)

                If Not IsNull(RsNotes("DueDate").value) Then
                    .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsNotes("DueDate").value)
                Else
                    .TextMatrix(i, .ColIndex("DueDate")) = ""
                End If

                RsNotes.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
        SumChecks
    End With

    '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
    If XPTxtValue(1).Tag <> "" Then
        StrSQL = "Select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
        Set RsTest = New ADODB.Recordset
        RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            CmdINSTALLMENT.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
            Else
                CmdINSTALLMENT.Caption = "View"
            End If

            LngPartID = RsTest("PartID").value
            Me.LblPrecenType.Tag = RsTest("InterestType").value

            If RsTest("InterestType").value = 0 Then
                LblPrecenType.Caption = "‰”»… „∆ÊÌ…"
            ElseIf RsTest("InterestType").value = 1 Then
                LblPrecenType.Caption = "ﬁÌ„… À«» …"
            ElseIf RsTest("InterestType").value = 2 Then
                LblPrecenType.Caption = "·«ÌÊÃœ"
            End If

            Me.LblPrecenValue.Caption = RsTest("InterestVal").value
            Me.LblInstallTotal.Caption = RsTest("Total").value
            Me.LblInstallCount.Caption = RsTest("InstallCount").value
            Me.LblFirstInstallDate.Caption = DisplayDate(RsTest("FirstInstallDate").value)
            Me.LblInstallmentType.Tag = RsTest("InstallmentType").value

            If RsTest("InstallmentType").value = 0 Then
                LblInstallmentType.Caption = "ÌÊ„"
            ElseIf RsTest("InstallmentType").value = 1 Then
                LblInstallmentType.Caption = "‘Â—"
            ElseIf RsTest("InstallmentType").value = 2 Then
                LblInstallmentType.Caption = "”‰…"
            End If

            Me.LblInstallSeprator.Caption = RsTest("InstallSeprator").value
            Me.LblStartValue.Caption = IIf(IsNull(RsTest("StartValue").value), "", RsTest("StartValue").value)
            Set RsPartDetails = New ADODB.Recordset
            StrSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
            RsPartDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
                RsPartDetails.MoveFirst

                With Me.FgInstallments
                    .rows = .FixedRows + RsPartDetails.RecordCount

                    For i = .FixedRows To .rows - 1
                        .TextMatrix(i, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").value), "", RsPartDetails("QestID").value)
                        .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsPartDetails("Value").value), "", RsPartDetails("Value").value)

                        If Not IsNull(RsPartDetails("DueDate").value) Then
                            .TextMatrix(i, .ColIndex("Due_Date")) = DisplayDate(RsPartDetails("DueDate").value)
                        Else
                            .TextMatrix(i, .ColIndex("Due_Date")) = ""
                        End If
 
                        RsPartDetails.MoveNext
                    Next i

                End With

            End If

        Else
            CmdINSTALLMENT.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
            Else
                CmdINSTALLMENT.Caption = "calc"
            End If
        End If
    End If

    NewGrid.Calculate 1, , , True
    Dim SngRelatedNotesValues As Single
    Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.Text), 0, SngRelatedNotesValues)
    Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.Text), 0, SngRelatedNotesValues)
    Me.CmdRetruns.Tag = SngRelatedNotesValues
    
    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    TxtFillData.Text = "F"
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    fill_bill_items_table
    Command3_Click
    
    '«” —Ã«⁄ «·„’—Ê›«  «· ﬁœÌ—ÌÂ
    fillExpensesFactoryGrid
     If Me.TxtModFlg.Text = "R" Or Me.TxtModFlg.Text = "" Then
          FillGridWithDataSalesPayment
          RetriveValueAdded
          
          End If
    TXT_order_no.Text = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))
    'Command4_Click
     'Command2_Click
    ChAddToTotal_Click
    RelinVatGrid
    ChAddToTotal_Click
    '  FillVoucherGrid
 mIsFinishSave = True
 chkTaxExempt_Click
    Exit Sub
ErrTrap:
    Msg = "Œÿ« ›Ï ≈” —Ã«⁄ «·»Ì«‰« ..!!!"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault

End Sub
Private Sub CmdNos_Click(Index As Integer)
  If Index <= 9 Then
LBLPayVal.Caption = LBLPayVal.Caption & Index

ElseIf Index = 10 Then
LBLPayVal.Caption = LBLPayVal.Caption & "00"

ElseIf Index = 11 Then
LBLPayVal.Caption = LBLPayVal.Caption & "."

ElseIf Index = 12 Then 'ar
If Len(LBLPayVal.Caption) > 1 Then
LBLPayVal.Caption = mId(LBLPayVal.Caption, 1, Len(LBLPayVal.Caption) - 1)
Else
LBLPayVal.Caption = ""
End If
ElseIf Index = 13 Then 'ar
 LBLPayVal.Caption = ""

TxtPayedValue2.Text = ""
cleargrid

ElseIf Index = 14 Then
TxtPayedValue2.Text = val(LBLPayVal)

 
        With Grid22
          .TextMatrix(.Row, .ColIndex("Value")) = TxtPayedValue2.Text
          End With
    ReLineGrid2
     
 TxtRemainValue2.Text = val(Me.TxtPayedValue2.Text) - val(Me.TxtNetValue2.Text)
 

End If

 ReLineGrid2
 
End Sub
Private Sub cleargrid()
    On Error Resume Next
    Dim i As Integer
 
  With Grid22

       ' For I = .FixedRows To .Rows - 1

         .TextMatrix(.Row, .ColIndex("value")) = 0
          
       ' Next I

    End With
     TxtPayedValue2 = 0
    
End Sub
Private Sub CmdValue_Click(Index As Integer)
LBLPayVal.Caption = 0
'TxtPayedValue.text = CmdValue(Index).Caption
LBLPayVal.Caption = CmdValue(Index).Caption
        With Grid22
          .TextMatrix(.Row, .ColIndex("Value")) = LBLPayVal.Caption
          End With
     ReLineGrid2
End Sub
Sub SaveSalesPayment(Optional TransID As Double)
Dim Rs3 As ADODB.Recordset
Dim i As Integer
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = "select * from TblSalesPayment where 1=-1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Grid22
For i = 1 To .rows - 1
If (.TextMatrix(i, .ColIndex("PaymentName"))) <> "" Then
Rs3.AddNew
Rs3("TransID").value = TransID
Rs3("PaymentID").value = val(.TextMatrix(i, .ColIndex("PaymentID")))
Rs3("Value").value = val(.TextMatrix(i, .ColIndex("Value")))
Rs3("CardNo").value = (.TextMatrix(i, .ColIndex("CardNo")))
Rs3("maxvalue").value = val((.TextMatrix(i, .ColIndex("MaxValue"))))

Rs3.update
End If
Next i
End With
End Sub
Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
                Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
                Msg = "Confirm Undo"
            End If

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
        
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
                Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
                Msg = "Confirm Undo"
            End If
  
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub
Function ChekPaymet() As Boolean

   If SystemOptions.CanChanegeLinkedPurcahsenvoice = True Then
     ChekPaymet = False
    Exit Function
     End If

Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
ChekPaymet = False
sql = "select * from  TblNotesBillBuyPayment where NoteID=" & val(Me.XPTxtBillID.Text) & " "
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
ChekPaymet = True
Else
ChekPaymet = False
End If
End Function
Private Sub Del_TransAction()
    Dim Msg As String
    Dim StrSQL As String
    Dim BegainTrans As Boolean
    Dim order_no As String
    order_no = Me.TXT_order_no
    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
    If ChekPaymet() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "·«Ì„ﬂ‰ «·”„«Õ »Õ–› Â–Â «·⁄„·Ì…"
Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ   "
Else
Msg = "Can not be allowed to delete this process"
Msg = Msg & CHR(13) & "There repayment process   "
End If
MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
Exit Sub
End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·›« Ê—…  —ﬁ„ " & CHR(13)
        Msg = Msg + TxtNoteSerial1.Text & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
        Msg = "Confirm Delete Invoice No: " & CHR(13)
        Msg = Msg + TxtNoteSerial1.Text & CHR(13)
        Msg = Msg + " yes/no?"


End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If AvailableDeal = True Or AvailableDeal = False Then
                If Not rs.RecordCount < 1 Then
                    Cn.BeginTrans
                    BegainTrans = True
                            DeleteTransactiomsVoucher val(Text1.Text)
                            
                deletelinktoVoucher
                Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.Text) & ""
                 StrSQL = "delete From TblSalesPayment where TransID=" & val(Me.XPTxtBillID.Text) 'Val(rs("Transaction_ID").value)
               Cn.Execute StrSQL, , adExecuteNoRecords
                    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
                    Cn.Execute StrSQL, , adExecuteNoRecords
                ''//////////////
                  Cn.Execute "delete from Transaction_Details where Transaction_ID in ( " & GetTrnasectionID(val(XPTxtBillID.Text), 22) & " )"
                  Cn.Execute "delete from Transactions where Transaction_ID in ( " & GetTrnasectionID(val(XPTxtBillID.Text), 22) & " )"
                  Cn.Execute "delete from DOUBLE_ENTREY_VOUCHERS where Transaction_ID in ( " & GetTrnasectionID(val(XPTxtBillID.Text), 22) & " )"
                  Cn.Execute "Delete From TblTransctionIDES where Transaction_Type=22 and MainTransaction_ID=" & val(XPTxtBillID.Text) & " "
                ''////////
                    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & get_transaction_id(rs("nots").value, 20)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
                    StrSQL = "Delete From Transactions  " & "Where Transaction_ID=" & get_transaction_id(rs("nots").value, 20)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
                    StrSQL = "update Notes set  Transaction_ID1=Null , ItemID=NUll, buy = null Where   (Transaction_ID1=" & val(Me.XPTxtBillID.Text) & ")"
                    Cn.Execute StrSQL
            
                    StrSQL = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=Null ,  ItemID=NUll, buy = null Where  ( Transaction_ID1=" & val(Me.XPTxtBillID.Text) & ")"
                    Cn.Execute StrSQL
            
                    StrSQL = "delete From Notes where  NoteType= 150 and  noteid=" & val(TXTNoteID.Text)
    
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
            
                    CuurentLogdata ("D")
                    rs.delete
                    Cn.CommitTrans
                    BegainTrans = False
                    rs.MoveFirst

        
                    close_order2 order_no
                                             With Me.Grid4
            .rows = .FixedRows
   
        End With
                    If rs.RecordCount < 1 Then
                        clear_all Me
                        TxtModFlg_Change
                        XPTxtCurrent.Caption = 0
                        XPTxtCount.Caption = 0
                         VatGrid.Clear flexClearScrollable, flexClearEverything
                         VatGrid.rows = 1
                    Else
                        Retrive
                    End If
                End If
            End If
        End If

    Else
        clear_all Me
         VatGrid.Clear flexClearScrollable, flexClearEverything
           VatGrid.rows = 1
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Set TTP = New clstooltip
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ‘—«¡ ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «·‘—«¡" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·‘—«¡ «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·‘—«¡" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ⁄„·Ì«  «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… ‘—«¡" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ‘—«¡" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— Ctrl + X", True
    End With

    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnAdd, _
    '    "≈÷«›… «·√’‰«› ..." & Wrap & _
    '    " ·«÷«›… ’‰› ÃœÌœ" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— F2", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnRemove, _
    '    "Õ–› ’‰› ..." & Wrap & _
    '    "·Õ–› √Õœ «·√’‰«›" & Wrap & _
    '    " ÕœœÂ Ê«÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— F3", True
    'End With
    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPFillData, _
    '    " ⁄»∆… »Ì«‰«  «·√’‰«›" & Wrap & _
    '    "· ⁄»∆… »Ì«‰«  «·√’‰«› ›Ì" & Wrap & _
    '    "›Ì ‰«›–… ÕÊ«—" & Wrap & _
    '    "  ≈÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— Ctrl + Space", True
    'End With
    With TTP
        .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Function Closeorders()
    On Error Resume Next
    Dim i As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset

    Dim sql As String
 
    Dim differnt As Integer
    Dim order_qty As Integer
    Dim QTYRecived As Integer
    Dim close_order As Boolean

    Dim j As Integer

    With Grid2

        For i = 1 To Grid2.rows - 1
            close_order = True

            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "select * from QRY_items_orders_data where order_no='" & Grid2.TextMatrix(i, Grid2.ColIndex("order_no")) & "'"
                Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Rs3.RecordCount = 0 Then GoTo ll

                For j = 1 To Rs3.RecordCount
                    order_qty = IIf(IsNull(Rs3("Quantity").value), 0, Rs3("Quantity").value)
                    QTYRecived = IIf(IsNull(Rs3("QTYRSV").value), 0, Rs3("QTYRSV").value)
                    differnt = order_qty - QTYRecived

                    If differnt <= 0 Then
                        close_order = False
                    End If
                
                    Rs3.MoveNext
                Next j
           
                If close_order = True Then
                    sql = "select * from Transactions where Transaction_Type=6 and order_no='" & Grid2.TextMatrix(i, Grid2.ColIndex("order_no")) & "'"
                    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    Rs4("Closed").value = 1
                    Rs4.update
                    Rs4.Close
                   
                End If
               
            End If
       
            Rs3.Close
ll:
        Next
       
    End With
       
End Function

Function SaveNewGl()
If ChkCompsBill.value = vbUnchecked Then Exit Function
 Dim supplierGL As New ADODB.Recordset
  Dim StrSQL As String
  Dim RowNum As Integer
  Dim Account_Code_dynamic As String
  Dim StrTempDes  As String
  Dim CommissionAccount As String
 Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
   
   
   Dim SupplierAccount  As String
     Account_Code_dynamic = get_account_code_branch(4, my_branch)
    CommissionAccount = get_account_code_branch(96, my_branch)
    SupplierAccount = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
    LngDevNO = 2
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            End If

   
    Dim SngTemp  As Variant
        Dim SngTempe  As Variant
 

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    

  StrSQL = " SELECT     dbo.TblCustemers.CusName, dbo.TblCustemers.Account_Code, SUM(( dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty)-Transaction_Details.discountvalue) AS totals"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & " dbo.TblCustemers ON dbo.Transaction_Details.SupplierID = dbo.TblCustemers.CusID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_ID = " & val(XPTxtBillID.Text) & ")"
StrSQL = StrSQL & " GROUP BY dbo.TblCustemers.CusName, dbo.TblCustemers.Account_Code"
 
     supplierGL.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
Dim value As Variant
Dim Account_code As String
Dim CusName As String
Dim TotalValue As Variant
TotalValue = 0
    For RowNum = 1 To supplierGL.RecordCount

    value = IIf(IsNull(supplierGL("totals").value), 0, supplierGL("totals").value)
    Account_code = IIf(IsNull(supplierGL("Account_Code").value), "", supplierGL("Account_Code").value)
    CusName = IIf(IsNull(supplierGL("CusName").value), "", supplierGL("CusName").value)
    
    If value > 0 And Account_code <> "" Then
  '  value = Round(value, SystemOptions.SysDefCurrencyForamt)
    SngTemp = Round(value * val(txt_Currency_rate.Text), SystemOptions.SysDefCurrencyForamt)
   
     SngTempe = value
     LngDevNO = LngDevNO + 1
     TotalValue = TotalValue + SngTempe
              If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text & " «À»«  „‘ —Ì«  «·„Ê—œ " & CusName
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            End If
            
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
    End If
         supplierGL.MoveNext
    Next RowNum
   
     LngDevNO = LngDevNO + 1
     
    SngTemp = LblCommision * val(txt_Currency_rate.Text)
    SngTempe = LblCommision
    SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
   If SngTemp > 0 Then


            If ModAccounts.AddNewDev(LngDevID, LngDevNO, CommissionAccount, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
       
            
             LngDevNO = LngDevNO + 1
             
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, SupplierAccount, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
            
    End If
    
   supplierGL.Close
   Set supplierGL = Nothing
    
    
    
  '' SngTemp = (LblTotalAll - LblDiscountsTotal) * val(txt_Currency_rate.text)
   ' SngTempe = (LblTotalAll - LblDiscountsTotal)
    
    TotalValue = TotalValue * val(txt_Currency_rate.Text) '+ Round(LblCommision * val(txt_Currency_rate.text), SystemOptions.SysDefCurrencyForamt)
      LngDevNO = 1
   If TotalValue > 0 Then


    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic, TotalValue, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
    End If
    
    
    Exit Function
ErrTrap:


End Function
Function SaveNewGl2()
'If SystemOptions.PoCreateVoucher = True And CboPaymentType.ListIndex = 1 Then
'If txt_ORDER_NO.text = "" Then Exit Function

'End If

 
 If SystemOptions.PoCreateVoucher = True And CboPayMentType.ListIndex = 1 Then
        If TXT_order_no.Text = "" Then
         Exit Function
        Else
         
        End If

Else

Exit Function
End If





 Dim supplierGL As New ADODB.Recordset
  Dim StrSQL As String
  Dim RowNum As Integer
  Dim Account_Code_dynamic As String
  Dim StrTempDes  As String
  Dim CommissionAccount As String
 Dim LngDevID As Long
    Dim LngDevNO  As Integer
 


 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            End If

   
    Dim SngTemp  As Variant
        Dim SngTempe  As Variant
 

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    
 
Dim value As Variant
Dim Account_code As String
Dim CusName As String
Dim TotalValue As Variant
Dim SngTempe2 As Variant
Dim value2 As Variant

TotalValue = 0
   LngDevNO = 1
   value = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
   value2 = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
 
  '  Account_Code = Account_Code_dynamic101
 '   CusName = IIf(IsNull(supplierGL("CusName").value), "", supplierGL("CusName").value)
    
    If value > 0 Then
  '  value = Round(value, SystemOptions.SysDefCurrencyForamt)
    SngTemp = Round(value, SystemOptions.SysDefCurrencyForamt)
     SngTempe = value
     SngTempe2 = value2
     LngDevNO = LngDevNO + 1
     TotalValue = TotalValue + SngTempe
              If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text & " «À»«  „‘ —Ì«  «·„Ê—œ " & Me.DBCboClientName.Text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            End If
            
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic102, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTempe2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
    End If
         
 
   
     LngDevNO = LngDevNO + 1
  
   If value > 0 Then


            If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic101, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTempe2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
    
    End If
    
    Exit Function
ErrTrap:


End Function
Function CheckGeidExpensss() As Boolean
Dim i As Integer
CheckGeidExpensss = False
With Fg_Journal
For i = 1 To .rows - 1
If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And .TextMatrix(i, .ColIndex("Accountcode2")) = "" And .TextMatrix(i, .ColIndex("FlgVat")) = "" Then
CheckGeidExpensss = True
Exit Function
End If
Next i
End With
End Function
Function CheckBeforSave() As Boolean
    CheckBeforSave = True
Dim Msg As String
    If Not IsSaveWithOutMsg Then

        If Not Checks Then
            CheckBeforSave = False
            Exit Function
        End If
            
'        If Due_Date > DtpDelayDate.value Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                MsgBox "ÌÃ» «‰ ÌﬂÊ‰  «—ÌŒ «·«” Õﬁ«ﬁ «ﬂ»—   „‰ «Ê Ì”«ÊÌ     «—ÌŒ «Œ— ﬁ”ÿ"
'            Else
'                MsgBox "installment Date Must be Graeter than  or equal todya"
'
'            End If
'
'            CheckBeforSave = False
'            Exit Function
'        End If
        
        If CboPayMentType.ListIndex = 1 Then
            Me.XPChkPayType(1).value = 1
            ' hany  XPTxtValue(1).text = Val(LblTotalAll.Caption)
        End If
        
        If Trim(Me.TxtTransSerial.Text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ ›« Ê—… «·‘—«¡..!!!"
            Else
                Msg = "Must Enter Bill No."
           
            End If
           
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.TxtTransSerial.SetFocus
            CheckBeforSave = False
            Exit Function
        End If
        
        '«· √ﬂœ „‰ ⁄œ„  ﬂ—«— —ﬁ„ «·”‰œ
        Dim BolTemp As Boolean
        
        If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22) = "" Then
            If Me.TxtModFlg.Text = "N" Then
           
                BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.Text), 22, , val(dcBranch.BoundText))
            ElseIf Me.TxtModFlg.Text = "E" Then
                BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.Text), 22, val(Me.XPTxtBillID.Text), val(dcBranch.BoundText))
            End If
        
            If BolTemp = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "—ﬁ„ «·”‰œ „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & CHR(13)
                    Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·”‰œ"
                Else
                    Msg = "This Bill No Already Exist" & CHR(13)
               
                End If
        
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtNoteSerial1.SetFocus
                Screen.MousePointer = vbDefault
                CheckBeforSave = False
                Exit Function
            End If
           
        End If
        
        '‰Â«Ì… «· √ﬂœ
        
        Screen.MousePointer = vbArrowHourglass
        
        If DBCboClientName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ"
            Else
                Msg = "Select Customer Name"
           
            End If
        
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            CheckBeforSave = False
            Exit Function
        End If
        
        If DCboStoreName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ Õœœ «”„ «·„Œ“‰"
            Else
                Msg = "Select Inventory First"
            End If
        
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            CheckBeforSave = False
            Exit Function
        End If
        
        If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
            If XPTxtDiscountVal.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ»  ÕœÌœ ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—…"
                Else
               
                    Msg = "Specify Total Discount"
                End If
        
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtDiscountVal.SetFocus
                Screen.MousePointer = vbDefault
                CheckBeforSave = False
                Exit Function
            End If
        
            If Not IsNumeric(XPTxtDiscountVal.Text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                Else
                    Msg = "Discount Value Must be Numeric"
                End If
        
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtDiscountVal.SetFocus
                Screen.MousePointer = vbDefault
                CheckBeforSave = False
                Exit Function
            End If
        
            XPTxtDiscountVal.SetFocus
        End If
        
        If CboPayMentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
            Else
                Msg = "Specify Payment Method"
           
            End If
        
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            CheckBeforSave = False
            Exit Function
        End If
        
        If XPChkPayType(0).value = vbChecked And Not ((CboPayMentType.ListIndex = 2 And SystemOptions.AllowPurchasesMultyPayed = False) Or (SystemOptions.AllowPurchasesMultyPayed = True And CboPayMentType.ListIndex = 3)) Then

            '
            If Me.DcboBox.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!"
                Else
                    Msg = "Specify Box Name "
                End If
        
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Screen.MousePointer = vbDefault
                CheckBeforSave = False
                Exit Function
            End If
        
            If Me.TxtModFlg.Text = "N" Then
                If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).Text), Me.XPDtbBill.value) = False Then
                    Screen.MousePointer = vbDefault
                    CheckBeforSave = False
                    Exit Function
                End If
        
            ElseIf Me.TxtModFlg.Text = "E" Then
        
                If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).Text), Me.XPDtbBill.value, , , val(Me.XPTxtValue(0).Tag)) = False Then
                    Screen.MousePointer = vbDefault
                    CheckBeforSave = False
                    Exit Function
                End If
            End If
        End If
        
        If ((CboPayMentType.ListIndex = 2 And SystemOptions.AllowPurchasesMultyPayed = False) Or (SystemOptions.AllowPurchasesMultyPayed = True And CboPayMentType.ListIndex = 3)) Then

            '
            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ...!!!"
                Else
                    Msg = "Specify Box Name "
                End If
        
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Screen.MousePointer = vbDefault
                CheckBeforSave = False
                Exit Function
            End If
              
        End If

        If val(Me.XPTxtValue(1).Text) > 0 Then
            If ChkInstall.value = vbChecked Then
                If val(Me.LblInstallTotal.Caption) = 0 Then
                    Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Me.XPTab301.CurrTab = 1
                    Screen.MousePointer = vbDefault
                    CheckBeforSave = False
                    Exit Function
                End If
             
            End If
        End If
    End If

End Function

Private Sub SaveData()
    Dim usedaccount          As Integer
    Dim RSTransDetails       As ADODB.Recordset
    Dim RsNotes              As ADODB.Recordset
    Dim RsNotesGeneral       As ADODB.Recordset
    Dim RsTemp               As New ADODB.Recordset
    Dim Msg                  As String
    Dim RowNum               As Integer
    Dim StrSQL               As String
    Dim StrSqlDel            As String
    Dim SearchResault        As Integer
    Dim note_id              As Long
    Dim RsDetalis            As ADODB.Recordset
    Dim BeginTrans           As Boolean
    Dim LnItemID             As Long
    Dim i                    As Long
    Dim StrCurrentItemName   As String
    Dim DblNotesTotal        As Double

    Dim IntLineNO            As Integer
    Dim StrAccountCode       As String
    Dim TotalBillDiscount    As Double
    Dim TotalDiscountPerLine As Double
    ' On Error GoTo ErrTrap

   '*************************
   If Not CheckBeforSave() Then Exit Sub
    '**********************************
    If XPChkPayType(2).value = vbChecked Then
        If val(Me.lbl(18).Caption) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
            Else
                Msg = "Enter Cheques Data Before Save"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Dcbanks.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = Msg + "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ     " & CHR(13)
            Else
                Msg = Msg + " Specify Bank NAme     " & CHR(13)
            End If
        
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '            Dcbanks.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
    
            Dim rsbank As New ADODB.Recordset
            Set rsbank = New ADODB.Recordset
            rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
            If Not (rsbank.EOF Or rsbank.BOF) Then
                If rsbank!banks_Accounts = True Then
                    bank_account = get_bank_Account(val(Me.Dcbanks.BoundText), "Account_Code2")
                Else
                    bank_account = get_bank_Account(val(Me.Dcbanks.BoundText), "Account_Code")
                End If
            End If
        
        End If
    
    End If

    XPTab301.CurrTab = 0

    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If

    If CheckMyData = False Then
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    If Me.TxtModFlg.text = "E" Then
    '        If EditTransStatus(Val(Me.XPTxtBillID.text), "E", NewGrid) = False Then
    '            Exit Sub
    '        End If
    '    End If
    '---------------------------------------------------------------

    If NewGrid.Calculate(1, , False, True) = False Then
        Screen.MousePointer = vbDefault
        '  Exit Sub
    End If

    '-------------------------------
 
    DblNotesTotal = DblNotesTotal + val(Me.XPTxtValue(1).Text)
 
    DblNotesTotal = val(Me.XPTxtValue(0).Text) + val(Me.XPTxtValue(1).Text) + val(lbl(18).Caption)

    If CboPayMentType.ListIndex = 1 Then
        Me.XPChkPayType(1).value = 1
        '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
    End If

    '---------Start Saving------------------------------------------------
 
    '---------Notes ID ------------------------------------------------
    'Create big notes
    GoTo xll

    If TxtNoteSerial.Text = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                TxtNoteSerial.Text = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
        
    Dim NoteSerial1str As String
        
    If TxtNoteSerial1.Text = "" Then
    
        NoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22, , val(DCboStoreName.BoundText))

        If NoteSerial1str = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ›« Ê—… „‘ —Ì«  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                                   
            If NoteSerial1str = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… „‘ —Ì«   ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                TxtNoteSerial1.Text = NoteSerial1str
            End If
        End If
    End If
     
xll:
    
    '---------Start Saving------------------------------------------------
 
    'Õ›Ÿ «·„’—Ê›«  «·«÷«›Ì… Ê«·›Ê« Ì— «·„«·Ì…
    
    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    BeginTrans = True
    
    Cn.Execute "delete DOUBLE_ENTREY_VOUCHERS where Transaction_ID = " & val(Text2.Text)
    Save_Financial_invoice
    save_expenses

    Set RSTransDetails = New ADODB.Recordset
    Set RsNotes = New ADODB.Recordset
    '  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '  RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
    RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
    RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Me.TxtModFlg.Text = "N" Then
        XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs.AddNew
        rs("Transaction_ID").value = val(XPTxtBillID.Text)

        If TxtNoteSerial1 = "" Then
            TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22, , val(DCboStoreName.BoundText))
        End If
      
        Me.TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=21"))
         
        Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
          
    ElseIf Me.TxtModFlg.Text = "E" Then

        If rs("Transaction_ID").value <> val(XPTxtBillID.Text) Then
            rs.Find "Transaction_ID=" & val(XPTxtBillID.Text), , adSearchForward, 1
        End If

        Cn.Execute "delete from Transaction_Details where Transaction_ID in ( " & GetTrnasectionID(val(XPTxtBillID.Text), 22) & " )"
        Cn.Execute "delete from Transactions where Transaction_ID in ( " & GetTrnasectionID(val(XPTxtBillID.Text), 22) & " )"
        Cn.Execute "delete from DOUBLE_ENTREY_VOUCHERS where Transaction_ID in ( " & GetTrnasectionID(val(XPTxtBillID.Text), 22) & " )"
        Cn.Execute "Delete From TblTransctionIDES where Transaction_Type=22 and MainTransaction_ID=" & val(XPTxtBillID.Text) & " "
  
        Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.Text) & ""
        StrSqlDel = "delete From TblSalesPayment where TransID=" & val(Me.XPTxtBillID.Text) 'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
          
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.Text)
        Cn.Execute StrSQL
        
        StrSqlDel = "delete From Notes where   NoteType= 150 and noteid=" & val(TXTNoteID.Text)
        Cn.Execute StrSqlDel
        
        general_noteid = val(TXTNoteID.Text)

        If TxtNoteSerial.Text = "" Then
            TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22, , val(DCboStoreName.BoundText))
        End If
        
    End If
    
    If ChkCompsBill.value = vbChecked Then
        rs("CompsBill").value = 1
    Else
        rs("CompsBill").value = 0
    End If
 
    If VstReverse.value = vbChecked Then
        rs("VstReverse").value = 1
    Else
        rs("VstReverse").value = 0
    End If
 
    rs("txtManulaVat").value = val(txtManulaVat.Text)

    rs("VATCustoms").value = val(TxtVATCustoms.Text)
    rs("VATCustoms1").value = val(TxtVATCustoms1.Text)
    rs("VATNO").value = IIf(Trim(Me.TxtVATNO.Text) = "", Null, Trim(Me.TxtVATNO.Text))
    rs("VAT").value = val(TxtValueAdded.Text)
    rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, val(DcboEmp.BoundText))
    rs("ManualNo1").value = IIf(TxtManualNo1.Text = "", Null, val(TxtManualNo1.Text))
    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.Text) = "", Null, Trim(Me.TxtNoteSerial.Text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
    rs("NoteId").value = val(TXTNoteID.Text)
    rs("order_no").value = IIf((TXT_order_no.Text) = "", Null, TXT_order_no.Text)
    rs("ContainerNo").value = IIf(txtContainerNo.Text = "", Null, Trim(txtContainerNo.Text))
    rs("poTransaction_ID").value = val(poTransaction_ID)
    
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.Text), 1, txt_Currency_rate.Text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.Text) = "", Null, Trim(Me.TxtTransSerial.Text))
    rs("Transaction_Date").value = XPDtbBill.value
    rs("ArrivalDate").value = DTArrivalDate.value
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs("Transaction_Type").value = BillType
    rs("DueDate").value = DtpDelayDate.value
    rs("UserID").value = user_id
    rs("nots").value = Text1.Text
    rs("Typ").value = val(Me.DcbTyp.ListIndex)
    rs("ResonVAT").value = TXtResonVAT.Text
    rs("BLDate").value = BLDate.value
    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))

    If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
    Else
        rs("CashCustomerName").value = Null
    End If

    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If

    If XPCboDiscountType.ListIndex = -1 Or XPCboDiscountType.ListIndex = 0 Then
        rs("Trans_Discount").value = Null
 
    Else
        rs("Trans_Discount").value = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text))
    End If

    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else

        If SystemOptions.AllowPurchasesMultyPayed = False And val(CboPayMentType.ListIndex) = 2 Then
            rs("PaymentType").value = 3
        Else
            rs("PaymentType").value = val(CboPayMentType.ListIndex)
        End If
    End If

    If Trim$(Me.TxtPhone.Text) <> "" Then
        rs("CashCustomerPhone").value = Trim$(Me.TxtPhone.Text)
    Else
        rs("CashCustomerPhone").value = Null
    End If

    rs("AddValue").value = IIf(val(txtAddValue.Caption) = 0, Null, val(txtAddValue.Caption))

    If ChAddToTotal.value = vbChecked Then
        rs("AddToTotal").value = 1
    Else
        rs("AddToTotal").value = Null
    End If
     
    rs("project_id").value = IIf(dcproject.BoundText = "", Null, (dcproject.BoundText))
    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, (DBCboClientName.BoundText))
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, (DCboStoreName.BoundText))
    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.Text = "", Null, val(XPTxtTaxValue.Text))
    rs("ToTAlELSHahn").value = IIf(Not IsNumeric(TXTToTAlELSHahn.Text), 0, Me.TXTToTAlELSHahn.Text)
    
    rs("total_expenses").value = IIf(Not IsNumeric(Txt_EXport.Text), 0, Txt_EXport.Text)
    rs("total_payments").value = IIf(Not IsNumeric(txt_total_bill.Text), 0, txt_total_bill.Text)
    rs("LcNo").value = IIf(TxtLcNo.Text = "", Null, (TxtLcNo.Text))
    rs("Transaction_NetValue").value = val(LblTotal.Caption)


      If chkTaxExempt.value = vbChecked Then
            rs("chkTaxExempt").value = 1
        Else
            rs("chkTaxExempt").value = 0
        End If

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.Text) > 0 Then
        rs("TaxAddValue").value = val(Me.TxtTaxAddValue.Text)
    Else
        rs("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.Text) > 0 Then
        rs("TaxStampValue").value = val(Me.TxtTaxStampValue.Text)
    Else
        rs("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.Text) > 0 Then
        rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.Text)
    Else
        rs("TaxServiceValue").value = 0
    End If

    'rs("Shipment_no").value = IIf(txt_Shipment_no.text = "", Null, (txt_Shipment_no.text))
    rs("order_no").value = IIf(TXT_order_no.Text = "", Null, (TXT_order_no.Text))
    rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, val(Dccurrency.BoundText))
    rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    '
    rs("BankID").value = IIf(Me.DcboBankName.BoundText = "", Null, val(DcboBankName.BoundText))
    rs("ManualNO").value = IIf(TxtManualNO.Text = "", Null, (TxtManualNO.Text))

    If XPCboDiscountType.ListIndex = -1 Then
        rs("CBoBasedON").value = 0
    Else
        rs("CBoBasedON").value = val(CBoBasedON.ListIndex)
    End If
    
    If BillBasedOn(0).value = True Then
        rs("BillBasedOn").value = 0
    ElseIf BillBasedOn(1).value = True Then
        rs("BillBasedOn").value = 1
    ElseIf BillBasedOn(2).value = True Then
        rs("BillBasedOn").value = 2
    End If
    
    rs.update

    Dim NoteID     As Long
    Dim NoteDate   As Date
    Dim NoteSerial As String
    Dim Notevalue  As Double
    Dim des        As String

    If Me.TxtNoteSerial.Text <> "" Then
        NoteSerial = Me.TxtNoteSerial.Text
    End If
        
    If ChAddToTotal.value = vbUnchecked Then
        Notevalue = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
    Else
        '  If ChSameCurrncy.value = vbChecked Then
        '  Notevalue = (val(TxtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
        '  Else
        Notevalue = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
        '  End If
    End If

    Notevalue = Notevalue + val(LblValueAdded.Caption) * val(txt_Currency_rate.Text)
         
    SaveSalesPayment val(Me.XPTxtBillID.Text)

    CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 150, Notevalue, NoteSerial, TxtNoteSerial1, "Transactions", "Transaction_ID", val(XPTxtBillID.Text), TxtNoteSerial1.Text, ToHijriDate(XPDtbBill.value), TxtManualNO.Text
    TXTNoteID.Text = NoteID
    general_noteid = NoteID
    
    Dim mTaxExemptTotal As Double
   
    For RowNum = 1 To FG.rows - 1

        'Check Repeat Serial
        If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
            StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
            StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.Text
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                    Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                    Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                Else
                    Msg = "Item Serial" & CHR(13)
                    Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                    Msg = Msg + "Already Exist in this bill"
            
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                RsTemp.Close
                XPTab301.CurrTab = 0
                FG.Row = RowNum
                FG.Col = FG.ColIndex("name")
                FG.ShowCell RowNum, FG.ColIndex("name")
                FG.SetFocus
                Screen.MousePointer = vbDefault
                BeginTrans = False
                Cn.RollbackTrans
                Exit Sub
            End If

            RsTemp.Close
        End If

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            RSTransDetails.AddNew
            RSTransDetails("BranchId").value = Me.dcBranch.BoundText
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Me.XPDtbBill.value, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
            RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text)
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
            RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            RSTransDetails("StoreID2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("StoreID2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("StoreID2"))))
                If Me.FG.ColIndex("SalesValue") <> -1 Then
                RSTransDetails("SalesValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("SalesValue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("SalesValue"))))
            End If
             RSTransDetails("Account_Code").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Account_Code")) = ""), Null, Trim(FG.TextMatrix(RowNum, FG.ColIndex("Account_Code"))))

            If SystemOptions.IsGeometricProportions = True Then
                RSTransDetails("UnitPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("UnitPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("UnitPrice"))))
            End If
            
            '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) _
            '            = ""), "", Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
            If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                                
                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If RsTemp("HaveSerial").value = True Then
                        RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Serial"))))
                    End If
                End If
                                
                RsTemp.Close
            End If


            
            
            
           
             RSTransDetails("projectid").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("projectid")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("projectid")))
            
            RSTransDetails("LineExpenses").value = ((val(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / val(LblTotalAll.Caption))) / RSTransDetails("Quantity"))
            RSTransDetails("TypeVAT").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT")))
            RSTransDetails("ItemsDetailsNewidea").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")))
            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
            RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
        
            RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
         
            '.TextMatrix(LngRow, .ColIndex("ColorID")) = 1
            '.TextMatrix(LngRow, .ColIndex("ItemSize")) = 0
        
            RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
            RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
         
            RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            RSTransDetails("order_no").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))

            '     RSTransDetails("OrderArrivalDate").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("OrderArrivalDate")) = ""), Me.XPDtbBill.value, Fg.TextMatrix(RowNum, Fg.ColIndex("OrderArrivalDate")))
            If (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) = Empty Then
              
                Dim newUnitId As Long
                GetDefaultItemUnit RSTransDetails("Item_ID").value, newUnitId
                RSTransDetails("UnitID").value = newUnitId
            End If
             
            RSTransDetails("LineShahn").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("LineShahn")) = ""), 0, FG.TextMatrix(RowNum, FG.ColIndex("LineShahn")))
             
            Dim RsUnitData   As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID    As Long
            Dim DblQty       As Double
        
            LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))

            If (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) = Empty Then
            
                LngUnitID = newUnitId
            Else
                LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            End If
            
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            If LngUnitID = 0 Then

            End If

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
            End If

            '          RSTransDetails("Price").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ShowPrice")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("ShowPrice"))))
            '     RSTransDetails("price").value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").value, 2)
            RSTransDetails("price").value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Price")) / RSTransDetails("QtyBySmalltUnit").value, 15)
            RSTransDetails("OpeningBurcahseQty").value = RSTransDetails("Quantity").value
            RSTransDetails("OpeningBurcahseValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
            
            Dim Rate As Single
         
            RSTransDetails("discountvalue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("discountvalue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("discountvalue")))) / RSTransDetails("Quantity").value
      
            'RATE = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ScurrencyID")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("ScurrencyID"))))
            RSTransDetails("rate").value = val(txt_Currency_rate.Text)

            If val(LblTotal.Caption) = 0 Then LblTotal.Caption = 1
            ' RSTransDetails("ToTAlELSHahn") = Round((((RSTransDetails("showPrice") * _
            ' RSTransDetails("ShowQty")) / Val(LblTotal.Caption)) * _
            ' Val(TXTToTAlELSHahn.text)) / RSTransDetails("ShowQty"), 2)   ' / RSTransDetails("ShowQty")
            Dim TotalShahnPerLine As Double

            If SystemOptions.ExpensesByQtyOnly = True Then
            
                TotalShahnPerLine = ((((IIf(IsNull(RSTransDetails("price")), 0, 1) * IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity")) / (LblTotalQty.Caption))) * val(TXTToTAlELSHahn.Text)) / IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity")))
            
            Else
                TotalShahnPerLine = ((((IIf(IsNull(RSTransDetails("price")), 0, RSTransDetails("price")) * IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity")) / (LblTotalAll.Caption))) * val(TXTToTAlELSHahn.Text)) / IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity")))
            End If
            
            TotalShahnPerLine = Round(TotalShahnPerLine, 15) 'Val(Format(TotalShahnPerLine, "." & String(Abs(18), "#")))
            RSTransDetails("ToTAlELSHahn") = TotalShahnPerLine
         
            If Me.XPCboDiscountType.ListIndex = 1 Then
                TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text))
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.Text <> "" Then
                    TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text)) * val(LblTotalAll.Caption) / 100
                             
                Else
                    TotalBillDiscount = 0
                End If
            End If

            RSTransDetails("price") = IIf(IsNull(RSTransDetails("price")), 0, RSTransDetails("price"))
            TotalDiscountPerLine = ((((RSTransDetails("price") * RSTransDetails("Quantity") / (LblTotalAll.Caption))) * val(TotalBillDiscount)) / RSTransDetails("Quantity"))
            RSTransDetails("TotalDiscountPerLine") = TotalDiscountPerLine
            
            If VstReverse.value = vbChecked Then
                RSTransDetails("VstReverse").value = 1
            Else
                RSTransDetails("VstReverse").value = 0
            End If
            
            '*******************************************************************test
 
            '        TotalDiscountPerLine = FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / (LBLGross) * TotalBillDiscount
           
            '      TotalDiscountPerLine = Round(TotalDiscountPerLine, 20)
            '          RSTransDetails("TotalDiscountPerLine") = TotalDiscountPerLine
            '*******************************************************************test
            
            ' RSTransDetails.update
     
            RSTransDetails("OUTR").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("OUTR")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("OUTR"))))
            RSTransDetails("INR").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("INR")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("INR"))))
        
            RSTransDetails("NoCount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NoCount")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("NoCount"))))
        
            RSTransDetails("Height").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Height"))))
            RSTransDetails("length").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("length")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("length"))))
            RSTransDetails("Width").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Width"))))

            RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
            RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
            RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
            RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            
            RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
            '    RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(Fg.TextMatrix(RowNum, Fg.ColIndex("OrderArrivalDate"))), Null, Fg.TextMatrix(RowNum, Fg.ColIndex("OrderArrivalDate")))
            RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
            RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
            RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
            RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
            RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
            RSTransDetails("GoldDetails").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails")))
            RSTransDetails("Wages").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Wages")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("Wages")))

            '******************************
            RSTransDetails("ExtraType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExtraType")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ExtraType"))))
            RSTransDetails("ExtraVal").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExtraVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ExtraVal"))))
            'RSTransDetails("Commisionvalue").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Commisionvalue")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Commisionvalue"))))
            RSTransDetails("Commisionvalue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Commisionvalue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Commisionvalue")))) / RSTransDetails("Quantity").value

            RSTransDetails("SBillNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("SBillNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("SBillNO")))

            RSTransDetails("SupplierID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("SupplierID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("SupplierID"))))
            RSTransDetails("Scurrenyrate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Scurrenyrate")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Scurrenyrate"))))
            RSTransDetails("ScurrencyID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ScurrencyID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ScurrencyID"))))
            RSTransDetails("Vat").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vat")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vat"))))
            RSTransDetails("Vatyo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vatyo"))))
      
            '******************************
            Dim OldQty  As Double
            Dim OldCost As Double
            Dim NewQty  As Double
            Dim NewCost As Double
               
            'getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.Text), OldQty, OldCost, NewQty, NewCost
            '  RSTransDetails("OldQty").value = NewQty
            '  RSTransDetails("OldCost").value = NewCost
       
            ' RSTransDetails("NewQty").value = RSTransDetails("Quantity").value + RSTransDetails("OldQty").value
            '  RSTransDetails("NewCost").value = ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       
            RSTransDetails.update
        End If

        'wael
        If FG.ValueMatrix(RowNum, FG.ColIndex("chkTaxExempt")) = True Then
            mTaxExemptTotal = mTaxExemptTotal + ((val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) * val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) - val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountValue")))
                
        End If

    Next RowNum

    LblValueAdded.Tag = mTaxExemptTotal
    rs!TotalTaxExempt = mTaxExemptTotal
    '------------------------------------------------------------------------------
     
    '------------------------------------------------------------------------------
   

    If Me.XPChkPayType(1).value = Checked Then
        RsNotes.AddNew
        RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        note_id = RsNotes("NoteID").value
        RsNotes("NoteDate").value = XPDtbBill.value
 
        RsNotes("remark").value = Me.TxtNoteSerial1.Text
        RsNotes("NoteSerial").value = Null

        RsNotes("Transaction_ID").value = val(XPTxtBillID.Text)
        RsNotes("NoteType").value = 1
        RsNotes("Note_Value").value = IIf(XPTxtValue(1).Text = "", Null, val(XPTxtValue(1).Text))
        RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        RsNotes("BankID").value = Null
        RsNotes("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText)) 'Null SALIM MY BE ERROR
        RsNotes("DueDate").value = DtpDelayDate.value
        RsNotes.update
 
    End If
chkTaxExempt_Click
    If Me.XPChkPayType(2).value = Checked Then

        With Me.FgCheques

            For i = .FixedRows To .rows - 1
                '--------------------------------------------------------------------------
            Next i

        End With

    End If

    'Õ›Ÿ «·√›”«ÿ
    If Me.XPChkPayType(1).value = Checked Then
        If ChkInstall.value = vbChecked Then
            'Save installment Data
            Set RsTemp = New ADODB.Recordset
            
            '      RsTemp.Open "InstallMent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                 
            StrSQL = " SELECT       * FROM  dbo.InstallMent WHERE     (PartID = - 1)"
            RsTemp.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
            RsTemp.AddNew
            RsTemp("PartID").value = CStr(new_id("InstallMent", "PartID", "", True))
            RsTemp("NoteID").value = note_id
            RsTemp("BasicAmmount").value = IIf(XPTxtValue(1).Text = "", 0, val(XPTxtValue(1).Text))
            RsTemp("InterestType").value = val(Me.LblPrecenType.Tag)
            RsTemp("InterestVal").value = val(LblPrecenValue.Caption)
            RsTemp("Total").value = val(LblInstallTotal.Caption)
            RsTemp("InstallCount").value = val(LblInstallCount.Caption)
            RsTemp("FirstInstallDate").value = CDate(Me.LblFirstInstallDate.Caption)

            If val(LblInstallmentType.Tag) = 0 Then
                RsTemp("InstallmentType").value = 0
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                RsTemp("InstallmentType").value = 1
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                RsTemp("InstallmentType").value = 2
            End If

            RsTemp("InstallSeprator").value = val(Me.LblInstallSeprator.Caption)
            RsTemp("StartValue").value = IIf(val(Me.LblStartValue.Caption) = 0, Null, val(Me.LblStartValue.Caption))
            RsTemp("CustID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsTemp("Type").value = 1
            RsTemp.update
            'save installment Details
            Set RsDetalis = New ADODB.Recordset
            '   RsDetalis.Open "InstallMentDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            StrSQL = " SELECT       * FROM  dbo.InstallMentDetails WHERE     (PartID = - 1)"
            RsDetalis.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
            With Me.FgInstallments

                For RowNum = 1 To .rows - 1
                    RsDetalis.AddNew
                    RsDetalis("QestID").value = CStr(new_id("InstallMentDetails", "QestID", "", True))
                    RsDetalis("PartID").value = RsTemp("PartID").value
                    RsDetalis("QeqtNum").value = IIf(.TextMatrix(RowNum, .ColIndex("Serial")) = "", "", .TextMatrix(RowNum, .ColIndex("Serial")))
                    RsDetalis("Value").value = IIf(.TextMatrix(RowNum, .ColIndex("Value")) = "", "", val(.TextMatrix(RowNum, .ColIndex("Value"))))
                    RsDetalis("DueDate").value = IIf(.TextMatrix(RowNum, .ColIndex("Due_Date")) = "", "", .TextMatrix(RowNum, .ColIndex("Due_Date")))
                    RsDetalis("Receipt").value = False
                    RsDetalis.update
                Next RowNum

            End With

        End If
    End If

    Dim LngDevID             As Long, LngDevNO            As Integer, StrTempAccountCode  As String, StrTempDes          As String
    Dim SngTemp              As Variant
    Dim SngTemp2             As Variant
    Dim Account_Code_dynamic As String
    '    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
    'If SystemOptions.PoCreateVoucher = True And CBoBasedON.ListIndex = 1 And TXT_order_no.text <> "" Then GoTo NewGL2

    If SystemOptions.PoCreateVoucher = True And CboPayMentType.ListIndex = 1 Then
        If TXT_order_no.Text = "" Then
 
        Else
            GoTo NewGL2
        End If

    End If

    Dim Note_Value   As Double
    Dim Note_Value2  As Double
    Dim Account_code As String
   
    ''
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    If ChkCompsBill.value = vbChecked Then GoTo NewGL
    If ChAddToTotal.value = vbUnchecked Then
    
        With Fg_Journal
            Dim mDisc As String

            For i = 1 To .rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And val(.TextMatrix(i, .ColIndex("value"))) <> 0 And .TextMatrix(i, .ColIndex("AccountCode2")) <> "" Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.Text & CHR(13) & Trim(.TextMatrix(i, .ColIndex("des")))
                    Else
                        StrTempDes = "Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.Text & CHR(13) & Trim(.TextMatrix(i, .ColIndex("des")))
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_code = .TextMatrix(i, .ColIndex("AccountCode"))
                    Note_Value = val(.TextMatrix(i, .ColIndex("value"))) * val(txt_Currency_rate.Text)
                    Note_Value2 = val(.TextMatrix(i, .ColIndex("value")))
                    mDisc = Trim(.TextMatrix(i, .ColIndex("des")))

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , mDisc, , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , mDisc) = False Then
                        GoTo ErrTrap
                    End If

                    LngDevNO = LngDevNO + 1
                    Account_code = .TextMatrix(i, .ColIndex("AccountCode2"))

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , mDisc) = False Then
                        GoTo ErrTrap
                    End If
                    
                End If
        
            Next

        End With

    End If

    '«·ÿ—› «·„œÌ‰
    
    
    

    
    '    SngTemp = (Val(Me.lblTotal.Caption) * Val(txt_Currency_rate.text))
    If ChAddToTotal.value = vbChecked Then
        SngTemp2 = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
        SngTemp = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
    Else
        SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
        SngTemp2 = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
    End If

    SngTemp = SngTemp ' val(LblValueAdded.Caption)
    SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
    SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

            Account_Code_dynamic = get_account_code_branch(4, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "Branch Not Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„‘ —Ì«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Purchase  Account   Not Defined in this Branch", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If

            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ··›« Ê—…", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 Then
        
                    StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
            End If
            
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            End If

            LngDevNO = LngDevNO + 1
    
            If TxtManualNO.Text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & TxtManualNO
                Else
                    StrTempDes = StrTempDes & " Supp Bill# " & TxtManualNO
                End If
            
            End If

            Dim Material_account As String
            Dim project_id       As Integer

            If SystemOptions.NotCrtResvVouchProjects = True And dcproject.BoundText <> "" Then
                project_id = val(dcproject.BoundText)
                Material_account = get_project_Account(project_id, "Material_account")
                
                If Material_account <> "" Then

                    StrTempAccountCode = Material_account

                End If

            End If

            ''////////////
            If Me.XPChkPayType(1).value = vbChecked And CboPayMentType.ListIndex <> 2 Then
                '«·√Ã·
                OtherInformation.NextAccount_Code = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))

            ElseIf Me.XPChkPayType(0).value = vbChecked Then
                OtherInformation.NextAccount_Code = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
            End If

            ''////////
            If SystemOptions.CreateEntryBillItems Then
                i = 1
                For i = 1 To FG.rows - 1
                    'If Trim(Fg.TextMatrix(i, Fg.ColIndex("projectid"))) = "" Then
                    '    MsgBox "«Œ — «·„‘—Ê⁄"
                    '    GoTo ErrTrap
                    'End If
                    If Trim(FG.TextMatrix(i, FG.ColIndex("Account_Code"))) = "" Then
                    
                        If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»     «·„‘ —Ì«  ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Purchase  Account Not Defined"
                            End If

                            GoTo ErrTrap
                    Else
                        LngDevNO = LngDevNO + 1

                        mValue = val((val(FG.TextMatrix(i, FG.ColIndex("Count"))) * val(FG.TextMatrix(i, FG.ColIndex("Price"))))) - val(FG.TextMatrix(i, FG.ColIndex("DiscountValue")))
                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Trim(FG.TextMatrix(i, FG.ColIndex("Account_Code"))), mValue, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , mValue, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , val(dcproject.BoundText), , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                            GoTo ErrTrap
                        End If
                
                    
                    End If
                    
                    
                Next
                
        
            Else
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , val(dcproject.BoundText), , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                    GoTo ErrTrap
                End If
        
            End If
            
      

            '''/////////////////////
            If val(TxtValueAdded.Text) > 0 Then
                Dim AccountVATCreit As String
                GetValueAddedAccount XPDtbBill.value, AccountVATCreit, , 1, 22
                LngDevNO = LngDevNO + 1

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & " «·ﬁÌ„… «·„÷«›…  "
                Else
                    StrTempDes = StrTempDes & " VAT "
                End If

                SngTemp = Round(val(TxtValueAdded.Text) * val(txt_Currency_rate.Text), 2)
                SngTemp2 = val(TxtValueAdded.Text)

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                    GoTo ErrTrap
                End If
            End If

            '''///////////////////
        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value   As Single

            With FG

                For i = 1 To FG.rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 4)

                        'groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»     «·„‘ —Ì«  ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Purchase  Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * val(txt_Currency_rate.Text) * FG.TextMatrix(i, FG.ColIndex("Count"))
                        SngTemp2 = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
                        Else
                            StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
                        End If
                            
                        If TxtManualNO.Text <> "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & TxtManualNO
                            Else
                                StrTempDes = StrTempDes & " Supp Bill# " & TxtManualNO
                            End If
            
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If
    
        'Œ’„ ⁄·Ï „” ÊÏ «·”ÿ—
        If detect_inventory_work_type = 3 Then

            With FG

                For i = 1 To FG.rows - 1
 
                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("DiscountType"))) <> 1 Then
    
                        groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 13)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  Œ’„ „ﬂ ”»  ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Discount Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = (FG.TextMatrix(i, FG.ColIndex("Price")) * val(txt_Currency_rate.Text) * FG.TextMatrix(i, FG.ColIndex("Count"))) - FG.TextMatrix(i, FG.ColIndex("Valu"))
                        SngTemp2 = (FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))) - FG.TextMatrix(i, FG.ColIndex("Valu"))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
                        Else
                            StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With
    
        End If

    End If

    '«·œ«∆‰
    If Me.XPChkPayType(0).value = vbChecked Then
        If val(CboPayMentType.ListIndex) = 0 Then

            '«·Œ“Ì‰…
            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··›« Ê—…", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 Then
        
                    StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
                End If

            Else
                StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            End If
    
            If TxtManualNO.Text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & TxtManualNO
                Else
                    StrTempDes = StrTempDes & " Supp Bill# " & TxtManualNO
                End If
            
            End If

            '  SngTemp = (Val(Me.XPTxtValue(0).text) * Val(txt_Currency_rate.text))
            '  SngTemp = (Val(Me.lblTotal.Caption) * Val(txt_Currency_rate.text))
            ' SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) * Val(txt_Currency_rate.text)
            If ChAddToTotal.value = vbChecked Then
                SngTemp2 = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
                SngTemp = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
            Else
                SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
                SngTemp2 = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
            End If

            SngTemp = SngTemp + val(LblValueAdded.Caption) * val(txt_Currency_rate.Text)
            SngTemp2 = SngTemp2 + val(LblValueAdded.Caption)
            '   SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
            SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
            SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)
            LngDevNO = LngDevNO + 1

            If Trim(TxtLcNo) <> "" Then
                StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLcNo.Text)
            End If

            OtherInformation.NextAccount_Code = get_account_code_branch(4, my_branch)

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

        ElseIf ((CboPayMentType.ListIndex = 2 And SystemOptions.AllowPurchasesMultyPayed = False) Or (SystemOptions.AllowPurchasesMultyPayed = True And CboPayMentType.ListIndex = 3)) Then

            '
            '«·»‰ﬂ
            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··›« Ê—…", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 Then
        
                    StrTempAccountCode = GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
                End If

            Else
                StrTempAccountCode = GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
            End If

         
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
            
    
            If TxtManualNO.Text <> "" Then
           
                    StrTempDes = StrTempDes & " Supp Bill# " & TxtManualNO
                
            
            End If

            '  SngTemp = (Val(Me.XPTxtValue(0).text) * Val(txt_Currency_rate.text))
            '  SngTemp = (Val(Me.lblTotal.Caption) * Val(txt_Currency_rate.text))
            ' SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) * Val(txt_Currency_rate.text)
            If ChAddToTotal.value = vbChecked Then
                SngTemp2 = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
                SngTemp = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
            Else
                SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
                SngTemp2 = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
            End If

            SngTemp = SngTemp + val(LblValueAdded.Caption) * val(txt_Currency_rate.Text)
            SngTemp2 = SngTemp2 + val(LblValueAdded.Caption)
            '   SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
            SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
            SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)
            LngDevNO = LngDevNO + 1

            If Trim(TxtLcNo) <> "" Then
                StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLcNo.Text)
            End If

            OtherInformation.NextAccount_Code = get_account_code_branch(4, my_branch)

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
            
'
            If val(XPTxtDiscountVal) <> 0 Then
                SngTemp = val(LblDiscountsTotal)
                LngDevNO = LngDevNO + 1
                StrTempAccountCode = get_account_code_branch(13, my_branch)
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If


            End If

        Else

            With Grid22
                Dim StrMSG   As String
                Dim ValuGird As Double

                For i = 1 To .rows - 1
                    StrMSG = ""

                    If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
                        ValuGird = val(.TextMatrix(i, .ColIndex("Value"))) * val(txt_Currency_rate.Text)
                        SngTemp2 = val(.TextMatrix(i, .ColIndex("Value")))
                        StrMSG = " " & (.TextMatrix(i, .ColIndex("PaymentName")))

                        If val(.TextMatrix(i, .ColIndex("PaymentID"))) = 0 Then
                            If val(DCDocTypes.BoundText) > 0 Then
                                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

                                If StrTempAccountCode = "" And usedaccount = 1 Then
                                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··›« Ê—…", vbCritical
                                    GoTo ErrTrap
                                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                                ElseIf usedaccount = 0 Then
        
                                    StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
                                End If

                            Else
                                StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
                            End If

                            
                                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
                            
    
                            If TxtManualNO.Text <> "" Then
                                
                                    StrTempDes = StrTempDes & " Supp Bill# " & TxtManualNO
                                
            
                            End If

                            SngTemp = Round(ValuGird, 2)
                            SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)
                            LngDevNO = LngDevNO + 1

                            If Trim(TxtLcNo) <> "" Then
                                StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLcNo.Text)
                            End If
        
                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If

                        ElseIf val(.TextMatrix(i, .ColIndex("PaymentID"))) > 0 Then
                            StrTempAccountCode = .TextMatrix(i, .ColIndex("bankAccount_Code"))
                            SngTemp = Round(ValuGird, 2)
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes & " " & StrMSG, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
                        End If
                    End If

                Next i

            End With

        End If
    End If

    If Me.XPChkPayType(1).value = vbChecked And CboPayMentType.ListIndex <> 2 Then
    
        '«·√Ã·
        If ChAddToTotal.value = vbChecked Then
            SngTemp2 = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
            SngTemp = (val(txtAddValue.Caption) + NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
        Else
            SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.Text)
            SngTemp2 = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
        End If

        SngTemp = SngTemp + val(LblValueAdded.Caption) * val(txt_Currency_rate.Text)
        SngTemp2 = SngTemp2 + val(LblValueAdded.Caption)
        SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
        SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)

        If val(DCDocTypes.BoundText) > 0 Then
            getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

            If StrTempAccountCode = "" And usedaccount = 1 Then
                MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··›« Ê—…", vbCritical
                GoTo ErrTrap
            ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
            ElseIf usedaccount = 0 Then
        
                StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            End If

        Else
            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
        Else
            StrTempDes = "Purchase Invoice NO: " & Me.TxtNoteSerial1.Text & " " & TxtBillComment.Text
        End If

        LngDevNO = LngDevNO + 1
    
        If TxtManualNO.Text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & TxtManualNO
            Else
                StrTempDes = StrTempDes & " Supp Bill# " & TxtManualNO
            End If
            
        End If

        If Trim(TxtLcNo) <> "" Then
            StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLcNo.Text)
        End If

        OtherInformation.NextAccount_Code = get_account_code_branch(4, my_branch)

        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
            GoTo ErrTrap
        End If
        
        
        
        '
            If val(XPTxtDiscountVal) <> 0 Then
                SngTemp = val(LblDiscountsTotal)
                LngDevNO = LngDevNO + 1
                StrTempAccountCode = get_account_code_branch(13, my_branch)
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If


            End If
    End If

    If Me.XPChkPayType(2).value = vbChecked Then
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) * val(txt_Currency_rate.Text)
        SngTemp2 = NewGrid.GetItemsTotal(ItemsGoodType)
        SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
        SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)
        StrTempAccountCode = bank_account  '‘Ìﬂ«  „ƒÃ·…

        '    StrTempAccountCode = "a2a3a2" '√Ê—«ﬁ «·œ›⁄
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "⁄œœ " & Me.lbl(19).Caption & "  ‘Ìﬂ«  " & CHR(13)
            StrTempDes = StrTempDes & "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.Text
        Else
            StrTempDes = "Count " & Me.lbl(19).Caption & "  Cheque " & CHR(13)
            StrTempDes = StrTempDes & "Purchase Invoice No:" & Me.TxtNoteSerial1.Text
    
        End If

        LngDevNO = LngDevNO + 1

        If Trim(TxtLcNo) <> "" Then
            StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLcNo.Text)
        End If
        
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , SngTemp2, Dccurrency.Text, val(txt_Currency_rate.Text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
            GoTo ErrTrap
        End If
    End If

    'new GE
    Dim NoteID2 As Double
    NoteID2 = general_noteid
    updateNotesValueAndNobytext NoteID2

  
    
    
    SaveFlow

NewGL:
    SaveNewGl
NewGL2:

SaveNewGl2
    CloseIssueVoucher
    
    If SystemOptions.autoReseiveVoucher = True Then
        IsVouc = False

        If Not CreateRecieveVouchers Then BeginTrans = True: MsgBox "ÕœÀ Œÿ√ «À‰«¡ «‰‘«¡ «–‰ «·«” ·«„ ": GoTo ErrTrap
            
    End If
    
    SaveValueAdded
       
    close_order2 Me.TXT_order_no
   

    Cn.CommitTrans
    BeginTrans = False:      XPTxtCurrent.Caption = rs.AbsolutePosition:     XPTxtCount.Caption = rs.RecordCount
    
    If IsSaveWithOutMsg Then Exit Sub

    '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
      
    If invoiceSerach = True Then
        StrSQL = "SELECT * FROM Transactions WHERE Transaction_ID=" & val(Me.XPTxtBillID.Text) & "" ' & InvType
    Else
        StrSQL = "SELECT * FROM Transactions WHERE  Transaction_Type=" & BillType
    End If
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.Retrive val(Me.XPTxtBillID.Text)
    '----------------------------------------------------------------

    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & CHR(13)
    
            End If
    
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Changes was Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If

            lbl(67).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    
    End Select

    'Closeorders
    TxtModFlg.Text = "R"
    Command4_Click

    Screen.MousePointer = vbDefault
    Command2.Enabled = True
    Txt_EXport.Enabled = True
    'Grid.Visible = False
    Exit Sub
ErrTrap:
    
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans

        If rs.State = 0 Then
     rs.Resync adAffectCurrent
        End If
    End If

    Screen.MousePointer = vbDefault

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry....Error During Saving" & CHR(13)
    End If
 
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Function Checks() As Boolean
Dim StrSQL  As String
    If CBoBasedON.ListIndex = 3 Then
         If SystemOptions.MaintOrderCantRepeatBillBuy Then
            Dim rs2 As New ADODB.Recordset
            

            StrSQL = "SELECT NoteSerial1 FROM Transactions where Transaction_Type = 22 and IsNull(order_no,0)  = '" & val(TXT_order_no.Text) & "' and Transaction_ID <> " & val(XPTxtBillID)
            rs2.Open StrSQL, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rs2.EOF Then
                MsgBox "Â–« «·«„— ·« Ì„ﬂ‰ «œ—«ÃÂ ›ﬁœ «œ—Ã „‰ ﬁ»· ›Ï «·›« Ê—… —ﬁ„" & rs2!NoteSerial1 & ""
                TXT_order_no = ""
                Cmd(2).Enabled = True
                Checks = False
                Exit Function
            End If
        End If
    End If
    Checks = True
End Function
Private Sub XPBtnNewClients_Click()

    'With FrmAddNewCustemer
    '    '    .Tag = "x"
    '    .DealingForm = PurchaseTransaction
    '    Set .DcboCustomers = DBCboClientName
    '    .Caption = "≈÷«›… „Ê—œ ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·„Ê—œ"
    '    .lbl(0).Caption = "«”„ «·„Ê—œ"
    '    .AddType = 2
    '    .show vbModal
    ''End With

End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
 
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
        lbl(11).Enabled = False
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.Text = ""
    Else
        lbl(11).Enabled = True
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.Text = ""
    End If

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1
        End If
    End If

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    If XPCboDiscountType.ListIndex = 0 Then
        ' lbl(8).Visible = False
        XPTxtDiscountVal.Visible = False
        lbl(11).Visible = False
    Else
        lbl(11).Visible = True
        XPTxtDiscountVal.Visible = True
        lbl(11).Visible = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap
    Exit Sub

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.Text = "N" Then
                    XPTxtValue(0).Text = ""
                    XPTxtSerial(0).Text = ""
                End If

                If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).Text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.Text = "N" Then
                    XPTxtValue(1).Text = ""
                    DtpDelayDate.value = Date
                End If

                If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

                Me.ChkInstall.Enabled = True
            Else
                XPTxtValue(1).Enabled = False
                XPTxtValue(1).Text = ""
                Me.ChkInstall.Enabled = False
            End If

        Case 2
Dim noteseialDuplicates As String

        If SystemOptions.DontDuplicateManulaNoInPurchase = True Then
        
                             If checkManulanoisExist(22, val(Me.XPTxtBillID.Text), val(Me.DBCboClientName.BoundText), TxtManualNO.Text, noteseialDuplicates) = True Then
                                         If SystemOptions.UserInterface = ArabicInterface Then
                                             MsgBox "—ﬁ„ ﬁ« Ê—… «·„Ê—œ „ﬂ—— ›Ì ›« Ê—… —ﬁ„ :" & noteseialDuplicates
                                       Else
                                               MsgBox " Duplicate Manual No in Invoce: " & noteseialDuplicates
                                        End If
                                        'DCDocTypes.SetFocus
                   Exit Sub
                  End If
                  
        End If
        
            If XPChkPayType(2).value = Checked And Me.TxtModFlg.Text <> "R" Then
                Me.CmdCheque.Enabled = True
            Else
                Me.CmdCheque.Enabled = False
                Me.lbl(18).Caption = 0
                Me.lbl(19).Caption = 0
                Me.FgCheques.rows = Me.FgCheques.FixedRows
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        lbl(22).Enabled = True
        lbl(45).Enabled = True
    Else
        XPTxtTaxValue.Text = ""
        XPTxtTaxValue.Enabled = False
        lbl(22).Enabled = False
        lbl(45).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    CurrentVoucherNo = ""
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    DateChanged = True
End Sub

Private Sub XPTab301_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub printing()
    On Error GoTo ErrTrap

    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)
   
    If ShowType = True Then
        If Not XPTxtBillID.Text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowBuyData XPTxtBillID.Text, 1, True, Round(LblTotal.Caption * val(txt_Currency_rate), 2), TxtManualNO.Text, Me.Dccurrency.Text, val(Me.dcBranch.BoundText)
           
            
        End If

    Else

        If Not XPTxtBillID.Text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowBuyDataShort XPTxtBillID.Text, val(Me.dcBranch.BoundText)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Function AvailableDeal() As Boolean
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long

    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))

            '        If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) <> "" Then
            '            If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = True Then
            If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If

                '            End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

                    With FrmAlarm
                        .Tag = "x"
                        .DealingForm = PurchaseTransaction
                        .show vbModal
                     End With

                    AvailableDeal = False
                    Exit Function
                    '                End If
                    RsTemp.Close
                Else
                    Set RsTemp = New ADODB.Recordset
                    LngItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.Text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("QTY").value) < val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then

                            With FrmAlarm
                                .DealingForm = PurchaseTransaction
                                .show vbModal
                            End With

                            AvailableDeal = False
                            Exit Function
                        End If
                    End If

                    RsTemp.Close
                End If
            End If

            RsSalle.Close
        End If

    Next RowNum

    AvailableDeal = True
    Exit Function
ErrTrap:
End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String

    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "" Then Exit Sub
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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub
Function FillGridWithDataSalesPayment() As Boolean

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = " SELECT     TOP 100 PERCENT dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, "
    My_SQL = My_SQL & "                   dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code, dbo.TblSalesPayment.[Value],"
    My_SQL = My_SQL & "                   dbo.TblSalesPayment.CardNo, dbo.TblSalesPayment.PaymentID , dbo.TblSalesPayment.TransID,dbo.TblSalesPayment.MaxValue"
    My_SQL = My_SQL & "      FROM         dbo.TblPaymentType RIGHT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.TblSalesPayment ON dbo.TblPaymentType.PaymentID = dbo.TblSalesPayment.PaymentID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
    My_SQL = My_SQL & "     Where (dbo.TblSalesPayment.TransID = " & val(XPTxtBillID.Text) & ")"
    My_SQL = My_SQL & "   ORDER BY dbo.TblPaymentType.PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid22
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
        FillGridWithDataSalesPayment = True
            .rows = rs.RecordCount + 2
            rs.MoveFirst
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "‰ﬁœÌ", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "Cash", rs.Fields("PaymentNamee").value)
               End If
               .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs.Fields("Value").value), "", rs.Fields("Value").value)
               .TextMatrix(i, .ColIndex("CardNo")) = IIf(IsNull(rs.Fields("CardNo").value), "", rs.Fields("CardNo").value)
               .TextMatrix(i, .ColIndex("MaxValue")) = IIf(IsNull(rs.Fields("MaxValue").value), 0, rs.Fields("MaxValue").value)
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), "", rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(rs.Fields("BankId").value), "", rs.Fields("BankId").value)
            
            .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(rs.Fields("Accountcom").value), "", rs.Fields("Accountcom").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(rs.Fields("commision").value), "", rs.Fields("commision").value)
           .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs.Fields("bankAccount_Code").value), "", rs.Fields("bankAccount_Code").value)
            
                rs.MoveNext
            Next

            rs.Close
            Else
            FillGridWithDataSalesPayment = False
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Function
Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap
FramePay.Visible = False
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If CboPayMentType.ListIndex = 0 Then
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPChkPayType(0).value = Checked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).Text = XPTxtSum.Text
            XPTxtValue(1).Text = 0
            '        DBCboClientName.Enabled = False
'            DBCboClientName.Text = ""
            
           '
            DcboBox.Enabled = True
            DcboBox.Visible = True
            DcboBankName.Text = ""
            DcboBankName.Visible = False
            lbl(2).Caption = "«·’‰œÊﬁ"
        ElseIf CboPayMentType.ListIndex = 1 Then
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).Text = 0
            XPTxtValue(1).Text = XPTxtSum.Text
            '         DBCboClientName.Enabled = True
          '
            DcboBox.Enabled = False
            DcboBankName.Enabled = False
            
            DcboBox.Text = ""
      '
        ElseIf ((CboPayMentType.ListIndex = 2 And SystemOptions.AllowPurchasesMultyPayed = False) Or (SystemOptions.AllowPurchasesMultyPayed = True And CboPayMentType.ListIndex = 3)) Then
                    
            DcboBox.Visible = False
            DcboBox.Text = ""
            DcboBankName.Visible = True
            DcboBankName.Enabled = True
            lbl(2).Caption = "«·»‰ﬂ"
      Else
    If Me.TxtModFlg.Text <> "R" Then
     If Me.TxtModFlg.Text = "N" Then
     If val(LblTotal.Caption) > 0 Then
     FramePay.Visible = True
     FillGridWithData222
     LBLPayVal.Caption = 0
LBLPayVal.Caption = val(LblTotal.Caption)
TxtNetValue2.Text = val(LblTotal.Caption)
    With Grid22
          .TextMatrix(.Row, .ColIndex("Value")) = 0
    End With
     ReLineGrid2
     End If
     Else
      FramePay.Visible = True
    If FillGridWithDataSalesPayment() = True Then
     LBLPayVal.Caption = val(LblTotal.Caption)
     TxtNetValue2.Text = val(LblTotal.Caption)
     ReLineGrid2
     Else
     '''/////////////
          If val(LblTotal.Caption) > 0 Then
     FramePay.Visible = True
     FillGridWithData222
     LBLPayVal.Caption = 0
LBLPayVal.Caption = val(LblTotal.Caption)
TxtNetValue2.Text = val(LblTotal.Caption)
    With Grid22
          .TextMatrix(.Row, .ColIndex("Value")) = 0
    End With
     ReLineGrid2
     End If
     End If
     ''///////////
     End If
           XPChkPayType(0).Enabled = False
        XPChkPayType(1).Enabled = False
        XPChkPayType(2).Enabled = False
        XPChkPayType(0).value = Checked
        XPChkPayType(1).value = Unchecked
        XPChkPayType(2).value = Unchecked
        XPTxtValue(0).Text = XPTxtSum.Text
        XPTxtValue(1).Text = ""
        DcboBox.Enabled = True
        Frame1.Visible = True
      '  DCPaymentNet.Enabled = True
        End If
    End If
End If
    Exit Sub
ErrTrap:

End Sub
Private Sub ReLineGrid2()
    On Error Resume Next
    Dim i As Integer
    Dim IntCounter As Integer
    Dim totalPayed As Double
    Dim visapayed As Double
 totalPayed = 0
 visapayed = 0
  With Grid22
        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("value")) <> "" Then
               ' IntCounter = IntCounter + 1
                totalPayed = totalPayed + .TextMatrix(i, .ColIndex("value"))
                If totalPayed > val(Me.TxtNetValue2.Text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬂ»— „‰ «·«Ã„«·Ì"
                Else
                 MsgBox "ERROR Incorrect Value" & CHR(13)
                End If
                .TextMatrix(i, .ColIndex("value")) = 0
                Exit Sub
                End If
            End If

        Next i

    End With
  TxtPayedValue2.Text = totalPayed
    TxtRemainValue2.Text = val(Me.TxtPayedValue2.Text) - val(Me.TxtNetValue2.Text)
End Sub
Public Sub FillGridWithData222()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "SELECT     dbo.TblPaymentType.PaymentID, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, "
My_SQL = My_SQL & "  dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code ,dbo.TblPaymentType.MaxValue"
My_SQL = My_SQL & " FROM         dbo.TblPaymentType LEFT OUTER JOIN"
My_SQL = My_SQL & " dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
My_SQL = My_SQL & " where (dbo.TblPaymentType.TypTran=2  or dbo.TblPaymentType.TypTran=1) "
If SystemOptions.LinkUsersWithPayment = True Then
My_SQL = My_SQL & " and dbo.TblPaymentType.PaymentID in (SELECT     PaynetID"
My_SQL = My_SQL & " From dbo.TblPaymentUser"
My_SQL = My_SQL & " Where (UserID = " & user_id & "))"
End If
My_SQL = My_SQL & " order by PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid22
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 2
            rs.MoveFirst
      If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(1, .ColIndex("PaymentName")) = " ‰ﬁœÌ"
               Else
               .TextMatrix(1, .ColIndex("PaymentName")) = " Cash"
               End If
               
                .TextMatrix(1, .ColIndex("PaymentID")) = 0
           
           
            For i = 2 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "", rs.Fields("PaymentNamee").value)
               End If
               
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), "", rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(rs.Fields("BankId").value), "", rs.Fields("BankId").value)
            
            .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(rs.Fields("Accountcom").value), "", rs.Fields("Accountcom").value)
            .TextMatrix(i, .ColIndex("MaxValue")) = IIf(IsNull(rs.Fields("MaxValue").value), 0, rs.Fields("MaxValue").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(rs.Fields("commision").value), "", rs.Fields("commision").value)
           .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs.Fields("bankAccount_Code").value), "", rs.Fields("bankAccount_Code").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub XPTxtDiscountVal_Change()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.Text, 0)
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap

    If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 2 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).Text = XPTxtSum.Text
    End If
If ChAddToTotal.value = vbChecked Then
    Me.LblTotal.Caption = val(XPTxtSum.Text) + val(txtAddValue.Caption) + val(TxtValueAdded.Text)
    Else
    Me.LblTotal.Caption = val(XPTxtSum.Text) + val(TxtValueAdded.Text)
    End If
    Exit Sub
ErrTrap:
End Sub

Public Function RepeatSerial(StrSerial As String, _
                             IntTransType As Integer, _
                             Optional IntTransID As Long = 0, _
                             Optional LngCusID As Long = 0) As Boolean

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    RepeatSerial = False

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryTransactionsTotal.Transaction_ID," & "QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Serial, " & "QryTransactionsTotal.Transaction_Date , QryTransactionsTotal.Transaction_Type," & "dbo.TblCustemers.CusName"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN " & "dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + " Where QryTransactionsTotal.Transaction_Serial ='" & StrSerial & "'"
        StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Type=" & BillType & ""

        If LngCusID <> 0 Then
            StrSQL = StrSQL + " AND dbo.TblCustemers.CusID=" & LngCusID & ""
        End If

        If IntTransID <> 0 Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_ID <> " & IntTransID & ""
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            Msg = "—ﬁ„ «·›« Ê—… „ÊÃÊœ „”»ﬁ« ›Ï «·»—‰«„Ã øø" & CHR(13)
            Msg = Msg + "„⁄·Ê„«  ⁄‰ «·›« Ê—… «·„”Ã·…:-" & CHR(13)
        
            Msg = Msg + "—ﬁ„ «·›« Ê—… ›Ï «·»—‰«„Ã:" & rs("Transaction_ID").value & CHR(13)
            Msg = Msg + "„”·”· «·›« Ê—…:" & rs("Transaction_Serial").value & CHR(13)
            Msg = Msg + " «—ÌŒ  ”ÃÌ· «·›« Ê—…:" & rs("Transaction_Date").value & CHR(13)
            Msg = Msg + "«”„ «·⁄„Ì· «Ê «·„Ê—œ:" & rs("CusName").value & CHR(13)
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            RepeatSerial = True
        End If

        rs.Close
        Set rs = Nothing

    End If

End Function

Private Sub SetDefaults()
    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

    If SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer Then
        XPDtbBill.value = Date
    ElseIf SystemOptions.SysPurDateTakeType = InvDateFromServerComputer Then
        StrTemp = "select Getdate() as ServerDate"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not IsNull(RsTemp("ServerDate").value) Then
                XPDtbBill.value = Format(RsTemp("ServerDate").value, "yyyy/M/d")
            End If

            'XPDtbBill.Value = IIf(IsNull(RsTemp("ServerDate").Value), Date, (RsTemp("ServerDate").Value))
        End If

        RsTemp.Close
        Set RsTemp = Nothing
    End If

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast

        If SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate Then
            XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), Date, (rs("Transaction_Date").value))
        End If
    End If

    Me.DcboBox.BoundText = 1
End Sub
Private Sub SaveFlow()
  If Text1.Text <> "" Then
        Cn.Execute "update Transactions set nots =' " & TxtTransSerial.Text & "' where Transaction_Type= 20 and Transaction_Serial=" & Text1.Text & ""
    End If

    Cn.Execute "update Transactions set NoteSerial =' " & Trim(Me.TxtNoteSerial.Text) & "' where Transaction_ID=" & val(Me.XPTxtBillID.Text)

    'Õ›Ÿ «·„’—Êﬁ«  «· ﬁœÌ—Ì…
    Dim FactoryExpenses As New ADODB.Recordset

    If Me.TxtModFlg.Text = "E" Then
        
        Cn.Execute "Delete TblProductOrderFactoryexpenses where Transaction_ID=" & val(XPTxtBillID.Text)
    End If
    Dim StrSQL As String
    StrSQL = "Select * from TblProductOrderFactoryexpenses where Transaction_ID=" & val(XPTxtBillID.Text)
    FactoryExpenses.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim RowNum As Long
    For RowNum = 1 To Fg_Journal.rows - 2

        If Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName")) <> "" Then
            FactoryExpenses.AddNew
            FactoryExpenses("Transaction_ID").value = val(XPTxtBillID.Text)
            FactoryExpenses("Accountcode2").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Accountcode2"))
            FactoryExpenses("Accountcode").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Accountcode"))
            FactoryExpenses("AccountName").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName"))
            FactoryExpenses("value").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("value")))
            FactoryExpenses("Price").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Price")))
            
            FactoryExpenses("CurrRow").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("CurrRow")))
            FactoryExpenses("FlgVat").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("FlgVat")))
            FactoryExpenses("Vatyo").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Vatyo")))
            FactoryExpenses("Vat").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Vat")))
            FactoryExpenses("ForcedFlg").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("ForcedFlg")))
            FactoryExpenses("PriceTotal").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("PriceTotal")))

            
       
                 
            If Fg_Journal.cell(flexcpChecked, RowNum, Fg_Journal.ColIndex("ChSameCurrncy")) = flexChecked Then
                FactoryExpenses("ChSameCurrncy").value = 1
            Else
                FactoryExpenses("ChSameCurrncy").value = 0
            End If
            
            FactoryExpenses("des").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("des"))
            FactoryExpenses.update
        End If
         
    Next RowNum
End Sub
