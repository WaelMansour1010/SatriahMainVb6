VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmInpout2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”‰œ «” ·«„"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15870
   HelpContextID   =   100
   Icon            =   "FrmInpout2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmInpout2.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   15870
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      TabIndex        =   50
      Top             =   9060
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
      Height          =   9060
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15870
      _cx             =   27993
      _cy             =   15981
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
      GridRows        =   6
      GridCols        =   6
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmInpout2.frx":2B2C
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   4860
         Left            =   15
         TabIndex        =   1
         Top             =   2940
         Width           =   15825
         _cx             =   27914
         _cy             =   8572
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
         BackColor       =   14871017
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|«·√Ê—«ﬁ «·„«·Ì…|„·«ÕŸ«  ⁄·Ï «·›« Ê—…|”‰œ«  «·’—›|«·ÿ·»Ì« |›Ê« Ì— „«·Ì…|„’—Ê›«   ﬁœÌ—ÌÂ|«·„—›ﬁ« |«·‘Õ‰"
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
         Picture(0)      =   "FrmInpout2.frx":2BD2
         Picture(1)      =   "FrmInpout2.frx":2F6C
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   4395
            Left            =   18270
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
               TabIndex        =   211
               TabStop         =   0   'False
               Top             =   0
               Width           =   15750
               _cx             =   27781
               _cy             =   8467
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
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   480
                  TabIndex        =   223
                  Top             =   840
                  Width           =   1410
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   219
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
                  TabIndex        =   218
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
                  TabIndex        =   217
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
                  TabIndex        =   216
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
                  TabIndex        =   212
                  Top             =   1800
                  Width           =   4695
                  Begin VB.TextBox TxtNoteSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   1560
                     RightToLeft     =   -1  'True
                     TabIndex        =   213
                     Top             =   600
                     Width           =   2625
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     CausesValidation=   0   'False
                     Height          =   375
                     Index           =   10
                     Left            =   120
                     TabIndex        =   214
                     Top             =   600
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
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
                     TabIndex        =   215
                     Top             =   240
                     Width           =   2175
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid GRID1 
                  Height          =   2085
                  Left            =   5160
                  TabIndex        =   220
                  Tag             =   "1"
                  Top             =   840
                  Width           =   9255
                  _cx             =   16325
                  _cy             =   3678
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
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmInpout2.frx":3306
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
                  TabIndex        =   221
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmInpout2.frx":3453
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
                  TabIndex        =   230
                  Top             =   3360
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   661
                  ButtonPositionImage=   1
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
                  TabIndex        =   224
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
                  TabIndex        =   222
                  Top             =   120
                  Width           =   2160
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   4395
            Left            =   17970
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
               TabIndex        =   201
               TabStop         =   0   'False
               Top             =   0
               Width           =   15750
               _cx             =   27781
               _cy             =   8467
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
               Begin VB.TextBox TXTFactoryExpenses 
                  Alignment       =   2  'Center
                  Height          =   405
                  Left            =   7920
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   202
                  Top             =   2880
                  Width           =   1215
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
                  Height          =   2340
                  Left            =   1800
                  TabIndex        =   203
                  Top             =   480
                  Width           =   12600
                  _cx             =   22225
                  _cy             =   4128
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   0   'False
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
                  Rows            =   1
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmInpout2.frx":3546
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
                     TabIndex        =   204
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
                        TabIndex        =   205
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
                        TabIndex        =   206
                        Top             =   0
                        Width           =   2445
                     End
                  End
                  Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   207
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
                        Name            =   "MS Sans Serif"
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
                        Name            =   "MS Sans Serif"
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
                     Picture         =   "FrmInpout2.frx":36A6
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
                  TabIndex        =   208
                  Top             =   2880
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmInpout2.frx":3C40
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì  «·„’«—Ì› «· ﬁœÌ—ÌÂ"
                  Height          =   375
                  Left            =   9120
                  RightToLeft     =   -1  'True
                  TabIndex        =   210
                  Top             =   3000
                  Width           =   2055
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Œ Ì«— «·„’—Ê›«  «· ﬁœÌ—ÌÂ"
                  Height          =   255
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   209
                  Top             =   120
                  Width           =   3855
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4395
            Index           =   0
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
            GridRows        =   3
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmInpout2.frx":41DA
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   3255
               Left            =   30
               TabIndex        =   3
               Top             =   735
               Width           =   15675
               _cx             =   27649
               _cy             =   5741
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
               Rows            =   2
               Cols            =   28
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmInpout2.frx":422A
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
               Left            =   30
               TabIndex        =   4
               Top             =   4005
               Width           =   15675
               _ExtentX        =   27649
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   4
               Left            =   30
               TabIndex        =   264
               TabStop         =   0   'False
               Top             =   30
               Width           =   15675
               _cx             =   27649
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
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   735
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   268
                  Top             =   315
                  Width           =   1440
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   3975
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   267
                  Top             =   315
                  Width           =   3480
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2175
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   266
                  Top             =   315
                  Width           =   1800
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   7515
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   265
                  Top             =   315
                  Width           =   2460
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   10050
                  TabIndex        =   269
                  Top             =   315
                  Width           =   3585
                  _ExtentX        =   6324
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   13680
                  TabIndex        =   270
                  Top             =   315
                  Width           =   1950
                  _ExtentX        =   3440
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   345
                  Left            =   45
                  TabIndex        =   271
                  Top             =   315
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
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
                  ButtonImage     =   "FrmInpout2.frx":46C6
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬂ·›…"
                  Height          =   270
                  Index           =   26
                  Left            =   735
                  RightToLeft     =   -1  'True
                  TabIndex        =   277
                  Top             =   30
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   270
                  Index           =   27
                  Left            =   2175
                  RightToLeft     =   -1  'True
                  TabIndex        =   276
                  Top             =   30
                  Width           =   1800
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   270
                  Index           =   28
                  Left            =   3975
                  RightToLeft     =   -1  'True
                  TabIndex        =   275
                  Top             =   30
                  Width           =   3480
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   270
                  Index           =   29
                  Left            =   7515
                  RightToLeft     =   -1  'True
                  TabIndex        =   274
                  Top             =   30
                  Width           =   2460
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   270
                  Index           =   30
                  Left            =   10050
                  RightToLeft     =   -1  'True
                  TabIndex        =   273
                  Top             =   30
                  Width           =   3585
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   270
                  Index           =   31
                  Left            =   13680
                  RightToLeft     =   -1  'True
                  TabIndex        =   272
                  Top             =   30
                  Width           =   1950
               End
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   360
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   4005
               Width           =   15675
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4395
            Index           =   2
            Left            =   16470
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
            _GridInfo       =   $"FrmInpout2.frx":4A60
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1950
               Index           =   10
               Left            =   0
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   2445
               Width           =   15735
               _cx             =   27755
               _cy             =   3440
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
               _GridInfo       =   $"FrmInpout2.frx":4AD1
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   375
                  Index           =   14
                  Left            =   15
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   661
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
                     Height          =   300
                     Index           =   2
                     Left            =   12645
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   75
                     Width           =   1110
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   300
                     Left            =   6390
                     TabIndex        =   10
                     Top             =   75
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   529
                     Caption         =   " ”ÃÌ· «·‘Ìﬂ« "
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
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin MSDataListLib.DataCombo dcbanks 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   63
                     Top             =   0
                     Width           =   2370
                     _ExtentX        =   4180
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·»‰ﬂ"
                     Height          =   285
                     Left            =   2370
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   135
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   300
                     Index           =   18
                     Left            =   7095
                     RightToLeft     =   -1  'True
                     TabIndex        =   14
                     Top             =   75
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
                     Height          =   300
                     Index           =   16
                     Left            =   7920
                     RightToLeft     =   -1  'True
                     TabIndex        =   13
                     Top             =   75
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
                     Height          =   300
                     Index           =   17
                     Left            =   10980
                     RightToLeft     =   -1  'True
                     TabIndex        =   12
                     Top             =   75
                     Visible         =   0   'False
                     Width           =   1110
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   300
                     Index           =   19
                     Left            =   10005
                     RightToLeft     =   -1  'True
                     TabIndex        =   11
                     Top             =   75
                     Width           =   975
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   1140
                  Left            =   15
                  TabIndex        =   125
                  Top             =   405
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   2011
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
                  FormatString    =   $"FrmInpout2.frx":4B6E
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
               Height          =   2085
               Index           =   7
               Left            =   0
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   360
               Width           =   15735
               _cx             =   27755
               _cy             =   3678
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
               _GridInfo       =   $"FrmInpout2.frx":4CA2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   405
                  Index           =   12
                  Left            =   15
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   714
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
                     Height          =   675
                     Index           =   1
                     Left            =   12645
                     RightToLeft     =   -1  'True
                     TabIndex        =   20
                     Top             =   -195
                     Width           =   975
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   600
                     Index           =   1
                     Left            =   10290
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   19
                     Top             =   45
                     Width           =   1245
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   720
                     Index           =   1
                     Left            =   7365
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   18
                     Top             =   45
                     Width           =   1815
                  End
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   195
                     Left            =   2775
                     RightToLeft     =   -1  'True
                     TabIndex        =   17
                     Top             =   75
                     Width           =   1260
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   300
                     Left            =   135
                     TabIndex        =   21
                     Top             =   75
                     Width           =   1665
                     _ExtentX        =   2937
                     _ExtentY        =   529
                     ButtonPositionImage=   1
                     Caption         =   "Õ”«» «·√ﬁ”«ÿ"
                     BackColor       =   14871017
                     Enabled         =   0   'False
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmInpout2.frx":4D3F
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
                     Height          =   930
                     Index           =   21
                     Left            =   4440
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   1005
                     Width           =   975
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   615
                     Index           =   14
                     Left            =   9315
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   75
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   660
                     Index           =   15
                     Left            =   11820
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   75
                     Width           =   405
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   810
                  Left            =   15
                  TabIndex        =   100
                  Top             =   435
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   1429
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
                  FormatString    =   $"FrmInpout2.frx":50D9
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
                  Height          =   195
                  Index           =   13
                  Left            =   15
                  TabIndex        =   109
                  TabStop         =   0   'False
                  Top             =   1875
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   344
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
                     Height          =   135
                     Index           =   37
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   124
                     Top             =   30
                     Width           =   1665
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   135
                     Left            =   135
                     RightToLeft     =   -1  'True
                     TabIndex        =   123
                     Top             =   30
                     Width           =   285
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   135
                     Left            =   3195
                     RightToLeft     =   -1  'True
                     TabIndex        =   122
                     Top             =   30
                     Width           =   420
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   135
                     Left            =   12225
                     RightToLeft     =   -1  'True
                     TabIndex        =   121
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
                     Height          =   135
                     Index           =   35
                     Left            =   12645
                     RightToLeft     =   -1  'True
                     TabIndex        =   120
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
                     Height          =   135
                     Index           =   34
                     Left            =   14460
                     RightToLeft     =   -1  'True
                     TabIndex        =   119
                     Top             =   30
                     Width           =   1110
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   135
                     Left            =   13335
                     RightToLeft     =   -1  'True
                     TabIndex        =   118
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
                     Height          =   135
                     Index           =   36
                     Left            =   10695
                     RightToLeft     =   -1  'True
                     TabIndex        =   117
                     Top             =   30
                     Width           =   1395
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   135
                     Left            =   9735
                     RightToLeft     =   -1  'True
                     TabIndex        =   116
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
                     Height          =   135
                     Index           =   38
                     Left            =   8205
                     RightToLeft     =   -1  'True
                     TabIndex        =   115
                     Top             =   30
                     Width           =   1530
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   135
                     Left            =   7650
                     RightToLeft     =   -1  'True
                     TabIndex        =   114
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
                     Height          =   135
                     Index           =   40
                     Left            =   6525
                     RightToLeft     =   -1  'True
                     TabIndex        =   113
                     Top             =   30
                     Width           =   1125
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   135
                     Left            =   5145
                     RightToLeft     =   -1  'True
                     TabIndex        =   112
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
                     Height          =   135
                     Index           =   42
                     Left            =   3615
                     RightToLeft     =   -1  'True
                     TabIndex        =   111
                     Top             =   30
                     Width           =   1530
                  End
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   135
                     Left            =   2085
                     RightToLeft     =   -1  'True
                     TabIndex        =   110
                     Top             =   30
                     Width           =   1110
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   360
               Index           =   11
               Left            =   0
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   0
               Width           =   15735
               _cx             =   27755
               _cy             =   635
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
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   1125
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   0
                  Width           =   150
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   810
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   60
                  Width           =   210
               End
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   345
                  Index           =   0
                  Left            =   1410
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   0
                  Width           =   75
               End
               Begin MSDataListLib.DataCombo DcboCurrency 
                  Height          =   315
                  Left            =   270
                  TabIndex        =   102
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   556
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
                  Left            =   540
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   345
                  Index           =   13
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   90
                  Width           =   90
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   345
                  Index           =   12
                  Left            =   1005
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   90
                  Width           =   105
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4395
            Index           =   15
            Left            =   16770
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
            _GridInfo       =   $"FrmInpout2.frx":51AA
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   615
               Index           =   18
               Left            =   15
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1020
               Width           =   15705
               _cx             =   27702
               _cy             =   1085
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
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   195
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   420
                  Left            =   975
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   49
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   390
                  Index           =   43
                  Left            =   1095
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   120
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
                  Left            =   870
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   45
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   615
               Index           =   17
               Left            =   15
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   1020
               Visible         =   0   'False
               Width           =   15705
               _cx             =   27702
               _cy             =   1085
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
                  Left            =   1290
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   45
                  Width           =   165
               End
               Begin VB.TextBox TxtTaxStampValue 
                  Alignment       =   1  'Right Justify
                  Height          =   225
                  Left            =   975
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   33
                  Left            =   255
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   195
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   195
                  Index           =   41
                  Left            =   1095
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   45
                  Width           =   120
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
                  Left            =   870
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   45
                  Width           =   30
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   660
               Index           =   16
               Left            =   15
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   345
               Visible         =   0   'False
               Width           =   15705
               _cx             =   27702
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
                  Left            =   1215
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.TextBox TxtTaxAddValue 
                  Alignment       =   1  'Right Justify
                  Height          =   480
                  Left            =   975
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   75
                  Width           =   150
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   32
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
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
                  Left            =   1125
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   90
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
                  Left            =   810
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   90
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   315
               Index           =   8
               Left            =   15
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   15
               Visible         =   0   'False
               Width           =   15705
               _cx             =   27702
               _cy             =   556
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
                  Left            =   975
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   45
                  Width           =   150
               End
               Begin VB.CheckBox XPChkTAX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   150
                  Left            =   1215
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   150
                  Index           =   25
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
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
                  Left            =   1065
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   135
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
                  Left            =   780
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   135
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
               Height          =   615
               Index           =   44
               Left            =   15
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1020
               Width           =   15705
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   4395
            Left            =   17070
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
               TabIndex        =   180
               Top             =   0
               Width           =   15750
               Begin VB.TextBox TxtMode 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   4680
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   248
                  Text            =   "0"
                  Top             =   3480
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.TextBox TxtTotals 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   840
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   247
                  Text            =   "0"
                  Top             =   3720
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.CommandButton Command6 
                  Caption         =   "Command6"
                  Height          =   375
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  TabIndex        =   184
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
                  TabIndex        =   183
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
                  TabIndex        =   182
                  Top             =   2880
                  Width           =   1890
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "⁄—÷ «·„’—Ê›« "
                  Height          =   480
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   181
                  Top             =   3240
                  Width           =   2220
               End
               Begin VSFlex8UCtl.VSFlexGrid Grid 
                  Height          =   2325
                  Left            =   240
                  TabIndex        =   185
                  Tag             =   "1"
                  Top             =   480
                  Width           =   15255
                  _cx             =   26908
                  _cy             =   4101
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
                  Rows            =   50
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmInpout2.frx":5222
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
                  Caption         =   "«Ã„«·Ì «·„’—Ê›« "
                  Height          =   285
                  Index           =   60
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   188
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
                  TabIndex        =   187
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
                  TabIndex        =   186
                  Top             =   3000
                  Width           =   1920
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   4395
            Left            =   17370
            TabIndex        =   176
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
               TabIndex        =   189
               Top             =   0
               Width           =   15750
               Begin VB.CommandButton Command7 
                  Caption         =   "Command7"
                  Height          =   195
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   191
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "⁄—÷ ÿ·»«  «·‘—«¡"
                  Height          =   480
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   190
                  Top             =   3000
                  Width           =   2010
               End
               Begin VSFlex8UCtl.VSFlexGrid GRID2 
                  Height          =   2205
                  Left            =   5040
                  TabIndex        =   192
                  Tag             =   "1"
                  Top             =   600
                  Width           =   7695
                  _cx             =   13573
                  _cy             =   3889
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmInpout2.frx":53F8
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
                  TabIndex        =   193
                  Top             =   240
                  Width           =   4440
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   4395
            Left            =   17670
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
               TabIndex        =   194
               Top             =   0
               Width           =   15750
               Begin VB.CommandButton Command5 
                  Caption         =   " Œ’Ì’"
                  Height          =   480
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   197
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   2220
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "⁄—÷ «·›Ê« Ì— «·„«·Ì…"
                  Height          =   480
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   196
                  Top             =   2880
                  Width           =   2220
               End
               Begin VB.TextBox txt_total_bill 
                  Height          =   405
                  Left            =   10200
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   195
                  Top             =   2880
                  Width           =   1770
               End
               Begin VSFlex8UCtl.VSFlexGrid grid4 
                  Height          =   2325
                  Left            =   240
                  TabIndex        =   198
                  Tag             =   "1"
                  Top             =   480
                  Width           =   14055
                  _cx             =   24791
                  _cy             =   4101
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
                  Rows            =   50
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmInpout2.frx":5500
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
                  Caption         =   "«·›Ê« Ì— «·„«·ÌÂ"
                  Height          =   285
                  Index           =   64
                  Left            =   10800
                  RightToLeft     =   -1  'True
                  TabIndex        =   200
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
                  TabIndex        =   199
                  Top             =   3000
                  Width           =   2040
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4395
            Index           =   20
            Left            =   18570
            TabIndex        =   249
            TabStop         =   0   'False
            Top             =   45
            Width           =   15735
            _cx             =   27755
            _cy             =   7752
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
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   12585
               RightToLeft     =   -1  'True
               TabIndex        =   254
               Top             =   1335
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   12585
               RightToLeft     =   -1  'True
               TabIndex        =   253
               Top             =   780
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   12585
               RightToLeft     =   -1  'True
               TabIndex        =   252
               Top             =   225
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox Text9 
               Alignment       =   1  'Right Justify
               Height          =   0
               Left            =   135
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   251
               Top             =   360
               Width           =   0
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì… Œœ„…"
               Height          =   0
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   250
               Top             =   495
               Width           =   0
            End
            Begin MSDataListLib.DataCombo DCboStoreName2 
               Height          =   315
               Left            =   8490
               TabIndex        =   255
               Top             =   255
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "7"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCCar 
               Height          =   315
               Left            =   8490
               TabIndex        =   256
               Top             =   765
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "7"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCDriver 
               Height          =   315
               Left            =   8490
               TabIndex        =   257
               Top             =   1290
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "7"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”«∆ﬁ"
               Height          =   210
               Index           =   83
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   263
               Top             =   1305
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„⁄œÂ/«·”Ì«—…"
               Height          =   210
               Index           =   82
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   780
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ «·„Œ“‰"
               Height          =   210
               Index           =   81
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   270
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   1095
               Index           =   80
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   360
               Width           =   135
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
               Height          =   1095
               Index           =   79
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Top             =   360
               Width           =   135
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   1455
               Index           =   78
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   258
               Top             =   360
               Width           =   135
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   600
         Index           =   6
         Left            =   15
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   15
         Width           =   15840
         _cx             =   27940
         _cy             =   1058
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
         Caption         =   "”‰œ «” ·«„ ÃœÌœ"
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
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2955
            RightToLeft     =   -1  'True
            TabIndex        =   133
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   120
            Visible         =   0   'False
            Width           =   630
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   390
            Left            =   5055
            TabIndex        =   56
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpout2.frx":56C4
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2700
            TabIndex        =   57
            Top             =   105
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpout2.frx":5A5E
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
            TabIndex        =   58
            Top             =   105
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpout2.frx":5DF8
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
            TabIndex        =   59
            Top             =   105
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpout2.frx":6192
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
            TabIndex        =   60
            Top             =   105
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpout2.frx":652C
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
            Left            =   6315
            TabIndex        =   61
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpout2.frx":68C6
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   480
            Left            =   11295
            TabIndex        =   62
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpout2.frx":6E60
            ButtonImageHover=   "FrmInpout2.frx":7B3A
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   11880
            Picture         =   "FrmInpout2.frx":8814
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
            TabIndex        =   134
            Top             =   120
            Width           =   6855
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2295
         Index           =   5
         Left            =   15
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   630
         Width           =   15825
         _cx             =   27914
         _cy             =   4048
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
         Begin VB.TextBox TxtReciveOrderO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4245
            RightToLeft     =   -1  'True
            TabIndex        =   244
            Top             =   720
            Width           =   1380
         End
         Begin VB.TextBox TxtPolicyNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   243
            Top             =   720
            Width           =   1230
         End
         Begin VB.TextBox txtEmpCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2655
            RightToLeft     =   -1  'True
            TabIndex        =   238
            Top             =   840
            Width           =   630
         End
         Begin VB.CheckBox ChkCompsBill 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›« Ê—… „Ã„⁄Â"
            Height          =   255
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   237
            Top             =   1920
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox TxtPhone 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   10065
            TabIndex        =   232
            Top             =   1080
            Width           =   1650
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   13275
            TabIndex        =   228
            Top             =   1440
            Width           =   1305
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   735
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   226
            Top             =   1560
            Width           =   3855
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   660
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   227
               Top             =   240
               Width           =   3600
            End
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12345
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   1125
            Width           =   2265
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   13275
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   780
            Width           =   1305
         End
         Begin VB.TextBox TxtLCNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   720
            Locked          =   -1  'True
            MaxLength       =   55
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   390
            Width           =   1365
         End
         Begin VB.TextBox txtManualNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11145
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   0
            Width           =   1140
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmInpout2.frx":C47C
            Left            =   11115
            List            =   "FrmInpout2.frx":C47E
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   360
            Width           =   1185
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   5595
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   -180
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13275
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   0
            Width           =   1305
         End
         Begin VB.TextBox TXT_order_no 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9180
            MaxLength       =   55
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   390
            Width           =   1230
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   6960
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   390
            Width           =   1230
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   6960
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   1485
            Width           =   1230
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4200
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   1485
            Width           =   1365
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   15795
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1725
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   10335
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   -240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   13275
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   1905
            Width           =   1305
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   13380
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   855
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
            Caption         =   " ÕÊÌ· «·Ï  «–‰ «÷«›… "
            Height          =   255
            Left            =   -345
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   1605
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   210
            Left            =   675
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   1605
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox txt_Currency_rate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   285
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Text            =   "1"
            Top             =   15
            Width           =   960
         End
         Begin MSDataListLib.DataCombo DcCurrency 
            Height          =   315
            Left            =   1245
            TabIndex        =   68
            Top             =   0
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   9705
            TabIndex        =   77
            Top             =   780
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   9705
            TabIndex        =   78
            Top             =   1890
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   13275
            TabIndex        =   79
            Top             =   390
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102039553
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   285
            Left            =   14490
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   825
            Visible         =   0   'False
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmInpout2.frx":C480
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCproject 
            Height          =   315
            Left            =   4170
            TabIndex        =   93
            Top             =   1890
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTArrivalDate 
            Height          =   315
            Left            =   4245
            TabIndex        =   96
            Top             =   1140
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102039553
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   6930
            TabIndex        =   98
            Top             =   0
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   3000
            TabIndex        =   129
            Top             =   390
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   3000
            TabIndex        =   135
            Top             =   30
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   315
            Left            =   315
            TabIndex        =   139
            Top             =   390
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷"
            BackColor       =   12632256
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Height          =   315
            Left            =   6960
            TabIndex        =   140
            Top             =   1140
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102039553
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   9705
            TabIndex        =   229
            Top             =   1440
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   375
            Left            =   9585
            TabIndex        =   231
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   1080
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmInpout2.frx":C81A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Height          =   315
            Left            =   0
            TabIndex        =   239
            Top             =   840
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmpDepartments 
            Height          =   315
            Left            =   0
            TabIndex        =   240
            Top             =   1200
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ê·Ì’… «·‘Õ‰"
            Height          =   270
            Index           =   77
            Left            =   8145
            RightToLeft     =   -1  'True
            TabIndex        =   246
            Top             =   720
            Width           =   1350
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰œ  Ê’Ì· »÷«⁄Â"
            Height          =   390
            Index           =   76
            Left            =   5700
            RightToLeft     =   -1  'True
            TabIndex        =   245
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ÊŸ›"
            Height          =   240
            Index           =   75
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   242
            Top             =   840
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«œ«—… "
            Height          =   210
            Index           =   74
            Left            =   3225
            RightToLeft     =   -1  'True
            TabIndex        =   241
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ·Ì›Ê‰"
            Height          =   315
            Index           =   84
            Left            =   11625
            TabIndex        =   233
            Top             =   1185
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„‰œÊ»"
            Height          =   210
            Index           =   72
            Left            =   14625
            RightToLeft     =   -1  'True
            TabIndex        =   225
            Top             =   1530
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   195
            Index           =   71
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   900
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ «·‰ﬁœÌ"
            Height          =   195
            Index           =   70
            Left            =   14625
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   1155
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            Height          =   315
            Left            =   8385
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   1140
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«⁄ „«œ"
            Height          =   195
            Index           =   68
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   390
            Width           =   810
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5820
            TabIndex        =   136
            Top             =   30
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            Height          =   195
            Index           =   53
            Left            =   12225
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰œÊﬁ"
            Height          =   225
            Index           =   2
            Left            =   5775
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   390
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   270
            Index           =   66
            Left            =   10275
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   390
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   270
            Index           =   65
            Left            =   12480
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   390
            Width           =   675
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10545
            TabIndex        =   99
            Top             =   75
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  Ê’Ê· «·‘Õ‰Â"
            Height          =   315
            Index           =   56
            Left            =   5475
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1140
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‘—Ê⁄"
            Height          =   210
            Index           =   58
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1920
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·Œ’„"
            Height          =   315
            Index           =   11
            Left            =   5850
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1485
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·”œ«œ"
            Height          =   270
            Index           =   10
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   390
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   270
            Index           =   7
            Left            =   14565
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   390
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
            TabIndex        =   87
            Top             =   75
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   195
            Index           =   8
            Left            =   14640
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   15
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„Ê—œ"
            Height          =   195
            Index           =   6
            Left            =   14640
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   885
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   315
            Index           =   5
            Left            =   8385
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   1485
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„Œ“‰"
            Height          =   315
            Index           =   4
            Left            =   14520
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   1890
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   55
            Left            =   3945
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   1485
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
            TabIndex        =   81
            Top             =   1065
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   8055
         Width           =   15825
         _cx             =   27914
         _cy             =   767
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
            Height          =   390
            Left            =   5550
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   -90
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3585
            TabIndex        =   145
            Top             =   90
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„Ê·« "
            Height          =   255
            Index           =   73
            Left            =   11010
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   90
            Width           =   615
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
            Height          =   345
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   235
            Top             =   -240
            Visible         =   0   'False
            Width           =   1560
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
            Left            =   9585
            TabIndex        =   234
            Top             =   0
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   495
            Index           =   0
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   90
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   660
            Index           =   63
            Left            =   7935
            TabIndex        =   160
            Top             =   615
            Visible         =   0   'False
            Width           =   585
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
            Left            =   7725
            TabIndex        =   159
            Top             =   0
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   330
            Index           =   24
            Left            =   8820
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   120
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
            Left            =   11790
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   255
            Index           =   50
            Left            =   12840
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   90
            Width           =   600
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
            Height          =   405
            Left            =   13590
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   15
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   255
            Index           =   23
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   90
            Width           =   150
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
            Height          =   405
            Left            =   7245
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   255
            Index           =   1
            Left            =   6105
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   90
            Width           =   795
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   90
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   255
            Index           =   3
            Left            =   15165
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   120
            Width           =   570
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
            Height          =   405
            Left            =   13740
            TabIndex        =   148
            Top             =   0
            Visible         =   0   'False
            Width           =   1110
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
            Height          =   345
            Left            =   11790
            TabIndex        =   147
            Top             =   0
            Visible         =   0   'False
            Width           =   990
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
            Height          =   405
            Left            =   7335
            TabIndex        =   146
            Top             =   0
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   8505
         Width           =   15840
         _cx             =   27940
         _cy             =   953
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
            Height          =   540
            Index           =   0
            Left            =   14145
            TabIndex        =   163
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   953
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   1
            Left            =   12360
            TabIndex        =   164
            Top             =   0
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   953
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
            Height          =   540
            Index           =   2
            Left            =   10575
            TabIndex        =   165
            Top             =   0
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   953
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
            Height          =   540
            Index           =   3
            Left            =   8865
            TabIndex        =   166
            Top             =   0
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   953
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
            Height          =   540
            Index           =   4
            Left            =   7065
            TabIndex        =   167
            Top             =   0
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   953
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
            Height          =   540
            Index           =   5
            Left            =   5310
            TabIndex        =   168
            Top             =   0
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   953
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
            Height          =   540
            Index           =   6
            Left            =   45
            TabIndex        =   169
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   953
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
            Height          =   540
            Index           =   7
            Left            =   3495
            TabIndex        =   170
            Top             =   0
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…"
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
            Height          =   540
            Left            =   1770
            TabIndex        =   171
            Top             =   0
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„”«⁄œ…"
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
End
Attribute VB_Name = "FrmInpout2"
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

Public BolPrint As Boolean
Dim WithEvents m_MnuShowNewItemsPrices As Menu
Attribute m_MnuShowNewItemsPrices.VB_VarHelpID = -1
Dim WithEvents m_MenuViewList As Menu
Attribute m_MenuViewList.VB_VarHelpID = -1
Dim WithEvents m_MenuShowItemCostEffect As Menu
Attribute m_MenuShowItemCostEffect.VB_VarHelpID = -1
Dim WithEvents m_FrmSearch As Form
Attribute m_FrmSearch.VB_VarHelpID = -1
Dim bank_account As String
Dim general_noteid As Long
Dim CurrentVoucherNo As String
Dim CurrentVoucherSerialNo As String
Dim DateChanged As Boolean
Dim StroreChanged As Boolean
Dim TxtNoteSerial1V As String
Dim Account_Code_dynamic101 As String
Dim Account_Code_dynamic102 As String

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
Sub ClacultePrice(Optional ind As Integer = 0)
If ind <> 0 Then
    If val(TXTFactoryExpenses.text) = 0 And val(txt_total_bill.text) = 0 And val(Txt_EXport.text) = 0 Then
         FG.TextMatrix(ind, FG.ColIndex("CostPriceTk")) = val(FG.TextMatrix(ind, FG.ColIndex("Price")))
         Else
         TxtTotals.text = val(TXTFactoryExpenses.text) + val(txt_total_bill.text) + val(Txt_EXport.text)
         If val(LblTotalAll.Caption) <> 0 And val(FG.TextMatrix(ind, FG.ColIndex("Price"))) <> 0 Then
         TxtMode.text = ((val(FG.TextMatrix(ind, FG.ColIndex("Price"))) / val(LblTotalAll.Caption)) * val(TxtTotals.text)) + val(FG.TextMatrix(ind, FG.ColIndex("Price")))
         FG.TextMatrix(ind, FG.ColIndex("CostPriceTk")) = Round(val(TxtMode.text), 2)
         End If
             End If
        End If
End Sub

Function SaveItemsData(Optional Transaction_ID As String = 0)
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
    Cn.Execute "delete ItemsDetails   where Transaction_ID= " & val(Me.XPTxtBillID.text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  ItemsDetails Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim Unitname As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.Rows - 1

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
                         RsgGrantee("unitid").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 1, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))  ' val(astrSplitItems1(3))
                         RsgGrantee("ColorID").value = val(astrSplitItems1(4))
                         RsgGrantee("sizeid").value = val(astrSplitItems1(5))
                         RsgGrantee("ClassId").value = val(astrSplitItems1(6))
                         RsgGrantee("ProductionDate").value = IIf(IsDate((astrSplitItems1(7))), astrSplitItems1(7), Null)
                         RsgGrantee("ExpireDate").value = IIf(IsDate((astrSplitItems1(8))), astrSplitItems1(8), Null)
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
                        RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                       RsgGrantee("EffectN").value = 1
                    RsgGrantee.update
                                    Next intX
                Else
                RsgGrantee.AddNew
              RsgGrantee("ParrtNoCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))
            RsgGrantee("count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
            RsgGrantee("unitid").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          RsgGrantee("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RsgGrantee("sizeid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RsgGrantee("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
           RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
          RsgGrantee("ItemDetailedCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))
          RsgGrantee("EffectN").value = 1
           RsgGrantee.update
                  
                   
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
    Cn.Execute "delete TblGoldDetail   where Transaction_ID= " & val(Me.XPTxtBillID.text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  TblGoldDetail Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim Unitname As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.Rows - 1

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
         RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
         
      
  
               
                         
                         RsgGrantee.update
                  
                       
                Next intX
                    
            End If

        End If

    Next RowNum

End Function

Public Sub RetriveSerials(ItemID As String, _
                          itemname As String, _
                          seriallist As String, _
                          currentrow As Long)
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
    ' For i = 1 To Fg.Rows - 2
    '        If Fg.TextMatrix(i, Fg.ColIndex("Code")) = ItemID Then
    '         Me.Fg.RemoveItem (i)
    '         i = 1
    '        End If
    'NewGrid.Grid_AfterEdit Num, Fg.ColIndex("Code")
    ' Next i
   
    Num = currentrow

    '  For Num = currentrow To UBound(astrSplitItems)+currentrow
    For intX = 0 To UBound(astrSplitItems)
   
        FG.TextMatrix(Num, FG.ColIndex("Code")) = ItemID
        NewGrid.Grid_AfterEdit Num, FG.ColIndex("Code")
        ' FG.TextMatrix(Num, FG.ColIndex("Name")) = itemname
        FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
        FG.TextMatrix(Num, FG.ColIndex("Serial")) = astrSplitItems(intX)
  
        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.Rows = FG.Rows + 1
 
        Num = Num + 1
    Next
 
    TxtFillData.text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Function CheckMyData() As Boolean
    CheckMyData = True

    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": GoTo ErrTrap
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": GoTo ErrTrap
            Else
                TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
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
     If TxtNoteSerial1.text = "" Then
    
    NoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 20, , val(DCboStoreName.BoundText))
                    If NoteSerial1str = "error" Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ›« Ê—… „‘ —Ì«  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": GoTo ErrTrap
                    Else
                                   
                        If NoteSerial1str = "" Then
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… „‘ —Ì«   ÌœÊÌ« ﬂ„« Õœœ   ": GoTo ErrTrap
                        Else
                            TxtNoteSerial1.text = NoteSerial1str
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
        .Rows = .FixedRows
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
        My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_ID= " & val(Text1.text)
    Else
        'My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_ID= " & Val(Text1.text)
        My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where   ( (nots='" & Me.XPTxtBillID.text & "' and  Transaction_Type=20) or ( Transaction_Type=20   and  closed =0 and (nots='' or nots is null) ) and  (dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText) & ")) "
    End If

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID1
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
             
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
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

    With GRID1

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2='" & Me.TxtNoteSerial1.text & "' where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
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

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                'sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2='" & Me.TxtNoteSerial1.text & "' where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
           sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) ' & "nots=" & "" & "nots2=" & ""
             
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function
Private Sub CBoBasedON_Change()
   
  '      .AddItem "»·«"
  '      .AddItem "√„— ‘—¡"
  '      .AddItem "›« Ê—… „»œ∆ÌÂ"
  '      .AddItem "”‰œ ’—›"
  '      .AddItem "ÿ·» «— Ã«⁄"
  '      .AddItem "›« Ê—… ‘—«¡"
  '      .AddItem " ”ÊÌ«  Ã—œÌ…  "
  ' .AddItem "ÿ·» œ«Œ·Ì"
    
    
    If Me.CBoBasedON.ListIndex = 0 Then

    ElseIf Me.CBoBasedON.ListIndex = 1 Then
    If SystemOptions.UserInterface = EnglishInterface Then
    lbl(55).Caption = "Order NO"
    Else
    
        lbl(55).Caption = "—ﬁ„ «·«„—"
  End If
  
    ElseIf Me.CBoBasedON.ListIndex = 2 Or Me.CBoBasedON.ListIndex = 5 Then
       If SystemOptions.UserInterface = EnglishInterface Then
    lbl(55).Caption = "Bill NO"
    Else
        lbl(55).Caption = "—ﬁ„ «·ﬁ« Ê—… "
    End If
    
     ElseIf Me.CBoBasedON.ListIndex = 3 Or Me.CBoBasedON.ListIndex = 6 Then
       If SystemOptions.UserInterface = EnglishInterface Then
    lbl(55).Caption = "Vchr NO"
    Else
     lbl(55).Caption = "—ﬁ„ «·”‰œ"
     End If
     
        ElseIf Me.CBoBasedON.ListIndex = 4 Or Me.CBoBasedON.ListIndex = 7 Then
       If SystemOptions.UserInterface = EnglishInterface Then
    lbl(55).Caption = "Order NO"
    Else
     lbl(55).Caption = "—ﬁ„ «·ÿ·»"
     End If
     
      ElseIf Me.CBoBasedON.ListIndex = 8 Or Me.CBoBasedON.ListIndex = 10 Then
          If SystemOptions.UserInterface = EnglishInterface Then
    lbl(55).Caption = "Order NO"
    Else
     lbl(55).Caption = "—ﬁ„ «·«–‰"
     End If
    End If

End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
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
        TxtTaxAddValue.text = ""
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
        TxtTaxServiceValue.text = ""
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
        TxtTaxStampValue.text = ""
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

    With Me.Fg_Journal
        Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With

    ReLineGrid

End Function

Private Sub Cmd_Click(Index As Integer)
       On Error GoTo ErrTrap
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String

    BolPrint = True
 
    Select Case Index
    
        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
       With Me.grid4
            .Rows = .FixedRows
   
        End With
        
            Command2.Enabled = True
            Txt_EXport.Enabled = True
            '  Grid.Visible = True
            clear_all Me
            TxtModFlg.text = "N"
            ' Me.TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
            BillBasedOn(0).value = True
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))

            If BillType = 20 Then
                TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=20"))
            End If
        
            If BillType = 1 Then
                TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=1"))
            End If

            '      TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "",  True  )
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", 1)
            DCboStoreName.BoundText = intDef
            XPTab301.CurrTab = 0
            '        FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1
            Command2_Click
            
            
            
Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = False
              '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
'                TxtStoreID.Enabled = True
            End If
                    
                    
        

      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                                   Me.dcBranch.Enabled = False
                                   Else
                                    Me.dcBranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = False
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If

            
            Me.dcBranch.BoundText = Current_branch
            Me.CBoBasedON.ListIndex = 0
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
            Fg_Journal.Enabled = True
            GRID1.Rows = 1
            GRID1.Enabled = True
          
            DcCurrency.BoundText = 1
TxtNoteSerial1V = ""

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                If AvailableDeal = False Then
                    Exit Sub
                End If
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
        
            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            DateChanged = False
            CuurentLogdata
    StroreChanged = False
    Command4_Click
    
        Case 2
    If CheckCompositeAccount = False Then
        
            
End If

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«  "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
 
 
 
 
             If (DBCboClientName.BoundText) = 1 And CboPayMentType.ListIndex = 1 Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = " Cash Vendor can't be credit  "
                Else
                    Msg = "«·„Ê—œ «·‰ﬁœÌ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ «Ã·"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              ' Dcbranch.SetFocus
              '  SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            
            
            my_branch = Me.dcBranch.BoundText

            '   If Me.TxtModFlg.text = "N" Then
             
            ' End If
      
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

         '   If SystemOptions.usertype = UserNormal Then
         '       Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
         '       MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
         '       Exit Sub
         '   End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            If m_FrmSearch Is Nothing Then
                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = INVENTORYIN
            If SystemOptions.UserInterface = ArabicInterface Then
                m_FrmSearch.Caption = "«·»ÕÀ ⁄‰   ”‰œ«  «” ·«„"
             Else
             m_FrmSearch.Caption = "Search Recive Vchr"
             End If
                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal
            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’À… »‘«‘…     ”‰œ «·«” ·«„ «·Õ«·Ì… "
                Msg = Msg & Chr(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘…  ”‰œ «” ·«„"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                m_FrmSearch.Visible = True
                m_FrmSearch.ZOrder 0
                m_FrmSearch.SetFocus
            End If


        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
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
            ShowGL_cc TxtNoteSerial.text, , 200
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
     On Error Resume Next
ShowAttachments TxtNoteSerial1, "0812201403"

End Sub

Private Sub CmdCheque_Click()

    If Me.TxtModFlg.text = "R" Then
        Exit Sub
    End If

    Load FrmChecks
    FrmChecks.TxtModFlg.text = Me.TxtModFlg.text
    FrmChecks.XPTxtBillID.text = Me.XPTxtBillID.text
    Set FrmChecks.PutFg = Me.FgCheques
    FrmChecks.show vbModal
    SumChecks

End Sub

Private Sub SumChecks()

    With Me.FgCheques

        If .Rows > 1 Then
            Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .Rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .Rows - 1, .ColIndex("CheckValue"))
        Else
            Me.lbl(19).Caption = 0
            Me.lbl(18).Caption = 0
        End If

    End With

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdInfo_Click()
    Me.PopupMenu mdifrmmain.MnuInvPurchase
End Sub

Private Sub CmdINSTALLMENT_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim i As Integer

    If XPTxtValue(1).text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        If XPTxtValue(1).Enabled = True Then
            XPTxtValue(1).SetFocus
        End If

        Exit Sub
    End If

    Load FrmInstallMent
    Set FrmInstallMent.Frm = Me

    With FrmInstallMent

        If Me.TxtModFlg.text = "R" Then
            .Tag = "R"
            .Retrive val(XPTxtValue(1).Tag)
        Else
            .Tag = "N"
            .Txt(1).text = XPTxtValue(1).text
            .LblNoteID.Caption = XPTxtSerial(1).text
            .CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            .Txt(3).text = val(LblPrecenValue.Caption)
            .Txt(5).text = val(LblInstallCount.Caption)

            If IsDate(Me.LblFirstInstallDate.Caption) Then
                .Dtp_First.value = Me.LblFirstInstallDate.Caption
            End If

            .Txt(7).text = val(LblInstallSeprator.Caption)

            If val(LblInstallmentType.Tag) = 0 Then
                .OptInt(0).value = True
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                .OptInt(1).value = True
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                .OptInt(2).value = True
            End If

            With .FG
                .Rows = Me.FgInstallments.Rows

                For i = 1 To Me.FgInstallments.Rows - 1
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
    ShowRelatedNotes val(Me.XPTxtBillID.text), 1
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

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchId As Integer)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim total_shahn As Single
        
    Dim usedaccount As Integer

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    SngTemp = ((NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption) + val(LblCommision.Caption)) * val(txt_Currency_rate.text) + val(TXTToTAlELSHahn.text))


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
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                End If
            
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
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
        
                    Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
                End If

            Else
                Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
            End If

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = 0

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.text)
    
                        total_shahn = Round((((line_value) / (val(LblTotal.Caption) * val(txt_Currency_rate.text))) * val(TXTToTAlELSHahn.text)), 2)  'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                        line_value = line_value + total_shahn + val(FG.TextMatrix(i, FG.ColIndex("LineShahn")))
                        line_value = Round(line_value, 2)
     
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                        Else
                            StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                        End If
   
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(Me.LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text) '+ Val(TXTToTAlELSHahn.text)

        If SngTemp > 0 Then
            
            If SystemOptions.PoCreateVoucher = True And CboPayMentType.ListIndex = 1 Then
            If TXT_order_no.text <> "" Then
               StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            GoTo NewGl3
            End If
            
         
            End If
            
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

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
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
         
         
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.Rows - 1

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
                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.text)
                            '  total_shahn = Round((((line_value) / (Val(LblTotal.Caption) * Val(txt_Currency_rate.text))) * Val(TXTToTAlELSHahn.text)), 2)  'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                            '  line_value = line_value + total_shahn + Val(FG.TextMatrix(I, FG.ColIndex("LineShahn")))
                            line_value = Round(line_value, 2)
     
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                            Else
                                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If

        'ﬁÌœ «·„’—Ê›« 
        Dim Account_Code As String
        Dim Note_Value As Double

        With Grid

            For i = 1 To Grid.Rows - 1

                If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                    Else
                        StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_Code = Grid.TextMatrix(i, Grid.ColIndex("Account_code"))
                    Note_Value = Grid.TextMatrix(i, Grid.ColIndex("Note_value"))

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
        
            Next

        End With

        'ﬁÌœ «·›Ê« Ì—
        With grid4

            For i = 1 To grid4.Rows - 1

                If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                                            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                    Else
                        StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                    End If
                                                        
                    LngDevNO = LngDevNO + 1
                    Account_Code = grid4.TextMatrix(i, grid4.ColIndex("Account_code"))
                    Note_Value = grid4.TextMatrix(i, grid4.ColIndex("Note_value"))

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
         
            Next
   
        End With

        '«·„’—Ê›«  «·„»«‘—…
        With Fg_Journal

            For i = 1 To .Rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And val(.TextMatrix(i, .ColIndex("value"))) <> 0 Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                    Else
                        StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_Code = .TextMatrix(i, .ColIndex("AccountCode"))
                    Note_Value = val(.TextMatrix(i, .ColIndex("value")))

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
        
            Next

        End With

    End If



''
Dim CommissionAccount As String
        LngDevNO = LngDevNO + 1
                  CommissionAccount = get_account_code_branch(96, my_branch)
  
                    
                    Note_Value = val(LblCommision.Caption) * val(txt_Currency_rate.text)
If Note_Value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, CommissionAccount, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
   End If
ErrTrap:
End Function

'Function CreateRecieveVouchers()
'
'    If BillBasedOn(1).value = True Then Exit Function
'   '  On Error GoTo errortrap
'    Dim MYWAER As String
'    Dim StrSQL As String
''    Dim RsNotes As ADODB.Recordset
 '   Dim MYinvnum As String
 '   Dim note_id As Long
'
''    Dim RSTransDetails As ADODB.Recordset
 '   Dim RsTemp As New ADODB.Recordset
 ''   Dim RowNum As Integer
'    Dim StrSqlDel As String
'    Dim SearchResault As Integer
'    'Dim Note_ID As Long
'    Dim RsDetalis  As ADODB.Recordset
'    Dim BeginTrans As Boolean
'    Dim LnItemID As Long
'    Dim i As Long
'    Dim StrCurrentItemName As String
'    Dim DblNotesTotal As Double
'    Dim rs As ADODB.Recordset
'    Dim IntLineNO As Integer
'    Dim StrAccountCode As String
'    '  Dim RowNum As Integer
'    Dim Frm As Form
'    Dim Msg As String
'    Dim mytext As Integer
'    '>>>>>>>>>>>>>>>>>>>>>>>>>
'    CurrentVoucherNo = ""
'    CurrentVoucherSerialNo = ""
'    CurrentVoucherNo = Trim(GetVoucherGLNO(Val(Text1.text), CurrentVoucherSerialNo))
'
'TxtNoteSerial1V = ""
'    DeleteTransactiomsVoucher Val(Text1.text)
'
'    ' rs.Close
'
'    '        rs.Close
'
'    Set rs = New ADODB.Recordset
'
'    StrSQL = "select * from Transactions where Transaction_ID = " & Val(XPTxtBillID.text)   ' & TxtTransSerial.text & " and Transaction_type = 22"
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'
'    mytext = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=20"))
'
'    Dim Transaction_ID As Long
'
'    my_branch = Val(Me.dcBranch.BoundText)
'    Dim general_noteid As Long
'    Dim RsNotesGeneral As ADODB.Recordset
'    Dim TxtNoteSerialV As String
' 'Dim txtNoteSerial1V As String
'
'    my_branch = Val(Me.dcBranch.BoundText)
'
'    If TxtNoteSerialV = "" Then
'        If Notes_coding(Val(my_branch), XPDtbBill.value) = "error" Then
'            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Function
'        Else
'
'            If Notes_coding(Val(my_branch), XPDtbBill.value) = "" Then
'                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Function
'            Else
'                TxtNoteSerialV = Notes_coding(Val(my_branch), XPDtbBill.value)
'            End If
'        End If
'    End If
'
'        Dim TxtNoteSerial1Vstr As String
'
'
'    If TxtNoteSerial1V = "" Then
'    TxtNoteSerial1Vstr = Voucher_coding(Val(my_branch), XPDtbBill.value, 9, 160, , 20, , Val(DCboStoreName.BoundText))
'        If TxtNoteSerial1Vstr = "error" Then
'            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «÷«›… ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Function
'        Else
'
'            If TxtNoteSerial1Vstr = "" Then
'                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «÷«›…  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Function
'            Else
'                TxtNoteSerial1V = TxtNoteSerial1Vstr
'            End If
'        End If
'    End If
'
'    If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
'        TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
'                    TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
'                If StroreChanged <> True Then
'                    TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
'
'
'                End If
''
 '   Else
 '           TxtNoteSerial1V = Voucher_coding(Val(my_branch), XPDtbBill.value, 9, 160, , 20, , Val(DCboStoreName.BoundText))
 '           CurrentVoucherNo = ""
 '
 '
 '   End If
 '
 '
 '
 '
 '
 '    Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
 '    Cn.Execute "INSERT INTO  Transactions (CBoBasedON,Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,BranchId,nots2)SELECT 5," & Transaction_ID & "," & mytext & ",Transaction_Date,Transaction_Type = 20,CusID,StoreID,UserID,Emp_ID,nots='" & TxtNoteSerial1.text & "',NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId ," & TxtNoteSerial1 & " From Transactions Where Transaction_ID =" & XPTxtBillID.text + " And Transaction_Type = 22"
 '
 '   rs!nots = Transaction_ID
 '   rs.update
 '
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
'    Dim sql As String
'
'       sql = "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO) " & "SELECT   ( ( (round(Commisionvalue,2)+showPrice-( round(discountvalue,2)+TotalDiscountPerLine)*QtyBySmalltUnit)*" & Val(txt_Currency_rate.text) & ")+(ToTAlELSHahn+LineShahn)*QtyBySmalltUnit) ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, (( ( round(Commisionvalue,2)+ Price-(round(discountvalue,2)+TotalDiscountPerLine))*" & Val(txt_Currency_rate.text) & ")+(ToTAlELSHahn+LineShahn) ), ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID," & Me.XPDtbBill.value & ",ExpiryDate,LotNO  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
'
'
'    Cn.Execute sql
'
'
'    SaveItemsData (Transaction_ID)
' Dim NoteID As Long
'  Dim NoteDate As Date
'    Dim NoteSerial As String
'    Dim Notevalue As Double
'    Dim des As String
'If CurrentVoucherNo <> "" Then
'NoteSerial = CurrentVoucherNo
'End If
''TxtNoteSerialV
'
'CreateNotes NoteID, (XPDtbBill.value), Val(dcBranch.BoundText), 160, 0, NoteSerial, TxtNoteSerial1V, "Transactions", "Transaction_ID", Transaction_ID, TxtNoteSerial1V, ToHijriDate(XPDtbBill.value)
'          ' TxtNoteID.text = NoteID
'           general_noteid = NoteID
'
'
'     CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, Val(Me.dcBranch.BoundText)
'
''
'E 'rrTrap:
'
''End Function

Private Sub Command1_Click()
    'CreateRecieveVouchers
End Sub

Private Sub Command2_Click()

    '    ⁄»∆… «·«–Ê‰ «·„’—Ê›« 
    If CBoBasedON.ListIndex = 0 Or CBoBasedON.ListIndex = 1 Or TXT_order_no.text = "" Then

        With Me.Grid
            .Rows = .FixedRows
   
        End With

   '     Exit Sub

    End If

    With Me.Grid
        .Rows = .FixedRows
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

    'My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3   and order_no='" & Me.Txt_order_no.text & "' and(Transaction_ID1 is null or Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ")  )  "
'    My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where  dbo.Notes.NoteType = 3   and order_no='" & Me.TXT_order_no.text & "'"
    'My_SQL = ""
My_SQL = " SELECT     dbo.Notes.NoteID, dbo.Notes.Buy, dbo.Notes.NoteSerial, dbo.Notes.ItemID, dbo.Notes.Note_Value, dbo.ExpensesType.Name, dbo.ExpensesType.Account_Code, "
My_SQL = My_SQL & "   dbo.notes_all.BasedONID"
My_SQL = My_SQL & "  FROM         dbo.Notes INNER JOIN"
My_SQL = My_SQL & "  dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID INNER JOIN"
My_SQL = My_SQL & "    dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"
My_SQL = My_SQL & "  WHERE     (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.Notes.NoteType = 3) AND ( dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2 )"


    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim StrSQL  As String

    With Me.Grid
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                   
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsExp.Fields("ItemID").value), "", RsExp.Fields("ItemID").value)
    
                StrSQL = "select * from TblItems where ItemID=" & val(.TextMatrix(i, .ColIndex("ItemID")))
                Dim rs As New ADODB.Recordset
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
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
           
                .TextMatrix(i, .ColIndex("Select")) = 1
               
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

    With Me.GRID2
        .Rows = .FixedRows
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

    With Me.GRID2
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("order_no").value), "", RsExp.Fields("order_no").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    GRID2.Visible = True

End Sub

Private Sub Command4_Click()
    'If Not Fg.TextMatrix(Fg.Row, Fg.ColIndex("Code")) = "" Then
    '    ⁄»∆… «·«–Ê‰ «·„’—Ê›« 

    'Frame2.Caption = FG.TextMatrix(FG.Row, FG.ColIndex("name"))

    If CBoBasedON.ListIndex = 0 Or CBoBasedON.ListIndex = 1 Or TXT_order_no.text = "" Then

        With Me.grid4
            .Rows = .FixedRows
   
        End With

    '    Exit Sub

    End If

    With Me.grid4
        .Rows = .FixedRows
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
My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.buy , dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1, dbo.notes_all.BasedONID"
My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
My_SQL = My_SQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
My_SQL = My_SQL + " dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"

If Me.TxtModFlg.text = "R" Then
            If CBoBasedON.ListIndex = 0 Then
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 80)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.text) & ""
            Else
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2) and  dbo.DOUBLE_ENTREY_VOUCHERS.buy=1"
            End If

ElseIf Me.TxtModFlg.text = "E" Then


            If CBoBasedON.ListIndex = 0 Then
            My_SQL = My_SQL + " WHERE   (   dbo.Notes.NoteType = 80   AND  dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0 and     dbo.notes_all.BasedONID = 0   and   ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) ) or  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & val(Me.XPTxtBillID.text)
            Else
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2)"
            End If

ElseIf Me.TxtModFlg.text = "N" Then


            If CBoBasedON.ListIndex = 0 Then
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 80)  AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 0 ) and   ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) "
            Else
            My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TXT_order_no & "') AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) and  (  dbo.notes_all.BasedONID = 1 or dbo.notes_all.BasedONID = 2) and     ( dbo.DOUBLE_ENTREY_VOUCHERS.buy is null or dbo.DOUBLE_ENTREY_VOUCHERS.buy=0) "
            End If
            
End If


My_SQL = My_SQL + "  order by dbo.DOUBLE_ENTREY_VOUCHERS.buy desc ,dbo.Notes.NoteSerial1"
    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    Dim StrSQL As String
    Dim rs As New ADODB.Recordset

    With Me.grid4
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
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
 
                ' .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("buy").value), _
                  0, RsExp.Fields("buy").value)
            '    .TextMatrix(i, .ColIndex("Select")) = 1
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    grid4.Visible = True

    ' End If
  
    update_finincial_invoice_total
       
End Sub

Private Sub save_expenses()
   Dim Item_ID As Integer
    Dim i As Integer
    Dim sql As String
    ' ﬁÊ„ » ÕÌÀ ﬂ· ”ÿ— ›Ì «·ﬁÌœ »Ê÷⁄ —ﬁ„ «·⁄„·Ì… Ê—ﬁ„ «·’‰› Ê buy - Notes
    'Item_ID = Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code")))
    ' ›—Ì€  ﬂ·›… «·‘Õ‰ ⁄·Ï „” ÊÏ «·”ÿ—

    With Grid

        For i = 1 To Grid.Rows - 1
      
            Cn.BeginTrans
 
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                check_item_Exist_in_Grid val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("Note_value"))), True
        
                sql = "update notes set Transaction_ID1=" & val(Me.XPTxtBillID.text) & " , buy='1',itemid=" & IIf(val(.TextMatrix(i, .ColIndex("itemid"))) = 0, "Null", val(.TextMatrix(i, .ColIndex("itemid")))) & " where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
        
            Else
                sql = "update notes set Transaction_ID1=null ,  buy=Null,itemid=Null  where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))

            End If

            Cn.Execute sql

            Cn.CommitTrans

        Next

    End With

    Expenses_update_total

End Sub

Function Expenses_update_total()
    Dim i As Integer
    On Error Resume Next
    Txt_EXport.text = 0

    If Grid.Rows = 1 Then Exit Function

    With Grid

        For i = 1 To Grid.Rows - 1
        
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked And Grid.TextMatrix(i, Grid.ColIndex("ItemID")) = "" Then
            
                Txt_EXport.text = val(Txt_EXport.text) + val(Grid.TextMatrix(i, Grid.ColIndex("note_value")))
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

        For i = 1 To FG.Rows - 1
        
            .TextMatrix(i, .ColIndex("LineShahn")) = 0
      
        Next i

    End With

    With grid4
 
        For i = 1 To grid4.Rows - 1
      
            Cn.BeginTrans
 
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                check_item_Exist_in_Grid val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("Note_value")))
        
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=" & val(Me.XPTxtBillID.text) & " , buy='1',itemid=" & IIf(val(grid4.TextMatrix(i, grid4.ColIndex("itemid"))) = 0, "Null", val(grid4.TextMatrix(i, grid4.ColIndex("itemid")))) & " where Double_Entry_Vouchers_ID=" & val(grid4.TextMatrix(i, grid4.ColIndex("Double_Entry_Vouchers_ID")))
        
            Else
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=null , buy=Null,itemid=Null where Double_Entry_Vouchers_ID=" & val(grid4.TextMatrix(i, grid4.ColIndex("Double_Entry_Vouchers_ID")))

            End If

            Cn.Execute sql

            Cn.CommitTrans

        Next

    End With

    update_finincial_invoice_total

    '    DoEvents
    '    Command4_Click
End Sub

Function update_finincial_invoice_total()
    On Error Resume Next
    Dim i As Integer
    txt_total_bill.text = 0

    If grid4.Rows = 1 Then Exit Function

    With grid4

        For i = 1 To grid4.Rows - 1
        
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked And grid4.TextMatrix(i, grid4.ColIndex("ItemID")) = "" Then
                txt_total_bill.text = val(txt_total_bill.text) + val(grid4.TextMatrix(i, grid4.ColIndex("note_value")))
  
            End If
            
            If val(grid4.TextMatrix(i, grid4.ColIndex("select"))) = 0 Then
                grid4.TextMatrix(i, grid4.ColIndex("ItemID")) = ""
                grid4.TextMatrix(i, grid4.ColIndex("ItemCode")) = ""
                grid4.TextMatrix(i, grid4.ColIndex("ItemName")) = ""
            
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
    Dim RsTemp As ADODB.Recordset

    On Error GoTo ErrTrap
    Dim fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 2
    TxtSearchCode.text = fullcode

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
                '   CboPayMentType.locked = True
                '   CboPayMentType.ListIndex = 0
            Else
                '   CboPayMentType.locked = False
            End If
        End If
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
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
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                ElseIf RsTemp("Trans_DiscountTypePur").value = 2 Then
                    Me.XPCboDiscountType.ListIndex = 2
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                End If

            Else
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.text = 0
            End If

        Else
            Me.XPCboDiscountType.ListIndex = 0
            '     mina   Me.XPTxtDiscountVal.text = 0
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
        FrmCompanySearch.show vbModal
        FrmCompanySearch.lblSearchtype.Caption = 1
    End If
    
    
        If KeyCode = vbKeyF5 Then
        reloadCombos

    End If
    
 

End Sub
Function reloadCombos()
             Dim Dcombos As New ClsDataCombos
 Dim StrSQL As String
 
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True
    Dcombos.GetBranches dcBranch
    Dcombos.GetDocTypebyid Me.DCDocTypes, 20, val(Me.dcBranch.BoundText)
    StrSQL = "  select  BankID,BankName  from BanksData   "
    fill_combo dcbanks, StrSQL
 
    StrSQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, StrSQL

    StrSQL = " select id,Project_name from projects"
 
    fill_combo Me.DCproject, StrSQL
 Dcombos.GetStores Me.DCboStoreName
    
    
    
End Function

Private Sub dcbanks_KeyUp(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF5 Then
        reloadCombos

    End If
End Sub

Private Sub DcboBox_KeyUp(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF5 Then
        reloadCombos

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

TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 6) = True Or CheckStoreCoding(val(dcBranch.BoundText), 9) = True Then
     TxtNoteSerial1V = ""
     
StroreChanged = True



  CurrentVoucherNo = ""
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    DateChanged = True
    
    
     End If
     
    End If


End Sub

 

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF5 Then
        reloadCombos

    End If
    
    
End Sub

Private Sub dcBranch_Change()
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        Dcombos.GetDocTypebyid Me.DCDocTypes, 20, val(Me.dcBranch.BoundText)
    End If

    If dcBranch.BoundText = "" Then TxtNoteSerial1.locked = True: Exit Sub

    If Voucher_coding(val(Me.dcBranch.BoundText), XPDtbBill.value, 6, 150, 20) = "" Then
        TxtNoteSerial1.locked = True
    Else
        TxtNoteSerial1.locked = False
 
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    dcBranch_Change

    If Voucher_coding(val(val(Me.dcBranch.BoundText)), XPDtbBill.value, 6, 150, , 20) = "" Then Exit Sub
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    
         TxtNoteSerial1V = ""
     




  CurrentVoucherNo = ""
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
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
        reloadCombos

    End If
    
    

End Sub

Private Sub DcCurrency_Change()

    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.DcCurrency.BoundText <> "" Then
        txt_Currency_rate.text = get_currency_rate(Me.DcCurrency.BoundText)
    Else
        txt_Currency_rate.text = 1
    End If

End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim StrSQL As String
        StrSQL = " select id,code from currency"
 
        fill_combo Me.DcCurrency, StrSQL
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
        reloadCombos

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
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), Me.TxtNoteSerial, Me.TxtNoteSerial1, 150

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
End Sub

Private Sub Fg_Click()
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

        For RowNum = 1 To FG.Rows - 1

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

            Case "value"
                Dim sgl As String
               
                Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

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

            Case "value"
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
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.text = ""
        Else
            '
            TxtDes.text = Fg_Journal.Cell(flexcpData, r, c)
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
                    FrmExpensesSearch.RetrunType = 3
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

Function fillExpensesFactoryGrid()
 
    '  «·’‰«⁄Ì…   ⁄»∆… «·«–Ê‰ «·„’—Ê›« 
    With Me.Fg_Journal
        .Rows = .FixedRows
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
    My_SQL = "SELECT * from TblProductOrderFactoryexpenses where Transaction_ID=" & val(Me.XPTxtBillID.text)

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim StrSQL  As String

    With Me.Fg_Journal
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                   
                .TextMatrix(i, .ColIndex("LineNo")) = i
                
                .TextMatrix(i, .ColIndex("Accountcode")) = IIf(IsNull(RsExp.Fields("Accountcode").value), "", RsExp.Fields("Accountcode").value)
            
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsExp.Fields("AccountName").value), "", RsExp.Fields("AccountName").value)
               
                .TextMatrix(i, .ColIndex("value")) = IIf(Not IsNumeric(RsExp.Fields("value").value), 0, RsExp.Fields("value").value)
            
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsExp.Fields("des").value), "", RsExp.Fields("des").value)
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    With Me.Fg_Journal
        Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows - 1, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With
 
End Function

Private Sub Form_Activate()
    Set m_MnuShowNewItemsPrices = mdifrmmain.MnuInvPurchaseMnu2
    Set m_MenuViewList = mdifrmmain.MnuInvPurchaseMnu1
    Set m_MenuShowItemCostEffect = mdifrmmain.MnuInvPurchaseMnu4

End Sub

Private Sub CmdRetruns_Click()
    ShowRelatedTransactions val(Me.XPTxtBillID.text), 1
End Sub

Private Sub Form_Resize()
'  Me.WindowState = 2
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
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
        If Row = .Rows - 1 Then
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
        Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("NoteSerial"))
        Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("name"))
        Frm.txtValue = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Note_value"))
        
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
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim i As Integer
    Dim result As Integer
    result = 1
    StrSQL = "select * from  items_qty_not_recieved_in_order where  order_no='" & order_no & "'"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then

        For i = 1 To RsDetails.RecordCount

            If IsNull(RsDetails("net").value) Then result = 0: GoTo ll
            If RsDetails("net").value <> 0 Then
                result = 0
                GoTo ll
            End If

            RsDetails.MoveNext
        Next i
 
    End If

ll:
    Dim sql As String
    sql = "update Transactions Set closed = " & result & " Where Transaction_Type = 6 and order_no='" & Me.TXT_order_no & "'"
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


Function Retrive_orders_data(Transaction_ID As Integer, Optional Transaction_Type As Integer = 0)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim row_count As Integer
    Dim Num As Integer

    StrSQL = "Select * from transactions where Transaction_ID=" & Transaction_ID
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Function
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
         Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
          Me.XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
          
    End If

    If rs.EOF Or rs.BOF Then
        Exit Function
    End If

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.Rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.Rows - 1 'RsDetails.RecordCount
    
         '   Fg.TextMatrix(Num, Fg.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
          '  Fg.TextMatrix(Num, Fg.ColIndex("OrderArrivalDate")) = DTArrivalDate.value
         
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '          FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            
          If Transaction_Type <> 55 Then
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Showqty")), "", (RsDetails("Showqty").value))
          Else
          FG.TextMatrix(Num, FG.ColIndex("ShipedQty")) = IIf(IsNull(RsDetails("ShipedQty")), "", (RsDetails("ShipedQty").value))
          
          End If
            
            
 If Transaction_Type = 38 Then
FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value)) - IIf(IsNull(RsDetails("ItemBalance")), 0, (RsDetails("ItemBalance").value))
End If

            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassId")) = IIf(IsNull(RsDetails("ClassId")), 1, (RsDetails("ClassId").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
       
                    FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), "", (RsDetails("ShowPrice").value))

            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If

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
                StrComboList = grid4.BuildComboList(rs, "ItemName", "ItemID")
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
                    .Rows = 1
       
                End With
 
                fillVchr

            Case 7
                FrmInpout.Retrive val(.TextMatrix(.Row, 1))

            Case 8
                ShowGL_cc .TextMatrix(.Row, .ColIndex("NoteSerial")), , 200

        End Select

    End With

End Sub

Function fillVchr()
Dim str As String
Dim Transaction_ID As Double
str = ""

    Dim i As Integer
        
    With GRID1

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
            
            
            Transaction_ID = val(.TextMatrix(i, .ColIndex("Transaction_ID")))
                
           str = val(.TextMatrix(i, .ColIndex("Transaction_ID"))) & "," & str
            
            End If

        Next i

    End With

If str <> "" Then
str = Mid(str, 1, Len(str) - 1)
End If

Retrive_orders_data Transaction_ID, str

End Function

Private Sub GRID2_Click()

    With GRID2

        If .Cell(flexcpChecked, .Row, .ColIndex("select")) = flexChecked Then
            Retrive_orders_data (val(GRID2.TextMatrix(GRID2.Row, GRID2.ColIndex("Transaction_ID"))))
            
        End If

    End With

End Sub
 
Private Function check_item_Exist_in_Grid(ItemID As Integer, _
                                          value As Single, _
                                          Optional addition As Boolean)
    Dim i As Integer
    On Error Resume Next

    With FG

        For i = 1 To FG.Rows - 1

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
       
    With grid4

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
        If Row = .Rows - 1 Then
            '    .Rows = .Rows + 1
        End If

    End With

    update_finincial_invoice_total
End Sub

Private Sub grid4_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With grid4

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

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With grid4

        Select Case .ColKey(Col)

            Case "ItemName"
       
                StrSQL = "Select * from QRY_temp_bill_items"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = grid4.BuildComboList(rs, "ItemName", "ItemID")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
    FrmLC.show
    FrmLC.Retrive Trim(Me.TxtLCNO.text)

End Sub

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalview.Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblTotal_Change()

    If CboPayMentType.ListIndex = 1 Then
        XPTxtValue(1).text = LblTotal.Caption
    ElseIf CboPayMentType.ListIndex = 0 Then
        XPTxtValue(0).text = LblTotal.Caption
    End If
         
    LblTotalview.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblTotal_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Private Sub LblTotalAll_Change()
    LblTotalAllview.Caption = Format(val(LblTotalAll.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub m_FrmSearch_Unload(Cancel As Integer)
    Set m_FrmSearch = Nothing
End Sub

Private Sub m_MenuShowItemCostEffect_Click()

    If Me.TxtModFlg.text = "R" Then
        ShowItemCostEffectForTrans 1, , Trim$(Me.TxtTransSerial.text)
    End If

End Sub

Private Sub m_MenuViewList_Click()
    Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.vsFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.vsFlexGrid

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
    FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
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

Private Sub Txt_EXport_Change()
    Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, val(Txt_EXport.text)) + IIf(Not IsNumeric(txt_total_bill.text), 0, val(txt_total_bill.text)) + IIf(Not IsNumeric(TXTFactoryExpenses.text), 0, val(TXTFactoryExpenses.text))
End Sub

Private Sub Txt_order_no_Change()
'    Retrive_Expenses_Vouchers
 
    Dim Transaction_ID As String
    Dim Transaction_Type As Integer
    Dim Transaction_Type2 As Integer
    Transaction_Type2 = 0
    If CBoBasedON.ListIndex = 1 Then
        Transaction_Type = 29
    ElseIf CBoBasedON.ListIndex = 2 Then
        Transaction_Type = 17
    ElseIf CBoBasedON.ListIndex = 3 Then
        Transaction_Type = 19
        ElseIf CBoBasedON.ListIndex = 4 Then
        Transaction_Type = 0 ' ”‰œ «—Ã«⁄
        
            ElseIf CBoBasedON.ListIndex = 5 Then
        Transaction_Type = 22
            ElseIf CBoBasedON.ListIndex = 6 Then
        Transaction_Type = 16
        Transaction_Type2 = 15
    
                ElseIf CBoBasedON.ListIndex = 7 Then
        Transaction_Type = 38
       
        
                        ElseIf CBoBasedON.ListIndex = 9 Then
        Transaction_Type = 21
                       ElseIf CBoBasedON.ListIndex = 10 Then
        Transaction_Type = 55
        
    Else
        Transaction_Type = 0
        Exit Sub
    End If

    Transaction_ID = get_transactionData("noteserial1", TXT_order_no.text, "Transaction_ID", Transaction_Type, Transaction_Type2)

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        Retrive_orders_data val(Transaction_ID), Transaction_Type
    End If

'    With Me.grid4
'        .Rows = .FixedRows
'
'    End With
'
'    With Me.Grid
'        .Rows = .FixedRows
'
'    End With
'
'    If TXT_order_no.text = "" Then
'        txt_total_bill.text = ""
'        Txt_EXport.text = ""
'    End If
'
'    Command4_Click
'    Command2_Click
'    Command3_Click
'    Dim Transaction_ID As String
'    Dim Transaction_Type As Integer
'
''    If CBoBasedON.ListIndex = 1 Then
'        Transaction_Type = 29
'    ElseIf CBoBasedON.ListIndex = 2 Then
'        Transaction_Type = 17
'    Else
'        Transaction_Type = 0
'        Exit Sub
'    End If
'
'    Transaction_ID = get_transactionData("order_no", TXT_order_no.text, "Transaction_ID", Transaction_Type)
'
  '  If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
  '      Retrive_orders_data (Val(Transaction_ID))
  '  End If

End Sub

Private Sub TXT_total_payments_Change()
    'Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, Val(Txt_EXport.text)) + IIf(Not IsNumeric(TXT_total_payments.text), 0, Val(TXT_total_payments.text))
End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

  '  If KeyCode = vbKeyF3 Then
'
'        If CBoBasedON.ListIndex = 0 Then
'            Exit Sub
'
'        Else
'
'            TXT_order_no.text = ""
'            Order_no_search.show
'            Order_no_search.RetrunType = 3
'            Order_no_search.lblSpecificsearch.Caption = Val(CBoBasedON.ListIndex)
'        End If

'    End If
      If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
    If CBoBasedON.ListIndex = 1 Then
     
                 If KeyCode = vbKeyF3 Then
                           TXT_order_no.text = ""
                               Order_no_search.show
                                Order_no_search.RetrunType = 18
                                Order_no_search.lblSpecificsearch.Caption = val(CBoBasedON.ListIndex)
                     Order_no_search.DCboStoreName.BoundText = val(DCboStoreName.BoundText)
                    End If
ElseIf CBoBasedON.ListIndex = 7 Then
                               If KeyCode = vbKeyF3 Then
                          FrmBuySearch.Index = 20
                             FrmBuySearch.DealingForm = GridTransType.internalorder
                            
                                      FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…"
                                       FrmBuySearch.show vbModal
                               End If
             


ElseIf CBoBasedON.ListIndex = 10 Then
                    If KeyCode = vbKeyF3 Then
                    
            'Load ShippingissueSearch
            ShippingissueSearch.TType = 3
            
ShippingissueSearch.show
                End If
                





            End If
End If

End Sub

Private Sub txt_total_bill_Change()
    Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, val(Txt_EXport.text)) + IIf(Not IsNumeric(txt_total_bill.text), 0, val(txt_total_bill.text)) + IIf(Not IsNumeric(TXTFactoryExpenses.text), 0, val(TXTFactoryExpenses.text))

End Sub

Private Sub TXTFactoryExpenses_Change()
    Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, val(Txt_EXport.text)) + IIf(Not IsNumeric(txt_total_bill.text), 0, val(txt_total_bill.text)) + IIf(Not IsNumeric(TXTFactoryExpenses.text), 0, val(TXTFactoryExpenses.text))

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
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
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 2
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreId
    End If
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

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
        If Me.TxtModFlg.text = "R" Then
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
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
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
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
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
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
       
    ScreenNameArabic = "  ›« Ê—… „‘ —Ì«  "
    ScreenNameEnglish = " Purchase Invoice  "
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1", 150

    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim Dcombos As ClsDataCombos
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset

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
 
        Me.dcBranch.Enabled = False
    End If

    SetDtpickerDate XPDtbBill
    Set NewGrid = New ClsGrid
    NewGrid.GridTrans = INVENTORYIN
  '  NewGrid.GridTrans = PurchaseTransaction
    Set NewGrid.Grid = Me.FG
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtModFlag = Me.TxtModFlg
    Set NewGrid.TxtTotal = Me.XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.STORENAME = Me.DCboStoreName
    Set NewGrid.GrdTBar = Me.TBar
    '-----------------------------------------------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    '-----------------------------------------------------------------------------
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DcboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
        Set NewGrid.STORENAME = DCboStoreName
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
Set NewGrid.customer = Me.DBCboClientName
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
    End With

    With Me.CBoBasedON

        .Clear
        .AddItem "»·«"
        .AddItem "√„— ‘—¡"
        .AddItem "›« Ê—… „»œ∆ÌÂ"
        .AddItem "”‰œ ’—›"
        .AddItem "ÿ·» «— Ã«⁄"
        .AddItem "›« Ê—… ‘—«¡"
        .AddItem " ”ÊÌ«  Ã—œÌ…  "
                .AddItem "ÿ·» œ«Œ·Ì"
                .AddItem " «” ·«„ ‘Õ‰"
                .AddItem "›« Ê—… „»Ì⁄« "
                .AddItem "«–‰ ‘Õ‰ /  ”·Ì„"
                           .AddItem "«” ·«„ Â«·ﬂ"
    End With

    NewGrid.FillGrid
    Set Dcombos = New ClsDataCombos
     Dcombos.GetEmpDepartments Me.DcboEmpDepartments
     Dcombos.GetEmployees Me.DcboEmpName
     Dcombos.GetSalesRepDatapurchase Me.DcboEmp
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True
    Dcombos.GetBranches dcBranch
    Dcombos.GetDocTypebyid Me.DCDocTypes, 20, val(Me.dcBranch.BoundText)
   
    StrSQL = "  select  BankID,BankName  from BanksData   "
    fill_combo dcbanks, StrSQL
 
    StrSQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, StrSQL

    StrSQL = " select id,Project_name from projects"
 
    fill_combo Me.DCproject, StrSQL
      
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
        .Rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCheques
        .Rows = .FixedRows
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
    Dim Msg As String
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!checkinpo = True Then
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=20 Order by Transaction_ID"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.text = "R"
              'Resize_Form Me, TransactionSize
            BillType = 20
    
        Else
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=1"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.text = "R"
            '  Resize_Form Me, TransactionSize
            BillType = 1
            Exit Sub
        End If
    End If

    Me.TxtModFlg.text = "R"
    Command2_Click
    Command4_Click

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

ErrTrap:
End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    txtEmpCode.text = EmpCode
    
End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtEmpCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
    
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
CmdAttach.Caption = "Attachments"
lbl(75).Caption = "Employee"
lbl(74).Caption = "Dept"
lbl(81).Caption = "From Store"
lbl(82).Caption = "Car"
lbl(83).Caption = "Driver"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(69).Caption = "I. Manual No"
    lbl(68).Caption = "LC No"
    Label4.Caption = "Doc Type"
    lbl(77).Caption = "Order No:"
    'Label3.Caption = "Shioment No."
    'Label4.Caption = "Order No."
    Ele(13).Visible = False
    Frame5.Caption = "Notes"
    lbl(84).Caption = "Tel"
    lbl(72).Caption = "Employee"
    lbl(76).Caption = "LC No:"
    Command4.Caption = "Financial Invoice"
    XPCboDiscountType.Clear
    XPCboDiscountType.AddItem "NO"
    XPCboDiscountType.AddItem "Value"
    XPCboDiscountType.AddItem "Percent"
    CboPayMentType.Clear
    CboPayMentType.AddItem "Cash"
    CboPayMentType.AddItem "Credit"
    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(1) = "Securities"
    Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.XPTab301.TabCaption(5) = "Expenses Vouchers"
    Me.XPTab301.TabCaption(4) = "Purchase Orders and Performa Invoices"
    Me.XPTab301.TabCaption(3) = "Fn invoices"
    Me.XPTab301.TabCaption(6) = "Estimated Expenses"
    Me.XPTab301.TabCaption(7) = " Linked voucher"
    Me.XPTab301.TabCaption(8) = "Shipping"
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
       .Clear
        .AddItem "NA"
        .AddItem "PO"
        .AddItem "Performa Inv."
        .AddItem "Issue Voucher"
        .AddItem "Return Request"
        .AddItem "Purchase Inv"
        .AddItem "Adjustement "
                .AddItem "Internal Order"
                .AddItem "Shioment Vchr "
                .AddItem "Sales Invoice"
                .AddItem "Shipment Vchr"
                .AddItem "dipp Re"

    End With

    ' lbl(53).Caption = "Order No:"
    lbl(54).Caption = "Expenses"
    '  lbl(56).Caption = "Payment Voucher"
    '  lbl(57).Caption = "Total Payment"
    lbl(60).Caption = "Total"
 
    lbl(51).Caption = "Total Expenses"
    Command3.Caption = "View P.O. For Vendor"
  Me.Caption = "Recive Voucher"
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
        .TextMatrix(0, .ColIndex("ShipedQty")) = "Shiped Qty"
 
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
    lbl(8).Caption = "No#"
    lbl(9).Caption = "Currency"

    lbl(5).Caption = "Discount type"
    lbl(11).Caption = "Discount Value"
    lbl(10).Caption = "Payment Method"
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
        .TextMatrix(0, .ColIndex("value")) = "value"
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

    With Me.GRID2
 
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

    With Me.grid4
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
        .TextMatrix(0, .ColIndex("value")) = "value"

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
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·›« Ê—…   " & TxtNoteSerial1.text & Chr(13) & " —ﬁ„ ›« Ê—… «·„Ê—œ   " & txtManualNO.text & Chr(13) & " «· «—ÌŒ " & XPDtbBill.value & Chr(13) & " «·Œ“Ì‰… " & DcboBox.text & Chr(13) & " «·„Œ“‰  " & DCboStoreName.text & Chr(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.text & Chr(13) & "‰Ê⁄ «·”‰œ " & DCDocTypes & Chr(13) & "»‰«¡ ⁄·Ï " & CBoBasedON & "»—ﬁ„   " & TXT_order_no & Chr(13) & "ÿ—Ìﬁ… «·œ›⁄ " & CboPayMentType & Chr(13) & "‰Ê⁄ «·Œ’„ " & XPCboDiscountType & Chr(13) & "ﬁÌ„… «·Œ’„ " & XPTxtDiscountVal & Chr(13) & "  Ê’Ê· «·‘Õ‰… " & DTArrivalDate & Chr(13) & "  «·«” Õﬁ«ﬁ " & DtpDelayDate & Chr(13) & " «·⁄„·Â " & DcCurrency & Chr(13) & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Bill No " & TxtNoteSerial1.text & Chr(13) & "Supplier Bill No " & txtManualNO.text & Chr(13) & " Date " & XPDtbBill.value & Chr(13) & " Box " & DcboBox.text & Chr(13) & " Store  " & DCboStoreName.text & Chr(13) & " Supplier/Cuxtomer" & DBCboClientName.text & Chr(13) & "Doc Type" & DCDocTypes & Chr(13) & "Based On" & CBoBasedON & "No :   " & TXT_order_no & Chr(13) & "Payment Type" & CboPayMentType & Chr(13) & "Discount Type  " & XPCboDiscountType & Chr(13) & " Discount Vaalue   " & XPTxtDiscountVal & Chr(13) & " Shipment Arival Date" & DTArrivalDate & Chr(13) & "Due Date " & DtpDelayDate & Chr(13) & " Currency " & DcCurrency & Chr(13) & " GE NO" & TxtNoteSerial
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 150, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, "", , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 150, Date, Time, LogTextA, LogTextE, Me.name, "D", "", , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, 150

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

    Select Case Me.TxtModFlg.text

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
                Me.XPBtnMove(2).Enabled = False
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
            FG.Rows = 2
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

            If CboPayMentType.ListIndex = 0 Then
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

      On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting
'Me.TxtModFlg.text = "R"

    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""
    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxSerivce.value = vbUnchecked
    Me.TxtTaxServiceValue.text = ""

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
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    Me.TxtReciveOrderO.text = IIf(IsNull(rs("ReciveOrderO").value), "", (rs("ReciveOrderO").value))
     Me.TxtPolicyNo.text = IIf(IsNull(rs("PolicyNo").value), "", (rs("PolicyNo").value))
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    Me.DCproject.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(67).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    If Not (IsNull(rs("CashCustomerPhone").value)) Then
        Me.TxtPhone.text = rs("CashCustomerPhone").value
    Else
        Me.TxtPhone.text = ""
    End If


    If Not (IsNull(rs("CompsBill").value)) Then
         
                If (rs("CompsBill").value) = 0 Then
                          ChkCompsBill.value = vbUnchecked
                Else
                          ChkCompsBill.value = vbChecked
                End If
         
    Else
      ChkCompsBill.value = vbUnchecked
    End If
    
 
    
    
    TXT_order_no.text = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))
    txtManualNO.text = IIf(IsNull(rs("ManualNO").value), "", (rs("ManualNO").value))

    TxtManualNo1.text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))

    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    
    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    DTArrivalDate.value = IIf(IsNull(rs("ArrivalDate").value), Date, (rs("ArrivalDate").value))

    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), 0, rs("Trans_DiscountType").value)

    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", Trim(rs("Trans_Discount").value))
    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.text = ""
    End If

DcboEmpDepartments.BoundText = IIf(IsNull(rs("DepartementID").value), "", rs("DepartementID").value)
DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    TXTToTAlELSHahn.text = IIf(Not IsNumeric(rs("ToTAlELSHahn").value), 0, rs("ToTAlELSHahn").value)
    Txt_EXport.text = IIf(Not IsNumeric(rs("total_expenses").value), 0, rs("total_expenses").value)
    txt_total_bill.text = IIf(Not IsNumeric(rs("total_payments").value), 0, rs("total_payments").value)
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)

    TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))

    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)

    Text1.text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
    'Text1.text = IIf(IsNull(Rs("nots").Value), "", (Rs("nots").Value))
Me.DCboStoreName2.BoundText = IIf(IsNull(rs("storeid1").value), "", rs("storeid1").value)
    
    Me.DCDriver.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)
Me.DCCar.BoundText = IIf(IsNull(rs("CarId").value), "", rs("CarId").value)
    'txt_Shipment_no.text = IIf(IsNull(Rs("Shipment_no").value), "", Trim(Rs("Shipment_no").value))
    'Txt_order_no.text = IIf(IsNull(Rs("order_no").value), "", Trim(Rs("order_no").value))
    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)

    TxtLCNO.text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))

    '÷—»Ì… «·„»Ì⁄« 
    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(rs("TaxAddValue").value) Then
        If rs("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.text = rs("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(rs("TaxStampValue").value) Then
        If rs("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.text = rs("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(rs("TaxServiceValue").value) Then
        If rs("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.text = rs("TaxServiceValue").value
        End If
    End If

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    XPTxtSum.text = ""
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

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + "  where Transaction_ID=" & val(rs("Transaction_ID").value) & " order by Transaction_Details.id "
    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
     
            FG.TextMatrix(Num, FG.ColIndex("LineShahn")) = IIf(IsNull(RsDetails("LineShahn")), 0, (RsDetails("LineShahn").value))
            FG.TextMatrix(Num, FG.ColIndex("CostPriceTk")) = IIf(IsNull(RsDetails("CostPriceTk")), "", (RsDetails("CostPriceTk").value))
            FG.TextMatrix(Num, FG.ColIndex("CostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            'FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").Value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))

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

            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 0, (RsDetails("ItemSize").value))
           FG.TextMatrix(Num, FG.ColIndex("ItemsDetailsNewidea")) = IIf(IsNull(RsDetails("ItemsDetailsNewidea")), "", (RsDetails("ItemsDetailsNewidea").value))
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
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


      
            RsDetails.MoveNext

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

        '  FG.AutoSize 0, FG.Cols - 1, False
    End If

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""

    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    DtpDelayDate.value = Date
    StrSQL = "select * From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                'Me.TxtNoteID(0).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                'Me.TxtNoteID(1).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
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
        .Rows = .FixedRows

        If Not (RsNotes.BOF Or RsNotes.EOF) Then
            .Rows = .FixedRows + RsNotes.RecordCount

            For i = .FixedRows To .Rows - 1
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
                    .Rows = .FixedRows + RsPartDetails.RecordCount

                    For i = .FixedRows To .Rows - 1
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
    Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdRetruns.Tag = SngRelatedNotesValues
    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    TxtFillData.text = "F"
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    fill_bill_items_table
    Command3_Click
    '«” —Ã«⁄ «·„’—Ê›«  «· ﬁœÌ—ÌÂ
    fillExpensesFactoryGrid
    Command4_Click
    
    '  FillVoucherGrid

    Exit Sub
ErrTrap:
    Msg = "Œÿ« ›Ï ≈” —Ã«⁄ «·»Ì«‰« ..!!!"
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault

End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
                Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
                Msg = "Confirm Undo"
            End If

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
        
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
                Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
                Msg = "Confirm Undo"
            End If
  
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    Dim Msg As String
    Dim StrSQL As String
    Dim BegainTrans As Boolean
    Dim order_no As String
    order_no = Me.TXT_order_no
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·›« Ê—…  —ﬁ„ " & Chr(13)
        Msg = Msg + TxtNoteSerial1.text & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If AvailableDeal = True Or AvailableDeal = False Then
                If Not rs.RecordCount < 1 Then
                    Cn.BeginTrans
                    BegainTrans = True
                deletelinktoVoucher
                
                    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
                    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & get_transaction_id(rs("nots").value, 20)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
                    StrSQL = "Delete From Transactions  " & "Where Transaction_ID=" & get_transaction_id(rs("nots").value, 20)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
                    StrSQL = "update Notes set  Transaction_ID1=Null , ItemID=NUll, buy = null Where   (Transaction_ID1=" & val(Me.XPTxtBillID.text) & ")"
                    Cn.Execute StrSQL
            
                    StrSQL = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=Null ,  ItemID=NUll, buy = null Where  ( Transaction_ID1=" & val(Me.XPTxtBillID.text) & ")"
                    Cn.Execute StrSQL
            
                    StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.text)
    
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
                    DeleteTransactiomsVoucher val(Text1.text)
                    CuurentLogdata ("D")
                    rs.delete
                    Cn.CommitTrans
                    BegainTrans = False
                    rs.MoveFirst

        
                    close_order2 order_no
                                             With Me.grid4
            .Rows = .FixedRows
   
        End With
                    If rs.RecordCount < 1 Then
                        clear_all Me
                        TxtModFlg_Change
                        XPTxtCurrent.Caption = 0
                        XPTxtCount.Caption = 0
                    Else
                        Retrive
                    End If
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title

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
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ‘—«¡ ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «·‘—«¡" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·‘—«¡ «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·‘—«¡" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… ‘—«¡" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ‘—«¡" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
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
        .Create Me.hwnd, "»Ì«‰« ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
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
        .Create Me.hwnd, "»Ì«‰«  ”‰œ «” ·«„", 1, 15204351, -2147483630
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

    Dim J As Integer

    With GRID2

        For i = 1 To GRID2.Rows - 1
            close_order = True

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "select * from QRY_items_orders_data where order_no='" & GRID2.TextMatrix(i, GRID2.ColIndex("order_no")) & "'"
                Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Rs3.RecordCount = 0 Then GoTo ll

                For J = 1 To Rs3.RecordCount
                    order_qty = IIf(IsNull(Rs3("Quantity").value), 0, Rs3("Quantity").value)
                    QTYRecived = IIf(IsNull(Rs3("QTYRSV").value), 0, Rs3("QTYRSV").value)
                    differnt = order_qty - QTYRecived

                    If differnt <= 0 Then
                        close_order = False
                    End If
                
                    Rs3.MoveNext
                Next J
           
                If close_order = True Then
                    sql = "select * from Transactions where Transaction_Type=6 and order_no='" & GRID2.TextMatrix(i, GRID2.ColIndex("order_no")) & "'"
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
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            End If

   
    Dim SngTemp  As Variant
        Dim SngTempe  As Variant
 

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    
    
    
    
 
    
   
  StrSQL = " SELECT     dbo.TblCustemers.CusName, dbo.TblCustemers.Account_Code, SUM(( dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty)-Transaction_Details.discountvalue) AS totals"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & " dbo.TblCustemers ON dbo.Transaction_Details.SupplierID = dbo.TblCustemers.CusID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_ID = " & val(XPTxtBillID.text) & ")"
StrSQL = StrSQL & " GROUP BY dbo.TblCustemers.CusName, dbo.TblCustemers.Account_Code"
 
    supplierGL.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
Dim value As Variant
Dim Account_Code As String
Dim CusName As String
Dim totalvalue As Variant
totalvalue = 0
    For RowNum = 1 To supplierGL.RecordCount

    value = IIf(IsNull(supplierGL("totals").value), 0, supplierGL("totals").value)
    Account_Code = IIf(IsNull(supplierGL("Account_Code").value), "", supplierGL("Account_Code").value)
    CusName = IIf(IsNull(supplierGL("CusName").value), "", supplierGL("CusName").value)
    
    If value > 0 Then
  '  value = Round(value, SystemOptions.SysDefCurrencyForamt)
    SngTemp = Round(value * val(txt_Currency_rate.text), SystemOptions.SysDefCurrencyForamt)
     SngTempe = value
     LngDevNO = LngDevNO + 1
     totalvalue = totalvalue + SngTempe
              If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text & " «À»«  „‘ —Ì«  «·„Ê—œ " & CusName
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            End If
            
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
    End If
         supplierGL.MoveNext
    Next RowNum
   
     LngDevNO = LngDevNO + 1
     
    SngTemp = LblCommision * val(txt_Currency_rate.text)
    SngTempe = LblCommision
    SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
   If SngTemp > 0 Then


            If ModAccounts.AddNewDev(LngDevID, LngDevNO, CommissionAccount, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
       
            
             LngDevNO = LngDevNO + 1
             
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, SupplierAccount, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
            
    End If
    
   supplierGL.Close
   Set supplierGL = Nothing
    
    
    
  '' SngTemp = (LblTotalAll - LblDiscountsTotal) * val(txt_Currency_rate.text)
   ' SngTempe = (LblTotalAll - LblDiscountsTotal)
    
    totalvalue = totalvalue * val(txt_Currency_rate.text) '+ Round(LblCommision * val(txt_Currency_rate.text), SystemOptions.SysDefCurrencyForamt)
      LngDevNO = 1
   If totalvalue > 0 Then


    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic, totalvalue, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
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
        If TXT_order_no.text = "" Then
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
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            End If

   
    Dim SngTemp  As Variant
        Dim SngTempe  As Variant
 

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    
 
Dim value As Variant
Dim Account_Code As String
Dim CusName As String
Dim totalvalue As Variant
totalvalue = 0
   LngDevNO = 1
   value = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text)
   
 
  '  Account_Code = Account_Code_dynamic101
 '   CusName = IIf(IsNull(supplierGL("CusName").value), "", supplierGL("CusName").value)
    
    If value > 0 Then
  '  value = Round(value, SystemOptions.SysDefCurrencyForamt)
    SngTemp = Round(value, SystemOptions.SysDefCurrencyForamt)
     SngTempe = value
     LngDevNO = LngDevNO + 1
     totalvalue = totalvalue + SngTempe
              If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text & " «À»«  „‘ —Ì«  «·„Ê—œ " & Me.DBCboClientName.text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            End If
            
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic102, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
    End If
         
 
   
     LngDevNO = LngDevNO + 1
  
   If value > 0 Then


            If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic101, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTempe, , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
       
            
        
            
    End If
    
  
    
 
    
    Exit Function
ErrTrap:


End Function

Private Sub SaveData()
    Dim usedaccount As Integer
    Dim RSTransDetails As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim RsNotesGeneral As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim Msg As String
    Dim RowNum As Integer
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    Dim note_id As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    Dim TotalBillDiscount As Double
    Dim TotalDiscountPerLine As Double
 On Error GoTo ErrTrap

If SystemOptions.PoCreateVoucher = True Then

        If CheckAcconts = False Then
        
        
                Exit Sub
        
        End If

End If

    If CboPayMentType.ListIndex = 1 Then
        XPChkPayType(1).value = 1
        '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
        XPTxtValue(1).text = val(LblTotal.Caption)

    Else
        XPChkPayType(0).value = 1
        '  XPTxtValue(0).text = Val(LblTotalAll.Caption)
        XPTxtValue(0).text = val(LblTotal.Caption)

    End If
        
        
        
If ChkCompsBill.value = vbChecked Then
CboPayMentType.ListIndex = 1

    If DBCboClientName.BoundText = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„  «·„Ê—œ   «·«Ã·"
        Else
            Msg = "Select Customer Name"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
End If

        
        
    If DcCurrency.BoundText = "" Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«Œ — «·⁄„·… «Ê·« "
        Else
            Msg = "Select Currency First"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcCurrency.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    
    End If

    If Due_Date > DtpDelayDate.value Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÌÃ» «‰ ÌﬂÊ‰  «—ÌŒ «·«” Õﬁ«ﬁ «ﬂ»—   „‰ «Ê Ì”«ÊÌ     «—ÌŒ «Œ— ﬁ”ÿ"
        Else
            MsgBox "installment Date Must be Graeter than  or equal todya"
    
        End If

        Exit Sub
    End If

    If CboPayMentType.ListIndex = 1 Then
        Me.XPChkPayType(1).value = 1
        ' hany  XPTxtValue(1).text = Val(LblTotalAll.Caption)
    End If

    If Trim(Me.TxtTransSerial.text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ﬂ «»… —ﬁ„ ›« Ê—… «·‘—«¡..!!!"
        Else
            Msg = "Must Enter Bill No."
    
        End If
    
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtTransSerial.SetFocus
        Exit Sub
    End If

    'If Me.TxtModFlg.text = "N" Then
    '    If RepeatSerial(Trim(Me.TxtTransSerial.text), 1, 0, val(Me.DBCboClientName.BoundText)) = True Then
    '        Exit Sub
    '    End If

    'ElseIf Me.TxtModFlg.text = "E" Then

    '    If RepeatSerial(Trim(Me.TxtTransSerial.text), 1, val(Me.XPTxtBillID.text), val(Me.DBCboClientName.BoundText)) = True Then
    '        Exit Sub
    '    End If
    'End If

    '«· √ﬂœ „‰ ⁄œ„  ﬂ—«— —ﬁ„ «·”‰œ
    Dim BolTemp As Boolean

    If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 20) = "" Then
        If Me.TxtModFlg.text = "N" Then
    
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.text), 20, , val(dcBranch.BoundText))
        ElseIf Me.TxtModFlg.text = "E" Then
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.text), 20, val(Me.XPTxtBillID.text), val(dcBranch.BoundText))
        End If
 
        If BolTemp = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ «·”‰œ „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & Chr(13)
                Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·”‰œ"
            Else
                Msg = "This Bill No Already Exist" & Chr(13)
        
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtNoteSerial1.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    End If

    '‰Â«Ì… «· √ﬂœ

    Screen.MousePointer = vbArrowHourglass

    If DBCboClientName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ"
        Else
            Msg = "Select Customer Name"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If DCboStoreName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "„‰ ›÷·ﬂ Õœœ «”„ «·„Œ“‰"
        Else
            Msg = "Select Inventory First"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboStoreName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—…"
            Else
        
                Msg = "Specify Total Discount"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtDiscountVal.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Not IsNumeric(XPTxtDiscountVal.text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            Else
                Msg = "Discount Value Must be Numeric"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtDiscountVal.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        XPTxtDiscountVal.SetFocus
    End If

    If CboPayMentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        Else
            Msg = "Specify Payment Method"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboPayMentType.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If XPChkPayType(0).value = vbChecked Then
        If Me.DcboBox.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!"
            Else
                Msg = "Specify Box Name "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Me.TxtModFlg.text = "N" Then
            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).text), Me.XPDtbBill.value) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

        ElseIf Me.TxtModFlg.text = "E" Then

            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).text), Me.XPDtbBill.value, , , val(Me.XPTxtValue(0).Tag)) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If

    If val(Me.XPTxtValue(1).text) > 0 Then
        If ChkInstall.value = vbChecked Then
            If val(Me.LblInstallTotal.Caption) = 0 Then
                Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.XPTab301.CurrTab = 1
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            '  hany      If Val(Me.LblInstallTotal.Caption) <> Val(Me.XPTxtValue(1).text) Then
            '            Me.XPTxtValue(1).text = Val(Me.LblInstallTotal.Caption)
            '        End If
        End If
    End If

    If XPChkPayType(2).value = vbChecked Then
        If val(Me.lbl(18).Caption) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
            Else
                Msg = "Enter Cheques Data Before Save"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If dcbanks.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = Msg + "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ     " & Chr(13)
            Else
                Msg = Msg + " Specify Bank NAme     " & Chr(13)
            End If
        
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '            Dcbanks.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
    
            Dim rsbank As New ADODB.Recordset
            Set rsbank = New ADODB.Recordset
            rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
            If Not (rsbank.EOF Or rsbank.BOF) Then
                If rsbank!banks_Accounts = True Then
                    bank_account = get_bank_Account(val(Me.dcbanks.BoundText), "Account_Code2")
                Else
                    bank_account = get_bank_Account(val(Me.dcbanks.BoundText), "Account_Code")
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
    Cn.Execute "delete DOUBLE_ENTREY_VOUCHERS where Transaction_ID = " & val(Text2.text)

    If NewGrid.Calculate(1, , False, True) = False Then
        Screen.MousePointer = vbDefault
        '  Exit Sub
    End If

    '-------------------------------
    '    If Me.XPChkPayType(0).value = vbChecked Then
    '        DblNotesTotal = Val(Me.XPTxtValue(0).text)
    '    End If

    '    If Me.XPChkPayType(1).value = vbChecked Then
    '        Me.XPTxtValue(1).text = LblTotal.Caption
    DblNotesTotal = DblNotesTotal + val(Me.XPTxtValue(1).text)
    '    End If

    '    If Me.XPChkPayType(2).value = vbChecked Then
    '        DblNotesTotal = DblNotesTotal + Val(Me.lbl(18).Caption)
    '    End If
    DblNotesTotal = val(Me.XPTxtValue(0).text) + val(Me.XPTxtValue(1).text) + val(lbl(18).Caption)

    If CboPayMentType.ListIndex = 1 Then
        Me.XPChkPayType(1).value = 1
        '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
    End If

    '   If CboPayMentType.ListIndex = 0 Then
    '       Me.XPChkPayType(0).value = 1
    '         XPTxtValue(0).text = Val(LblTotalAll.Caption)
    '   End If
     
    If DblNotesTotal <> val(LblTotal.Caption) Then
        If SystemOptions.UserInterface = ArabicInterface Then
     '       Msg = "≈Ã„«·Ï «·√Ê—«ﬁ «·„«·Ì… €Ì— „ ”«ÊÏ „⁄ ≈Ã„«·Ï «·›« Ê—…...!!!"
        Else
     '       Msg = "Error In total ...!!!"
        End If

     '   MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     '   Exit Sub
    End If

    '---------Start Saving------------------------------------------------
    '    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    'Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    'Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))

    'Õ›Ÿ «·„’—Ê›«  «·«÷«›Ì… Ê«·›Ê« Ì— «·„«·Ì…
    Save_Financial_invoice
    save_expenses

    '---------Notes ID ------------------------------------------------
    'Create big notes
    GoTo xll

    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
        
        Dim NoteSerial1str As String
        
    If TxtNoteSerial1.text = "" Then
    
    NoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 20, , val(DCboStoreName.BoundText))
                    If NoteSerial1str = "error" Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ›« Ê—… „‘ —Ì«  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                                   
                        If NoteSerial1str = "" Then
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… „‘ —Ì«   ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            TxtNoteSerial1.text = NoteSerial1str
                        End If
                    End If
    End If
     
xll:
 '   Set RsNotesGeneral = New ADODB.Recordset
 '      StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
 '  RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'
'
    
'    RsNotesGeneral.AddNew
'    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
'    general_noteid = RsNotesGeneral("NoteID").value
'    TxtNoteID.text = general_noteid
'
'    RsNotesGeneral("NoteDate").value = XPDtbBill.value
'    RsNotesGeneral("NoteType").value = 150
'    RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
'    RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
'
'    RsNotesGeneral("Remark").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
'    RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
'
'    RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
'    RsNotesGeneral("numbering_type1").value = sand_numbering_type(6) '  ›« Ê—… ‘ƒ«¡
'    RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
'    RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
'
'    RsNotesGeneral("branch_no").value = val(Me.Dcbranch.BoundText)
'    RsNotesGeneral.update

    '---------Start Saving------------------------------------------------

    Set RSTransDetails = New ADODB.Recordset
    Set RsNotes = New ADODB.Recordset
  '  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  '  RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    
      StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    BeginTrans = True

    If Me.TxtModFlg.text = "N" Then
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs.AddNew
        rs("Transaction_ID").value = val(XPTxtBillID.text)
      TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 20, , val(DCboStoreName.BoundText))
     Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=21"))
         
         
          Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
          
          
    ElseIf Me.TxtModFlg.text = "E" Then

        If rs("Transaction_ID").value <> val(XPTxtBillID.text) Then
            rs.find "Transaction_ID=" & val(XPTxtBillID.text), , adSearchForward, 1
        End If

        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
          
            StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL
        
        StrSqlDel = "delete From Notes where noteid=" & val(TXTNoteID.text)
        Cn.Execute StrSqlDel
        
        
        general_noteid = val(TXTNoteID.text)
         If TxtNoteSerial.text = "" Then
          TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 20, , val(DCboStoreName.BoundText))
          End If
          
        
    End If
    
 


    
    If ChkCompsBill.value = vbChecked Then
     rs("CompsBill").value = 1
    Else
        rs("CompsBill").value = 0
    End If
     rs("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
    rs("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
     rs("storeid1").value = IIf(Me.DCboStoreName2.BoundText = "", Null, (Me.DCboStoreName2.BoundText))
    rs("ReciveOrderO").value = IIf(Trim(Me.TxtReciveOrderO.text) = "", Null, Trim(Me.TxtReciveOrderO.text))
    rs("PolicyNo").value = IIf(Trim(Me.TxtPolicyNo.text) = "", Null, Trim(Me.TxtPolicyNo.text))
    rs("Emp_ID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
    rs("DepartementID").value = IIf(DcboEmpDepartments.BoundText = "", Null, val(DcboEmpDepartments.BoundText))
    rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
    rs("ManualNo1").value = IIf(TxtManualNo1.text = "", Null, val(TxtManualNo1.text))
    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    rs("NoteId").value = val(TXTNoteID.text)
    rs("order_no").value = IIf((TXT_order_no.text) = "", Null, TXT_order_no.text)
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", Null, Trim(Me.TxtTransSerial.text))
    rs("Transaction_Date").value = XPDtbBill.value
    rs("ArrivalDate").value = DTArrivalDate.value
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs("Transaction_Type").value = BillType
    rs("UserID").value = user_id
    rs("nots").value = Text1.text
    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))
        If Trim$(Me.TxtCashCustomerName.text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
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
        rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
    End If

    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If
   If Trim$(Me.TxtPhone.text) <> "" Then
        rs("CashCustomerPhone").value = Trim$(Me.TxtPhone.text)
    Else
        rs("CashCustomerPhone").value = Null
    End If
    
    rs("project_id").value = IIf(DCproject.BoundText = "", Null, (DCproject.BoundText))
    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, (DBCboClientName.BoundText))
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, (DCboStoreName.BoundText))
    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))
    rs("ToTAlELSHahn").value = IIf(Not IsNumeric(TXTToTAlELSHahn.text), 0, Me.TXTToTAlELSHahn.text)
    
    rs("total_expenses").value = IIf(Not IsNumeric(Txt_EXport.text), 0, Txt_EXport.text)
    rs("total_payments").value = IIf(Not IsNumeric(txt_total_bill.text), 0, txt_total_bill.text)
    rs("LcNo").value = IIf(TxtLCNO.text = "", Null, (TxtLCNO.text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.text) > 0 Then
        rs("TaxAddValue").value = val(Me.TxtTaxAddValue.text)
    Else
        rs("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.text) > 0 Then
        rs("TaxStampValue").value = val(Me.TxtTaxStampValue.text)
    Else
        rs("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.text) > 0 Then
        rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.text)
    Else
        rs("TaxServiceValue").value = 0
    End If

    'rs("Shipment_no").value = IIf(txt_Shipment_no.text = "", Null, (txt_Shipment_no.text))
    rs("order_no").value = IIf(TXT_order_no.text = "", Null, (TXT_order_no.text))
    rs("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    rs("ManualNO").value = IIf(txtManualNO.text = "", Null, (txtManualNO.text))

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

     Dim NoteID As Long
  Dim NoteDate As Date
    Dim NoteSerial As String
    Dim Notevalue As Double
    Dim des As String
    If Me.TxtNoteSerial.text <> "" Then
NoteSerial = Me.TxtNoteSerial.text
End If

 CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 150, 0, NoteSerial, TxtNoteSerial1, "Transactions", "Transaction_ID", val(XPTxtBillID.text), TxtNoteSerial1.text, ToHijriDate(XPDtbBill.value)
           TXTNoteID.text = NoteID
           general_noteid = NoteID


    
    
   
    For RowNum = 1 To FG.Rows - 1

        'Check Repeat Serial
        If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
            StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
            StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.text
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                    Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & Chr(13)
                    Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                Else
                    Msg = "Item Serial" & Chr(13)
                    Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & Chr(13)
                    Msg = Msg + "Already Exist in this bill"
            
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
          ClacultePrice RowNum
          
            RSTransDetails("CostPriceTk").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CostPriceTk")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("CostPriceTk"))))
            RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("CostPrice"))))
            RSTransDetails("BranchId").value = Me.dcBranch.BoundText
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Me.XPDtbBill.value, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
            RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
            RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

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
            
            RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
         
            RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            RSTransDetails("order_no").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
       '     RSTransDetails("OrderArrivalDate").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("OrderArrivalDate")) = ""), Me.XPDtbBill.value, Fg.TextMatrix(RowNum, Fg.ColIndex("OrderArrivalDate")))
             If (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) = Empty Then
              
             Dim newUnitId As Long
             GetDefaultItemUnit RSTransDetails("Item_ID").value, newUnitId
               RSTransDetails("UnitID").value = newUnitId
             End If
             
            RSTransDetails("LineShahn").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("LineShahn")) = ""), 0, FG.TextMatrix(RowNum, FG.ColIndex("LineShahn")))
             
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            If (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) = Empty Then
            
          LngUnitID = newUnitId
            Else
           LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            End If
            
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))
If LngUnitID = 0 Then

End If

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
            End If

            '          RSTransDetails("Price").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ShowPrice")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("ShowPrice"))))
            '     RSTransDetails("price").value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").value, 2)
            RSTransDetails("price").value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Price")) / RSTransDetails("QtyBySmalltUnit").value, 15)
    
            RSTransDetails("OpeningBurcahseQty").value = RSTransDetails("Quantity").value
            RSTransDetails("OpeningBurcahseValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
            
         Dim rate As Single
         
            RSTransDetails("discountvalue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("discountvalue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("discountvalue")))) / RSTransDetails("Quantity").value
      
      'RATE = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ScurrencyID")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("ScurrencyID"))))
            RSTransDetails("rate").value = val(txt_Currency_rate.text)

            If val(LblTotal.Caption) = 0 Then LblTotal.Caption = 1
            ' RSTransDetails("ToTAlELSHahn") = Round((((RSTransDetails("showPrice") * _
            ' RSTransDetails("ShowQty")) / Val(LblTotal.Caption)) * _
            ' Val(TXTToTAlELSHahn.text)) / RSTransDetails("ShowQty"), 2)   ' / RSTransDetails("ShowQty")
            Dim TotalShahnPerLine As Double
            TotalShahnPerLine = ((((RSTransDetails("price") * RSTransDetails("Quantity") / (LblTotalAll.Caption))) * val(TXTToTAlELSHahn.text)) / RSTransDetails("Quantity"))
            TotalShahnPerLine = Round(TotalShahnPerLine, 15) 'Val(Format(TotalShahnPerLine, "." & String(Abs(18), "#")))
            RSTransDetails("ToTAlELSHahn") = TotalShahnPerLine
         
            If Me.XPCboDiscountType.ListIndex = 1 Then
                TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.text <> "" Then
                    TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text)) * val(LblTotalAll.Caption) / 100
                             
                Else
                    TotalBillDiscount = 0
                End If
            End If
           
            TotalDiscountPerLine = ((((RSTransDetails("price") * RSTransDetails("Quantity") / (LblTotalAll.Caption))) * val(TotalBillDiscount)) / RSTransDetails("Quantity"))
            RSTransDetails("TotalDiscountPerLine") = TotalDiscountPerLine
            
            ' RSTransDetails.update
        End If

        RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
        RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
        RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
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

    
      
'******************************

        RSTransDetails.update
    Next RowNum

    '------------------------------------------------------------------------------
     
    '------------------------------------------------------------------------------
    If Me.XPChkPayType(0).value = Checked Then
        'RsNotes.AddNew
        'RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        'note_id = RsNotes("NoteID").value

        If Me.TxtModFlg.text = "N" Then
            '    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
            '    XPTxtSerial(0).text = RsNotes("NoteSerial").value
        ElseIf Trim(XPTxtSerial(0).text) <> "" Then
            '    RsNotes("NoteSerial").value = Trim(XPTxtSerial(0).text)
        Else
            '    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
            '    XPTxtSerial(0).text = RsNotes("NoteSerial").value
        End If

        '--------------------------------------------------------------------------
    End If

    If Me.XPChkPayType(1).value = Checked Then
        RsNotes.AddNew
        RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        note_id = RsNotes("NoteID").value
        RsNotes("NoteDate").value = XPDtbBill.value
 
        RsNotes("remark").value = Me.TxtNoteSerial1.text
        RsNotes("NoteSerial").value = Null

        RsNotes("Transaction_ID").value = val(XPTxtBillID.text)
        RsNotes("NoteType").value = 1
        RsNotes("Note_Value").value = IIf(XPTxtValue(1).text = "", Null, val(XPTxtValue(1).text))
        RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        RsNotes("BankID").value = Null
        RsNotes("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText)) 'Null SALIM MY BE ERROR
        RsNotes("DueDate").value = DtpDelayDate.value
        RsNotes.update
 
    End If

    If Me.XPChkPayType(2).value = Checked Then

        With Me.FgCheques

            For i = .FixedRows To .Rows - 1
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
            RsTemp("BasicAmmount").value = IIf(XPTxtValue(1).text = "", 0, val(XPTxtValue(1).text))
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

                For RowNum = 1 To .Rows - 1
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

    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
'If SystemOptions.PoCreateVoucher = True And CBoBasedON.ListIndex = 1 And TXT_order_no.text <> "" Then GoTo NewGL2


If SystemOptions.PoCreateVoucher = True And CboPayMentType.ListIndex = 1 Then
If TXT_order_no.text = "" Then
 
Else
GoTo NewGL2
End If

End If


If ChkCompsBill.value = vbChecked Then GoTo NewGL

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text)
    '    SngTemp = (Val(Me.lblTotal.Caption) * Val(txt_Currency_rate.text))
      
    'SngTemp =  (SngTemp, SystemOptions.SysDefCurrencyForamt)
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
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            End If

            LngDevNO = LngDevNO + 1
    
            If txtManualNO.text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & txtManualNO
                Else
                    StrTempDes = StrTempDes & " Supp Bill# " & txtManualNO
                End If
            
            End If

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
   
        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

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

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * val(txt_Currency_rate.text) * FG.TextMatrix(i, FG.ColIndex("Count"))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
                        Else
                            StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
                        End If
                            
                        If txtManualNO.text <> "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & txtManualNO
                            Else
                                StrTempDes = StrTempDes & " Supp Bill# " & txtManualNO
                            End If
            
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If
    
        'Œ’„ ⁄·Ï „” ÊÏ «·”ÿ—
        If detect_inventory_work_type = 3 Then

            With FG

                For i = 1 To FG.Rows - 1
 
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

                        line_value = (FG.TextMatrix(i, FG.ColIndex("Price")) * val(txt_Currency_rate.text) * FG.TextMatrix(i, FG.ColIndex("Count"))) - FG.TextMatrix(i, FG.ColIndex("Valu"))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
                        Else
                            StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With
    
        End If

    End If

 

    '«·œ«∆‰
    If Me.XPChkPayType(0).value = vbChecked Then

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
            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        Else
            StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        End If
    
        If txtManualNO.text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & txtManualNO
            Else
                StrTempDes = StrTempDes & " Supp Bill# " & txtManualNO
            End If
            
        End If

        '  SngTemp = (Val(Me.XPTxtValue(0).text) * Val(txt_Currency_rate.text))
        '  SngTemp = (Val(Me.lblTotal.Caption) * Val(txt_Currency_rate.text))
        ' SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) * Val(txt_Currency_rate.text)
        SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text)
        '   SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
    
        LngDevNO = LngDevNO + 1

        If Trim(TxtLCNO) <> "" Then
            StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLCNO.text)
        End If
        
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
            GoTo ErrTrap
        End If
    End If

    If Me.XPChkPayType(1).value = vbChecked Then
    
        '«·√Ã·
        SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text)

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
            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        Else
            StrTempDes = "Purchase Invoice NO: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        End If

        LngDevNO = LngDevNO + 1
    
        If txtManualNO.text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & txtManualNO
            Else
                StrTempDes = StrTempDes & " Supp Bill# " & txtManualNO
            End If
            
        End If

        If Trim(TxtLCNO) <> "" Then
            StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLCNO.text)
        End If
        
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
            GoTo ErrTrap
        End If
    End If

    If Me.XPChkPayType(2).value = vbChecked Then
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) * val(txt_Currency_rate.text)
  
        StrTempAccountCode = bank_account  '‘Ìﬂ«  „ƒÃ·…

        '    StrTempAccountCode = "a2a3a2" '√Ê—«ﬁ «·œ›⁄
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "⁄œœ " & Me.lbl(19).Caption & "  ‘Ìﬂ«  " & Chr(13)
            StrTempDes = StrTempDes & "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text
        Else
            StrTempDes = "Count " & Me.lbl(19).Caption & "  Cheque " & Chr(13)
            StrTempDes = StrTempDes & "Purchase Invoice No:" & Me.TxtNoteSerial1.text
    
        End If

        LngDevNO = LngDevNO + 1

        If Trim(TxtLCNO) <> "" Then
            StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLCNO.text)
        End If
        
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
            GoTo ErrTrap
        End If
    End If

    
    
    
  'new GE
    
    
    If Text1.text <> "" Then
        Cn.Execute "update Transactions set nots =' " & TxtTransSerial.text & "' where Transaction_Type= 20 and Transaction_Serial=" & Text1.text & ""
    End If

    Cn.Execute "update Transactions set NoteSerial =' " & Trim(Me.TxtNoteSerial.text) & "' where Transaction_ID=" & val(Me.XPTxtBillID.text)


'*************************
'******************
    'Õ›Ÿ «·„’—Êﬁ«  «· ﬁœÌ—Ì…
    Dim FactoryExpenses As New ADODB.Recordset

    If Me.TxtModFlg.text = "E" Then
        StrSQL = "Delete TblProductOrderFactoryexpenses where Transaction_ID=" & val(XPTxtBillID.text)
        Cn.Execute StrSQL
    End If

    StrSQL = "Select * from TblProductOrderFactoryexpenses where Transaction_ID=" & val(XPTxtBillID.text)
    FactoryExpenses.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For RowNum = 1 To Fg_Journal.Rows - 2

        If Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName")) <> "" Then
            FactoryExpenses.AddNew
            FactoryExpenses("Transaction_ID").value = val(XPTxtBillID.text)
         
            FactoryExpenses("Accountcode").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Accountcode"))
            FactoryExpenses("AccountName").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName"))
            FactoryExpenses("value").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("value")))
            FactoryExpenses("des").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("des"))
            FactoryExpenses.update
        End If
         
    Next RowNum
NewGL:
 SaveNewGl
NewGL2:
 SaveNewGl2
 
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    


    CloseIssueVoucher

 '   If SystemOptions.autoReseiveVoucher = True Then
 '       CreateRecieveVouchers
 '   End If
 
 'SaveItemsData
    '----------------------------------------------------------------
    '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=" & BillType
         
    If SystemOptions.usertype <> UserAdminAll Then
        'StrSQL = StrSQL & " and Transaction_ID=0  AND   BranchId=" & branch_id
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.Retrive val(Me.XPTxtBillID.text)
    '----------------------------------------------------------------

    CuurentLogdata

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & Chr(13)
    
            End If
    
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Changes was Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

            lbl(67).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    
    End Select

    close_order2 Me.TXT_order_no
    'Closeorders
    TxtModFlg.text = "R"
    Command4_Click
    'UpdateTransCost val(Me.XPTxtBillID.text)

    If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
        '›Ï Õ«·… «‰  ﬂÊ‰ ÿ—Ìﬁ… Õ”«» „ Ê”ÿ «· ﬂ·›…
        'ÂÊ
        '    ModernWeightAverage
        '·«»œ «‰ ÌﬁÊ„ «·»—‰«„Ã » ⁄œÌ· ﬁÌ„… „ Ê”ÿ «· ﬂ·›… ··√’‰«›
        '«·„ÊÃÊœ… ›Ï «·›« Ê—…
    End If

    Screen.MousePointer = vbDefault
    Command2.Enabled = True
    Txt_EXport.Enabled = True
    'Grid.Visible = False
    Exit Sub
ErrTrap:

    'Stop
    'Resume
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
        Msg = "Sorry....Error During Saving" & Chr(13)
    End If

    Msg = Msg & Err.description & Chr(13)
    Msg = Msg & Err.Number & Chr(13)
    Msg = Msg & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub XPBtnNewClients_Click()

    With FrmAddNewCustemer
        '    .Tag = "x"
        .DealingForm = PurchaseTransaction
        Set .DcboCustomers = DBCboClientName
        .Caption = "≈÷«›… „Ê—œ ÃœÌœ"
        .lbl(1).Caption = "ﬂÊœ «·„Ê—œ"
        .lbl(0).Caption = "«”„ «·„Ê—œ"
        .AddType = 2
        .show vbModal
    End With

End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
 
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
        lbl(11).Enabled = False
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    Else
        lbl(11).Enabled = True
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
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
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(0).text = ""
                    XPTxtSerial(0).text = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(1).text = ""
                    DtpDelayDate.value = Date
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

                Me.ChkInstall.Enabled = True
            Else
                XPTxtValue(1).Enabled = False
                XPTxtValue(1).text = ""
                Me.ChkInstall.Enabled = False
            End If

        Case 2

            If XPChkPayType(2).value = Checked And Me.TxtModFlg.text <> "R" Then
                Me.CmdCheque.Enabled = True
            Else
                Me.CmdCheque.Enabled = False
                Me.lbl(18).Caption = 0
                Me.lbl(19).Caption = 0
                Me.FgCheques.Rows = Me.FgCheques.FixedRows
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
        XPTxtTaxValue.text = ""
        XPTxtTaxValue.Enabled = False
        lbl(22).Enabled = False
        lbl(45).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    CurrentVoucherNo = ""
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    DateChanged = True
End Sub

Private Sub XPTab301_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub printing()
    On Error GoTo ErrTrap
    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowRecieveVoucherData XPTxtBillID.text, , CBoBasedON.text, DCboStoreName2.text
        End If

    Else

        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowRecieveVoucherData XPTxtBillID.text, True, CBoBasedON.text, DCboStoreName2.text
        End If
    End If

    Exit Sub
ErrTrap:
End Sub
'Private Sub printing()
'    On Error GoTo ErrTrap
'
'    Dim ShowType As Boolean
'    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)
'
'    If ShowType = True Then
'        If Not XPTxtBillID.text Then
'            Set BuyReport = New ClsBuyReport
'            BuyReport.ShowBuyData XPTxtBillID.text, 1, True, Round(LblTotal.Caption * val(txt_Currency_rate), 2), txtManualNO.text, Me.DcCurrency.text
'
'
'        End If
'
'    Else
'
'        If Not XPTxtBillID.text Then
'            Set BuyReport = New ClsBuyReport
'            BuyReport.ShowBuyDataShort XPTxtBillID.text
'        End If
'    End If
'
'    Exit Sub
'ErrTrap:
'End Sub

Private Function AvailableDeal() As Boolean
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))

            '        If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) <> "" Then
            '            If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = True Then
            If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If

                '            End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

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
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.text))

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

    If Me.TxtModFlg.text = "" Then Exit Sub
    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If CboPayMentType.ListIndex = 0 Then
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPChkPayType(0).value = Checked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = XPTxtSum.text
            XPTxtValue(1).text = 0
            '        DBCboClientName.Enabled = False
'            DBCboClientName.Text = ""
            DcboBox.Enabled = True
        Else
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = 0
            XPTxtValue(1).text = XPTxtSum.text
            '         DBCboClientName.Enabled = True
            DcboBox.Enabled = False
            DcboBox.text = ""
        End If
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.text, 0)
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap

    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).text = XPTxtSum.text
    End If

    Me.LblTotal.Caption = XPTxtSum.text
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
            Msg = "—ﬁ„ «·›« Ê—… „ÊÃÊœ „”»ﬁ« ›Ï «·»—‰«„Ã øø" & Chr(13)
            Msg = Msg + "„⁄·Ê„«  ⁄‰ «·›« Ê—… «·„”Ã·…:-" & Chr(13)
        
            Msg = Msg + "—ﬁ„ «·›« Ê—… ›Ï «·»—‰«„Ã:" & rs("Transaction_ID").value & Chr(13)
            Msg = Msg + "„”·”· «·›« Ê—…:" & rs("Transaction_Serial").value & Chr(13)
            Msg = Msg + " «—ÌŒ  ”ÃÌ· «·›« Ê—…:" & rs("Transaction_Date").value & Chr(13)
            Msg = Msg + "«”„ «·⁄„Ì· «Ê «·„Ê—œ:" & rs("CusName").value & Chr(13)
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
