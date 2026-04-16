VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmBasicDataINvArch 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7575
   ClientLeft      =   5745
   ClientTop       =   2430
   ClientWidth     =   12180
   Icon            =   "FrmBasicDataINvArch.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   12180
   Visible         =   0   'False
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
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12180
      _cx             =   21484
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7290
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   11925
         _cx             =   21034
         _cy             =   12859
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
         Caption         =   "»Ì«‰«  «·√—‘Ì›|»Ì«‰«  «·€—›|»Ì«‰«  ’‰«œÌﬁ «·Õ›Ÿ|»Ì«‰«  «·√—››|»Ì«‰«  √‰Ê«⁄ «·„⁄«„·« "
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   6870
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   11835
            _cx             =   20876
            _cy             =   12118
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
            Begin C1SizerLibCtl.C1Elastic EleHeader 
               Height          =   585
               Left            =   -120
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   0
               Width           =   11835
               _cx             =   20876
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
               Appearance      =   0
               MousePointer    =   0
               Version         =   801
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "»Ì«‰«  «·√—‘Ì›«    "
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   2
               ChildSpacing    =   1
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
               CaptionStyle    =   1
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
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   8040
                  Visible         =   0   'False
                  Width           =   855
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic12 
               Height          =   6615
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   600
               Width           =   11775
               _cx             =   20770
               _cy             =   11668
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
               Begin VB.TextBox CodeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   5460
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   3660
                  Width           =   3975
               End
               Begin VB.TextBox NameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   3630
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   4080
                  Width           =   5805
               End
               Begin VB.TextBox NameeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   3630
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   4395
                  Width           =   5805
               End
               Begin ImpulseButton.ISButton btnAdd 
                  Height          =   405
                  Index           =   0
                  Left            =   1395
                  TabIndex        =   9
                  Top             =   3660
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   714
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":038A
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnDelete 
                  Height          =   390
                  Index           =   0
                  Left            =   1395
                  TabIndex        =   10
                  Top             =   4650
                  Width           =   720
                  _ExtentX        =   1270
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":0724
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnEdit 
                  Height          =   330
                  Index           =   0
                  Left            =   1395
                  TabIndex        =   11
                  Top             =   4155
                  Width           =   750
                  _ExtentX        =   1323
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":6F86
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2775
                  Index           =   0
                  Left            =   120
                  TabIndex        =   12
                  Top             =   360
                  Width           =   11505
                  _cx             =   20294
                  _cy             =   4895
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
                  Rows            =   50
                  Cols            =   7
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataINvArch.frx":7320
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   0   'False
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
               Begin MSDataListLib.DataCombo EmpDepDC 
                  Height          =   315
                  Index           =   0
                  Left            =   3630
                  TabIndex        =   13
                  Top             =   4770
                  Width           =   5805
                  _ExtentX        =   10239
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   34
                  Left            =   9645
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   4080
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   300
                  Index           =   33
                  Left            =   9615
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   3705
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   330
                  Index           =   32
                  Left            =   9645
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   4395
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ”„"
                  Height          =   315
                  Index           =   0
                  Left            =   9765
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   4770
                  Width           =   855
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   6870
            Left            =   12570
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Width           =   11835
            _cx             =   20876
            _cy             =   12118
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic9 
               Height          =   6255
               Left            =   0
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   600
               Width           =   11655
               _cx             =   20558
               _cy             =   11033
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
               Begin VB.TextBox CodeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   312
                  Index           =   1
                  Left            =   6900
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   3660
                  Width           =   2745
               End
               Begin VB.TextBox NameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Index           =   1
                  Left            =   3780
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   4035
                  Width           =   5865
               End
               Begin VB.TextBox NameeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   1
                  Left            =   3780
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   4500
                  Width           =   5865
               End
               Begin ImpulseButton.ISButton btnAdd 
                  Height          =   450
                  Index           =   1
                  Left            =   1320
                  TabIndex        =   23
                  Top             =   3900
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   794
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":7413
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnDelete 
                  Height          =   435
                  Index           =   1
                  Left            =   1320
                  TabIndex        =   24
                  Top             =   4935
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   767
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":77AD
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnEdit 
                  Height          =   270
                  Index           =   1
                  Left            =   1320
                  TabIndex        =   25
                  Top             =   4500
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":E00F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2955
                  Index           =   1
                  Left            =   90
                  TabIndex        =   26
                  Top             =   360
                  Width           =   11895
                  _cx             =   20981
                  _cy             =   5212
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
                  Rows            =   50
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataINvArch.frx":E3A9
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   0   'False
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
               Begin MSDataListLib.DataCombo EmpDepDC 
                  Height          =   315
                  Index           =   1
                  Left            =   3780
                  TabIndex        =   27
                  Top             =   4935
                  Width           =   5865
                  _ExtentX        =   10345
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo ArchDC 
                  Height          =   315
                  Index           =   0
                  Left            =   3780
                  TabIndex        =   28
                  Top             =   5355
                  Width           =   5865
                  _ExtentX        =   10345
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   405
                  Index           =   23
                  Left            =   9825
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   4035
                  Width           =   960
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   315
                  Index           =   24
                  Left            =   9825
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   3660
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   375
                  Index           =   25
                  Left            =   9825
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   4500
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ”„ "
                  Height          =   360
                  Index           =   4
                  Left            =   9825
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   4935
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·√—‘Ì›"
                  Height          =   300
                  Index           =   8
                  Left            =   9825
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   5355
                  Width           =   975
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic10 
               Height          =   585
               Left            =   -120
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   0
               Width           =   11835
               _cx             =   20876
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
               Appearance      =   0
               MousePointer    =   0
               Version         =   801
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "»Ì«‰«  «·€—›   "
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   2
               ChildSpacing    =   1
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
               CaptionStyle    =   1
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
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   855
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic16 
            Height          =   6870
            Left            =   12870
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   45
            Width           =   11835
            _cx             =   20876
            _cy             =   12118
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic18 
               Height          =   585
               Left            =   -240
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   0
               Width           =   12075
               _cx             =   21299
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
               Appearance      =   0
               MousePointer    =   0
               Version         =   801
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "»Ì«‰«  ’‰«œÌﬁ «·Õ›Ÿ   "
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   2
               ChildSpacing    =   1
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
               CaptionStyle    =   1
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
               Begin VB.TextBox Text25 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   855
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic22 
               Height          =   6255
               Left            =   0
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   600
               Width           =   11775
               _cx             =   20770
               _cy             =   11033
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
               Begin VB.TextBox CodeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   312
                  Index           =   2
                  Left            =   6765
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   3480
                  Width           =   3420
               End
               Begin VB.TextBox NameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   2
                  Left            =   4605
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   3885
                  Width           =   5580
               End
               Begin VB.TextBox NameeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   2
                  Left            =   4605
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   4200
                  Width           =   5580
               End
               Begin ImpulseButton.ISButton btnAdd 
                  Height          =   390
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   43
                  Top             =   3720
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":E4DD
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnDelete 
                  Height          =   390
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   44
                  Top             =   4680
                  Width           =   720
                  _ExtentX        =   1270
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":E877
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnEdit 
                  Height          =   330
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   45
                  Top             =   4200
                  Width           =   750
                  _ExtentX        =   1323
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":150D9
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2715
                  Index           =   2
                  Left            =   120
                  TabIndex        =   46
                  Top             =   360
                  Width           =   11865
                  _cx             =   20929
                  _cy             =   4789
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
                  Rows            =   50
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataINvArch.frx":15473
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   0   'False
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
               Begin MSDataListLib.DataCombo EmpDepDC 
                  Height          =   315
                  Index           =   2
                  Left            =   4605
                  TabIndex        =   47
                  Top             =   4560
                  Width           =   5580
                  _ExtentX        =   9843
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo ArchDC 
                  Height          =   315
                  Index           =   1
                  Left            =   4605
                  TabIndex        =   48
                  Top             =   4920
                  Width           =   5580
                  _ExtentX        =   9843
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo RoomDC 
                  Height          =   315
                  Index           =   0
                  Left            =   4605
                  TabIndex        =   49
                  Top             =   5280
                  Width           =   5580
                  _ExtentX        =   9843
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   62
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   3885
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   63
                  Left            =   10290
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   3525
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   64
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   4200
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ”„"
                  Height          =   315
                  Index           =   9
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   4560
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·√—‘Ì›"
                  Height          =   315
                  Index           =   10
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   4920
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·€—›…"
                  Height          =   315
                  Index           =   11
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   5280
                  Width           =   975
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6870
            Left            =   13170
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   45
            Width           =   11835
            _cx             =   20876
            _cy             =   12118
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   585
               Left            =   -240
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   0
               Width           =   12075
               _cx             =   21299
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
               Appearance      =   0
               MousePointer    =   0
               Version         =   801
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "»Ì«‰«  «·√—››   "
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   2
               ChildSpacing    =   1
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
               CaptionStyle    =   1
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
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   855
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   6495
               Left            =   0
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   600
               Width           =   11775
               _cx             =   20770
               _cy             =   11456
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
               Begin VB.TextBox NameeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   3
                  Left            =   4080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   4080
                  Width           =   5460
               End
               Begin VB.TextBox NameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   3
                  Left            =   4080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   3765
                  Width           =   5460
               End
               Begin VB.TextBox CodeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   312
                  Index           =   3
                  Left            =   6360
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   3390
                  Width           =   3180
               End
               Begin ImpulseButton.ISButton btnAdd 
                  Height          =   390
                  Index           =   3
                  Left            =   1440
                  TabIndex        =   63
                  Top             =   3840
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":155E8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnDelete 
                  Height          =   390
                  Index           =   3
                  Left            =   1440
                  TabIndex        =   64
                  Top             =   4800
                  Width           =   720
                  _ExtentX        =   1270
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":15982
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnEdit 
                  Height          =   330
                  Index           =   3
                  Left            =   1440
                  TabIndex        =   65
                  Top             =   4320
                  Width           =   750
                  _ExtentX        =   1323
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":1C1E4
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2715
                  Index           =   3
                  Left            =   120
                  TabIndex        =   66
                  Top             =   360
                  Width           =   11865
                  _cx             =   20929
                  _cy             =   4789
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
                  Rows            =   50
                  Cols            =   13
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataINvArch.frx":1C57E
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   0   'False
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
               Begin MSDataListLib.DataCombo EmpDepDC 
                  Height          =   315
                  Index           =   3
                  Left            =   4080
                  TabIndex        =   67
                  Top             =   4440
                  Width           =   5460
                  _ExtentX        =   9631
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo ArchDC 
                  Height          =   315
                  Index           =   2
                  Left            =   4080
                  TabIndex        =   68
                  Top             =   4800
                  Width           =   5460
                  _ExtentX        =   9631
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo RoomDC 
                  Height          =   315
                  Index           =   1
                  Left            =   4080
                  TabIndex        =   69
                  Top             =   5160
                  Width           =   5460
                  _ExtentX        =   9631
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo BoxDC 
                  Height          =   315
                  Index           =   0
                  Left            =   4080
                  TabIndex        =   70
                  Top             =   5520
                  Width           =   5460
                  _ExtentX        =   9631
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   1
                  Left            =   10050
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   4080
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   2
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   3405
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   3
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   3765
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ”„"
                  Height          =   315
                  Index           =   12
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   4440
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·√—‘Ì›"
                  Height          =   315
                  Index           =   13
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   4800
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·€—›…"
                  Height          =   315
                  Index           =   14
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   5160
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’‰œÊﬁ"
                  Height          =   315
                  Index           =   15
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   5520
                  Width           =   975
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   6870
            Left            =   13470
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   45
            Width           =   11835
            _cx             =   20876
            _cy             =   12118
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
               Height          =   585
               Left            =   -240
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   0
               Width           =   12075
               _cx             =   21299
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
               Appearance      =   0
               MousePointer    =   0
               Version         =   801
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "√‰Ê«⁄ «·„⁄«„·«    "
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   2
               ChildSpacing    =   1
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
               CaptionStyle    =   1
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
               Begin VB.TextBox Text13 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   855
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   6255
               Left            =   0
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   600
               Width           =   11775
               _cx             =   20770
               _cy             =   11033
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
               Begin VB.TextBox NameeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   4
                  Left            =   2280
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   3810
                  Width           =   7320
               End
               Begin VB.TextBox NameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   4
                  Left            =   2280
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   3510
                  Width           =   7320
               End
               Begin VB.TextBox CodeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   312
                  Index           =   4
                  Left            =   4620
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   3120
                  Width           =   4980
               End
               Begin VB.TextBox TimeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   3900
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   5550
                  Width           =   1200
               End
               Begin VB.ComboBox TimeUnitCB 
                  Height          =   315
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Text            =   "Combo1"
                  Top             =   5550
                  Width           =   1515
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2595
                  Index           =   4
                  Left            =   120
                  TabIndex        =   87
                  Top             =   360
                  Width           =   11385
                  _cx             =   20082
                  _cy             =   4577
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
                  Rows            =   50
                  Cols            =   17
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataINvArch.frx":1C730
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   0   'False
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
               Begin ImpulseButton.ISButton btnAdd 
                  Height          =   390
                  Index           =   4
                  Left            =   900
                  TabIndex        =   88
                  Top             =   3690
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":1C96B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnDelete 
                  Height          =   390
                  Index           =   4
                  Left            =   900
                  TabIndex        =   89
                  Top             =   4620
                  Width           =   720
                  _ExtentX        =   1270
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":1CD05
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnEdit 
                  Height          =   300
                  Index           =   4
                  Left            =   900
                  TabIndex        =   90
                  Top             =   4170
                  Width           =   750
                  _ExtentX        =   1323
                  _ExtentY        =   529
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
                  ButtonImage     =   "FrmBasicDataINvArch.frx":23567
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo EmpDepDC 
                  Height          =   315
                  Index           =   4
                  Left            =   2280
                  TabIndex        =   91
                  Top             =   4170
                  Width           =   7320
                  _ExtentX        =   12912
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo ArchDC 
                  Height          =   315
                  Index           =   3
                  Left            =   2280
                  TabIndex        =   92
                  Top             =   4500
                  Width           =   7320
                  _ExtentX        =   12912
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo RoomDC 
                  Height          =   315
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   93
                  Top             =   4860
                  Width           =   7320
                  _ExtentX        =   12912
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo BoxDC 
                  Height          =   315
                  Index           =   1
                  Left            =   2280
                  TabIndex        =   94
                  Top             =   5190
                  Width           =   7320
                  _ExtentX        =   12912
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo ShelfDC 
                  Height          =   315
                  Left            =   6360
                  TabIndex        =   95
                  Top             =   5550
                  Width           =   3240
                  _ExtentX        =   5715
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   5
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   3810
                  Width           =   2235
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   6
                  Left            =   8490
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   3165
                  Width           =   2265
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   300
                  Index           =   7
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   3510
                  Width           =   2235
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ”„"
                  Height          =   285
                  Index           =   16
                  Left            =   9780
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   4170
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·√—‘Ì›"
                  Height          =   315
                  Index           =   17
                  Left            =   9780
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   4500
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·€—›…"
                  Height          =   285
                  Index           =   18
                  Left            =   9780
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   4860
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’‰œÊﬁ"
                  Height          =   315
                  Index           =   19
                  Left            =   9780
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   5190
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—›"
                  Height          =   300
                  Index           =   20
                  Left            =   9780
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   5550
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  «·„⁄«„·…"
                  Height          =   300
                  Index           =   21
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   5550
                  Width           =   975
               End
            End
         End
      End
   End
End
Attribute VB_Name = "FrmBasicDataINvArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Indx As Integer
Dim rs_Arch As ADODB.Recordset
Dim rs_Room As ADODB.Recordset
Dim rs_Box As ADODB.Recordset
Dim rs_Shelf As ADODB.Recordset
Dim rs_DocType As ADODB.Recordset
Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos

    On Error GoTo ErrTrap
    
     C1Tab1.TabVisible(0) = False
     C1Tab1.TabVisible(1) = False
     C1Tab1.TabVisible(2) = False
     C1Tab1.TabVisible(3) = False
     C1Tab1.TabVisible(4) = False
  
     C1Tab1.TabVisible(Indx) = True
     C1Tab1.CurrTab = Indx
    
    Dcombos.GetEmpDepartments Me.EmpDepDC(0)
    Dcombos.GetEmpDepartments Me.EmpDepDC(1)
    Dcombos.GetEmpDepartments Me.EmpDepDC(2)
    Dcombos.GetEmpDepartments Me.EmpDepDC(3)
    Dcombos.GetEmpDepartments Me.EmpDepDC(4)
    
    'Dcombos.GetArch Me.ArchDC(1)
    'Dcombos.GetArch Me.ArchDC(2)
    'Dcombos.GetArch Me.ArchDC(3)
    
    'Dcombos.GetRoom Me.RoomDC(0)
    'Dcombos.GetRoom Me.RoomDC(1)
    'Dcombos.GetRoom Me.RoomDC(2)
    
    'Dcombos.GetBox Me.BoxDC(0)
    'Dcombos.GetBox Me.BoxDC(1)
    
    'Dcombos.GetShelf Me.ShelfDC
    
    If SystemOptions.UserInterface = ArabicInterface Then
        With TimeUnitCB
            .Clear
            .AddItem ("œﬁÌﬁ…")
            .AddItem ("”«⁄…")
            .AddItem ("ÌÊ„")
            .AddItem ("‘Â—")
        End With
    Else
        With TimeUnitCB
            .Clear
            .AddItem ("Min")
            .AddItem ("Hour")
            .AddItem ("Day")
            .AddItem ("Month")
        End With
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
            SetInterface Me
        ChangeLang
        
    End If
    Retrive_Arch
    Retrive_Room
    Retrive_Box
    Retrive_Shelf
    Retrive_DocType

ErrTrap:
End Sub
Private Sub ChangeLang()
    EleHeader.Caption = "Archives Data"
    C1Elastic10.Caption = "Rooms Data"
    C1Elastic18.Caption = "Boxes Data"
    C1Elastic3.Caption = "Shelfs Data"
    C1Elastic6.Caption = "Document Types"
    lbl(21).Caption = "Interval"
    EleHeader.Caption = "Archives Data"
    C1Elastic10.Caption = "Rooms Data"
    C1Elastic18.Caption = "Boxes Data"
    C1Elastic3.Caption = "Shelfs Data"
    C1Elastic6.Caption = "Document Types"
    
    C1Tab1.TabCaption(0) = "Archives Data"
    C1Tab1.TabCaption(1) = "Rooms Data"
    C1Tab1.TabCaption(2) = "Boxes Data"
    C1Tab1.TabCaption(3) = "Shelfs Data"
    C1Tab1.TabCaption(4) = "Document Types"
    
    lbl(33).Caption = "Code"
    lbl(24).Caption = "Code"
    lbl(63).Caption = "Code"
    lbl(2).Caption = "Code"
    lbl(6).Caption = "Code"
    
    lbl(34).Caption = "Arabic Name"
    lbl(23).Caption = "Arabic Name"
    lbl(62).Caption = "Arabic Name"
    lbl(3).Caption = "Arabic Name"
    lbl(7).Caption = "Arabic Name"
    
    lbl(32).Caption = "English Name"
    lbl(25).Caption = "English Name"
    lbl(64).Caption = "English Name"
    lbl(1).Caption = "English Name"
    lbl(5).Caption = "English Name"
    
    lbl(0).Caption = "Department"
    lbl(4).Caption = "Department"
    lbl(9).Caption = "Department"
    lbl(12).Caption = "Department"
    lbl(16).Caption = "Department"
    
    lbl(8).Caption = "Archive"
    lbl(10).Caption = "Archive"
    lbl(13).Caption = "Archive"
    lbl(17).Caption = "Archive"
    
    lbl(11).Caption = "Room"
    lbl(14).Caption = "Room"
    lbl(18).Caption = "Room"
    
    lbl(15).Caption = "Box"
    lbl(19).Caption = "Box"
    
    lbl(20).Caption = "Shelf"
    
    btnAdd(0).Caption = "Add"
    btnAdd(1).Caption = "Add"
    btnAdd(2).Caption = "Add"
    btnAdd(3).Caption = "Add"
    btnAdd(4).Caption = "Add"
    
    btnEdit(0).Caption = "Edit"
    btnEdit(1).Caption = "Edit"
    btnEdit(2).Caption = "Edit"
    btnEdit(3).Caption = "Edit"
    btnEdit(4).Caption = "Edit"
    
    btnDelete(0).Caption = "Delete"
    btnDelete(1).Caption = "Delete"
    btnDelete(2).Caption = "Delete"
    btnDelete(3).Caption = "Delete"
    btnDelete(4).Caption = "Delete"
    
    With Grid(0)
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("Name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("NameE")) = "English Name"
        .TextMatrix(0, .ColIndex("Dep")) = "Department"
    End With
    
    With Grid(1)
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("Name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("NameE")) = "English Name"
        .TextMatrix(0, .ColIndex("Dep")) = "Department"
        .TextMatrix(0, .ColIndex("Arch")) = "Archive"
    End With
    
    With Grid(2)
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("Name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("NameE")) = "English Name"
        .TextMatrix(0, .ColIndex("Dep")) = "Department"
        .TextMatrix(0, .ColIndex("Arch")) = "Archive"
        .TextMatrix(0, .ColIndex("Room")) = "Room"
    End With
    
    With Grid(3)
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("Name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("NameE")) = "English Name"
        .TextMatrix(0, .ColIndex("Dep")) = "Department"
        .TextMatrix(0, .ColIndex("Arch")) = "Archive"
        .TextMatrix(0, .ColIndex("Room")) = "Room"
        .TextMatrix(0, .ColIndex("Box")) = "Box"
    End With
    
    With Grid(4)
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("Name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("NameE")) = "English Name"
        .TextMatrix(0, .ColIndex("Dep")) = "Department"
        .TextMatrix(0, .ColIndex("Arch")) = "Archive"
        .TextMatrix(0, .ColIndex("Room")) = "Room"
        .TextMatrix(0, .ColIndex("Box")) = "Box"
        .TextMatrix(0, .ColIndex("Shelf")) = "Shelf"
    End With
    
End Sub
Private Sub EmpDepDC_Change(Index As Integer)
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    On Error GoTo ErrTrap
    Select Case Index
        Case 0
        Case 1
            Dcombos.GetArch ArchDC(0), EmpDepDC(1).BoundText
        Case 2
            Dcombos.GetArch ArchDC(1), EmpDepDC(2).BoundText
            Dcombos.GetRoom Me.RoomDC(0), 0
        Case 3
            Dcombos.GetArch ArchDC(2), EmpDepDC(3).BoundText
            Dcombos.GetRoom Me.RoomDC(1), 0
            Dcombos.GetBox Me.BoxDC(0), 0
        Case 4
            Dcombos.GetArch ArchDC(3), EmpDepDC(4).BoundText
            Dcombos.GetRoom Me.RoomDC(2), 0
            Dcombos.GetBox Me.BoxDC(1), 0
            Dcombos.GetShelf Me.ShelfDC, 0
    End Select
ErrTrap:
End Sub
Private Sub ArchDC_Change(Index As Integer)
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    On Error GoTo ErrTrap
    Select Case Index
        Case 0
        Case 1
            Dcombos.GetRoom Me.RoomDC(0), ArchDC(1).BoundText
        Case 2
            Dcombos.GetRoom Me.RoomDC(1), ArchDC(2).BoundText
            Dcombos.GetBox Me.BoxDC(0), Me.RoomDC(1).BoundText
        Case 3
            Dcombos.GetRoom Me.RoomDC(2), ArchDC(3).BoundText
            Dcombos.GetBox Me.BoxDC(1), 0
            Dcombos.GetShelf Me.ShelfDC, 0
    End Select
ErrTrap:
End Sub
Private Sub RoomDC_Change(Index As Integer)
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    On Error GoTo ErrTrap
    Select Case Index
        Case 0
        Case 1
            Dcombos.GetBox Me.BoxDC(0), Me.RoomDC(1).BoundText
        Case 2
            Dcombos.GetBox Me.BoxDC(1), Me.RoomDC(2).BoundText
            Dcombos.GetShelf Me.ShelfDC, 0
    End Select
ErrTrap:
End Sub
Private Sub BoxDC_Change(Index As Integer)
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    On Error GoTo ErrTrap
    Select Case Index
        Case 0
        Case 1
            Dcombos.GetShelf Me.ShelfDC, Me.BoxDC(1).BoundText
    End Select
ErrTrap:
End Sub
Private Sub btnAdd_Click(Index As Integer)
    Select Case Index
        Case 0
            Add_Arch
        Case 1
            Add_Room
        Case 2
            Add_Box
        Case 3
            Add_Shelf
        Case 4
            Add_DocType
    End Select
End Sub
Private Sub BtnEdit_Click(Index As Integer)
Select Case Index
    Case 0
        Update_Arch
    Case 1
        Update_Room
    Case 2
        Update_Box
    Case 3
        Update_Shelf
    Case 4
        Update_DocType
End Select
End Sub
Private Sub btnDelete_Click(Index As Integer)
    Select Case Index
        Case 0
            Del_Arch
        Case 1
            Del_Room
        Case 2
            Del_Box
        Case 3
            Del_Shelf
        Case 4
            Del_DocType
    End Select
End Sub
Private Sub Grid_Click(Index As Integer)
    Select Case Index
        Case 0
            ArchGridClk
        Case 1
            RoomGridClk
        Case 2
            BoxGridClk
        Case 3
            ShelfGridClk
        Case 4
            DocTypeGridClk
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

ErrTrap:
End Sub
'############################################################### Arch Part ###################################################################
Private Sub Add_Arch()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    On Error GoTo errortrap
  
    If NameTxt(0).Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
    Else
        MsgBox ("Enter Name ")
    End If
    NameTxt(0).SetFocus
    Exit Sub
    End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_Arch = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblXXArch"
    rs_Arch.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_Arch.AddNew

    rs_Arch("Code") = IIf(CodeTxt(0).Text = "", Null, CodeTxt(0).Text)
    rs_Arch("Name") = IIf(NameTxt(0).Text = "", Null, NameTxt(0).Text)
    rs_Arch("Namee") = IIf(NameeTxt(0).Text = "", Null, NameeTxt(0).Text)
    rs_Arch("DepID") = IIf(val(EmpDepDC(0).BoundText) = 0, Null, val(EmpDepDC(0).BoundText))
    
    rs_Arch.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
        MsgBox ("Data saved successfully")
    End If
    Retrive_Arch
    Clear_Arch
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Retrive_Arch()
    Dim i As Integer
    Set rs_Arch = New ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = " SELECT TblXXArch.ID, TblXXArch.Code, TblXXArch.Name, TblXXArch.Namee, TblXXArch.DepID, TblEmpDepartments.DepartmentName, TblEmpDepartments.DepartmentNamee"
    StrSQL = StrSQL & " FROM TblXXArch INNER JOIN "
    StrSQL = StrSQL & " TblEmpDepartments ON TblXXArch.DepID = TblEmpDepartments.DeparmentID "
    
    rs_Arch.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Grid(0).Rows = 1
    
    If rs_Arch.RecordCount > 0 Then
        rs_Arch.MoveFirst
        With Grid(0)
            .Rows = rs_Arch.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs_Arch("ID").value), 0, rs_Arch("ID").value)
                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs_Arch("Code").value), "", rs_Arch("Code").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_Arch("name").value), "", rs_Arch("name").value)
                .TextMatrix(i, .ColIndex("nameE")) = IIf(IsNull(rs_Arch("namee").value), "", rs_Arch("namee").value)
                .TextMatrix(i, .ColIndex("DepID")) = IIf(IsNull(rs_Arch("DepID").value), 0, rs_Arch("DepID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_Arch("DepartmentName").value), "", rs_Arch("DepartmentName").value)
                Else
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_Arch("DepartmentNamee").value), "", rs_Arch("DepartmentNamee").value)
                End If
          rs_Arch.MoveNext
            Next
         End With
    End If
End Sub
Private Sub Update_Arch()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    Dim str As String, sr As String
    
    On Error GoTo errortrap

    If NameTxt(0).Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("«œŒ· «·«”„ «Ê·«")
        Else
            MsgBox ("Enter Name ")
        End If
        NameTxt(0).SetFocus
        Exit Sub
    End If
 
    str = Grid(0).TextMatrix(Grid(0).Row, Grid(0).ColIndex("ID"))
    sr = Grid(0).TextMatrix(Grid(0).Row, Grid(0).ColIndex("serial"))
        
    If str <> "" Then
        Cn.BeginTrans
        BeginTrans = True
        StrSQL = "Update TblXXArch Set  Namee='" & NameeTxt(0).Text & "',Name='" & NameTxt(0).Text & "',Code = '" & CodeTxt(0).Text & "',DepID=" & val(Me.EmpDepDC(0).BoundText) & " Where ID=" & val(str)
        Cn.Execute StrSQL, , adExecuteNoRecords
           
        Cn.CommitTrans
        BeginTrans = False
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
        Else
            MsgBox ("New data saved ")
        End If
        Retrive_Arch
        Clear_Arch
    End If
        
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    
End Sub
Private Sub Del_Arch()
    Dim Msg As String
    Dim StrSQL As String
    Dim str As String, sr As String
    Set rs_Arch = New ADODB.Recordset
    
    On Error GoTo ErrTrap
        
        str = Grid(0).TextMatrix(Grid(0).Row, Grid(0).ColIndex("ID"))
        sr = Grid(0).TextMatrix(Grid(0).Row, Grid(0).ColIndex("serial"))
        
        If str <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
                Msg = Msg + (sr) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Are you sure you want to delete  " & CHR(13)
                Msg = Msg + "Data in row No."
                Msg = Msg + (sr) & CHR(13)
            End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
                StrSQL = "SELECT  *  From TblXXArch"
                rs_Arch.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs_Arch.RecordCount < 1 Then
                
                    StrSQL = "delete From TblXXArch where  ID =" & val(str)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    
                    StrSQL = "SELECT  *  From TblXXArch"
                    rs_Arch.Close
                    rs_Arch.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If rs_Arch.RecordCount < 1 Then
                    Else
                        Retrive_Arch
                    End If
                    
                End If
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
            Else
                Msg = "This operation is not available due to lack of records"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Retrive_Arch
    Clear_Arch
    
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs_Arch.CancelUpdate
    'End If
End Sub
Private Sub ArchGridClk()
    With Grid(0)
        If .Row > 0 Then
            CodeTxt(0).Text = .TextMatrix(.Row, .ColIndex("Code"))
            NameTxt(0).Text = .TextMatrix(.Row, .ColIndex("Name"))
            NameeTxt(0).Text = .TextMatrix(.Row, .ColIndex("Namee"))
            EmpDepDC(0).BoundText = val(.TextMatrix(.Row, .ColIndex("DepID")))
        End If
    End With
End Sub
Private Sub Clear_Arch()
    CodeTxt(0).Text = ""
    NameTxt(0).Text = ""
    NameeTxt(0).Text = ""
    EmpDepDC(0).BoundText = 0
    EmpDepDC(0).Text = ""
End Sub
'############################################################### Room Part ###################################################################
Private Sub Add_Room()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    On Error GoTo errortrap
  
    If NameTxt(1).Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
    Else
        MsgBox ("Enter Name ")
    End If
    NameTxt(1).SetFocus
    Exit Sub
    End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_Room = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblXXRoom"
    rs_Room.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_Room.AddNew

    rs_Room("Code") = IIf(CodeTxt(1).Text = "", Null, CodeTxt(1).Text)
    rs_Room("Name") = IIf(NameTxt(1).Text = "", Null, NameTxt(1).Text)
    rs_Room("Namee") = IIf(NameeTxt(1).Text = "", Null, NameeTxt(1).Text)
    rs_Room("DepID") = IIf(val(EmpDepDC(1).BoundText) = 0, Null, val(EmpDepDC(1).BoundText))
    rs_Room("ArchID") = IIf(val(ArchDC(0).BoundText) = 0, Null, val(ArchDC(0).BoundText))
    
    rs_Room.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
        MsgBox ("Data saved successfully")
    End If
    Retrive_Room
    Clear_Room
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Retrive_Room()
    Dim i As Integer
    Set rs_Room = New ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = "SELECT TblXXRoom.ID, TblXXRoom.Code, TblXXRoom.Name, TblXXRoom.Namee, TblXXRoom.DepID, TblXXRoom.ArchID, TblEmpDepartments.DepartmentName, TblEmpDepartments.DepartmentNamee, TblXXArch.Name AS ArchName, "
    StrSQL = StrSQL & " TblXXArch.Namee AS ArchNamee "
    StrSQL = StrSQL & " FROM TblXXRoom INNER JOIN "
    StrSQL = StrSQL & " TblEmpDepartments ON TblXXRoom.DepID = TblEmpDepartments.DeparmentID INNER JOIN "
    StrSQL = StrSQL & " TblXXArch ON TblXXRoom.ArchID = TblXXArch.ID "
    
    rs_Room.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Grid(1).Rows = 1
    
    If rs_Room.RecordCount > 0 Then
        rs_Room.MoveFirst
        With Grid(1)
            .Rows = rs_Room.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs_Room("ID").value), 0, rs_Room("ID").value)
                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs_Room("Code").value), "", rs_Room("Code").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_Room("name").value), "", rs_Room("name").value)
                .TextMatrix(i, .ColIndex("nameE")) = IIf(IsNull(rs_Room("namee").value), "", rs_Room("namee").value)
                .TextMatrix(i, .ColIndex("DepID")) = IIf(IsNull(rs_Room("DepID").value), 0, rs_Room("DepID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_Room("DepartmentName").value), "", rs_Room("DepartmentName").value)
                Else
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_Room("DepartmentNamee").value), "", rs_Room("DepartmentNamee").value)
                End If
                .TextMatrix(i, .ColIndex("ArchID")) = IIf(IsNull(rs_Room("ArchID").value), 0, rs_Room("ArchID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Arch")) = IIf(IsNull(rs_Room("ArchName").value), "", rs_Room("ArchName").value)
                Else
                    .TextMatrix(i, .ColIndex("Arch")) = IIf(IsNull(rs_Room("ArchNamee").value), "", rs_Room("ArchNamee").value)
                End If
          rs_Room.MoveNext
            Next
         End With
    End If
End Sub
Private Sub Update_Room()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    Dim str As String, sr As String
    
    On Error GoTo errortrap

    If NameTxt(1).Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("«œŒ· «·«”„ «Ê·«")
        Else
            MsgBox ("Enter Name ")
        End If
        NameTxt(1).SetFocus
        Exit Sub
    End If
 
    str = Grid(1).TextMatrix(Grid(1).Row, Grid(1).ColIndex("ID"))
    sr = Grid(1).TextMatrix(Grid(1).Row, Grid(1).ColIndex("serial"))
        
    If str <> "" Then
        Cn.BeginTrans
        BeginTrans = True
        StrSQL = "Update TblXXRoom Set  Namee='" & NameeTxt(1).Text & "',Name='" & NameTxt(1).Text & "',Code = '" & CodeTxt(1).Text & "',DepID=" & val(Me.EmpDepDC(1).BoundText) & ",ArchID=" & val(Me.ArchDC(0).BoundText) & " Where ID=" & val(str)
        Cn.Execute StrSQL, , adExecuteNoRecords
           
        Cn.CommitTrans
        BeginTrans = False
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
        Else
            MsgBox ("New data saved ")
        End If
        Retrive_Room
        Clear_Room
    End If
        
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    
End Sub
Private Sub Del_Room()
    Dim Msg As String
    Dim StrSQL As String
    Dim str As String, sr As String
    Set rs_Room = New ADODB.Recordset
    
    On Error GoTo ErrTrap
        
        str = Grid(1).TextMatrix(Grid(1).Row, Grid(1).ColIndex("ID"))
        sr = Grid(1).TextMatrix(Grid(1).Row, Grid(1).ColIndex("serial"))
        
        If str <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
                Msg = Msg + (sr) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Are you sure you want to delete  " & CHR(13)
                Msg = Msg + "Data in row No."
                Msg = Msg + (sr) & CHR(13)
            End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
                StrSQL = "SELECT  *  From TblXXRoom"
                rs_Room.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs_Room.RecordCount < 1 Then
                
                    StrSQL = "delete From TblXXRoom where  ID =" & val(str)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    
                    StrSQL = "SELECT  *  From TblXXRoom"
                    rs_Room.Close
                    rs_Room.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If rs_Room.RecordCount < 1 Then
                    Else
                        Retrive_Room
                    End If
                    
                End If
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
            Else
                Msg = "This operation is not available due to lack of records"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Retrive_Room
    Clear_Room
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs_Room.CancelUpdate
    'End If
End Sub

Private Sub RoomGridClk()
    With Grid(1)
        If .Row > 0 Then
            CodeTxt(1).Text = .TextMatrix(.Row, .ColIndex("Code"))
            NameTxt(1).Text = .TextMatrix(.Row, .ColIndex("Name"))
            NameeTxt(1).Text = .TextMatrix(.Row, .ColIndex("Namee"))
            EmpDepDC(1).BoundText = val(.TextMatrix(.Row, .ColIndex("DepID")))
            ArchDC(0).BoundText = val(.TextMatrix(.Row, .ColIndex("ArchID")))
        End If
    End With
End Sub
Private Sub Clear_Room()
    CodeTxt(1).Text = ""
    NameTxt(1).Text = ""
    NameeTxt(1).Text = ""
    EmpDepDC(1).BoundText = 0
    EmpDepDC(1).Text = ""
    ArchDC(0).BoundText = 0
    ArchDC(0).Text = ""
End Sub
'############################################################### Box Part ###################################################################
Private Sub Add_Box()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    On Error GoTo errortrap
  
    If NameTxt(2).Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
    Else
        MsgBox ("Enter Name ")
    End If
    NameTxt(2).SetFocus
    Exit Sub
    End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_Box = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblXXBox"
    rs_Box.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_Box.AddNew

    rs_Box("Code") = IIf(CodeTxt(2).Text = "", Null, CodeTxt(2).Text)
    rs_Box("Name") = IIf(NameTxt(2).Text = "", Null, NameTxt(2).Text)
    rs_Box("Namee") = IIf(NameeTxt(2).Text = "", Null, NameeTxt(2).Text)
    rs_Box("DepID") = IIf(val(EmpDepDC(2).BoundText) = 0, Null, val(EmpDepDC(2).BoundText))
    rs_Box("ArchID") = IIf(val(ArchDC(1).BoundText) = 0, Null, val(ArchDC(1).BoundText))
    rs_Box("RoomID") = IIf(val(RoomDC(0).BoundText) = 0, Null, val(RoomDC(0).BoundText))
    
    rs_Box.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
        MsgBox ("Data saved successfully")
    End If
    Retrive_Box
    Clear_Box
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Retrive_Box()
    Dim i As Integer
    Set rs_Box = New ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = "SELECT TblXXBox.ID, TblXXBox.Code, TblXXBox.Name, TblXXBox.Namee, TblXXBox.DepID, TblXXBox.ArchID, TblXXBox.RoomID, TblEmpDepartments.DepartmentName, TblEmpDepartments.DepartmentNamee, "
    StrSQL = StrSQL & " TblXXArch.Name AS ArchName, TblXXArch.Namee AS ArchNamee, TblXXRoom.Name AS RoomName, TblXXRoom.Namee AS RoomNamee "
    StrSQL = StrSQL & "FROM TblXXBox INNER JOIN "
    StrSQL = StrSQL & " TblEmpDepartments ON TblXXBox.DepID = TblEmpDepartments.DeparmentID INNER JOIN "
    StrSQL = StrSQL & " TblXXArch ON TblXXBox.ArchID = TblXXArch.ID INNER JOIN "
    StrSQL = StrSQL & " TblXXRoom ON TblXXBox.RoomID = TblXXRoom.ID "
    
    rs_Box.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Grid(2).Rows = 1
    
    If rs_Box.RecordCount > 0 Then
        rs_Box.MoveFirst
        With Grid(2)
            .Rows = rs_Box.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs_Box("ID").value), 0, rs_Box("ID").value)
                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs_Box("Code").value), "", rs_Box("Code").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_Box("name").value), "", rs_Box("name").value)
                .TextMatrix(i, .ColIndex("nameE")) = IIf(IsNull(rs_Box("namee").value), "", rs_Box("namee").value)
                .TextMatrix(i, .ColIndex("DepID")) = IIf(IsNull(rs_Box("DepID").value), 0, rs_Box("DepID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_Box("DepartmentName").value), "", rs_Box("DepartmentName").value)
                Else
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_Box("DepartmentNamee").value), "", rs_Box("DepartmentNamee").value)
                End If
                .TextMatrix(i, .ColIndex("ArchID")) = IIf(IsNull(rs_Box("ArchID").value), 0, rs_Box("ArchID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Arch")) = IIf(IsNull(rs_Box("ArchName").value), "", rs_Box("ArchName").value)
                Else
                    .TextMatrix(i, .ColIndex("Arch")) = IIf(IsNull(rs_Box("ArchNamee").value), "", rs_Box("ArchNamee").value)
                End If
                .TextMatrix(i, .ColIndex("RoomID")) = IIf(IsNull(rs_Box("RoomID").value), 0, rs_Box("RoomID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Room")) = IIf(IsNull(rs_Box("RoomName").value), "", rs_Box("RoomName").value)
                Else
                    .TextMatrix(i, .ColIndex("Room")) = IIf(IsNull(rs_Box("RoomNamee").value), "", rs_Box("RoomNamee").value)
                End If
          rs_Box.MoveNext
            Next
         End With
    End If
End Sub
Private Sub Update_Box()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    Dim str As String, sr As String
    
    On Error GoTo errortrap

    If NameTxt(2).Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("«œŒ· «·«”„ «Ê·«")
        Else
            MsgBox ("Enter Name ")
        End If
        NameTxt(2).SetFocus
        Exit Sub
    End If
 
    str = Grid(2).TextMatrix(Grid(2).Row, Grid(2).ColIndex("ID"))
    sr = Grid(2).TextMatrix(Grid(2).Row, Grid(2).ColIndex("serial"))
        
    If str <> "" Then
        Cn.BeginTrans
        BeginTrans = True
        StrSQL = "Update TblXXBox Set  Namee='" & NameeTxt(2).Text & "',Name='" & NameTxt(2).Text & "',Code = '" & CodeTxt(2).Text & "',DepID=" & val(Me.EmpDepDC(2).BoundText) & ",ArchID=" & val(Me.ArchDC(1).BoundText) & ",RoomID=" & val(Me.RoomDC(0).BoundText) & " Where ID=" & val(str)
        Cn.Execute StrSQL, , adExecuteNoRecords
           
        Cn.CommitTrans
        BeginTrans = False
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
        Else
            MsgBox ("New data saved ")
        End If
        Retrive_Box
        Clear_Box
    End If
        
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    
End Sub
Private Sub Del_Box()
    Dim Msg As String
    Dim StrSQL As String
    Dim str As String, sr As String
    Set rs_Box = New ADODB.Recordset
    
    On Error GoTo ErrTrap
        
        str = Grid(2).TextMatrix(Grid(2).Row, Grid(2).ColIndex("ID"))
        sr = Grid(2).TextMatrix(Grid(2).Row, Grid(2).ColIndex("serial"))
        
        If str <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
                Msg = Msg + (sr) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Are you sure you want to delete  " & CHR(13)
                Msg = Msg + "Data in row No."
                Msg = Msg + (sr) & CHR(13)
            End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
                StrSQL = "SELECT  *  From TblXXBox"
                rs_Box.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs_Room.RecordCount < 1 Then
                
                    StrSQL = "delete From TblXXBox where  ID =" & val(str)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    
                    StrSQL = "SELECT  *  From TblXXBox"
                    rs_Box.Close
                    rs_Box.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If rs_Box.RecordCount < 1 Then
                    Else
                        Retrive_Box
                    End If
                    
                End If
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
            Else
                Msg = "This operation is not available due to lack of records"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Retrive_Box
    Clear_Box
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs_Box.CancelUpdate
    'End If
End Sub
Private Sub BoxGridClk()
    With Grid(2)
        If .Row > 0 Then
            CodeTxt(2).Text = .TextMatrix(.Row, .ColIndex("Code"))
            NameTxt(2).Text = .TextMatrix(.Row, .ColIndex("Name"))
            NameeTxt(2).Text = .TextMatrix(.Row, .ColIndex("Namee"))
            EmpDepDC(2).BoundText = val(.TextMatrix(.Row, .ColIndex("DepID")))
            ArchDC(1).BoundText = val(.TextMatrix(.Row, .ColIndex("ArchID")))
            RoomDC(0).BoundText = val(.TextMatrix(.Row, .ColIndex("RoomID")))
        End If
    End With
End Sub
Private Sub Clear_Box()
    CodeTxt(2).Text = ""
    NameTxt(2).Text = ""
    NameeTxt(2).Text = ""
    EmpDepDC(2).BoundText = 0
    EmpDepDC(2).Text = ""
    ArchDC(1).BoundText = 0
    ArchDC(1).Text = ""
    RoomDC(0).BoundText = 0
    RoomDC(0).Text = ""
End Sub
'############################################################### Shelf Part ###################################################################
Private Sub Add_Shelf()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    On Error GoTo errortrap
  
    If NameTxt(3).Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
    Else
        MsgBox ("Enter Name ")
    End If
    NameTxt(3).SetFocus
    Exit Sub
    End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_Shelf = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblXXShelf"
    rs_Shelf.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_Shelf.AddNew

    rs_Shelf("Code") = IIf(CodeTxt(3).Text = "", Null, CodeTxt(3).Text)
    rs_Shelf("Name") = IIf(NameTxt(3).Text = "", Null, NameTxt(3).Text)
    rs_Shelf("Namee") = IIf(NameeTxt(3).Text = "", Null, NameeTxt(3).Text)
    rs_Shelf("DepID") = IIf(val(EmpDepDC(3).BoundText) = 0, Null, val(EmpDepDC(3).BoundText))
    rs_Shelf("ArchID") = IIf(val(ArchDC(2).BoundText) = 0, Null, val(ArchDC(2).BoundText))
    rs_Shelf("RoomID") = IIf(val(RoomDC(1).BoundText) = 0, Null, val(RoomDC(1).BoundText))
    rs_Shelf("BoxID") = IIf(val(BoxDC(0).BoundText) = 0, Null, val(BoxDC(0).BoundText))
    
    rs_Shelf.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
        MsgBox ("Data saved successfully")
    End If
    Retrive_Shelf
    Clear_Shelf
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Retrive_Shelf()
    Dim i As Integer
    Set rs_Shelf = New ADODB.Recordset
    Dim StrSQL As String
        
    StrSQL = "SELECT TblXXShelf.ID, TblXXShelf.Code, TblXXShelf.Name, TblXXShelf.Namee, TblXXShelf.DepID, TblXXShelf.ArchID, TblXXShelf.RoomID, TblXXShelf.BoxID, TblEmpDepartments.DepartmentName, "
    StrSQL = StrSQL & " TblEmpDepartments.DepartmentNamee, TblXXArch.Name AS ArchName, TblXXArch.Namee AS ArchNamee, TblXXRoom.Name AS RoomName, TblXXRoom.Namee AS RoomNamee, TblXXBox.Name AS BoxName, "
    StrSQL = StrSQL & " TblXXBox.Namee AS BoxNamee "
    StrSQL = StrSQL & " FROM TblXXShelf INNER JOIN "
    StrSQL = StrSQL & " TblEmpDepartments ON TblXXShelf.DepID = TblEmpDepartments.DeparmentID INNER JOIN "
    StrSQL = StrSQL & " TblXXArch ON TblXXShelf.ArchID = TblXXArch.ID INNER JOIN "
    StrSQL = StrSQL & " TblXXRoom ON TblXXShelf.RoomID = TblXXRoom.ID INNER JOIN "
    StrSQL = StrSQL & " TblXXBox ON TblXXShelf.BoxID = TblXXBox.ID "
    
    rs_Shelf.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Grid(3).Rows = 1
    
    If rs_Shelf.RecordCount > 0 Then
        rs_Shelf.MoveFirst
        With Grid(3)
            .Rows = rs_Shelf.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs_Shelf("ID").value), 0, rs_Shelf("ID").value)
                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs_Shelf("Code").value), "", rs_Shelf("Code").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_Shelf("name").value), "", rs_Shelf("name").value)
                .TextMatrix(i, .ColIndex("nameE")) = IIf(IsNull(rs_Shelf("namee").value), "", rs_Shelf("namee").value)
                
                .TextMatrix(i, .ColIndex("DepID")) = IIf(IsNull(rs_Shelf("DepID").value), 0, rs_Shelf("DepID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_Shelf("DepartmentName").value), "", rs_Shelf("DepartmentName").value)
                Else
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_Shelf("DepartmentNamee").value), "", rs_Shelf("DepartmentNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("ArchID")) = IIf(IsNull(rs_Shelf("ArchID").value), 0, rs_Shelf("ArchID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Arch")) = IIf(IsNull(rs_Shelf("ArchName").value), "", rs_Shelf("ArchName").value)
                Else
                    .TextMatrix(i, .ColIndex("Arch")) = IIf(IsNull(rs_Shelf("ArchNamee").value), "", rs_Shelf("ArchNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("RoomID")) = IIf(IsNull(rs_Shelf("RoomID").value), 0, rs_Shelf("RoomID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Room")) = IIf(IsNull(rs_Shelf("RoomName").value), "", rs_Shelf("RoomName").value)
                Else
                    .TextMatrix(i, .ColIndex("Room")) = IIf(IsNull(rs_Shelf("RoomNamee").value), "", rs_Shelf("RoomNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs_Shelf("BoxID").value), 0, rs_Shelf("BoxID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Box")) = IIf(IsNull(rs_Shelf("BoxName").value), "", rs_Shelf("BoxName").value)
                Else
                    .TextMatrix(i, .ColIndex("Box")) = IIf(IsNull(rs_Shelf("BoxNamee").value), "", rs_Shelf("BoxNamee").value)
                End If
          rs_Shelf.MoveNext
            Next
         End With
    End If
End Sub
Private Sub Update_Shelf()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    Dim str As String, sr As String
    
    On Error GoTo errortrap

    If NameTxt(3).Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("«œŒ· «·«”„ «Ê·«")
        Else
            MsgBox ("Enter Name ")
        End If
        NameTxt(3).SetFocus
        Exit Sub
    End If
 
    str = Grid(3).TextMatrix(Grid(3).Row, Grid(3).ColIndex("ID"))
    sr = Grid(3).TextMatrix(Grid(3).Row, Grid(3).ColIndex("serial"))
        
    If str <> "" Then
        Cn.BeginTrans
        BeginTrans = True
        StrSQL = "Update TblXXShelf Set  Namee='" & NameeTxt(3).Text & "',Name='" & NameTxt(3).Text & "',Code = '" & CodeTxt(3).Text & "',DepID=" & val(Me.EmpDepDC(3).BoundText) & ",ArchID=" & val(Me.ArchDC(2).BoundText) & ",RoomID=" & val(Me.RoomDC(1).BoundText) & ",BoxID=" & val(Me.BoxDC(0).BoundText) & " Where ID=" & val(str)
        Cn.Execute StrSQL, , adExecuteNoRecords
           
        Cn.CommitTrans
        BeginTrans = False
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
        Else
            MsgBox ("New data saved ")
        End If
        Retrive_Shelf
        Clear_Shelf
    End If
        
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    
End Sub
Private Sub Del_Shelf()
    Dim Msg As String
    Dim StrSQL As String
    Dim str As String, sr As String
    Set rs_Shelf = New ADODB.Recordset
    
    On Error GoTo ErrTrap
        
        str = Grid(3).TextMatrix(Grid(3).Row, Grid(3).ColIndex("ID"))
        sr = Grid(3).TextMatrix(Grid(3).Row, Grid(3).ColIndex("serial"))
        
        If str <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
                Msg = Msg + (sr) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Are you sure you want to delete  " & CHR(13)
                Msg = Msg + "Data in row No."
                Msg = Msg + (sr) & CHR(13)
            End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
                StrSQL = "SELECT  *  From TblXXShelf"
                rs_Shelf.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs_Room.RecordCount < 1 Then
                
                    StrSQL = "delete From TblXXShelf where  ID =" & val(str)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    
                    StrSQL = "SELECT  *  From TblXXShelf"
                    rs_Shelf.Close
                    rs_Shelf.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If rs_Shelf.RecordCount < 1 Then
                    Else
                        Retrive_Shelf
                    End If
                    
                End If
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
            Else
                Msg = "This operation is not available due to lack of records"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Retrive_Shelf
    Clear_Shelf
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs_Shelf.CancelUpdate
    'End If
End Sub
Private Sub ShelfGridClk()
    With Grid(3)
        If .Row > 0 Then
            CodeTxt(3).Text = .TextMatrix(.Row, .ColIndex("Code"))
            NameTxt(3).Text = .TextMatrix(.Row, .ColIndex("Name"))
            NameeTxt(3).Text = .TextMatrix(.Row, .ColIndex("Namee"))
            EmpDepDC(3).BoundText = val(.TextMatrix(.Row, .ColIndex("DepID")))
            ArchDC(2).BoundText = val(.TextMatrix(.Row, .ColIndex("ArchID")))
            RoomDC(1).BoundText = val(.TextMatrix(.Row, .ColIndex("RoomID")))
            BoxDC(0).BoundText = val(.TextMatrix(.Row, .ColIndex("BoxID")))
        End If
    End With
End Sub
Private Sub Clear_Shelf()
    CodeTxt(3).Text = ""
    NameTxt(3).Text = ""
    NameeTxt(3).Text = ""
    EmpDepDC(3).BoundText = 0
    EmpDepDC(3).Text = ""
    ArchDC(2).BoundText = 0
    ArchDC(2).Text = ""
    RoomDC(1).BoundText = 0
    RoomDC(1).Text = ""
    BoxDC(0).BoundText = 0
    BoxDC(0).Text = ""
End Sub
'############################################################### DocType Part ###################################################################
Private Sub Add_DocType()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    On Error GoTo errortrap
  
    If NameTxt(4).Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
    Else
        MsgBox ("Enter Name ")
    End If
    NameTxt(4).SetFocus
    Exit Sub
    End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_DocType = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblXXArchDocType"
    rs_DocType.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_DocType.AddNew

    rs_DocType("Code") = IIf(CodeTxt(4).Text = "", Null, CodeTxt(4).Text)
    rs_DocType("Name") = IIf(NameTxt(4).Text = "", Null, NameTxt(4).Text)
    rs_DocType("Namee") = IIf(NameeTxt(4).Text = "", Null, NameeTxt(4).Text)
    rs_DocType("DepID") = IIf(val(EmpDepDC(4).BoundText) = 0, Null, val(EmpDepDC(4).BoundText))
    rs_DocType("ArchID") = IIf(val(ArchDC(3).BoundText) = 0, Null, val(ArchDC(3).BoundText))
    rs_DocType("RoomID") = IIf(val(RoomDC(2).BoundText) = 0, Null, val(RoomDC(2).BoundText))
    rs_DocType("BoxID") = IIf(val(BoxDC(1).BoundText) = 0, Null, val(BoxDC(1).BoundText))
    rs_DocType("ShelfID") = IIf(val(ShelfDC.BoundText) = 0, Null, val(ShelfDC.BoundText))
    rs_DocType("Time") = IIf(TimeTxt.Text = "", 0, val(TimeTxt.Text))
    rs_DocType("TimeUnitID") = IIf(val(TimeUnitCB.ListIndex) = -1, -1, val(TimeUnitCB.ListIndex))

    rs_DocType.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
        MsgBox ("Data saved successfully")
    End If
    Retrive_DocType
    Clear_DocType
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Retrive_DocType()
    Dim i As Integer
    Set rs_DocType = New ADODB.Recordset
    Dim StrSQL As String
    Dim TUnit As Integer
    
    
   StrSQL = "SELECT TblXXArchDocType.ID, TblXXArchDocType.Code, TblXXArchDocType.Name, TblXXArchDocType.Namee, TblXXArchDocType.ArchID, TblXXArchDocType.RoomID, TblXXArchDocType.BoxID, TblXXArchDocType.ShelfID, "
   StrSQL = StrSQL & " TblEmpDepartments.DepartmentName, TblEmpDepartments.DepartmentNamee, TblXXArchDocType.DepID, TblXXArch.Name AS ArchName, TblXXArch.Namee AS ArchNamee, TblXXRoom.Name AS RoomName, TblXXRoom.Namee AS RoomNamee, "
   StrSQL = StrSQL & " TblXXBox.Name AS BoxName, TblXXBox.Namee AS BoxNamee, TblXXShelf.Name AS ShelfName, TblXXShelf.Namee AS ShelfNamee,TblXXArchDocType.Time, TblXXArchDocType.TimeUnitID "
   StrSQL = StrSQL & " FROM TblXXArchDocType INNER JOIN "
   StrSQL = StrSQL & " TblEmpDepartments ON TblXXArchDocType.DepID = TblEmpDepartments.DeparmentID INNER JOIN "
   StrSQL = StrSQL & " TblXXArch ON TblXXArchDocType.ArchID = TblXXArch.ID INNER JOIN "
   StrSQL = StrSQL & " TblXXRoom ON TblXXArchDocType.RoomID = TblXXRoom.ID INNER JOIN "
   StrSQL = StrSQL & " TblXXBox ON TblXXArchDocType.BoxID = TblXXBox.ID INNER JOIN "
   StrSQL = StrSQL & " TblXXShelf ON TblXXArchDocType.ShelfID = TblXXShelf.ID "
   
   rs_DocType.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Grid(4).Rows = 1
    
    If rs_DocType.RecordCount > 0 Then
        rs_DocType.MoveFirst
        With Grid(4)
            .Rows = rs_DocType.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs_DocType("ID").value), 0, rs_DocType("ID").value)
                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs_DocType("Code").value), "", rs_DocType("Code").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_DocType("name").value), "", rs_DocType("name").value)
                .TextMatrix(i, .ColIndex("nameE")) = IIf(IsNull(rs_DocType("namee").value), "", rs_DocType("namee").value)
                
                .TextMatrix(i, .ColIndex("DepID")) = IIf(IsNull(rs_DocType("DepID").value), 0, rs_DocType("DepID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_DocType("DepartmentName").value), "", rs_DocType("DepartmentName").value)
                Else
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs_DocType("DepartmentNamee").value), "", rs_DocType("DepartmentNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("ArchID")) = IIf(IsNull(rs_DocType("ArchID").value), 0, rs_DocType("ArchID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Arch")) = IIf(IsNull(rs_DocType("ArchName").value), "", rs_DocType("ArchName").value)
                Else
                    .TextMatrix(i, .ColIndex("Arch")) = IIf(IsNull(rs_DocType("ArchNamee").value), "", rs_DocType("ArchNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("RoomID")) = IIf(IsNull(rs_DocType("RoomID").value), 0, rs_DocType("RoomID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Room")) = IIf(IsNull(rs_DocType("RoomName").value), "", rs_DocType("RoomName").value)
                Else
                    .TextMatrix(i, .ColIndex("Room")) = IIf(IsNull(rs_DocType("RoomNamee").value), "", rs_DocType("RoomNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs_DocType("BoxID").value), 0, rs_DocType("RoomID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Box")) = IIf(IsNull(rs_DocType("BoxName").value), "", rs_DocType("BoxName").value)
                Else
                    .TextMatrix(i, .ColIndex("Box")) = IIf(IsNull(rs_DocType("BoxNamee").value), "", rs_DocType("BoxNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("ShelfID")) = IIf(IsNull(rs_DocType("ShelfID").value), 0, rs_DocType("ShelfID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Shelf")) = IIf(IsNull(rs_DocType("ShelfName").value), "", rs_DocType("ShelfName").value)
                Else
                    .TextMatrix(i, .ColIndex("Shelf")) = IIf(IsNull(rs_DocType("ShelfNamee").value), "", rs_DocType("ShelfNamee").value)
                End If
                
                TUnit = IIf(IsNull(rs_DocType("TimeUnitID").value), -1, rs_DocType("TimeUnitID").value)
                .TextMatrix(i, .ColIndex("TUnitID")) = TUnit
                If SystemOptions.UserInterface = ArabicInterface Then
                    Select Case TUnit
                        Case -1
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & ""
                        Case 0
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & " œﬁÌﬁ… "
                        Case 1
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & " ”«⁄… "
                        Case 2
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & " ÌÊ„ "
                        Case 3
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & " ‘Â— "
                    End Select
                Else
                     Select Case TUnit
                        Case -1
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & ""
                        Case 0
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & " Min"
                        Case 1
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & " Hour"
                        Case 2
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & " Day"
                        Case 3
                            .TextMatrix(i, .ColIndex("Time")) = (IIf(IsNull(rs_DocType("Time").value), 0, rs_DocType("Time").value)) & " Month"
                    End Select
                End If
                
                
                
          rs_DocType.MoveNext
            Next
         End With
    End If
End Sub
Private Sub Update_DocType()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    Dim str As String, sr As String
    
    On Error GoTo errortrap

    If NameTxt(4).Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("«œŒ· «·«”„ «Ê·«")
        Else
            MsgBox ("Enter Name ")
        End If
        NameTxt(4).SetFocus
        Exit Sub
    End If
 
    str = Grid(4).TextMatrix(Grid(4).Row, Grid(4).ColIndex("ID"))
    sr = Grid(4).TextMatrix(Grid(4).Row, Grid(4).ColIndex("serial"))
        
    If str <> "" Then
        Cn.BeginTrans
        BeginTrans = True
        StrSQL = "Update TblXXArchDocType Set  Namee='" & NameeTxt(4).Text & "',Name='" & NameTxt(4).Text & "',Code = '" & CodeTxt(4).Text & "',DepID=" & val(Me.EmpDepDC(4).BoundText) & ",ArchID=" & val(Me.ArchDC(3).BoundText) & ",RoomID=" & val(Me.RoomDC(2).BoundText) & ",BoxID=" & val(Me.BoxDC(1).BoundText) & ",ShelfID=" & val(Me.ShelfDC.BoundText) & ",Time=" & val(TimeTxt.Text) & ",TimeUnitID=" & val(TimeUnitCB.ListIndex) & " Where ID=" & val(str)
        Cn.Execute StrSQL, , adExecuteNoRecords
           
        Cn.CommitTrans
        BeginTrans = False
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
        Else
            MsgBox ("New data saved ")
        End If
        Retrive_DocType
        Clear_DocType
    End If
        
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    
End Sub
Private Sub Del_DocType()
    Dim Msg As String
    Dim StrSQL As String
    Dim str As String, sr As String
    Set rs_DocType = New ADODB.Recordset
    
    On Error GoTo ErrTrap
        
        str = Grid(4).TextMatrix(Grid(4).Row, Grid(4).ColIndex("ID"))
        sr = Grid(4).TextMatrix(Grid(4).Row, Grid(4).ColIndex("serial"))
        
        If str <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
                Msg = Msg + (sr) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Are you sure you want to delete  " & CHR(13)
                Msg = Msg + "Data in row No."
                Msg = Msg + (sr) & CHR(13)
            End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
                StrSQL = "SELECT  *  From TblXXArchDocType"
                rs_DocType.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs_DocType.RecordCount < 1 Then
                
                    StrSQL = "delete From TblXXArchDocType where  ID =" & val(str)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    
                    StrSQL = "SELECT  *  From TblXXArchDocType"
                    rs_DocType.Close
                    rs_DocType.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If rs_DocType.RecordCount < 1 Then
                    Else
                        Retrive_DocType
                    End If
                    
                End If
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
            Else
                Msg = "This operation is not available due to lack of records"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Retrive_DocType
    Clear_DocType
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs_DocType.CancelUpdate
    'End If
End Sub
Private Sub DocTypeGridClk()
    With Grid(4)
        If .Row > 0 Then
            CodeTxt(4).Text = .TextMatrix(.Row, .ColIndex("Code"))
            NameTxt(4).Text = .TextMatrix(.Row, .ColIndex("Name"))
            NameeTxt(4).Text = .TextMatrix(.Row, .ColIndex("Namee"))
            EmpDepDC(4).BoundText = val(.TextMatrix(.Row, .ColIndex("DepID")))
            ArchDC(3).BoundText = val(.TextMatrix(.Row, .ColIndex("ArchID")))
            RoomDC(2).BoundText = val(.TextMatrix(.Row, .ColIndex("RoomID")))
            BoxDC(1).BoundText = val(.TextMatrix(.Row, .ColIndex("BoxID")))
            ShelfDC.BoundText = val(.TextMatrix(.Row, .ColIndex("BoxID")))
            TimeTxt.Text = val(.TextMatrix(.Row, .ColIndex("Time")))
            TimeUnitCB.ListIndex = val(.TextMatrix(.Row, .ColIndex("TUnitID")))
        End If
    End With
End Sub
Private Sub Clear_DocType()
    CodeTxt(4).Text = ""
    NameTxt(4).Text = ""
    NameeTxt(4).Text = ""
    EmpDepDC(4).BoundText = 0
    EmpDepDC(4).Text = ""
    ArchDC(3).BoundText = 0
    ArchDC(3).Text = ""
    RoomDC(2).BoundText = 0
    RoomDC(2).Text = ""
    BoxDC(1).BoundText = 0
    BoxDC(1).Text = ""
    ShelfDC.BoundText = 0
    ShelfDC.Text = ""
    TimeTxt.Text = ""
    TimeUnitCB.ListIndex = -1
End Sub
Private Sub TimeTxt_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TimeTxt.Text, 1)
End Sub

