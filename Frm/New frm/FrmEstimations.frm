VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEstimations 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  «·„Ê«“‰«  «· ﬁœÌ—Ì… "
   ClientHeight    =   8715
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   21210
   HelpContextID   =   580
   Icon            =   "FrmEstimations.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   21210
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
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8715
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   21210
      _cx             =   37412
      _cy             =   15372
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
         Height          =   8010
         Left            =   30
         TabIndex        =   1
         Top             =   0
         Width           =   21135
         _cx             =   37280
         _cy             =   14129
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
         Caption         =   "«·„Ê«“‰«  «· ﬁœÌ—Ì…|‘—Õ «·„Ê«“‰…"
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
            Height          =   7590
            Left            =   21780
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   45
            Width           =   21045
            _cx             =   37121
            _cy             =   13388
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   7455
               Left            =   15
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   0
               Width           =   21045
               _cx             =   37121
               _cy             =   13150
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
               Begin VB.TextBox TxtRemarks2 
                  Alignment       =   1  'Right Justify
                  Height          =   3990
                  Left            =   90
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   75
                  Top             =   165
                  Width           =   21690
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   525
                  Index           =   7
                  Left            =   255
                  TabIndex        =   76
                  Top             =   4500
                  Visible         =   0   'False
                  Width           =   2595
                  _ExtentX        =   4577
                  _ExtentY        =   926
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmEstimations.frx":038A
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   300
                  Index           =   8
                  Left            =   255
                  TabIndex        =   77
                  Top             =   4995
                  Visible         =   0   'False
                  Width           =   2565
                  _ExtentX        =   4524
                  _ExtentY        =   529
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmEstimations.frx":0724
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‘—Õ «·„Ê«“‰…"
                  Height          =   15
                  Index           =   12
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Tag             =   "53"
                  Top             =   5760
                  Width           =   21690
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid GridIntervals 
               Height          =   4275
               Left            =   7845
               TabIndex        =   41
               Top             =   1560
               Visible         =   0   'False
               Width           =   14370
               _cx             =   25347
               _cy             =   7541
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
               Cols            =   17
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEstimations.frx":0CBE
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
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7590
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   21045
            _cx             =   37121
            _cy             =   13388
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
               Height          =   7095
               Index           =   1
               Left            =   -120
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   480
               Width           =   22515
               _cx             =   39714
               _cy             =   12515
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
               Begin VB.TextBox TxtDesCription 
                  Alignment       =   1  'Right Justify
                  Height          =   1200
                  Left            =   5760
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   100
                  Top             =   735
                  Width           =   4140
               End
               Begin VB.ComboBox TypeEsstame 
                  Height          =   315
                  ItemData        =   "FrmEstimations.frx":0F2F
                  Left            =   11280
                  List            =   "FrmEstimations.frx":0F31
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Text            =   " "
                  Top             =   240
                  Width           =   1800
               End
               Begin VB.CheckBox OptMethod 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «· ﬁœÌ— „ Ê”ÿ „«”»ﬁ"
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   17340
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   2340
                  Width           =   3525
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic3 
                  Height          =   1335
                  Left            =   5805
                  TabIndex        =   71
                  TabStop         =   0   'False
                  Top             =   -1200
                  Visible         =   0   'False
                  Width           =   8865
                  _cx             =   15637
                  _cy             =   2355
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
                  Begin VSFlex8Ctl.VSFlexGrid GridOldEstimation 
                     Height          =   1050
                     Left            =   120
                     TabIndex        =   72
                     Top             =   210
                     Width           =   8655
                     _cx             =   15266
                     _cy             =   1852
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
                     Cols            =   4
                     FixedRows       =   1
                     FixedCols       =   2
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmEstimations.frx":0F33
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
                     Caption         =   "Õœœ «·„Ê«“‰«  «·”«»ﬁ…"
                     Height          =   165
                     Index           =   10
                     Left            =   6870
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   0
                     Width           =   2160
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   1455
                  Left            =   10680
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   8145
                  _cx             =   14367
                  _cy             =   2566
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
                  Begin VSFlex8Ctl.VSFlexGrid GridIntervals1 
                     Height          =   1065
                     Left            =   60
                     TabIndex        =   70
                     Top             =   300
                     Width           =   8070
                     _cx             =   14235
                     _cy             =   1879
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
                     Cols            =   8
                     FixedRows       =   1
                     FixedCols       =   2
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmEstimations.frx":0FD1
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
                     Caption         =   "⁄œœ ”‰Ê«  «·„ﬁ«—‰Â"
                     Height          =   510
                     Index           =   11
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   90
                     Width           =   3225
                  End
               End
               Begin VB.OptionButton OptAlarms 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «·Õ”«»"
                  Height          =   225
                  Index           =   1
                  Left            =   8895
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   2340
                  Width           =   1395
               End
               Begin VB.OptionButton OptAlarms 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " Õ–Ì— ›ﬁÿ"
                  Height          =   225
                  Index           =   0
                  Left            =   9915
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   2340
                  Width           =   1575
               End
               Begin VB.ComboBox OperatorsID 
                  Height          =   315
                  ItemData        =   "FrmEstimations.frx":1111
                  Left            =   15780
                  List            =   "FrmEstimations.frx":1121
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Text            =   " "
                  Top             =   2340
                  Width           =   1470
               End
               Begin VB.TextBox Percentage 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   13725
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Text            =   "0"
                  Top             =   2340
                  Width           =   1230
               End
               Begin VB.Frame Frame6 
                  Caption         =   "Õœœ «·„Ê«“‰«  «·”«»ﬁ…"
                  Height          =   1245
                  Left            =   -570
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   -3225
                  Width           =   9135
               End
               Begin VB.Frame Frame5 
                  Caption         =   "Õœœ ”‰Ê«  «·„ﬁ«—‰…"
                  Height          =   1275
                  Left            =   9705
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   -1740
                  Width           =   5805
               End
               Begin VB.Frame Frame1 
                  Caption         =   "«· Ê“Ì⁄ ⁄·Ï «Õ”«»« "
                  Height          =   885
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   7410
                  Width           =   18720
                  Begin VB.TextBox TxtRemarks1 
                     Alignment       =   1  'Right Justify
                     Height          =   615
                     Left            =   2160
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   49
                     Top             =   120
                     Width           =   3615
                  End
                  Begin VB.TextBox TxtPercentage 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   6840
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   240
                     Width           =   1215
                  End
                  Begin MSDataListLib.DataCombo DCAccountDist 
                     Height          =   315
                     Left            =   8760
                     TabIndex        =   44
                     Top             =   240
                     Width           =   3855
                     _ExtentX        =   6800
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
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   390
                     Index           =   20
                     Left            =   960
                     TabIndex        =   47
                     Top             =   240
                     Width           =   720
                     _ExtentX        =   1270
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "≈÷«›…"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmEstimations.frx":113D
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   390
                     Index           =   21
                     Left            =   240
                     TabIndex        =   48
                     Top             =   240
                     Width           =   690
                     _ExtentX        =   1217
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ–›"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmEstimations.frx":14D7
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   315
                     Index           =   9
                     Left            =   5760
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   240
                     Width           =   840
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·‰”»Â"
                     Height          =   315
                     Index           =   6
                     Left            =   8040
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Õ”«»"
                     Height          =   315
                     Index           =   5
                     Left            =   12720
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   240
                     Width           =   1080
                  End
               End
               Begin VB.TextBox TxtRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   1200
                  Left            =   480
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   39
                  Top             =   735
                  Width           =   4365
               End
               Begin MSDataListLib.DataCombo DCAccountMaster 
                  Height          =   315
                  Left            =   24840
                  TabIndex        =   37
                  Top             =   495
                  Width           =   6885
                  _ExtentX        =   12144
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
               Begin VB.Frame Frame3 
                  BackColor       =   &H00E2E9E9&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   765
                  Left            =   22515
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   975
                  Width           =   2895
                  Begin VB.OptionButton PercentagType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”» ÌœÊÌÂ"
                     Height          =   210
                     Index           =   1
                     Left            =   720
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   480
                     Width           =   1335
                  End
                  Begin VB.OptionButton PercentagType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”» «·ÌÂ"
                     Height          =   210
                     Index           =   0
                     Left            =   960
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   240
                     Width           =   1095
                  End
               End
               Begin VB.TextBox TxtTransID 
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
                  Height          =   330
                  Left            =   17760
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1845
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
                  Height          =   315
                  Index           =   0
                  Left            =   -5295
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   9645
                  Width           =   2910
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
                  Height          =   315
                  Left            =   510
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   -345
                  Visible         =   0   'False
                  Width           =   2895
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3900
                  Left            =   120
                  TabIndex        =   7
                  Top             =   2760
                  Width           =   20910
                  _cx             =   36883
                  _cy             =   6879
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
                  Rows            =   3
                  Cols            =   33
                  FixedRows       =   2
                  FixedCols       =   0
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEstimations.frx":1A71
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
               Begin MSComCtl2.DTPicker XPDtbTrans 
                  Height          =   330
                  Left            =   14730
                  TabIndex        =   12
                  Top             =   240
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   94633985
                  CurrentDate     =   41640
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00E2E9E9&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1005
                  Left            =   22635
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   345
                  Width           =   3180
                  Begin VB.OptionButton DistType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· Ê“Ì⁄ ⁄·Ï  «·›—Ê⁄"
                     Height          =   210
                     Index           =   2
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   720
                     Width           =   2055
                  End
                  Begin VB.OptionButton DistType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· Ê“Ì⁄ ⁄·Ï Õ”«»« "
                     Height          =   210
                     Index           =   0
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   240
                     Width           =   1815
                  End
                  Begin VB.OptionButton DistType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· Ê“Ì⁄ ⁄·Ï „—«ﬂ“  ﬂ·›…"
                     Height          =   210
                     Index           =   1
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   30
                     Top             =   480
                     Width           =   2055
                  End
               End
               Begin MSDataListLib.DataCombo DcBranch 
                  Height          =   315
                  Left            =   5820
                  TabIndex        =   54
                  Top             =   240
                  Width           =   4500
                  _ExtentX        =   7938
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
               Begin MSComCtl2.DTPicker Fromdate 
                  Height          =   330
                  Left            =   2910
                  TabIndex        =   56
                  Top             =   240
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   94633985
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker todate 
                  Height          =   330
                  Left            =   480
                  TabIndex        =   57
                  Top             =   240
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   94633985
                  CurrentDate     =   41640
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic6 
                  Height          =   675
                  Left            =   120
                  TabIndex        =   79
                  TabStop         =   0   'False
                  Top             =   735
                  Visible         =   0   'False
                  Width           =   5910
                  _cx             =   10425
                  _cy             =   1191
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
                  Begin VB.TextBox txtFile 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0E0FF&
                     Height          =   285
                     Left            =   3330
                     RightToLeft     =   -1  'True
                     TabIndex        =   89
                     Top             =   -375
                     Visible         =   0   'False
                     Width           =   1965
                  End
                  Begin VB.OptionButton OptActual 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Ì"
                     Height          =   225
                     Index           =   1
                     Left            =   3015
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   255
                     Width           =   1050
                  End
                  Begin VB.OptionButton OptActual 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÌœÊÌ"
                     Height          =   225
                     Index           =   0
                     Left            =   4155
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   255
                     Width           =   1335
                  End
                  Begin ImpulseButton.ISButton CmdImport 
                     Height          =   390
                     Left            =   120
                     TabIndex        =   88
                     Top             =   120
                     Width           =   1470
                     _ExtentX        =   2593
                     _ExtentY        =   688
                     Caption         =   " Õ„Ì· «·„·›"
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
                     ButtonImage     =   "FrmEstimations.frx":1F5B
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin ImpulseButton.ISButton CMDSelectFile 
                     Height          =   390
                     Left            =   1590
                     TabIndex        =   90
                     Top             =   120
                     Width           =   1410
                     _ExtentX        =   2487
                     _ExtentY        =   688
                     Caption         =   "Õœœ „”«— «·„·›"
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
                     ButtonImage     =   "FrmEstimations.frx":87BD
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin MSComDlg.CommonDialog CD1 
                     Left            =   240
                     Top             =   0
                     _ExtentX        =   847
                     _ExtentY        =   847
                     _Version        =   393216
                  End
                  Begin VB.Label Frame7 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈œŒ«· «·”‰Ê«  «·„«÷Ì…"
                     Height          =   315
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   0
                     Width           =   2535
                  End
               End
               Begin ALLButtonS.ALLButton BtnShow 
                  Height          =   375
                  Left            =   6405
                  TabIndex        =   85
                  Top             =   2190
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷ »Ì«‰«  «·„ﬁ«—‰…"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   16711680
                  BCOLO           =   16711680
                  FCOL            =   16777215
                  FCOLO           =   0
                  MCOL            =   192
                  MPTR            =   1
                  MICON           =   "FrmEstimations.frx":F01F
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   330
                  Left            =   16275
                  TabIndex        =   91
                  ToolTipText     =   "Õ–› «·ﬂ·"
                  Top             =   6720
                  Width           =   1455
                  _ExtentX        =   2566
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
                  ButtonImage     =   "FrmEstimations.frx":F03B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton ISButton3 
                  Height          =   330
                  Left            =   18945
                  TabIndex        =   92
                  ToolTipText     =   "Õ–› «·ﬂ·"
                  Top             =   6720
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·’› «·„Õœœ"
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
                  ButtonImage     =   "FrmEstimations.frx":1589D
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic7 
                  Height          =   1035
                  Left            =   18705
                  TabIndex        =   93
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   2460
                  _cx             =   4339
                  _cy             =   1826
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
                  Begin VB.OptionButton CompYear 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»ﬁ« ·”‰Ê«  «·⁄„·"
                     Height          =   345
                     Index           =   0
                     Left            =   -390
                     RightToLeft     =   -1  'True
                     TabIndex        =   96
                     Top             =   270
                     Width           =   2730
                  End
                  Begin VB.OptionButton CompYear 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Ì „  ÕœÌœ «·”‰Ê«  ÌœÊÌ"
                     Height          =   345
                     Index           =   1
                     Left            =   -690
                     RightToLeft     =   -1  'True
                     TabIndex        =   95
                     Top             =   630
                     Width           =   3030
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0E0FF&
                     Height          =   435
                     Left            =   1800
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   -570
                     Visible         =   0   'False
                     Width           =   1050
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ—Ìﬁ… «Œ Ì«— ”‰Ê«  «·„ﬁ«—‰…"
                     Height          =   240
                     Left            =   105
                     RightToLeft     =   -1  'True
                     TabIndex        =   97
                     Top             =   0
                     Width           =   2220
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic8 
                  Height          =   675
                  Left            =   120
                  TabIndex        =   102
                  TabStop         =   0   'False
                  Top             =   2040
                  Width           =   5910
                  _cx             =   10425
                  _cy             =   1191
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
                  Begin VB.TextBox Text2 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0E0FF&
                     Height          =   285
                     Left            =   3330
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   -375
                     Visible         =   0   'False
                     Width           =   1965
                  End
                  Begin ImpulseButton.ISButton ISButton7 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   105
                     ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                     Top             =   240
                     Width           =   975
                     _ExtentX        =   1720
                     _ExtentY        =   556
                     ButtonPositionImage=   1
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
                     ButtonImage     =   "FrmEstimations.frx":1C0FF
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin MSDataListLib.DataCombo DcbProject 
                     Height          =   315
                     Left            =   1440
                     TabIndex        =   106
                     Top             =   240
                     Width           =   4320
                     _ExtentX        =   7620
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
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈œŒ«· Õ”«»«  „‘—Ê⁄ „Õœœ"
                     Height          =   315
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   0
                     Width           =   2535
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ê’› «·„Ê«“‰…"
                  Height          =   675
                  Index           =   17
                  Left            =   9930
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   840
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·„Ê«“‰…"
                  Height          =   390
                  Index           =   15
                  Left            =   13185
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄‰œ „Œ«·›… «· ﬁœÌ—Ï"
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Index           =   16
                  Left            =   10530
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   2340
                  Width           =   2805
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "%"
                  Height          =   270
                  Left            =   13440
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   2340
                  Width           =   315
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰”»…"
                  Height          =   270
                  Left            =   14985
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   2340
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   450
                  Index           =   14
                  Left            =   2265
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   240
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   195
                  Index           =   13
                  Left            =   10095
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   240
                  Width           =   1305
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·› —… „‰"
                  Height          =   315
                  Index           =   0
                  Left            =   4695
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   240
                  Width           =   960
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ… Â«„…:-"
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
                  Height          =   270
                  Index           =   37
                  Left            =   -1545
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   1605
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   " Ê›— Â–… «·‘«‘… «„ﬂ«‰ÌÂ ⁄„· ŒÿÂ  Ê“Ì⁄ Õ”«» ⁄·Ï ⁄œÂ Õ”«»«  «Ê ⁄·Ï ⁄œÂ „—«ﬂ“  ﬂ·›… ›Ì › —«  „Œ ·›… »‰”» „Œ ·›Â   "
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
                  Height          =   930
                  Index           =   38
                  Left            =   -5265
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   -120
                  Visible         =   0   'False
                  Width           =   5325
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   315
                  Index           =   2
                  Left            =   4665
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   855
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·› —… „‰ "
                  Height          =   435
                  Index           =   4
                  Left            =   22065
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   495
                  Width           =   1110
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «· Ê“Ì⁄"
                  Height          =   345
                  Index           =   3
                  Left            =   22335
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   1215
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«‰‘«¡"
                  Height          =   390
                  Index           =   8
                  Left            =   16560
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”Ì‰«—ÌÊ"
                  Height          =   390
                  Index           =   7
                  Left            =   19770
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   240
                  Width           =   1200
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
                  Height          =   375
                  Left            =   18390
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   975
                  Width           =   1320
               End
               Begin VB.Shape Shape1 
                  BorderWidth     =   2
                  FillColor       =   &H00C0FFFF&
                  FillStyle       =   0  'Solid
                  Height          =   1110
                  Left            =   -5580
                  Top             =   495
                  Visible         =   0   'False
                  Width           =   5640
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   525
               Left            =   0
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   0
               Width           =   22500
               _cx             =   39688
               _cy             =   926
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   525
                  Index           =   5
                  Left            =   2160
                  TabIndex        =   107
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   18975
                  _cx             =   33470
                  _cy             =   926
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
                  Picture         =   "FrmEstimations.frx":22961
                  Caption         =   "  «·„Ê«“‰«  «· ﬁœÌ—Ì…      "
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
                  Begin VB.TextBox TxtRowNumber 
                     Alignment       =   1  'Right Justify
                     Height          =   375
                     Left            =   4410
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Text            =   "Text4"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   615
                  End
                  Begin ImpulseButton.ISButton XPBtnMove 
                     Height          =   375
                     Index           =   0
                     Left            =   1440
                     TabIndex        =   109
                     Top             =   90
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   661
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
                     ButtonImage     =   "FrmEstimations.frx":2363B
                     ColorButton     =   16777215
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
                     Height          =   375
                     Index           =   2
                     Left            =   645
                     TabIndex        =   110
                     Top             =   90
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   661
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
                     ButtonImage     =   "FrmEstimations.frx":239D5
                     ColorButton     =   16777215
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
                     Height          =   375
                     Index           =   1
                     Left            =   1965
                     TabIndex        =   111
                     Top             =   90
                     Width           =   510
                     _ExtentX        =   900
                     _ExtentY        =   661
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
                     ButtonImage     =   "FrmEstimations.frx":23D6F
                     ColorButton     =   16777215
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
                     Height          =   375
                     Index           =   3
                     Left            =   1140
                     TabIndex        =   112
                     Top             =   90
                     Width           =   255
                     _ExtentX        =   450
                     _ExtentY        =   661
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
                     ButtonImage     =   "FrmEstimations.frx":24109
                     ColorButton     =   16777215
                     ColorHighlight  =   4194304
                     ColorHoverText  =   16777215
                     ColorShadow     =   -2147483631
                     ColorOutline    =   -2147483631
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                     ColorToggledHoverText=   16777215
                     ColorTextShadow =   16777215
                  End
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   1
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   90
               Width           =   1260
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   435
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   8280
         Width           =   21210
         _cx             =   37412
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
         Align           =   2
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   75
            Left            =   12990
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   0
            Visible         =   0   'False
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   132
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
            ButtonImage     =   "FrmEstimations.frx":244A3
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   75
            Left            =   14025
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   0
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   132
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
            ButtonImage     =   "FrmEstimations.frx":2483D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   75
            Left            =   15360
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   132
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
            ButtonImage     =   "FrmEstimations.frx":24BD7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   165
            Index           =   0
            Left            =   19965
            TabIndex        =   19
            Top             =   75
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   291
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
            ButtonImage     =   "FrmEstimations.frx":24F71
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
            Height          =   165
            Index           =   1
            Left            =   17880
            TabIndex        =   20
            Top             =   75
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   291
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
            ButtonImage     =   "FrmEstimations.frx":2B7D3
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
            Height          =   165
            Index           =   2
            Left            =   16230
            TabIndex        =   21
            Top             =   75
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   291
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
            CausesValidation=   0   'False
            Height          =   165
            Index           =   3
            Left            =   13725
            TabIndex        =   22
            Top             =   75
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   291
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
            Height          =   165
            Index           =   4
            Left            =   11310
            TabIndex        =   23
            Top             =   75
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   291
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
            CausesValidation=   0   'False
            Height          =   165
            Index           =   6
            Left            =   390
            TabIndex        =   24
            Top             =   75
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   291
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
            Height          =   165
            Index           =   5
            Left            =   9060
            TabIndex        =   25
            Top             =   75
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   291
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   135
            Left            =   22845
            TabIndex        =   28
            Tag             =   "Delete Row"
            Top             =   0
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   238
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
            MICON           =   "FrmEstimations.frx":32035
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   165
            Index           =   9
            Left            =   5520
            TabIndex        =   68
            Top             =   75
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   291
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
            ButtonImage     =   "FrmEstimations.frx":32051
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdPrintAll 
            Height          =   165
            Left            =   3165
            TabIndex        =   86
            Top             =   75
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   291
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…  Õ·Ì·Ì"
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
            ButtonImage     =   "FrmEstimations.frx":388B3
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
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
            Height          =   60
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   75
            Width           =   1905
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
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
            Height          =   60
            Left            =   5415
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   75
            Width           =   1635
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
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
      ButtonImage     =   "FrmEstimations.frx":3F115
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmEstimations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Public LonRow As Long
Dim Yar As String
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long





Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub





Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
Dim i As Integer
    On Error GoTo ErrTrap

    If TxtTransID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
          Msg = "ConFirm Deleted " & Chr(13)
      
        
        End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblEstiation Where transID=" & val(Me.TxtTransID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                For i = 1 To Grid.Rows - 1
                
            StrSQL = "Delete From TblEstametChiled Where EstimatID=" & val(Grid.TextMatrix(i, Grid.ColIndex("id"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                Next i
                 StrSQL = "Delete From TblEstiationYaersDetails Where transID=" & val(Me.TxtTransID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
               StrSQL = "Delete From TblEstiationBudgetDetails Where TransID=" & val(Me.TxtTransID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblEstiationDetails Where transID=" & val(Me.TxtTransID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
          
       GridIntervals1.Clear flexClearScrollable, flexClearEverything
    GridIntervals1.Rows = 1
    
           GridOldEstimation.Clear flexClearScrollable, flexClearEverything
    GridOldEstimation.Rows = 1
                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                '    XPTxtCurrent.Caption = 0
                '    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
          
       GridIntervals1.Clear flexClearScrollable, flexClearEverything
    GridIntervals1.Rows = 1
    
           GridOldEstimation.Clear flexClearScrollable, flexClearEverything
    GridOldEstimation.Rows = 1
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This is Operation Not Allwo "
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub
Sub saveDetails(Optional i As Integer = 0, Optional EstamID As Integer = 0)
Dim RsDetails11 As ADODB.Recordset
Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim j As Integer
Dim st As String
Dim nElements As Integer
Dim k, m As Integer
If EstamID <> 0 Then
Set RsDetails11 = New ADODB.Recordset
    StrSQL = "SELECT  *  from dbo.TblEstametChiled Where (1 = -1)"
   RsDetails11.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


     If Grid.TextMatrix(i, Grid.ColIndex("StrEstametChiled")) <> "" Then
          st = Grid.TextMatrix(i, Grid.ColIndex("StrEstametChiled"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         m = 0
         For j = 0 To nElements - 1
          RsDetails11.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         For k = 1 To 12
         RsDetails11("EstimatID").value = EstamID
         RsDetails11("Dev" & k).value = val(astrSplit2tems2(m))
         m = m + 1
         RsDetails11("Act" & k).value = val(astrSplit2tems2(m))
         m = m + 1
         RsDetails11("Estim" & k).value = val(astrSplit2tems2(m))
         m = m + 1
Next k
              
         RsDetails11.update
       Next j
          End If

    

End If

End Sub
Sub RetrStrEstam(Optional EstValue As Double = 0, Optional ByRef str1 As String)
Dim str As String
Dim k As Integer
   

         For k = 1 To 12
  str = str & Trim(EstValue / 12) & "#"
 str = str & 0 & "#"
 str = str & 0 & "#"
Next k
 str = str & Trim("@")
  str = str & Chr(13)
  str1 = Trim(str)
  
End Sub
Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
End Sub

Private Sub SaveData()
    Dim i As Integer
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim RsDev1 As ADODB.Recordset
    Dim RsDev2 As ADODB.Recordset
    Dim str2 As String
    Dim LngDevID As Long

    On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
If val(TypeEsstame.ListIndex) = 1 Then
       If Trim(Me.DcBranch.BoundText) = "" Then
           Msg = "ÌÃ» ≈Œ Ì«— «·›—⁄..!!"
           MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          DcBranch.SetFocus
          SendKeys "{F4}"
         Exit Sub
       End If
 End If
    End If
With GridIntervals1
  If .Rows > 2 Then
  For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked Then
        If val(.TextMatrix(i, .ColIndex("TypeEnterYear"))) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ Ì—ÃÏ  ÕœÌœ ÿ—ﬁ «œŒ«· «·”‰Ê«  "
        Else
        MsgBox "Please Select Method Of Enter Years Aut/Manula"
        End If
        Exit Sub
        End If
       End If
    Next i
   End If
End With

 With Grid

     For i = 2 To .Rows - 1
       If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
         If (.TextMatrix(1, .ColIndex("Estimated1"))) <> "" And val(.TextMatrix(i, .ColIndex("Estimated1"))) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
               Msg = " ·«Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» ≈œŒ«· ﬁÌ„… «·”‰…"
               Msg = Msg & (.TextMatrix(1, .ColIndex("Estimated1")))
               Msg = Msg & "›Ì «·”ÿ— —ﬁ„ "
               Msg = Msg & i - 1
               MsgBox Msg
           Else
               MsgBox i - 1 & "Please Enter Value in Lin "
          End If
         Exit Sub
      End If
  If (.TextMatrix(1, .ColIndex("Estimated2"))) <> "" And val(.TextMatrix(i, .ColIndex("Estimated2"))) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 Msg = " ·«Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» ≈œŒ«· ﬁÌ„… «·”‰…"
 Msg = Msg & (.TextMatrix(1, .ColIndex("Estimated2")))
 Msg = Msg & "›Ì «·”ÿ— —ﬁ„ "
 Msg = Msg & i - 1
 MsgBox Msg
 Else
 MsgBox i - 1 & "Please Enter Value in Lin "
 End If
 Exit Sub
 End If
  If (.TextMatrix(1, .ColIndex("Estimated3"))) <> "" And val(.TextMatrix(i, .ColIndex("Estimated3"))) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 Msg = " ·«Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» ≈œŒ«· ﬁÌ„… «·”‰…"
 Msg = Msg & (.TextMatrix(1, .ColIndex("Estimated3")))
 Msg = Msg & "›Ì «·”ÿ— —ﬁ„ "
 Msg = Msg & i - 1
 MsgBox Msg
 Else
 MsgBox i - 1 & "Please Enter Value in Lin "
 End If
 Exit Sub
 End If
 
 End If
 Next i
 End With
  ''''''''''''////////////////////
    
    With Grid

 For i = 1 To .Rows - 1
 If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
 If val(.TextMatrix(i, .ColIndex("Distribution"))) <> 1 And val(.TextMatrix(i, .ColIndex("Distribution"))) <> 2 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox i & " ·«Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» ≈Œ Ì«— ‰Ê⁄  «· Ê“Ì⁄ ›Ì «·”ÿ—"
 Else
 MsgBox i & "Please Select Type Distribution "
 End If
 Exit Sub
 End If
 End If
 Next i
 End With
 With Grid
  For i = 1 To .Rows - 1
 If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
 If val(.TextMatrix(i, .ColIndex("Distribution"))) = 2 And (.TextMatrix(i, .ColIndex("StrEstametChiled"))) = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox i & " ·«Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» «œŒ«· «·»Ì«‰«  «· ﬁœÌ—ÌÂ «·„Ê“⁄Â ›Ì «·”ÿ— "
 Else
 MsgBox i & "Please Insert Distribution Data "
 End If
 Exit Sub
 End If
 End If
 Next i
 End With
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
    Me.TxtTransID.text = CStr(new_id("TblEstiation", "transID", "", True))
        rs.AddNew
           
    ElseIf Me.TxtModFlg.text = "E" Then
        Cn.Execute "delete TblEstiationDetails where transID=" & val(Me.TxtTransID.text)
        Cn.Execute "delete TblEstiationYaersDetails where transID=" & val(Me.TxtTransID.text)
      Cn.Execute "delete TblEstiationBudgetDetails where transID=" & val(Me.TxtTransID.text)
   
    End If
    
  rs("transID").value = TxtTransID.text
    rs("recordDate").value = XPDtbTrans.value
     
        rs("Fromdate").value = Fromdate.value
           rs("todate").value = todate.value
           
                   rs("Fromdateh").value = ToHijriDate(Fromdate.value)
           rs("todateh").value = ToHijriDate(todate.value)
           rs("BranchId").value = IIf(Me.DcBranch.BoundText = "", Null, val(Me.DcBranch.BoundText))
           rs("ProjectID").value = IIf(Me.DcbProject.BoundText = "", Null, val(Me.DcbProject.BoundText))
           
         rs("Percentage").value = val(Percentage.text)
         
         
    If OptAlarms(0).value = True Then
        rs("Alarms").value = 0
    Else
        rs("Alarms").value = 1
    End If
        If OptActual(0).value = True Then
        rs("ManualEntry").value = 0
    Else
       rs("ManualEntry").value = 1
    End If
        rs("OperatorsID").value = val(OperatorsID.ListIndex)
    
        If OptMethod.value = vbChecked Then
        rs("OptMethod").value = 1
        Else
        rs("OptMethod").value = 0
        End If
    
    'OptActual
    
 
 
      
    rs("Remarks").value = IIf(Me.TxtRemarks.text = "", "", Me.TxtRemarks.text)
    rs("FullRemarks").value = IIf(Me.TxtRemarks2.text = "", "", Me.TxtRemarks2.text)

    rs("DesCription").value = IIf(Me.TxtDesCription.text = "", "", Me.TxtDesCription.text)
    rs("TypeEsstame").value = IIf(Me.TypeEsstame.ListIndex = -1, Null, TypeEsstame.ListIndex)
    If CompYear(0).value = True Then
    rs("CompYear").value = 0
    ElseIf CompYear(1).value = True Then
    rs("CompYear").value = 1
    End If
    rs.update
    
    Set RsDev = New ADODB.Recordset
        
 '   RsDev.Open "TblAccountsDestributionsDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable

        
         StrSQL = "SELECT  *  from dbo.TblEstiationDetails Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
   
           
    With Me.Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                                         
                RsDev.AddNew
                 RsDev("transID").value = Me.TxtTransID.text
 
                RsDev("AccountCode").value = (.TextMatrix(i, .ColIndex("AccountCode")))
                RsDev("Distribution").value = val(.TextMatrix(i, .ColIndex("Distribution")))
                '  RsDev("year1").value = val(.TextMatrix(i, .ColIndex("year1")))
                '  RsDev("year2").value = val(.TextMatrix(i, .ColIndex("year2")))
                '  RsDev("year3").value = val(.TextMatrix(i, .ColIndex("year3")))
                  RsDev("StrEstametChiled").value = .TextMatrix(i, .ColIndex("StrEstametChiled"))
                  RsDev("Estimated1").value = val(.TextMatrix(i, .ColIndex("Estimated1")))
                   RsDev("Estimated2").value = val(.TextMatrix(i, .ColIndex("Estimated2")))
                 RsDev("Estimated3").value = val(.TextMatrix(i, .ColIndex("Estimated3")))
                  RsDev("Estimated").value = val(.TextMatrix(i, .ColIndex("Estimated")))
                  
                  RsDev("Estimated").value = val(.TextMatrix(i, .ColIndex("Estimated")))
                  RsDev("Acctual").value = val(.TextMatrix(i, .ColIndex("Acctual")))
                  RsDev("Diff").value = val(.TextMatrix(i, .ColIndex("Diff")))
                  RsDev("Varance").value = val(.TextMatrix(i, .ColIndex("Varance")))
                  RsDev("AllowVariance").value = val(.TextMatrix(i, .ColIndex("AllowVariance")))
                  RsDev("DiffVariance").value = val(.TextMatrix(i, .ColIndex("DiffVariance")))
                   If val(Grid.TextMatrix(i, Grid.ColIndex("Distribution"))) = 1 And val(Grid.TextMatrix(i, Grid.ColIndex("Estimated"))) <> 0 Then
                   RetrStrEstam val(Grid.TextMatrix(i, Grid.ColIndex("Estimated"))), str2
                    .TextMatrix(i, .ColIndex("StrEstametChiled")) = str2
                   End If
                   RsDev("StrEstametChiled").value = .TextMatrix(i, .ColIndex("StrEstametChiled"))
                  If Me.TxtModFlg.text = "E" Then
                          StrSQL = "Delete From TblEstametChiled Where EstimatID =" & val(.TextMatrix(i, .ColIndex("id"))) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
                  End If
         
                RsDev.update
                 saveDetails i, RsDev("id").value
            End If

        Next i

    End With
RsDev.Close
    Set RsDev1 = New ADODB.Recordset
        
   ' RsDev1.Open "TblEstiationYaersDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
           StrSQL = "SELECT  *  from dbo.TblEstiationYaersDetails Where (1 = -1)"
   RsDev1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
      
    With Me.GridIntervals1
'Selected
        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Remarks")) <> "" Then
                                         
                RsDev1.AddNew
                RsDev1("transID").value = Me.TxtTransID.text
                                                       
                   RsDev1("Remarks").value = val(.TextMatrix(i, .ColIndex("Remarks")))
                   RsDev1("YearId").value = val(.TextMatrix(i, .ColIndex("YearId")))
                   RsDev1("TypeEnterYear").value = val(.TextMatrix(i, .ColIndex("TypeEnterYear")))
                If .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked Then
                    RsDev1("Selected").value = 1
                Else
                    RsDev1("Selected").value = 0
                End If
                   RsDev1("datesatrt").value = val(.TextMatrix(i, .ColIndex("datesatrt")))
                   RsDev1("DateEnd").value = val(.TextMatrix(i, .ColIndex("DateEnd")))
                RsDev1.update
            End If

        Next i

    End With
 RsDev1.Close
 
 
 
    Set RsDev2 = New ADODB.Recordset
        
   ' rsdev2.Open "TblEstiationYaersDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
           StrSQL = "SELECT  *  from dbo.TblEstiationBudgetDetails Where (1 = -1)"
   RsDev2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
      
    With Me.GridOldEstimation
'Selected
        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("BudgetId")) <> "" Then
                                         
                RsDev2.AddNew
                RsDev2("transID").value = Me.TxtTransID.text
                                                       
                RsDev2("BudgetId").value = val(.TextMatrix(i, .ColIndex("BudgetId")))
                                                        
                                                    
                If .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked Then
                    RsDev2("Selected").value = 1
                Else
                    RsDev2("Selected").value = 0
                End If
 
                RsDev2.update
            End If

        Next i

    End With
 RsDev2.Close
     Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.text

        Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
           Else
           Msg = "This Record Already Saved "
           Msg = Msg + "You need enter another record"
           End If
            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
            MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.text = "R"
    Retrive val(TxtTransID.text)

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Function addInterval()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String

    If Me.DistType(0).value = True Then 'Õ”«»« 
        If SystemOptions.UserInterface = ArabicInterface Then
            des = "  «·Õ”«» "
        Else
            des = " Accounts "
        End If
 
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            des = " «·„—ﬂ“  "
        Else
            des = " CC "
        End If
    End If

    If val(Me.TxtRowNumber.text) <> 0 Then
        LngRow = val(Me.TxtRowNumber.text)
    Else
        Me.GridIntervals.Rows = Me.GridIntervals.Rows + 1
        LngRow = Me.GridIntervals.Rows - 1
    End If
 
    On Error Resume Next
 
    With Me.GridIntervals
  
      '  .TextMatrix(LngRow, .ColIndex("FromDate")) = dpFromDate.value
    
      '  .TextMatrix(LngRow, .ColIndex("ToDate")) = Me.DpToDate
  '
  '      .TextMatrix(LngRow, .ColIndex("Remarks")) = (Me.TxtRemarks2.text)
  '
  '      .AutoSize 0, .Cols - 1, False
    End With

    Me.DCAccountDist.BoundText = ""
 
    Me.TxtRemarks2.text = ""
  
    ReLineGrid
End Function

Function FillGrid()
 Dim rs As ADODB.Recordset
    Dim StrSQL As String
  
Dim AccountCode As String

   StrSQL = "Select * From Accounts Where  mowazna=1"
    StrSQL = StrSQL + " Order By Account_Serial"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

Dim i As Integer
    With Me.Grid
        .Clear flexClearScrollable, flexClearEverything
Grid.Rows = 2
        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount
'Account_Code ACode
'Account_Name ASerial  AName

            For i = .FixedRows To .Rows - 1
           '     .TextMatrix(i, .ColIndex("Account_ID")) = IIf(IsNull(rs("Account_ID").value), "", rs("Account_ID").value)
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                AccountCode = .TextMatrix(i, .ColIndex("AccountCode"))
                .TextMatrix(i, .ColIndex("ASerial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)

                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("AName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                Else
                    .TextMatrix(i, .ColIndex("AName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                End If
             
            '    CurrentAccount = Abs(val(GetActualAccountBalance(AccountCode, branch_id, FirstPeriod, Date, , False, opening_balance)))
   '  .TextMatrix(i, .ColIndex("OpenAccount")) = FormatNumber(Abs(opening_balance), SystemOptions.SysDefCurrencyForamt, True, True, True)
           
                If rs("last_account").value = True Then
 
                    .Cell(flexcpFontBold, i, .ColIndex("AName")) = False
                
                Else
                    
                    .Cell(flexcpFontBold, i, .ColIndex("AName")) = True
                    .Cell(flexcpFontName, i, .ColIndex("AName")) = "Tahoma"
                End If

                rs.MoveNext
            Next i

        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If

        .AutoSize 0, .Cols - 1, False
    End With


End Function



Private Sub BtnShow_Click()
If Me.TxtModFlg.text <> "R" Then
Dim i As Integer
Dim StrAccountCode As String
 Dim Balance As Double
With Grid
If Grid.Rows > 2 Then
For i = 1 To .Rows - 1
StrAccountCode = .TextMatrix(i, .ColIndex("AccountCode"))
If StrAccountCode <> "" Then
 Balance = GetActualAccountBalance(StrAccountCode, 0, Fromdate.value, todate.value)
.TextMatrix(i, .ColIndex("Acctual")) = Balance
.TextMatrix(i, .ColIndex("Diff")) = val(.TextMatrix(i, .ColIndex("Estimated"))) - val(.TextMatrix(i, .ColIndex("Acctual")))
End If
Next i
End If
End With
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap
    Select Case Index
        Case 0
             TxtModFlg.text = "N"
           clear_all Me
        OperatorsID.ListIndex = 0
       OptAlarms(0).value = True
       OptActual(1).value = True
            Me.XPDtbTrans.value = Date
            Me.Fromdate.value = Date
            Me.todate.value = Date
            
            Percentage.text = 0
            
             'Me.dbTodate.value = Date
       
            XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
           Grid.Rows = 3
            Grid.Enabled = True
            PercentagType(0).value = True
       '
           GridIntervals1.Clear flexClearScrollable, flexClearEverything
            GridIntervals1.Rows = 1
            GridIntervals1.Enabled = True
       '
          GridOldEstimation.Clear flexClearScrollable, flexClearEverything
            GridOldEstimation.Rows = 1
            GridOldEstimation.Enabled = True
       '
            DistType(0).value = True
          DcBranch.BoundText = Current_branch
FillGrid
FillNewGrid
If Grid.Rows <= 2 Then
Grid.Rows = 3
End If
Me.MergGrid
        Case 1
 
            TxtModFlg.text = "E"
                    Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

             Del_Trans
       
        Case 5

              If DoPremis(Do_Search, Me.name, True) = False Then
                  Exit Sub
              End If
           Load FrmEstimationsSearch
           FrmEstimationsSearch.show vbModal
        Case 6
            Unload Me

        Case 7
            addInterval

            '   ViewDataList
      Case 9
      print_report , 1
        Case 20
            addrow

        Case 21
            RemoveGridRow

        Case 8
            RemoveGridRow1
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow1()

    With Me.GridIntervals

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Sub addrow()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String

    If Me.DistType(0).value = True Then 'Õ”«»« 
        If SystemOptions.UserInterface = ArabicInterface Then
            des = "  «·Õ”«» "
        Else
            des = " Accounts "
        End If
 
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            des = " «·„—ﬂ“  "
        Else
            des = " CC "
        End If
    End If

    If (Me.DCAccountDist.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ  " & des & "   «·„—«œ  Ê“Ì⁄ ⁄·ÌÂ...!!!"
        Else
            Msg = "must select " & des & " To Desrtribute...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    'If Val(Me.TxtRowNumber.text) = 0 Then
    '    LngFindRow = Grid.FindRow(Val(Me.DCAccountDist.BoundText), _
    '    Grid.FixedRows, Grid.ColIndex("ACode"), False, True)
    '    If LngFindRow <> -1 Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        Msg = "·«Ì„ﬂ‰  ﬂ—«— " & Des & "  ...!!!"
    '    Else
    '        Msg = " Can't Repeat  " & Des & "  ...!!!"
    '    End If
    '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Exit Sub
    '    End If
    'End If

    If val(Me.TxtRowNumber.text) <> 0 Then
        LngRow = val(Me.TxtRowNumber.text)
    Else
        Me.Grid.Rows = Me.Grid.Rows + 1
        LngRow = Me.Grid.Rows - 1
    End If
 
    On Error Resume Next
 
    With Me.Grid
    
        If DistType(0).value = True Then
            .TextMatrix(LngRow, .ColIndex("Aid")) = val(GetID("ACCOUNTS", "Account_Code", "Account_ID", Me.DCAccountDist.BoundText))
            .TextMatrix(LngRow, .ColIndex("ASerial")) = val(GetID("ACCOUNTS", "Account_Code", "Account_Serial", Me.DCAccountDist.BoundText))
        Else
            .TextMatrix(LngRow, .ColIndex("Aid")) = val(GetID("markaas_taklefa", "account_no", "id", Me.DCAccountDist.BoundText))
            .TextMatrix(LngRow, .ColIndex("ASerial")) = Me.DCAccountDist.BoundText

        End If
  
        .TextMatrix(LngRow, .ColIndex("ACode")) = Me.DCAccountDist.BoundText
    
        .TextMatrix(LngRow, .ColIndex("AName")) = Me.DCAccountDist.text
    
        .TextMatrix(LngRow, .ColIndex("Percentage")) = val(Me.TxtPercentage.text)
    
        .TextMatrix(LngRow, .ColIndex("Remarks")) = (Me.TxtRemarks1.text)
     
        .AutoSize 0, .Cols - 1, False
    End With

    Me.DCAccountDist.BoundText = ""
    Me.TxtPercentage.text = ""
    Me.TxtRemarks1.text = ""
  
    ReLineGrid
 
End Sub

Private Sub Undo()
   ' On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive TxtTransID.text
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
Function GetTotalProject(Optional ByRef Project As Integer, Optional str As String) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
GetTotalProject = 0
Set rs2 = New ADODB.Recordset
sql = "SELECT     SUM(" & str & ") AS Total"
sql = sql & " From dbo.terms_operations"
sql = sql & " Where (project_id =" & Project & ")"
sql = sql & " GROUP BY project_id"
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetTotalProject = IIf(IsNull(rs2("Total").value), 0, rs2("Total").value)
End If
End Function
Function GetAccountProjects(Optional str As String, Optional AccountCode As String, Optional ByRef ProjectID As Integer = 0) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim cod As String
cod = Replace(Replace(AccountCode, Chr(10), ""), Chr(13), "")
GetAccountProjects = False
ProjectID = 0
sql = " SELECT    id, " & str & " from  dbo.projects where " & str & "='" & cod & "' "
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetAccountProjects = True
ProjectID = IIf(IsNull(rs2("id").value), 0, rs2("id").value)
End If

End Function

Sub FillProject_Graid(Optional ProjectID As Integer)
If ProjectID <> 0 Then
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim last As Integer
sql = " SELECT   Salary_account,Material_account,expanses_account  from  dbo.projects where id=" & ProjectID & ""
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Grid
If .Rows = 3 Then
.Rows = 2
End If
last = .Rows
.Rows = .Rows + rs2.RecordCount + 1
.TextMatrix(last, .ColIndex("AccountCode")) = IIf(IsNull(rs2("Salary_account").value), "", rs2("Salary_account").value)
Grid_AfterEdit last, .ColIndex("AccountCode")
last = last + 1
.TextMatrix(last, .ColIndex("AccountCode")) = IIf(IsNull(rs2("Material_account").value), "", rs2("Material_account").value)
Grid_AfterEdit last, .ColIndex("AccountCode")
last = last + 1
.TextMatrix(last, .ColIndex("AccountCode")) = IIf(IsNull(rs2("expanses_account").value), "", rs2("expanses_account").value)
Grid_AfterEdit last, .ColIndex("AccountCode")
.Rows = .Rows - 1
End With
End If
End If
End Sub



Sub ExilSheet()
CD1.ShowOpen
txtFile.text = CD1.FileName
If txtFile.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Õœœ «·„·› «Ê·«"
Exit Sub
Else
MsgBox "Select File"
Exit Sub
End If
End If
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim currentvalue As String
Dim j As Integer
Dim bol As Boolean
Dim code As String
Dim val1 As Double
Dim name As String
Dim k As Integer
'Dim CheqNo As String
bol = False

    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.Workbooks.Open txtFile.text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
    With ExcelSheet
    i = 2
    'Grid.Rows = 3
    Do Until .Cells(i, 2) & "" = ""
 '       Set l = lvwList.ListItems.Add(, , .Cells(i, 1))
   code = .Cells(i, 1)
    val1 = .Cells(i, 3)
         name = .Cells(i, 2)

 With Grid

   '  MsgBox .Rows
   
  
  For j = .FixedRows To .Rows - 1
  If code = .TextMatrix(j, .ColIndex("ASerial")) Then
  k = j
  bol = True
  End If
  Next j
  '''
  If bol = False Then
   If .Cell(flexcpText, 1, .ColIndex("Estimated1"), 1, .ColIndex("Estimated1")) = Yar Then
   .TextMatrix(i, .ColIndex("Estimated1")) = val(val1)
   ElseIf .Cell(flexcpText, 1, .ColIndex("Estimated2"), 1, .ColIndex("Estimated2")) = Yar Then
   .TextMatrix(i, .ColIndex("Estimated2")) = val(val1)
      ElseIf .Cell(flexcpText, 1, .ColIndex("Estimated3"), 1, .ColIndex("Estimated3")) = Yar Then
   .TextMatrix(i, .ColIndex("Estimated3")) = val(val1)
     
   Else
   Exit Sub
   End If
   .TextMatrix(i, .ColIndex("ASerial")) = code
  
   .TextMatrix(i, .ColIndex("AName")) = name
    Grid_AfterEdit i, .ColIndex("ASerial")
  ' Grid_AfterEdit i, .ColIndex("AName")
    Grid_AfterEdit i, .ColIndex("Estimated1")
   Grid_AfterEdit i, .ColIndex("Estimated2")
   Grid_AfterEdit i, .ColIndex("Estimated3")
   Else
    If .Cell(flexcpText, 1, .ColIndex("Estimated1"), 1, .ColIndex("Estimated1")) = Yar Then
   .TextMatrix(k, .ColIndex("Estimated1")) = val(val1)
   ElseIf .Cell(flexcpText, 1, .ColIndex("Estimated2"), 1, .ColIndex("Estimated2")) = Yar Then
   .TextMatrix(k, .ColIndex("Estimated2")) = val(val1)
      ElseIf .Cell(flexcpText, 1, .ColIndex("Estimated3"), 1, .ColIndex("Estimated3")) = Yar Then
   .TextMatrix(k, .ColIndex("Estimated3")) = val(val1)
    Else
   Exit Sub
   End If
   Grid_AfterEdit i, .ColIndex("Estimated1")
   Grid_AfterEdit i, .ColIndex("Estimated2")
   Grid_AfterEdit i, .ColIndex("Estimated3")
   End If

     
    
          


'  Fg_Journal_AfterEdit i, .ColIndex("BranchId")
'
'   If Val(DebitValue) > 0 Then
'      .TextMatrix(i, .ColIndex("DebitValue")) = Val(DebitValue)
'         Fg_Journal_AfterEdit i, .ColIndex("DebitValue")
'
'    End If
'
'       If Val(CreditValue) > 0 Then
'     .TextMatrix(i, .ColIndex("CreditValue")) = Val(CreditValue)
'     Fg_Journal_AfterEdit i, .ColIndex("CreditValue")
'      End If
      
   
 End With
        i = i + 1
      ' Grid.Rows = Grid.Rows + 1
        
    Loop

    End With
'Grid.Rows = Grid.Rows - 1
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
End Sub

Private Sub CmdImport_Click()
ExilSheet
txtFile.text = ""
End Sub

Private Sub CmdPrintAll_Click()
print_report , 0
End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    If Grid.Rows > 1 Then
        If Grid.Rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.Rows > 1 Then
                If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.Row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub
 
Private Sub CMDSelectFile_Click()

If SystemOptions.UserInterface = ArabicInterface Then

Yar = InputBox("≈œŒ· «·”‰… «·„«·Ì…")
If cked(Yar) = False Then
MsgBox "«·”‰… «·„œŒ·… €Ì— „ÊÃÊœÂ"
Exit Sub
End If
Else
Yar = InputBox("Enter Year")
If cked(Yar) = False Then
MsgBox "Year entered Not Found"
Exit Sub
End If
End If

CD1.ShowOpen
txtFile.text = CD1.FileName
End Sub
Function cked(Optional name As String = "") As Boolean
Dim i As Integer
With GridIntervals1
For i = .FixedRows To .Rows - 1
 If .Cell(flexcpChecked, 1, .ColIndex("Selected")) = flexChecked And .TextMatrix(i, .ColIndex("Remarks")) = name Then
 cked = True
 Exit Function
 End If
 Next i
 End With
 cked = False
End Function

Private Sub CompYear_Click(Index As Integer)
If Me.TxtModFlg.text <> "R" Then

Dim Msg As String
If Grid.Rows > 2 Then
If Grid.TextMatrix(2, Grid.ColIndex("AccountCode")) <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "Â·  —Ìœ  Õ–› ﬂ· «·»Ì«‰«  "
Else
Msg = "Confirm Delete Data"
End If
If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
Grid.Clear flexClearScrollable, flexClearEverything
 Grid.Rows = 3
 MergGrid
Else
Exit Sub
End If
End If
End If

If Index = 0 Then
FillNewGrid
ElseIf Index = 1 Then
FillGridManual
End If
End If
End Sub

Private Sub DCAccountDist_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If DistType(0).value = True Then
        If KeyCode = vbKeyF3 Then
            Unload Account_search
            Account_search.show
            Account_search.case_id = 178
            
        End If

    Else

        If KeyCode = vbKeyF3 Then
            CostCenterSearch.show
            CostCenterSearch.RetrunType = 178
        End If

    End If

End Sub

Private Sub DCAccountMaster_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 177

    End If

End Sub

Private Sub DistType_Click(Index As Integer)
    Dim Dcombos As ClsDataCombos

    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 3
          
    Select Case Index
        
        Case 0
            Frame1.Caption = "«· Ê“Ì⁄ ⁄·Ï «·Õ”«»«  "
            lbl(5).Caption = "«·Õ”«» "

            With Me.Grid
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("ASerial")) = "ﬂÊœ «·Õ”«»"
                .TextMatrix(0, .ColIndex("AName")) = "«”„ «·Õ”«»"
            Else
              .TextMatrix(0, .ColIndex("ASerial")) = " Account Code"
                .TextMatrix(0, .ColIndex("AName")) = "Account Name"
                End If
            End With
 
            Set Dcombos = New ClsDataCombos

            If SystemOptions.UserInterface = ArabicInterface Then
                Dcombos.GetAccountingCodes DCAccountDist, True
            Else
 
                Dcombos.GetAccountingCodesENg DCAccountDist, True

            End If

        Case 1
            Frame1.Caption = "«· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›Â "
            lbl(5).Caption = "«·„—ﬂ“ "

            With Me.Grid
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("ASerial")) = "ﬂÊœ «·„—ﬂ“"
                .TextMatrix(0, .ColIndex("AName")) = "«”„ «·„—ﬂ“"
                Else
                     .TextMatrix(0, .ColIndex("ASerial")) = " Code Center"
                .TextMatrix(0, .ColIndex("AName")) = "Name Center "
               End If
            End With
           
            
            Set Dcombos = New ClsDataCombos

            If SystemOptions.UserInterface = ArabicInterface Then
                Dcombos.getCC DCAccountDist
            Else
                Dcombos.getCC DCAccountDist

            End If

        Case 2
            Frame1.Caption = "«· Ê“Ì⁄ ⁄·Ï  «·›—Ê⁄   "
            lbl(5).Caption = " «·›— ⁄  "

            With Me.Grid
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("ASerial")) = "ﬂÊœ «·›— ⁄ "
                .TextMatrix(0, .ColIndex("AName")) = "«”„ «·›— ⁄ "
           Else
                  .TextMatrix(0, .ColIndex("ASerial")) = " Code  "
                .TextMatrix(0, .ColIndex("AName")) = "Name "
                End If
            End With
            
            Set Dcombos = New ClsDataCombos

            If SystemOptions.UserInterface = ArabicInterface Then
                Dcombos.GetBranches DCAccountDist
            Else
                Dcombos.GetBranches DCAccountDist

            End If

    End Select

End Sub

Private Sub Form_Load()
 MergGrid
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
  If SystemOptions.UserInterface = ArabicInterface Then
  TypeEsstame.AddItem "⁄«„…"
  TypeEsstame.AddItem "ÿ»ﬁ« ··›—⁄"
  
                Grid.ColComboList(Grid.ColIndex("Distribution")) = "#1;  «·Ì|#2; ÌœÊÌ"
                    GridIntervals1.ColComboList(GridIntervals1.ColIndex("TypeEnterYear")) = "#1;  «·Ì|#2; ÌœÊÌ"
            
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
            TypeEsstame.AddItem "General"
            TypeEsstame.AddItem "By Branch"
               Grid.ColComboList(Grid.ColIndex("Distribution")) = "#1;Auto  |#2;Manual "
               GridIntervals1.ColComboList(GridIntervals1.ColIndex("TypeEnterYear")) = "#1;Auto  |#2;Manual "
            End If
         
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With
    
    AdditemTocCmp
    If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = " select id,Project_name from projects"
    Else
    My_SQL = " select id,Project_nameE from projects"
    End If
    fill_combo DcbProject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic

    If SystemOptions.UserInterface = ArabicInterface Then
        Dcombos.GetAccountingCodes DCAccountMaster, True
        Dcombos.GetAccountingCodes DCAccountDist, True
    Else
        Dcombos.GetAccountingCodesENg DCAccountMaster, True
        Dcombos.GetAccountingCodesENg DCAccountDist, True

    End If
Dcombos.GetBranches DcBranch

  '  With Me.Grid
  '      .Rows = 1
  '      .ExplorerBar = flexExSortShowAndMove
  '      .RowHeightMin = 300
  '      .ExtendLastCol = True
  '  End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEstiation  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Retrive
    XPBtnMove_Click 2
MergGrid
    Me.TxtModFlg.text = "R"
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub
Sub MergGrid()
counter = 0
   With Grid
   .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
         .MergeCol(.ColIndex("Estimated1")) = True
        .MergeCol(.ColIndex("Estimated2")) = True
        .MergeCol(.ColIndex("Estimated3")) = True
        .ColWidth(.ColIndex("Estimated1")) = 1500
        .MergeCol(.ColIndex("Estimated")) = True
        .MergeCol(.ColIndex("Acctual")) = True
        .MergeCol(.ColIndex("Diff")) = True
        .ColWidth(.ColIndex("Estimated")) = 1500
        .ColWidth(.ColIndex("Acctual")) = 1500
        .ColWidth(.ColIndex("Diff")) = 1500
        
         .ColHidden(.ColIndex("Estimated1")) = True
       .ColHidden(.ColIndex("Estimated2")) = True
         .ColHidden(.ColIndex("Estimated3")) = True
         .ColWidth(.ColIndex("Estimated2")) = 1500
         .ColWidth(.ColIndex("Estimated3")) = 1500
        .Cell(flexcpAlignment, 0, .ColIndex("Estimated1"), 0, .ColIndex("Estimated3")) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, .ColIndex("Estimated"), 0, .ColIndex("Diff")) = flexAlignCenterCenter
              If GridIntervals1.Rows > 1 Then
      If GridIntervals1.Cell(flexcpChecked, 1, GridIntervals1.ColIndex("Selected")) = flexChecked Then
        .Cell(flexcpText, 1, .ColIndex("Estimated1"), 1, .ColIndex("Estimated1")) = GridIntervals1.TextMatrix(1, GridIntervals1.ColIndex("Remarks"))
         .ColHidden(.ColIndex("Estimated1")) = False
        counter = counter + 1
        Else
        .Cell(flexcpText, 1, .ColIndex("Estimated1"), 1, .ColIndex("Estimated1")) = ""
        End If
        Else
        .Cell(flexcpText, 1, .ColIndex("Estimated1"), 1, .ColIndex("Estimated1")) = ""
        End If
        If GridIntervals1.Rows > 2 Then
        If GridIntervals1.Cell(flexcpChecked, 2, GridIntervals1.ColIndex("Selected")) = flexChecked Then
        .Cell(flexcpText, 1, .ColIndex("Estimated2"), 1, .ColIndex("Estimated2")) = GridIntervals1.TextMatrix(2, GridIntervals1.ColIndex("Remarks"))
        .ColHidden(.ColIndex("Estimated2")) = False
        counter = counter + 1
        Else
        .Cell(flexcpText, 1, .ColIndex("Estimated2"), 1, .ColIndex("Estimated2")) = ""
        End If
        Else
        .Cell(flexcpText, 1, .ColIndex("Estimated2"), 1, .ColIndex("Estimated2")) = ""
        End If
         If GridIntervals1.Rows > 3 Then
        If GridIntervals1.Cell(flexcpChecked, 3, GridIntervals1.ColIndex("Selected")) = flexChecked Then
        .Cell(flexcpText, 1, .ColIndex("Estimated3"), 1, .ColIndex("Estimated3")) = GridIntervals1.TextMatrix(3, GridIntervals1.ColIndex("Remarks"))
        .ColHidden(.ColIndex("Estimated3")) = False
        counter = counter + 1
         Else
        .Cell(flexcpText, 1, .ColIndex("Estimated3"), 1, .ColIndex("Estimated3")) = ""
        End If
  Else
   .Cell(flexcpText, 1, .ColIndex("Estimated3"), 1, .ColIndex("Estimated3")) = ""
          End If
   If SystemOptions.UserInterface = ArabicInterface Then
        .Cell(flexcpText, 0, .ColIndex("Estimated1"), 0, .ColIndex("Estimated3")) = "”‰Ê«  ”«»ﬁ…"
        .Cell(flexcpText, 0, .ColIndex("Estimated"), 0, .ColIndex("Diff")) = " «·”‰… «·Õ«·Ì…" & "  " & year(Date)
        .Cell(flexcpText, 1, .ColIndex("Estimated"), 1, .ColIndex("Estimated")) = "«· ﬁœÌ—Ì"
           .Cell(flexcpText, 1, .ColIndex("Diff"), 1, .ColIndex("Diff")) = "«·›—ﬁ"
        .Cell(flexcpText, 1, .ColIndex("Acctual"), 1, .ColIndex("Acctual")) = "«·›⁄·Ì"
   Else
   .Cell(flexcpText, 0, .ColIndex("Estimated1"), 0, .ColIndex("Estimated3")) = "Last Years"
   .Cell(flexcpText, 0, .ColIndex("Estimated"), 0, .ColIndex("Diff")) = " Current Year" & "  " & year(Date)
          .Cell(flexcpText, 1, .ColIndex("Estimated"), 1, .ColIndex("Estimated")) = "Estimated"
           .Cell(flexcpText, 1, .ColIndex("Diff"), 1, .ColIndex("Diff")) = "Difference"
        .Cell(flexcpText, 1, .ColIndex("Acctual"), 1, .ColIndex("Acctual")) = "Actual"
   End If
   End With
End Sub
Private Sub ChangeLang()
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    lbl(15).Caption = "Type"
    ISButton7.Caption = "Add"
    Label4.Caption = "Select Project"
    Label3.Caption = "Method Of Select Years"
    CompYear(0).Caption = "Years of work"
    CompYear(1).Caption = "Select Manual Years"
    CmdPrintAll.Caption = "Analytic Print"
    ISButton3.Caption = "Delet Select"
    ISButton4.Caption = "Delet All"
    lbl(17).Caption = "Description"
    lbl(3).Caption = "Select "
    lbl(13).Caption = "Branch"
    lbl(14).Caption = "To"
    Frame5.Caption = "Select Comparison Year"
    Frame6.Caption = "Select Previous Budgets"
    Frame7.Caption = "Last Years Entering"
    OptActual(0).Caption = "Manual"
    OptActual(1).Caption = "Automatic"
    OptAlarms(0).Caption = "just Worning"
    OptAlarms(1).Caption = "Stop The Acount"
    lbl(16).Caption = " When the estimated offense "
    Label1.Caption = "Percentage"
    Me.OptMethod.Caption = " Average estimation method of the above "
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Me.Caption = " Budget"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(8).Caption = "Date"
    lbl(4).Caption = "Select Acc."
    lbl(0).Caption = "Dis .Type"
    DistType(0).Caption = "To Account"
    DistType(1).Caption = "To CC"
    DistType(2).Caption = "To Branches"
    lbl(3).Caption = "Dis Method"
    PercentagType(0).Caption = "Auto"
    PercentagType(1).Caption = "Manual"
    lbl(2).Caption = "Remarks"
    Frame1.Caption = "Dis To Account"
    lbl(5).Caption = "Sel Account"
    lbl(6).Caption = "%"
    lbl(9).Caption = "Remarks"
    Cmd(20).Caption = "Add"
    Cmd(21).Caption = "Del"
    lbl(37).Visible = False
    lbl(38).Visible = False
    Shape1.Visible = False
    CmdRemove.Caption = "Del Row"
    Me.C1Tab1.TabCaption(0) = "Budget Data "
    Me.C1Tab1.TabCaption(1) = "Remarks"
  '  Frame4.Caption = "Distributions Period"
    lbl(10).Caption = "From"
    lbl(11).Caption = "To"
    lbl(12).Caption = "Remarks"
    Cmd(7).Caption = "Add"
    Cmd(8).Caption = "Del"
    lbl(11).Caption = Frame5.Caption
    lbl(10).Caption = Frame6.Caption
    BtnShow.Caption = "Show Data"
   With GridIntervals1
         .TextMatrix(0, .ColIndex("ser")) = "Code"
        .TextMatrix(0, .ColIndex("Remarks")) = "Fiscal year"
        .TextMatrix(0, .ColIndex("Selected")) = "Selected"
        .TextMatrix(0, .ColIndex("TypeEnterYear")) = "TypeEnterYear"
        .TextMatrix(0, .ColIndex("selectfile")) = "Select File"
   End With

    With Me.GridOldEstimation
        .TextMatrix(0, .ColIndex("Ser")) = "Code"
        .TextMatrix(0, .ColIndex("BudgetId")) = "Scenario No."
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
       .TextMatrix(0, .ColIndex("Selected")) = "Selected"
    End With
    CMDSelectFile.Caption = "File Path"
    CmdImport.Caption = "Upload"
  
    With Grid
        .TextMatrix(0, .ColIndex("ch")) = "Select"
         .TextMatrix(0, .ColIndex("ser")) = "Code"
         .TextMatrix(0, .ColIndex("Distribution")) = "Distribution"
         .TextMatrix(0, .ColIndex("show")) = "Show"
        .TextMatrix(0, .ColIndex("ASerial")) = "Acounting Code"
        .TextMatrix(0, .ColIndex("AName")) = "Acounting Name"
       ' .TextMatrix(0, .ColIndex("year1")) = "year 2012"
       ' .TextMatrix(0, .ColIndex("year2")) = "year 2013"
       ' .TextMatrix(0, .ColIndex("year3")) = "year 2014"
        .TextMatrix(0, .ColIndex("Estimated1")) = "First Year "
        .TextMatrix(0, .ColIndex("Estimated2")) = "Second Year Year "
        .TextMatrix(0, .ColIndex("Estimated3")) = "Third year Year "
      '  .TextMatrix(0, .ColIndex("Estimated")) = "Estimated"
      '  .TextMatrix(0, .ColIndex("Acctual")) = "The actual"
      '  .TextMatrix(0, .ColIndex("Diff")) = "difference"
          .TextMatrix(0, .ColIndex("ASerial")) = "Account Code "
        .TextMatrix(0, .ColIndex("AName")) = "Account Name"

        .TextMatrix(0, .ColIndex("Varance")) = "actual deviation"
        .TextMatrix(0, .ColIndex("AllowVariance")) = "Allowable deviation"
        .TextMatrix(0, .ColIndex("DiffVariance")) = "Deviation difference"
      End With
       
    
    With Me.GridIntervals
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("IntervalSerial")) = "Interval NOl"
        .TextMatrix(0, .ColIndex("fromdate")) = "fromdate"
        .TextMatrix(0, .ColIndex("todate")) = "todate"
        .TextMatrix(0, .ColIndex("DistributedDone")) = "Done"
        .TextMatrix(0, .ColIndex("REMARKS")) = "REMARKS"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "NoteSerial"
        .TextMatrix(0, .ColIndex("PrintJL")) = "PrintJL"

    End With

End Sub
Sub FillGridManual()
Dim i As Integer
 Dim Num As Integer
 Dim k As Integer
    With Me.GridIntervals1
    k = 3
Num = year(Date)
        .Rows = 4
        .Clear flexClearScrollable

         
            For i = .FixedRows To 3
                .TextMatrix(i, .ColIndex("Remarks")) = Num - k
           k = k - 1
            Next i
 
            .AutoSize 0, .Cols - 1, False
    End With

End Sub

Public Sub FillNewGrid()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from TblyearsData order by Remarks  "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Me.GridIntervals1

        .Rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = .FixedRows To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("YearId")) = IIf(IsNull(Rs3.Fields("TblyearsDataid").value), "", Rs3.Fields("TblyearsDataid").value)
                       
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs3.Fields("Remarks").value), "", Rs3.Fields("Remarks").value)
                .TextMatrix(i, .ColIndex("datesatrt")) = IIf(IsNull(Rs3.Fields("datesatrt").value), "", Rs3.Fields("datesatrt").value)
                .TextMatrix(i, .ColIndex("DateEnd")) = IIf(IsNull(Rs3.Fields("DateEnd").value), "", Rs3.Fields("DateEnd").value)
                       '
               
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close


    sql = "Select * from TblEstiation "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With GridOldEstimation

        .Rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("BudgetId")) = IIf(IsNull(Rs3.Fields("transID").value), "", Rs3.Fields("transID").value)
                       
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs3.Fields("Remarks").value), "", Rs3.Fields("Remarks").value)
              '          '
               
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 Me.MergGrid
    Rs3.Close
End Sub


 
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

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
Public Sub hidcol()
With Grid
If Me.TxtModFlg.text = "R" Then
.ColHidden(.ColIndex("Acctual")) = True
.ColHidden(.ColIndex("Diff")) = True
Else
.ColHidden(.ColIndex("Acctual")) = False
.ColHidden(.ColIndex("Diff")) = False
End If
End With
''''////////////

End Sub
Sub Calculate(Optional Row1 As Long = 0)
If Row1 <> 0 Then
With Grid
.TextMatrix(Row1, .ColIndex("Estimated")) = val(.TextMatrix(Row1, .ColIndex("Estimated1"))) + val(.TextMatrix(Row1, .ColIndex("Estimated2"))) + val(.TextMatrix(Row1, .ColIndex("Estimated3")))
If counter <> 0 Then
.TextMatrix(Row1, .ColIndex("Estimated")) = Round(val(.TextMatrix(Row1, .ColIndex("Estimated"))) / counter, 2)
Else
.TextMatrix(Row1, .ColIndex("Estimated")) = 0
End If
Select Case OperatorsID.ListIndex
Case 0
.TextMatrix(Row1, .ColIndex("Estimated")) = val(.TextMatrix(Row1, .ColIndex("Estimated"))) + ((val(.TextMatrix(Row1, .ColIndex("Estimated"))) * val(Percentage.text)) / 100)
Case 1
.TextMatrix(Row1, .ColIndex("Estimated")) = val(.TextMatrix(Row1, .ColIndex("Estimated"))) - ((val(.TextMatrix(Row1, .ColIndex("Estimated"))) * val(Percentage.text)) / 100)
Case 2
.TextMatrix(Row1, .ColIndex("Estimated")) = val(.TextMatrix(Row1, .ColIndex("Estimated"))) * ((val(.TextMatrix(Row1, .ColIndex("Estimated"))) * val(Percentage.text)) / 100)
Case 3
.TextMatrix(Row1, .ColIndex("Estimated")) = val(.TextMatrix(Row1, .ColIndex("Estimated"))) / ((val(.TextMatrix(Row1, .ColIndex("Estimated"))) * val(Percentage.text)) / 100)
End Select
End With
End If
End Sub
Sub RtriveValuLatYear(Optional Row As Long = 0)
If CompYear(0).value = True Then
 Dim Balance As Double
 Dim StrAccountCode As String
  Dim Fromd As Date
  Dim Tod As Date
If Row <> 0 Then
With Grid
StrAccountCode = .TextMatrix(Row, .ColIndex("AccountCode"))
If StrAccountCode <> "" Then
If val(.TextMatrix(1, .ColIndex("Estimated1"))) <> 0 Then
 Tod = Format(CDate("31 / 12 / " & .TextMatrix(1, .ColIndex("Estimated1")) & ""), "yyyy/MM/dd")
 Fromd = Format(CDate("1 / 01 / " & .TextMatrix(1, .ColIndex("Estimated1")) & ""), "yyyy/MM/dd")
 Balance = GetActualAccountBalance(StrAccountCode, 0, Fromd, Tod)
.TextMatrix(Row, .ColIndex("Estimated1")) = Balance
Else
.TextMatrix(Row, .ColIndex("Estimated1")) = 0
End If
If val(.TextMatrix(1, .ColIndex("Estimated2"))) <> 0 Then
 Tod = Format(CDate("31 / 12 / " & .TextMatrix(1, .ColIndex("Estimated2")) & ""), "yyyy/MM/dd")
 Fromd = Format(CDate("1 / 01 / " & .TextMatrix(1, .ColIndex("Estimated2")) & ""), "yyyy/MM/dd")
 Balance = GetActualAccountBalance(StrAccountCode, 0, Fromd, Tod)
.TextMatrix(Row, .ColIndex("Estimated2")) = Balance
Else
.TextMatrix(1, .ColIndex("Estimated2")) = 0
 End If
 If val(.TextMatrix(1, .ColIndex("Estimated3"))) <> 0 Then
 Tod = Format(CDate("31 / 12 / " & .TextMatrix(1, .ColIndex("Estimated3")) & ""), "yyyy/MM/dd")
 Fromd = Format(CDate("1 / 01 / " & .TextMatrix(1, .ColIndex("Estimated3")) & ""), "yyyy/MM/dd")
 Balance = GetActualAccountBalance(StrAccountCode, 0, Fromd, Tod)
.TextMatrix(Row, .ColIndex("Estimated3")) = Balance
Else
.TextMatrix(Row, .ColIndex("Estimated3")) = 0
 End If
 Calculate Row
 End If
End With
End If
End If
End Sub
Public Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String
    Dim ProjectID As Integer
  Dim Tota  As Double
 
    With Grid

        Select Case .ColKey(Col)
        Case "Estimated1"
        If OptMethod.value = vbChecked Then
        Calculate Row
        End If
              Case "Estimated2"
        If OptMethod.value = vbChecked Then
        Calculate Row
        End If
              Case "Estimated3"
        If OptMethod.value = vbChecked Then
        Calculate Row
        End If
        
         Case "AName"
         StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
             If GetAccountProjects("expanses_account", .TextMatrix(Row, Trim(.ColIndex("AccountCode"))), ProjectID) = True Then
            Tota = GetTotalProject(ProjectID, "total_expenses")
             End If
               If GetAccountProjects("Salary_account", .TextMatrix(Row, .ColIndex("AccountCode")), ProjectID) = True Then
               Tota = GetTotalProject(ProjectID, "total_salary")
            
             End If
               If GetAccountProjects("Material_account", .TextMatrix(Row, .ColIndex("AccountCode")), ProjectID) = True Then
               Tota = GetTotalProject(ProjectID, "total_items")
             End If
           .TextMatrix(Row, .ColIndex("Estimated")) = Tota

              StrSQL = "select Account_Serial from ACCOUNTS where Account_Code ='" & StrAccountCode & "' "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If rs.RecordCount > 0 Then
 .TextMatrix(Row, .ColIndex("ASerial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
 End If
 
 ''''''''''''''''
RtriveValuLatYear Row
    
          Case "ASerial"
             
              StrSQL = "select Account_Name,Account_Code,Account_NameEng from ACCOUNTS where Account_Serial ='" & .TextMatrix(Row, .ColIndex("ASerial")) & "' "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If rs.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 .TextMatrix(Row, .ColIndex("AName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
 Else
 .TextMatrix(Row, .ColIndex("AName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
 End If
 .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
  If GetAccountProjects("expanses_account", .TextMatrix(Row, Trim(.ColIndex("AccountCode"))), ProjectID) = True Then
            Tota = GetTotalProject(ProjectID, "total_expenses")
             End If
               If GetAccountProjects("Salary_account", .TextMatrix(Row, .ColIndex("AccountCode")), ProjectID) = True Then
               Tota = GetTotalProject(ProjectID, "total_salary")
            
             End If
               If GetAccountProjects("Material_account", .TextMatrix(Row, .ColIndex("AccountCode")), ProjectID) = True Then
               Tota = GetTotalProject(ProjectID, "total_items")
             End If
           .TextMatrix(Row, .ColIndex("Estimated")) = Tota
 End If
 RtriveValuLatYear Row
 '''''''''''''''''''''''''''''''''''
     Case "AccountCode"
             
              StrSQL = "select Account_Name,Account_Serial ,Account_NameEng from ACCOUNTS where Account_Code ='" & .TextMatrix(Row, .ColIndex("AccountCode")) & "' "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If rs.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 .TextMatrix(Row, .ColIndex("AName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
 Else
 .TextMatrix(Row, .ColIndex("AName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
 End If
 .TextMatrix(Row, .ColIndex("ASerial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
  If GetAccountProjects("expanses_account", .TextMatrix(Row, Trim(.ColIndex("AccountCode"))), ProjectID) = True Then
            Tota = GetTotalProject(ProjectID, "total_expenses")
             End If
               If GetAccountProjects("Salary_account", .TextMatrix(Row, .ColIndex("AccountCode")), ProjectID) = True Then
               Tota = GetTotalProject(ProjectID, "total_salary")
            
             End If
               If GetAccountProjects("Material_account", .TextMatrix(Row, .ColIndex("AccountCode")), ProjectID) = True Then
               Tota = GetTotalProject(ProjectID, "total_items")
             End If
           .TextMatrix(Row, .ColIndex("Estimated")) = Tota
 End If
  RtriveValuLatYear Row

           Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 Case "Distribution"
 If val(.TextMatrix(Row, .ColIndex("Distribution"))) = 1 Then
 .Cell(flexcpBackColor, Row, .ColIndex("Ser"), Row, .ColIndex("StrEstametChiled")) = &H80000018
 End If
 If val(.TextMatrix(Row, .ColIndex("Distribution"))) = 2 Then
 LonRow = Row
 Unload FrmEstametChiled
 FrmEstametChiled.txttotal.text = val(.TextMatrix(Row, .ColIndex("Estimated")))
 Load FrmEstametChiled
 
 FrmEstametChiled.show vbModal
 
 End If
        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub
Private Sub ReLineGrid()
 On Error GoTo ErrTrap
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
 
    Dim Percenrage As Double
    With Me.Grid
        Percenrage = 100 / (.Rows - 1)
        For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("Aid")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("AName")) = IntCounter

                If PercentagType(0).value = True Then
                    .TextMatrix(i, .ColIndex("Percentage")) = Round(Percenrage, 2)
                End If
         
            End If

        Next i
   .AutoSize 0, .Cols - 1, False
    End With

    IntCounter = 0

    With Me.GridIntervals

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("FromDate")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                .TextMatrix(i, .ColIndex("IntervalSerial")) = IntCounter
         
            End If

        Next i
   
    End With
ErrTrap:
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "UnitName" Then
       
            .ComboList = ""
        End If

    End With

End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Me.TxtModFlg.text = "R" Then
With Grid

 LonRow = Row
 Unload FrmEstametChiled
 Load FrmEstametChiled
 FrmEstametChiled.txttotal.text = val(.TextMatrix(Row, .ColIndex("Estimated")))
 FrmEstametChiled.show vbModal

 End With
 End If
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.Grid

        Select Case .ColKey(Col)
        Case "show"
        .ColComboList(.ColIndex("show")) = "..."
        Case "AName"
        If SystemOptions.UserInterface = ArabicInterface Then
StrSQL = "select Account_Code,Account_Name from ACCOUNTS "
Else
StrSQL = "select Account_Code,Account_NameEng from ACCOUNTS "
End If
 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = Grid.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 rs.Close
            Case "UnitName"
                            LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemId")))

                'LngItemID = 1
                If LngItemID = 0 Then
                    Cancel = True
                Else
            
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
        End Select

    End With

End Sub
Function print_report(Optional NoteSerial As String, Optional ind As Integer = 0)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT     dbo.TblEstiation.transID, dbo.TblEstiation.recordDate, dbo.TblEstiation.Fromdate, dbo.TblEstiation.todate, dbo.TblEstiation.FromdateH, dbo.TblEstiation.todateH, "
MySQL = MySQL & "                       dbo.TblEstiation.Remarks, dbo.TblEstiation.FullRemarks, dbo.TblEstiation.Percentage, dbo.TblEstiation.OperatorsID, dbo.TblEstiation.BranchId,"
MySQL = MySQL & "                       dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEstiation.Alarms, dbo.TblEstiation.ManualEntry,"
MySQL = MySQL & "                       dbo.TblEstiationYaersDetails.YearId, dbo.TblEstiationYaersDetails.Selected, dbo.TblEstiationBudgetDetails.Selected AS SelectedBud,"
MySQL = MySQL & "                       dbo.TblEstiationBudgetDetails.BudgetId, dbo.TblEstiationDetails.year1, dbo.TblEstiationDetails.year2, dbo.TblEstiationDetails.year3,"
MySQL = MySQL & "                       dbo.TblEstiationDetails.Estimated1, dbo.TblEstiationDetails.Estimated2, dbo.TblEstiationDetails.Estimated3, dbo.TblEstiationDetails.Estimated,"
MySQL = MySQL & "                       dbo.TblEstiationDetails.Acctual, dbo.TblEstiationDetails.Diff, dbo.TblEstiationDetails.Varance, dbo.TblEstiationDetails.AllowVariance,"
MySQL = MySQL & "                       dbo.TblEstiationDetails.DiffVariance, dbo.TblEstiationDetails.AccountCode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial,"
MySQL = MySQL & "                       dbo.ACCOUNTS.Account_NameEng, dbo.TblEstiationDetails.id, dbo.TblEstiationBudgetDetails.id AS idBud, dbo.TblEstiationYaersDetails.id AS idYear,"
MySQL = MySQL & "                       dbo.TblEstametChiled.EstimatID, dbo.TblEstametChiled.Dev1, dbo.TblEstametChiled.Dev2, dbo.TblEstametChiled.Dev3, dbo.TblEstametChiled.Dev4,"
MySQL = MySQL & "                       dbo.TblEstametChiled.Dev5, dbo.TblEstametChiled.Dev6, dbo.TblEstametChiled.Dev7, dbo.TblEstametChiled.Dev8, dbo.TblEstametChiled.Dev9,"
MySQL = MySQL & "                       dbo.TblEstametChiled.Dev10, dbo.TblEstametChiled.Dev11, dbo.TblEstametChiled.Dev12, dbo.TblEstametChiled.Act1, dbo.TblEstametChiled.Act2,"
MySQL = MySQL & "                       dbo.TblEstametChiled.Act3, dbo.TblEstametChiled.Act4, dbo.TblEstametChiled.Act5, dbo.TblEstametChiled.Act6, dbo.TblEstametChiled.Act7,"
MySQL = MySQL & "                       dbo.TblEstametChiled.Act8, dbo.TblEstametChiled.Act9, dbo.TblEstametChiled.Act10, dbo.TblEstametChiled.Act11, dbo.TblEstametChiled.Act12,"
MySQL = MySQL & "                       dbo.TblEstametChiled.Estim1, dbo.TblEstametChiled.Estim2, dbo.TblEstametChiled.Estim3, dbo.TblEstametChiled.Estim4, dbo.TblEstametChiled.Estim5,"
MySQL = MySQL & "                       dbo.TblEstametChiled.Estim6, dbo.TblEstametChiled.Estim7, dbo.TblEstametChiled.Estim8, dbo.TblEstametChiled.Estim9, dbo.TblEstametChiled.Estim10,"
MySQL = MySQL & "                       dbo.TblEstametChiled.Estim11, dbo.TblEstametChiled.Estim12, dbo.TblEstiationDetails.Distribution, dbo.TblEstiationDetails.StrEstametChiled,"
MySQL = MySQL & "                       dbo.TblEstiationDetails.Auto_Manul, dbo.TblEstiationYaersDetails.Remarks AS RemarksYearH, dbo.TblEstiationYaersDetails.TypeEnterYear,"
MySQL = MySQL & "                       dbo.TblEstiation.OptMethod, dbo.TblEstiation.DesCription, dbo.TblEstiation.TypeEsstame, dbo.TblEstiation.CompYear, dbo.TblEstiationYaersDetails.datesatrt,"
MySQL = MySQL & "                       dbo.TblEstiationYaersDetails.DateEnd"
MySQL = MySQL & "  FROM         dbo.TblEstiationDetails LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEstametChiled ON dbo.TblEstiationDetails.id = dbo.TblEstametChiled.EstimatID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ON dbo.TblEstiationDetails.AccountCode = dbo.ACCOUNTS.Account_Code RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEstiation ON dbo.TblEstiationDetails.transID = dbo.TblEstiation.transID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEstiationBudgetDetails ON dbo.TblEstiation.transID = dbo.TblEstiationBudgetDetails.transID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEstiationYaersDetails ON dbo.TblEstiation.transID = dbo.TblEstiationYaersDetails.transID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblEstiation.BranchId = dbo.TblBranchesData.branch_id"

MySQL = MySQL & "  Where (dbo.TblEstiation.TransID = " & val(TxtTransID.text) & ")"
If ind = 1 Then
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEstmateTotal.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEstmateTotalE.rpt"
       End If
 Else
  If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEstmate.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEstmateEn.rpt"
       End If
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
    
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
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim RsDev1 As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
          
       GridIntervals1.Clear flexClearScrollable, flexClearEverything
    GridIntervals1.Rows = 1
    
           GridOldEstimation.Clear flexClearScrollable, flexClearEverything
    GridOldEstimation.Rows = 1
    
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtTransID.text = IIf(IsNull(rs("transID").value), "", rs("transID").value)
 
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
  DcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
  DcbProject.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
    Fromdate.value = IIf(IsNull(rs("Fromdate").value), Date, rs("Fromdate").value)
  '  fromdateH.value = IIf(IsNull(rs("FromDateh").value), ToHijriDate(Date), rs("FromDateh").value)
    
        todate.value = IIf(IsNull(rs("todate").value), Date, rs("todate").value)
 'todateH.value = IIf(IsNull(rs("todateH").value), ToHijriDate(Date), rs("todateH").value)
   
    If rs("OptMethod").value = True Then
    OptMethod.value = vbChecked
    Else
    OptMethod.value = vbUnchecked
    End If
   
    TxtRemarks.text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
TxtRemarks2.text = IIf(IsNull(rs("FullRemarks").value), 0, rs("FullRemarks").value)

OperatorsID.ListIndex = IIf(IsNull(rs("OperatorsID").value), -1, rs("OperatorsID").value)

Percentage.text = IIf(IsNull(rs("Percentage").value), 0, rs("Percentage").value)


    If (rs("Alarms").value) = 0 Then
        OptAlarms(0).value = True
    ElseIf (rs("Alarms").value) = 1 Then

        OptAlarms(1).value = True
   
    End If

    If (rs("ManualEntry").value) = 0 Then
        OptActual(0).value = True
    ElseIf (rs("ManualEntry").value) = 1 Then

        OptActual(1).value = True
   
    End If

Me.TxtDesCription.text = IIf(IsNull(rs("DesCription").value), "", rs("DesCription").value)
TypeEsstame.ListIndex = IIf(IsNull(rs("TypeEsstame").value), -1, rs("TypeEsstame").value)
   If (rs("CompYear").value) = 0 Then
        CompYear(0).value = True
    ElseIf (rs("CompYear").value) = 1 Then

        CompYear(1).value = True
   
    End If

    StrSQL = " SELECT  Varance,   dbo.TblEstiation.transID, dbo.TblEstiationDetails.AccountCode, dbo.TblEstiationDetails.year1, dbo.TblEstiationDetails.year2, dbo.TblEstiationDetails.year3, "
StrSQL = StrSQL & "   dbo.TblEstiationDetails.Estimated1, dbo.TblEstiationDetails.Estimated2, dbo.TblEstiationDetails.Estimated3, dbo.TblEstiationDetails.Estimated,"
StrSQL = StrSQL & "    dbo.TblEstiationDetails.Acctual, dbo.TblEstiationDetails.Diff, dbo.TblEstiationDetails.DiffVariance, dbo.TblEstiationDetails.AllowVariance,"
StrSQL = StrSQL & "   dbo.ACCOUNTS.account_name , dbo.ACCOUNTS.account_serial, dbo.ACCOUNTS.Account_NameEng,dbo.TblEstiationDetails.StrEstametChiled ,dbo.TblEstiationDetails.Distribution , dbo.TblEstiationDetails.ID"
StrSQL = StrSQL & "  FROM         dbo.TblEstiation INNER JOIN"
StrSQL = StrSQL & "   dbo.TblEstiationDetails ON dbo.TblEstiation.transID = dbo.TblEstiationDetails.transID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.ACCOUNTS ON dbo.TblEstiationDetails.AccountCode = dbo.ACCOUNTS.Account_Code"
 
StrSQL = StrSQL & "  WHERE     (dbo.TblEstiation.transID = " & val(Me.TxtTransID.text) & ") "
    'StrSQL = StrSQL & "  where transID=" & val(Me.TxtTransID.text)
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("StrEstametChiled")) = IIf(IsNull(RsDev("StrEstametChiled").value), "", RsDev("StrEstametChiled").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("id").value), "", RsDev("id").value)
                .TextMatrix(i, .ColIndex("ASerial")) = IIf(IsNull(RsDev("Account_Serial").value), "", RsDev("Account_Serial").value)
               .TextMatrix(i, .ColIndex("Distribution")) = IIf(IsNull(RsDev("Distribution").value), "", RsDev("Distribution").value)
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
                .Cell(flexcpBackColor, i, .ColIndex("Ser"), i, .ColIndex("StrEstametChiled")) = &H80000018
                 If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("AName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                  Else
                 .TextMatrix(i, .ColIndex("AName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                  End If
                  
             '   .TextMatrix(i, .ColIndex("year1")) = IIf(IsNull(RsDev("year1").value), 0, val(RsDev("year1").value))
             '   .TextMatrix(i, .ColIndex("year2")) = IIf(IsNull(RsDev("year2").value), 0, val(RsDev("year2").value))
             '   .TextMatrix(i, .ColIndex("year3")) = IIf(IsNull(RsDev("year3").value), 0, val(RsDev("year3").value))
                
                .TextMatrix(i, .ColIndex("Estimated1")) = IIf(IsNull(RsDev("Estimated1").value), 0, val(RsDev("Estimated1").value))
                .TextMatrix(i, .ColIndex("Estimated2")) = IIf(IsNull(RsDev("Estimated2").value), 0, val(RsDev("Estimated2").value))
                .TextMatrix(i, .ColIndex("Estimated3")) = IIf(IsNull(RsDev("Estimated3").value), 0, val(RsDev("Estimated3").value))
                .TextMatrix(i, .ColIndex("Estimated")) = IIf(IsNull(RsDev("Estimated").value), 0, val(RsDev("Estimated").value))
                
                .TextMatrix(i, .ColIndex("Acctual")) = IIf(IsNull(RsDev("Acctual").value), 0, val(RsDev("Acctual").value))
                .TextMatrix(i, .ColIndex("Diff")) = IIf(IsNull(RsDev("Acctual").value), 0, val(RsDev("Diff").value))
            .TextMatrix(i, .ColIndex("DiffVariance")) = IIf(IsNull(RsDev("DiffVariance").value), 0, val(RsDev("DiffVariance").value))
            .TextMatrix(i, .ColIndex("AllowVariance")) = IIf(IsNull(RsDev("AllowVariance").value), 0, val(RsDev("AllowVariance").value))
             .TextMatrix(i, .ColIndex("Varance")) = IIf(IsNull(RsDev("Varance").value), 0, val(RsDev("Varance").value))
             
                RsDev.MoveNext
            Next i
 
        End With

    End If
 RsDev.Close
'    StrSQL = " SELECT    * from   TblAccountsDestributionsIntervals "
'    StrSQL = StrSQL & "  where transID=" & val(Me.TxtTransID.text)

    StrSQL = "Select * From TblEstiationYaersDetails where transID=" & val(Me.TxtTransID.text) & " "

    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.GridIntervals1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
   
                .TextMatrix(i, .ColIndex("TypeEnterYear")) = IIf(IsNull(RsDev("TypeEnterYear").value), 0, (RsDev("TypeEnterYear").value))
                .TextMatrix(i, .ColIndex("YearId")) = IIf(IsNull(RsDev("YearId").value), 0, (RsDev("YearId").value))
            
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDev("Remarks").value), "", RsDev("Remarks").value)
         
                .TextMatrix(i, .ColIndex("datesatrt")) = IIf(IsNull(RsDev("datesatrt").value), "", RsDev("datesatrt").value)
            
                .TextMatrix(i, .ColIndex("DateEnd")) = IIf(IsNull(RsDev("DateEnd").value), "", RsDev("DateEnd").value)

                If RsDev("Selected").value = 1 Then
                    .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexUnchecked
                End If
  
                RsDev.MoveNext
            Next i
  .AutoSize 0, .Cols - 1, False
        End With

    End If
 RsDev.Close
 
 
StrSQL = "SELECT     dbo.TblEstiationBudgetDetails.Selected, dbo.TblEstiationBudgetDetails.transID, dbo.TblEstiation.Remarks"
StrSQL = StrSQL & " ,  dbo.TblEstiationBudgetDetails.BudgetId FROM         dbo.TblEstiation RIGHT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblEstiationBudgetDetails ON dbo.TblEstiation.transID = dbo.TblEstiationBudgetDetails.BudgetId"
'Where (dbo.TblEstiationBudgetDetails.TransID = 4)
StrSQL = StrSQL & "  WHERE     (dbo.TblEstiationBudgetDetails.transID = " & val(Me.TxtTransID.text) & ") "
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.GridOldEstimation
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
   
                .TextMatrix(i, .ColIndex("BudgetId")) = IIf(IsNull(RsDev("BudgetId").value), 0, (RsDev("BudgetId").value))
            
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDev("Remarks").value), "", RsDev("Remarks").value)
            
                
                If RsDev("Selected").value = 1 Then
                    .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexUnchecked
                End If
  
                RsDev.MoveNext
            Next i
   .AutoSize 0, .Cols - 1, False
        End With

    End If
 RsDev.Close
  
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub GridIntervals_Click()

    With GridIntervals
    
        Select Case .Col

            Case 7
                ShowGL_cc .TextMatrix(.Row, .ColIndex("NoteSerial")), , 200

        End Select

    End With

End Sub

Private Sub GridIntervals1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Long
With Grid
MergGrid
If .Rows > 2 Then
For i = 1 To .Rows - 1
RtriveValuLatYear i
Next i
BtnShow_Click
End If
End With
End Sub

Private Sub GridIntervals1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

With Me.GridIntervals1
Select Case .ColKey(Col)
Case "selectfile"
If val(.TextMatrix(Row, .ColIndex("TypeEnterYear"))) = 1 Then
Yar = .TextMatrix(Row, .ColIndex("Remarks"))
ExilSheet
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·«” Ì—«œ „‰ „·›  Ì—ÃÏ ≈Œ Ì«— ÿ—Ìﬁ…≈œŒ«· «·”‰Ê«  «·Ì"
Else
MsgBox "Please Select Aut "
End If
End If
End Select
MergGrid

End With
End Sub


Private Sub GridIntervals1_Click()
MergGrid
End Sub

Private Sub GridIntervals1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.GridIntervals1
Select Case .ColKey(Col)
Case "selectfile"
.ColComboList(.ColIndex("selectfile")) = "..."
End Select
End With
End Sub

Private Sub ISButton7_Click()
If val(DcbProject.BoundText) Then
FillProject_Graid val(DcbProject.BoundText)
RemoveEmpty
End If
End Sub

Private Sub OptActual_Click(Index As Integer)
If Index = 0 Then
CMDSelectFile.Enabled = False
CmdImport.Enabled = False
Else
CMDSelectFile.Enabled = True
CmdImport.Enabled = True
End If
End Sub

Private Sub OptMethod_Click()
'Grid_AfterEdit 1, 1
End Sub

Private Sub PercentagType_Click(Index As Integer)

    Select Case Index
        
        Case 0
            TxtPercentage.locked = True
            TxtPercentage.text = ""

        Case 1
            TxtPercentage.locked = False
            TxtPercentage.text = ""

    End Select

End Sub

Private Sub TxtModFlg_Change()
hidcol
    If Me.TxtModFlg.text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(9).Enabled = False
        Cmd(5).Enabled = False
        CmdPrintAll.Enabled = False
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        CmdPrintAll.Enabled = False
        Cmd(9).Enabled = False
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = True
        Grid.Enabled = True
        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        CmdPrintAll.Enabled = True
        Cmd(9).Enabled = True
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub TypeEsstame_Change()
If TypeEsstame.ListIndex = 0 Then
Me.DcBranch.BoundText = 0
DcBranch.Enabled = False
Else
DcBranch.Enabled = True
End If
End Sub

Private Sub TypeEsstame_Click()
TypeEsstame_Change
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If
    On Error GoTo ErrTrap
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
    Exit Sub
ErrTrap:
End Sub
''' aladein Add
Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    rs.find "transID=" & RecID, , adSearchForward, 1
    If Not (rs.EOF) Then
        Retrive
        End If
    Exit Function
ErrTrap:
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
       ' BtnUndo_Click
    End If
  End Function
Private Sub AdditemTocCmp()
 On Error GoTo ErrTrap
   If SystemOptions.UserInterface = EnglishInterface Then
    With Me.OperatorsID
        .Clear
        .AddItem "Addition"
        .AddItem "Subtraction"
        .AddItem "Multiplication"
        .AddItem "Division "
      End With
    Else
    With Me.OperatorsID
        .Clear
        .AddItem "«÷«›…"
        .AddItem "ÿ—Õ"
        .AddItem "÷—»"
        .AddItem " ﬁ”Ì„"
      End With
    End If
ErrTrap:
End Sub
Sub RemoveEmpty()
On Error Resume Next
Dim i As Integer
    With Me.Grid
    For i = 3 To .Rows
        If .Row <= 2 Then Exit Sub
        If .TextMatrix(i, .ColIndex("AName")) = "" Then
        .RemoveItem i
        End If
        Next i

    End With
End Sub

Private Sub RemoveGridRow2()
Dim StrSQL As String
Dim StrMSG As String
Dim ID As Double
Dim i As Integer
Dim k As Integer
If Me.TxtModFlg.text <> "R" Then

    With Me.Grid

        If Grid.Rows <= 2 Then Exit Sub
        If SystemOptions.UserInterface = ArabicInterface Then
        StrMSG = "”Ê› Ì „ Õ–› ﬂ· «·⁄„·Ì«  «·„— »ÿÂ »Â–« «·Õ”«» Â·  —Ìœ «·Õ–›"
        Else
        StrMSG = "It will be deleted all operations associated with this account"
        End If
        If MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title) = vbYes Then
        k = Grid.Rows - 1
        For i = .FixedRows To Grid.Rows - 1
        'If i <= val(Grid.Rows - 1) Then
       
        If .Cell(flexcpChecked, k, .ColIndex("ch")) = flexChecked Then
        ID = val(Grid.TextMatrix(k, Grid.ColIndex("id")))
            StrSQL = "Delete From TblEstametChiled Where EstimatID=" & ID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
          ' k = k - 1
       .RemoveItem k
    
       ' End If
        End If
        k = k - 1
        Next i
                
        Else
        Exit Sub
        End If
    End With

    ReLineGrid
    End If
End Sub
Private Sub RemoveGridRow11()
Dim StrSQL As String
Dim StrMSG As String
Dim ID As Double
Dim i As Integer
If Me.TxtModFlg.text <> "R" Then
    With Me.Grid
        If .Rows < 3 Then Exit Sub
        If SystemOptions.UserInterface = ArabicInterface Then
        StrMSG = "”Ê› Ì „ Õ–› ﬂ· «·⁄„·Ì«  «·„— »ÿÂ »Â–Â  «·Õ”«»«  Â·  —Ìœ «·Õ–›"
        Else
        StrMSG = "It will be deleted all operations associated with this accounts"
        End If
        If MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title) = vbYes Then
      For i = .FixedRows To Grid.Rows - 1
           ID = val(Grid.TextMatrix(i, Grid.ColIndex("id")))
            StrSQL = "Delete From TblEstametChiled Where EstimatID=" & ID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                      Next i
              Me.Grid.Clear flexClearScrollable, flexClearEverything
Me.Grid.Rows = 3
        Else
        Exit Sub
        End If
    End With

    ReLineGrid
    End If
End Sub

Private Sub ISButton3_Click()
RemoveGridRow2
End Sub

Private Sub ISButton4_Click()
 On Error Resume Next
 RemoveGridRow11
End Sub


