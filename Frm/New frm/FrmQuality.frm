VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmQuality 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9090
   ClientLeft      =   1410
   ClientTop       =   2970
   ClientWidth     =   17475
   Icon            =   "FrmQuality.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   17475
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9090
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   17475
      _cx             =   30824
      _cy             =   16034
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5655
         Left            =   0
         TabIndex        =   58
         Top             =   2160
         Width           =   17535
         _cx             =   30930
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   14871017
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "«·»Ì«‰«  «·«”«”Ì…|«·„Ê«œ «·Œ«„"
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
            Height          =   5280
            Left            =   45
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   45
            Width           =   17445
            _cx             =   30771
            _cy             =   9313
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   4560
               Left            =   -120
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   -120
               Width           =   17565
               _cx             =   30983
               _cy             =   8043
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
                  Index           =   0
                  Left            =   15855
                  TabIndex        =   61
                  Top             =   4185
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   450
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
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
                  ButtonImage     =   "FrmQuality.frx":6852
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   255
                  Index           =   1
                  Left            =   14175
                  TabIndex        =   62
                  Top             =   4185
                  Width           =   1530
                  _ExtentX        =   2699
                  _ExtentY        =   450
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› «·ﬂ·"
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
                  ButtonImage     =   "FrmQuality.frx":6DEC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid FG 
                  Height          =   3885
                  Left            =   120
                  TabIndex        =   63
                  Top             =   120
                  Width           =   17325
                  _cx             =   30559
                  _cy             =   6853
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
                  Cols            =   34
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmQuality.frx":7386
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
                  Caption         =   "0"
                  Height          =   315
                  Index           =   25
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   4155
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   315
                  Index           =   26
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   4155
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   315
                  Index           =   27
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   4155
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   315
                  Index           =   28
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   4155
                  Width           =   1050
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   315
                  Index           =   24
                  Left            =   5910
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   4155
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   315
                  Index           =   23
                  Left            =   7230
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   4155
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì« "
                  Height          =   330
                  Index           =   9
                  Left            =   11085
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   4185
                  Width           =   1440
               End
            End
            Begin C1SizerLibCtl.C1Elastic Els 
               Height          =   840
               Left            =   0
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   4440
               Width           =   17505
               _cx             =   30877
               _cy             =   1482
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
               Begin VB.TextBox TxtNoteID2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1080
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1920
               End
               Begin VB.TextBox TxtresiveVoucher 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   9765
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   120
                  Width           =   1920
               End
               Begin VB.CommandButton CmdResiveVoucher 
                  Caption         =   "«‰‘«¡ ”‰œ «” ·«„ ··ﬂ„Ì«  «·„—›Ê÷…"
                  Height          =   315
                  Left            =   13875
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   120
                  Width           =   2850
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "⁄—÷ «·«–‰"
                  Height          =   315
                  Left            =   7110
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   120
                  Width           =   2175
               End
               Begin VB.CommandButton Command7 
                  Caption         =   "⁄—÷ «·ﬁÌœ"
                  Height          =   315
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   120
                  Width           =   2880
               End
               Begin VB.TextBox TxtReceive_Voucher_Serial2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   9765
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   480
                  Width           =   1920
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "«‰‘«¡ ”‰œ «” ·«„ "
                  Height          =   315
                  Left            =   13875
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   480
                  Width           =   2850
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "⁄—÷ «·«–‰"
                  Height          =   315
                  Left            =   7110
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   480
                  Width           =   2175
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "⁄—÷ «·ﬁÌœ"
                  Height          =   315
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   480
                  Width           =   2880
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   -960
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1920
               End
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   -1080
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1920
               End
               Begin VB.TextBox Text2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   960
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1920
               End
               Begin MSComCtl2.DTPicker ReciveDate 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   83
                  Top             =   360
                  Width           =   1440
                  _ExtentX        =   2540
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   268566529
                  CurrentDate     =   38784
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·«–‰"
                  Height          =   255
                  Left            =   11835
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   195
                  Width           =   810
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·«–‰"
                  Height          =   255
                  Left            =   11835
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   555
                  Width           =   810
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ «·«” ·«„"
                  Height          =   255
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   360
                  Width           =   1050
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5280
            Index           =   7
            Left            =   18180
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   45
            Width           =   17445
            _cx             =   30771
            _cy             =   9313
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
            Begin VB.CommandButton Command8 
               Caption         =   "⁄—÷ «·ﬁÌœ"
               Height          =   315
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   4800
               Width           =   2190
            End
            Begin VB.CommandButton Command6 
               Caption         =   "⁄—÷ ”‰œ «·’—›"
               Height          =   315
               Left            =   8805
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   4800
               Width           =   2190
            End
            Begin VB.CommandButton Command5 
               Caption         =   "«‰‘«¡ ”‰œ ’—›"
               Height          =   315
               Left            =   14955
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   4800
               Width           =   2190
            End
            Begin VB.TextBox TxtIssueSerial 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   11475
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   4800
               Width           =   2190
            End
            Begin VB.TextBox Txtnots2 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   6960
               TabIndex        =   97
               Top             =   120
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   240
               Width           =   900
            End
            Begin VB.TextBox TxtTotalMaterials 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3420
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   4200
               Width           =   2190
            End
            Begin VB.TextBox txtCount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8805
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   4200
               Width           =   2190
            End
            Begin VB.TextBox TxtTotalQty 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   4200
               Width           =   2190
            End
            Begin VSFlex8UCtl.VSFlexGrid FG1 
               Height          =   3390
               Left            =   150
               TabIndex        =   91
               Top             =   720
               Width           =   17145
               _cx             =   30242
               _cy             =   5980
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
               Cols            =   17
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmQuality.frx":7890
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin MSDataListLib.DataCombo DcbStoreFinish 
               Height          =   315
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DateIssu 
               Height          =   315
               Left            =   3420
               TabIndex        =   104
               Top             =   4800
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   393216
               Format          =   268566529
               CurrentDate     =   38784
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ ”‰œ «·’—›"
               Height          =   255
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   4800
               Width           =   1290
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·«–‰"
               Height          =   315
               Left            =   13635
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   4800
               Width           =   1200
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Œ“‰ «·„Ê«œ «·Œ«„"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   33
               Left            =   4740
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   240
               Width           =   2040
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  «·„Ê«œ «·Œ«„"
               Height          =   375
               Left            =   10335
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   4200
               Width           =   2460
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰ »«·„Ê«œ «·Œ«„ «·„ÿ·Ê»… ·Â–« «·«„— Ê«· Ì ”Ì „ ”Õ»Â« „‰  „Œ“‰ «·„Ê«œ «·Œ«„"
               Height          =   255
               Left            =   10725
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   240
               Width           =   6030
            End
         End
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   0
         Width           =   17505
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   39
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
            ButtonImage     =   "FrmQuality.frx":7B3F
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   40
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
            ButtonImage     =   "FrmQuality.frx":7ED9
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   41
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
            ButtonImage     =   "FrmQuality.frx":8273
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   42
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
            ButtonImage     =   "FrmQuality.frx":860D
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "FrmQuality.frx":89A7
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÃÊœ…"
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
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   120
            Width           =   4080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   735
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   17520
         _cx             =   30903
         _cy             =   1296
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
         Begin VB.TextBox TxtBatchNo 
            Height          =   330
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   1485
         End
         Begin VB.ComboBox DcbTypeProcess 
            Height          =   315
            ItemData        =   "FrmQuality.frx":9DAC
            Left            =   3360
            List            =   "FrmQuality.frx":9DB3
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   2190
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   13920
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   2040
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   11625
            TabIndex        =   1
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   270860289
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmQuality.frx":9DBA
            Height          =   315
            Left            =   6720
            TabIndex        =   2
            Top             =   240
            Width           =   4440
            _ExtentX        =   7832
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·»« ‘"
            Height          =   285
            Index           =   5
            Left            =   1680
            TabIndex        =   57
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   285
            Index           =   3
            Left            =   5280
            TabIndex        =   48
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   13035
            TabIndex        =   18
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   285
            Index           =   4
            Left            =   16305
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   10560
            TabIndex        =   16
            Top             =   240
            Width           =   1620
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   615
         Left            =   0
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   7845
         Width           =   17475
         _cx             =   30824
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
            Height          =   405
            Left            =   360
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   105
            Width           =   4080
            _cx             =   7197
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
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   1050
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   120
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   2145
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   135
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   210
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   120
               Width           =   780
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   7785
            TabIndex        =   20
            Top             =   105
            Width           =   5925
            _ExtentX        =   10451
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
            Height          =   225
            Index           =   8
            Left            =   14010
            TabIndex        =   21
            Top             =   105
            Width           =   1485
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   495
         Left            =   0
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   8595
         Width           =   17475
         _cx             =   30824
         _cy             =   873
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
            Height          =   270
            Left            =   13635
            TabIndex        =   28
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   476
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
            ButtonImage     =   "FrmQuality.frx":9DCF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   270
            Left            =   10410
            TabIndex        =   29
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   90
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
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
            ButtonImage     =   "FrmQuality.frx":10631
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   270
            Left            =   12090
            TabIndex        =   30
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
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
            ButtonImage     =   "FrmQuality.frx":109CB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   270
            Left            =   8730
            TabIndex        =   31
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   476
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
            ButtonImage     =   "FrmQuality.frx":1722D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   270
            Left            =   7305
            TabIndex        =   32
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   476
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
            ButtonImage     =   "FrmQuality.frx":175C7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   270
            Left            =   3000
            TabIndex        =   33
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   476
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
            ButtonImage     =   "FrmQuality.frx":17B61
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   330
            Left            =   6105
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   90
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   582
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
            ButtonImage     =   "FrmQuality.frx":17EFB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   270
            Left            =   4545
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   90
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            ButtonImage     =   "FrmQuality.frx":1E75D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   615
         Left            =   0
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1440
         Width           =   17520
         _cx             =   30903
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
         Begin VB.TextBox TxtStoreID2 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   3765
            TabIndex        =   53
            Top             =   120
            Width           =   885
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   9885
            TabIndex        =   50
            Top             =   120
            Width           =   885
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmQuality.frx":1EAF7
            Left            =   14160
            List            =   "FrmQuality.frx":1EAF9
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   120
            Width           =   1830
         End
         Begin VB.TextBox TxtOderNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   12840
            TabIndex        =   45
            Top             =   120
            Width           =   1200
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   6000
            TabIndex        =   51
            Top             =   120
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName2 
            Height          =   315
            Left            =   240
            TabIndex        =   54
            Top             =   120
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Œ“‰ «·—›÷"
            Height          =   270
            Index           =   1
            Left            =   4575
            TabIndex        =   55
            Top             =   120
            Width           =   1230
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Œ“‰ «·«‰ «Ã «· «„"
            Height          =   270
            Index           =   0
            Left            =   10695
            TabIndex        =   52
            Top             =   120
            Width           =   1710
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   285
            Index           =   11
            Left            =   16200
            TabIndex        =   46
            Top             =   120
            Width           =   1620
         End
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmQuality.frx":1EAFB
      Left            =   18360
      List            =   "FrmQuality.frx":1EB0B
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   18600
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   18720
      TabIndex        =   8
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
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
      Left            =   18360
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   18480
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
            Picture         =   "FrmQuality.frx":1EB24
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQuality.frx":1EEBE
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQuality.frx":1F258
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQuality.frx":1F5F2
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQuality.frx":1F98C
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQuality.frx":1FD26
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQuality.frx":200C0
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQuality.frx":2065A
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   18480
      TabIndex        =   10
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
      ButtonImage     =   "FrmQuality.frx":209F4
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
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
      ButtonImage     =   "FrmQuality.frx":27256
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   19800
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
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
      ButtonImage     =   "FrmQuality.frx":2DAB8
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
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
      Left            =   18360
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmQuality"
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
 Dim Account_Code_dynamic As String
 Dim RevenueAccount As String
 Dim ii As Long
  Dim i As Long
 Dim TxtNoteSerialV As String
 Dim TxtNoteSerial1V As String
Private Sub Cmd_Click(index As Integer)
If Me.TxtModFlg.text <> "R" Then
Select Case index
Case 0
RemoveGridRow2
Case 1
 FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
 End Select
End If
End Sub
Sub CreateIssue()
Dim Msg As String
                If Trim(DcbStoreFinish.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Please Select Store"
                Else
                    Msg = "ÌÃ»  ÕœÌœ      „Œ“‰ «·„Ê«œ «·Œ«„"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                  DcbStoreFinish.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
           DeleteTransactiomsVoucher val(Txtnots2.text)
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

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
   ' Dim Msg As String
    Dim MYTEXT As Double
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ê› Ì „ «‰‘«¡ «–‰ ’—› „‰ Â–… «·Õ—ﬂ…  .."
        Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
    Else
        Msg = "Create ISSUE Voucher to this order ?"
    End If

    If MsgBox(Msg, vbYesNo, App.Title) = vbYes Then

        Dim Transaction_ID As Long
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
        Dim TxtNoteSerial1V As String
             
        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), DateIssu.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), DateIssu.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), DateIssu.value)
                End If
            End If
        End If
        
        If TxtNoteSerial1V = "" Then
            If Voucher_coding(val(my_branch), DateIssu.value, 10, 180, , 27) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), DateIssu.value, 10, 180, , 27) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1V = Voucher_coding(val(my_branch), DateIssu.value, 10, 180, , 27)
                End If
            End If
        End If
       MYTEXT = Transaction_ID
        Me.TxtIssueSerial = TxtNoteSerial1V
        Set RsNotesGeneral = New ADODB.Recordset
       StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
       RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        If Me.TxtModFlg.text = "N" Then
    
        Else
 
            general_noteid = val(TxtNoteID2.text)
        End If

        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TxtNoteID2.text = general_noteid
        RsNotesGeneral("branch_no").value = val(Dcbranch.BoundText)
        RsNotesGeneral("NoteDate").value = DateIssu.value
        RsNotesGeneral("NoteType").value = 240
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(10) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(DateIssu.value)
        RsNotesGeneral("sanad_month").value = Month(DateIssu.value)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim sql As String
  Dim rs2 As ADODB.Recordset
  sql = "select * from Transactions where 1=-1"
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
rs2.AddNew
rs2("Transaction_ID").value = Transaction_ID
rs2("Transaction_Date").value = DateIssu.value
rs2("Transaction_Type").value = 27
rs2("BranchID").value = val(Dcbranch.BoundText)
rs2("StoreID").value = val(Me.DcbStoreFinish.BoundText)
rs2("UserID").value = user_id
rs2("Transaction_Serial").value = MYTEXT
rs2("NoteSerial1").value = TxtNoteSerial1V
rs2("NoteSerial").value = TxtNoteSerialV
rs2.update

        Txtnots2.text = Transaction_ID
        Set RSTransDetails = New ADODB.Recordset
         StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
         RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        For RowNum = 1 To FG1.rows - 1
            If FG1.TextMatrix(RowNum, FG1.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                RSTransDetails("Item_ID").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("id")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("id"))))
                RSTransDetails("Quantity").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty"))))
                RSTransDetails("SHOWQTY").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty"))))
                RSTransDetails("showPrice").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("Cost")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("Cost"))))
                RSTransDetails("UnitID").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("unitid")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("unitid"))))
                          '«·ÊÕœ« 
           
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(FG1.TextMatrix(RowNum, FG1.ColIndex("id")))
            LngUnitID = val(FG1.TextMatrix(RowNum, FG1.ColIndex("unitid")))  ' val(Fg1.Cell(flexcpData, RowNum, Fg1.ColIndex("unitid")))
            DblQty = val(FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("Price").value = val(IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("Cost")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("Cost"))))) / RSTransDetails("QtyBySmalltUnit").value
            
            End If
            
                RSTransDetails.update
            End If

        Next RowNum
       UpdateTransactionsCost CStr(Transaction_ID)
        CREATE_VOUCHER_GEIssuVOucher Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid
        Cn.Execute " Update TblQuality set DateIssu=" & SQLDate(DateIssu.value, True) & ",  StoreFInishID=" & val(Me.DcbStoreFinish.BoundText) & ",Transaction_ID3 =" & Transaction_ID & " , IssueSerial = '" & TxtIssueSerial.text & "'  where ID=" & val(TxtSerial1.text) & ""
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „ «‰‘«¡ «·”‰œ"
       Else
       MsgBox "Create Successfully"
       End If

    End If
         RsSavRec.Resync
FiLLTXT
ErrTrap:


End Sub
Function CREATE_VOUCHER_GEIssuVOucher(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim TOTAL_COST As Variant
    TOTAL_COST = val(TxtTotalMaterials.text)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
    SngTemp = TOTAL_COST

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

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "»‰«¡ ⁄·Ï  «·ÃÊœ… —ﬁ„" & Me.TxtSerial1.text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(Me.DcbStoreFinish.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "»‰«¡ ⁄·Ï  «·ÃÊœ… —ﬁ„" & Me.TxtSerial1.text
            Else
                StrTempDes = "Quality Voucher No. " & Me.TxtSerial1.text
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Variant

            With FG1

                For i = 1 To FG1.rows - 1

                    If FG1.TextMatrix(i, FG1.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG1.TextMatrix(i, FG1.ColIndex("id")), Me.DcbStoreFinish.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        '         line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
                        line_value = val(FG1.TextMatrix(i, FG1.ColIndex("total")))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "»‰«¡ ⁄·Ï «·ÃÊœ… —ﬁ„ " & Me.TxtSerial1.text
                        Else
                            StrTempDes = "Quality  No. " & Me.TxtSerial1.text
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·„œÌ‰
   '     SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
 SngTemp = TOTAL_COST
        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ „’«—Ì› «‰ «Ã «·„Ê«œ    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  ÕœÌœ „’«—Ì› «‰ «Ã «·„Ê«
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "»‰«¡ ⁄·Ï «·ÃÊœ… —ﬁ„" & Me.TxtSerial1.text
                Else
                    StrTempDes = "Quality No. " & Me.TxtSerial1.text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            End If
      
        End If
    End If

ErrTrap:
End Function
Public Sub RetriveOrderProd(Optional order_no As String = "")
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
        FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    StrSQL = " Select * from transactions "
    StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_Type = 26) AND (dbo.Transactions.Transaction_Serial = N'" & order_no & "')"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    TxtBatchNo.text = IIf(IsNull(rs("BatchNo").value), "", rs("BatchNo").value)
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = " SELECT     dbo.Transaction_Details.Transaction_ID, Transaction_Details.*, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee,"
    StrSQL = StrSQL + "                  dbo.Transaction_Details.ShowQty , dbo.Transaction_Details.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee ,dbo.Transaction_Details.showPrice"
    StrSQL = StrSQL + "   FROM         dbo.Transaction_Details LEFT OUTER JOIN"
    StrSQL = StrSQL + "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    StrSQL = StrSQL + "      Where dbo.Transaction_Details.Transaction_ID =" & val(rs("Transaction_ID").value)
    
    
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
    With FG
        .rows = RsDetails.RecordCount + 1
        For Num = 1 To RsDetails.RecordCount
            .TextMatrix(Num, .ColIndex("ItemID")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            .TextMatrix(Num, .ColIndex("UnitId")) = IIf(IsNull(RsDetails("UnitId")), "", (RsDetails("UnitId").value))
            .TextMatrix(Num, .ColIndex("ItemCode")) = IIf(IsNull(RsDetails("Fullcode")), "", Trim(RsDetails("Fullcode").value))
            .TextMatrix(Num, .ColIndex("OriginalQty")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            .TextMatrix(Num, .ColIndex("ItemQty")) = val(.TextMatrix(Num, .ColIndex("ItemQty")))
            .TextMatrix(Num, .ColIndex("ItemQty")) = GetQty(val(.TextMatrix(Num, .ColIndex("ItemID"))))
            .TextMatrix(Num, .ColIndex("Cost")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
            .TextMatrix(Num, .ColIndex("Amount")) = val(.TextMatrix(Num, .ColIndex("ItemQty"))) * val(.TextMatrix(Num, .ColIndex("Cost")))
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))
                .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitName")), "", Trim(RsDetails("UnitName").value))
             Else
                .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(RsDetails("ItemNamee")), "", Trim(RsDetails("ItemNamee").value))
                .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitNamee")), "", Trim(RsDetails("UnitNamee").value))
            End If
            
            
             fg_StartEdit CLng(Num), FG.ColIndex("ItemSize"), False
            .TextMatrix(Num, .ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
               
                
             fg_StartEdit Num, .ColIndex("ColorID"), False
            .TextMatrix(Num, .ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            fg_StartEdit Num, .ColIndex("ClassID"), False
            .TextMatrix(Num, .ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            .TextMatrix(Num, .ColIndex("L")) = IIf(IsNull(RsDetails("L")), "", (RsDetails("L").value))
             .TextMatrix(Num, .ColIndex("W")) = IIf(IsNull(RsDetails("W")), "", (RsDetails("W").value))
             .TextMatrix(Num, .ColIndex("H1")) = IIf(IsNull(RsDetails("H1")), "", (RsDetails("H1").value))
             .TextMatrix(Num, .ColIndex("H2")) = IIf(IsNull(RsDetails("H2")), "", (RsDetails("H2").value))
             
             .TextMatrix(Num, .ColIndex("OldID")) = IIf(IsNull(RsDetails("OldID")), "", (RsDetails("OldID").value))
             .TextMatrix(Num, .ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
             .TextMatrix(Num, .ColIndex("Area")) = IIf(IsNull(RsDetails("Area")), "", (RsDetails("Area").value))
             .TextMatrix(Num, .ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
             .TextMatrix(Num, .ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))
             .TextMatrix(Num, .ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
             .TextMatrix(Num, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
            
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num
End With
    End If
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    RelainGrid
    With FG
        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("ItemCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
End Sub


Sub FillGrid()
     End Sub

Private Sub CmdResiveVoucher_Click()
If val(DCboStoreName2.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— „Œ“‰ «·—›÷"""
Else
MsgBox "Please Select Rejection Store"
End If
DCboStoreName2.SetFocus
Exit Sub
End If
If val(lbl(25).Caption) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·« ÊÃœ «Ì ﬂ„Ì«  „—›Ê÷…"
Else
MsgBox "No quantities are rejected"
End If
Exit Sub
End If
TxtNoteSerial1V = ""
    DoEvents
    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim rs2 As ADODB.Recordset
    Dim note_id As Long
    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    Dim sql As String
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String

    If SystemOptions.UserInterface = ArabicInterface Then
        
        Msg = "”Ê› Ì „ «‰‘«¡  ”‰œ  «÷«›…     .."
        Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        
    Else
        Msg = "Create Recieve Voucher to this bill ?"
    End If

    If MsgBox(Msg, vbYesNo, App.Title) = vbYes Then

        Dim Transaction_ID As Long
        

        'set rs!Transaction_Serial=  where Transaction_Type=20
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        TxtNoteSerial1V = ""
        TxtNoteSerialV = ""
        my_branch = val(Me.Dcbranch.BoundText)
        Dim NoteSerial As String
        Dim Vchr_result As String
        Dim notes_result As String
         DeleteTransactiomsVoucher val(Text1)

        If TxtresiveVoucher = "" Then
      
            If TxtNoteSerial1V = "" Then
                Vchr_result = Voucher_coding(val(my_branch), ReciveDate.value, 19, 250, , 28, , val(DCboStoreName2.BoundText))

                If Vchr_result = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «” ·«„ „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                       
                    If Vchr_result = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        TxtNoteSerial1V = Vchr_result
                    End If
                End If
            End If
                    
            If TxtNoteSerialV = "" Then
                notes_result = Notes_coding(val(my_branch), ReciveDate.value)

                If notes_result = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                       
                    If notes_result = "" Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        TxtNoteSerialV = notes_result
                    End If
                End If
            End If
            TxtresiveVoucher = TxtNoteSerial1V
        Else 'Õ«·… «· ⁄œÌ·
    
            TxtNoteSerial1V = TxtresiveVoucher
            TxtNoteSerialV = get_transaction_NoteSerial2(val(Text1.text))

            If Trim(TxtNoteSerialV) = "" Then
                TxtNoteSerialV = Notes_coding(val(my_branch), ReciveDate.value)
            End If
    
        End If

        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=28"))
        Set RsNotesGeneral = New ADODB.Recordset
        StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
        RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        general_noteid = CStr(new_id("Notes", "NoteID", "", True))
      
 Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
 sql = "select * from Transactions where 1=-1"
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
rs2.AddNew
rs2("Transaction_ID").value = Transaction_ID
rs2("Transaction_Date").value = ReciveDate.value
rs2("Transaction_Type").value = 28
rs2("BranchID").value = val(Dcbranch.BoundText)
rs2("StoreID").value = val(Me.DCboStoreName2.BoundText)
rs2("TypeTrans").value = 1

rs2("UserID").value = user_id
rs2("Transaction_Serial").value = MYTEXT
rs2("NoteSerial1").value = TxtNoteSerial1V
rs2("NoteSerial").value = TxtNoteSerialV
rs2.update
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
Set RSTransDetails = New ADODB.Recordset
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   With FG
        For RowNum = 1 To .rows - 1
            If val(.TextMatrix(RowNum, .ColIndex("ItemID"))) <> 0 Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                RSTransDetails("Item_ID").value = IIf((.TextMatrix(RowNum, .ColIndex("ItemID")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("ItemID"))))
                RSTransDetails("Quantity").value = IIf((.TextMatrix(RowNum, .ColIndex("RejectionQty")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("RejectionQty"))))
                RSTransDetails("SHOWQTY").value = IIf((.TextMatrix(RowNum, .ColIndex("RejectionQty")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("RejectionQty"))))
                RSTransDetails("showPrice").value = IIf((.TextMatrix(RowNum, .ColIndex("Cost")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("Cost"))))
                RSTransDetails("UnitID").value = IIf((.TextMatrix(RowNum, .ColIndex("UnitId")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("UnitId"))))
                
                i = RowNum
                    
                RSTransDetails("LineId").value = i
                RSTransDetails("L").value = IIf((FG.TextMatrix(i, FG.ColIndex("L")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("L"))))
                RSTransDetails("W").value = IIf((FG.TextMatrix(i, FG.ColIndex("W")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("W"))))
                RSTransDetails("H1").value = IIf((FG.TextMatrix(i, FG.ColIndex("H1")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("H1"))))
                RSTransDetails("H2").value = IIf((FG.TextMatrix(i, FG.ColIndex("H2")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("H2"))))
                RSTransDetails("NoCount").value = IIf((FG.TextMatrix(i, FG.ColIndex("NoCount")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("NoCount"))))
                RSTransDetails("Width").value = IIf((FG.TextMatrix(i, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Width"))))
                RSTransDetails("length").value = IIf((FG.TextMatrix(i, FG.ColIndex("length")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("length"))))
                RSTransDetails("Height").value = IIf((FG.TextMatrix(i, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Height"))))
                RSTransDetails("Area").value = IIf((FG.TextMatrix(i, FG.ColIndex("Area")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Area"))))
                RSTransDetails("OldID").value = val(FG.TextMatrix(i, FG.ColIndex("OldID")))
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(i, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(i, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(i, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(i, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(i, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(i, FG.ColIndex("ClassId"))))
                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(i, FG.ColIndex("Remarks")) = ""), "", Trim(FG.TextMatrix(i, FG.ColIndex("Remarks"))))
                
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(.TextMatrix(RowNum, .ColIndex("ItemID")))
            LngUnitID = val(.TextMatrix(RowNum, .ColIndex("UnitId")))  ' val(Fg1.Cell(flexcpData, RowNum, Fg1.ColIndex("unitid")))
            DblQty = val(.TextMatrix(RowNum, .ColIndex("RejectionQty")))
            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("Price").value = val(IIf((.TextMatrix(RowNum, .ColIndex("Cost")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("Cost"))))) / RSTransDetails("QtyBySmalltUnit").value
            End If
                RSTransDetails.update
            End If

        Next RowNum
        End With
      '  UpdateTransactionsCost CStr(Transaction_ID)
                
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = general_noteid ' CStr(new_id("Notes", "NoteID", "", True))
        TxtNoteID.text = general_noteid
        RsNotesGeneral("NoteDate").value = ReciveDate.value
        RsNotesGeneral("Branch_no").value = val(Me.Dcbranch.BoundText)
        RsNotesGeneral("NoteType").value = 250
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("remark").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(19) '«–‰ «÷«›…
        RsNotesGeneral("sanad_year").value = year(ReciveDate.value)
        RsNotesGeneral("sanad_month").value = Month(ReciveDate.value)
        RsNotesGeneral.update
       Cn.Execute " Update TblQuality set  Transaction_ID =" & Transaction_ID & " , Receive_Voucher_Serial = '" & TxtresiveVoucher.text & "'  where ID=" & val(TxtSerial1.text) & ""
        CREATE_VOUCHER_GE1 Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·”‰œ"
        Else
            MsgBox " Vouchers Created "
        End If
     End If
     RsSavRec.Resync
FiLLTXT
ErrTrap:
End Sub
Function CREATE_VOUCHER_GE1(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim total_shahn As Variant

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    '  SngTemp = NewGrid.GetItemsTotal(5)
    SngTemp = val(lbl(26).Caption)

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

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…

            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «·ÃÊœ… —ﬁ„" & TxtSerial1
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtSerial1
                End If
            
            Else
                StrTempDes = "ÒRecieve Voucher No. " & TxtNoteSerial1V & " From Quality NO:" & TxtSerial1
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName2.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «·ÃÊœ… —ﬁ„ —ﬁ„" & TxtSerial1 & CHR(13)
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From Quality NO:" & TxtSerial1 & CHR(13)
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Variant

            With FG

                For i = 1 To FG.rows - 1

                    If val(FG.TextMatrix(i, FG.ColIndex("ItemID"))) <> 0 Then
    
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName2.BoundText, 0)
                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = 0

                        line_value = val(FG.TextMatrix(i, FG.ColIndex("Cost"))) * val(FG.TextMatrix(i, FG.ColIndex("RejectionQty")))
                       ' line_value = Round(line_value, 0)

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V
                        Else
                            StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From Quality NO:" & TxtSerial1
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, Round(line_value, 0), 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        '   SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) '* Val(txt_Currency_rate.text) '+ Val(TXTToTAlELSHahn.text)
        SngTemp = val(lbl(26).Caption)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã „Ê«œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  „’«—Ì› «·«‰ «Ã „Ê«œ
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtSerial1
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtSerial1
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            End If
        End If
 
    End If

ErrTrap:
End Function
Function CREATE_VOUCHER_GE2(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim total_shahn As Variant

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    '  SngTemp = NewGrid.GetItemsTotal(5)
    SngTemp = val(lbl(28).Caption)

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

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…

            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «·ÃÊœ… —ﬁ„" & TxtSerial1
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtSerial1
                End If
            
            Else
                StrTempDes = "ÒRecieve Voucher No. " & TxtNoteSerial1V & " From Quality NO:" & TxtSerial1
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «·ÃÊœ… —ﬁ„ —ﬁ„" & TxtSerial1 & CHR(13)
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From Quality NO:" & TxtSerial1 & CHR(13)
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Variant

            With FG

                For i = 1 To FG.rows - 1

                    If val(FG.TextMatrix(i, FG.ColIndex("ItemID"))) <> 0 Then
    
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

                        line_value = val(FG.TextMatrix(i, FG.ColIndex("Cost"))) * val(FG.TextMatrix(i, FG.ColIndex("NetAmount")))
                       ' line_value = Round(line_value, 0)

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V
                        Else
                            StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From Quality NO:" & TxtSerial1
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, Round(line_value, 0), 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        '   SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) '* Val(txt_Currency_rate.text) '+ Val(TXTToTAlELSHahn.text)
        SngTemp = val(lbl(28).Caption)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã „Ê«œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  „’«—Ì› «·«‰ «Ã „Ê«œ
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtSerial1
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtSerial1
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            End If
        End If
 
    End If

ErrTrap:
End Function
Public Function get_transaction_NoteSerial2(Transaction_ID As Long) As String

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from Transactions where Transaction_ID=" & Transaction_ID
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        get_transaction_NoteSerial2 = ""
    Else
        get_transaction_NoteSerial2 = IIf(IsNull(rs("NoteSerial").value), 0, rs("NoteSerial").value)
    End If

End Function

Public Function get_transaction_id(NoteSerial1 As String, _
                                   Transaction_Type As Integer, _
                                   Transaction_Type_Sub As Integer) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select Transaction_ID,Transaction_Type,NoteSerial1,Transaction_Type_Sub from Transactions where NoteSerial1='" & NoteSerial1 & "' and  Transaction_Type= " & Transaction_Type '& " And Transaction_Type_Sub = " & Transaction_Type_Sub
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        get_transaction_id = 0
    Else
        get_transaction_id = IIf(IsNull(rs("Transaction_ID").value), 0, rs("Transaction_ID").value)
    End If

End Function


Private Sub Command1_Click()
If val(DCboStoreName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— „Œ“‰ «·«‰ «Ã «· «„"""
Else
MsgBox "Please Select  Store"
End If
DCboStoreName.SetFocus
Exit Sub
End If
If val(lbl(27).Caption) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·« ÊÃœ «Ì ﬂ„Ì«  "
Else
MsgBox "No Quantities "
End If
Exit Sub
End If
    DoEvents
    TxtNoteSerial1V = ""
    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim rs2 As ADODB.Recordset
    Dim note_id As Long
    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    Dim sql As String
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String

    If SystemOptions.UserInterface = ArabicInterface Then
        
        Msg = "”Ê› Ì „ «‰‘«¡  ”‰œ  «÷«›…     .."
        Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        
    Else
        Msg = "Create Recieve Voucher to this bill ?"
    End If

    If MsgBox(Msg, vbYesNo, App.Title) = vbYes Then

        Dim Transaction_ID As Long
        

        'set rs!Transaction_Serial=  where Transaction_Type=20
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        TxtNoteSerial1V = ""
        TxtNoteSerialV = ""
        my_branch = val(Me.Dcbranch.BoundText)
        Dim NoteSerial As String
        Dim Vchr_result As String
        Dim notes_result As String
         DeleteTransactiomsVoucher val(Text2.text)

        If TxtReceive_Voucher_Serial2 = "" Then
      
            If TxtNoteSerial1V = "" Then
                Vchr_result = Voucher_coding(val(my_branch), ReciveDate.value, 19, 250, , 28, , val(DCboStoreName.BoundText))

                If Vchr_result = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «” ·«„ „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                       
                    If Vchr_result = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        TxtNoteSerial1V = Vchr_result
                    End If
                End If
            End If
                    
            If TxtNoteSerialV = "" Then
                notes_result = Notes_coding(val(my_branch), ReciveDate.value)

                If notes_result = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                       
                    If notes_result = "" Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        TxtNoteSerialV = notes_result
                    End If
                End If
            End If
            TxtReceive_Voucher_Serial2 = TxtNoteSerial1V
        Else 'Õ«·… «· ⁄œÌ·
    
            TxtNoteSerial1V = TxtReceive_Voucher_Serial2
            TxtNoteSerialV = get_transaction_NoteSerial2(val(Text2.text))

            If Trim(TxtNoteSerialV) = "" Then
                TxtNoteSerialV = Notes_coding(val(my_branch), ReciveDate.value)
            End If
    
        End If

        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=28"))
        Set RsNotesGeneral = New ADODB.Recordset
        StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
        RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        general_noteid = CStr(new_id("Notes", "NoteID", "", True))
      
 Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
 sql = "select * from Transactions where 1=-1"
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
rs2.AddNew
rs2("Transaction_ID").value = Transaction_ID
rs2("Transaction_Date").value = ReciveDate.value
rs2("Transaction_Type").value = 28
rs2("TypeTrans").value = 0
rs2("BranchID").value = val(Dcbranch.BoundText)
rs2("StoreID").value = val(Me.DCboStoreName.BoundText)
rs2("UserID").value = user_id
rs2("Transaction_Serial").value = MYTEXT
rs2("NoteSerial1").value = TxtNoteSerial1V
rs2("NoteSerial").value = TxtNoteSerialV
rs2.update
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
Set RSTransDetails = New ADODB.Recordset
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   With FG
        For RowNum = 1 To .rows - 1
            If val(.TextMatrix(RowNum, .ColIndex("ItemID"))) <> 0 Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
                RSTransDetails("ColorID").value = 1
                RSTransDetails("NoSerial").value = IIf((.TextMatrix(RowNum, .ColIndex("NoSerial")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("NoSerial"))))
'                RSTransDetails("ItemSize").value = 1
'                RSTransDetails("ClassId").value = 1
                RSTransDetails("Item_ID").value = IIf((.TextMatrix(RowNum, .ColIndex("ItemID")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("ItemID"))))
                RSTransDetails("Quantity").value = IIf((.TextMatrix(RowNum, .ColIndex("NetQty")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("NetQty"))))
                RSTransDetails("SHOWQTY").value = IIf((.TextMatrix(RowNum, .ColIndex("NetQty")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("NetQty"))))
                RSTransDetails("showPrice").value = IIf((.TextMatrix(RowNum, .ColIndex("Cost")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("Cost"))))
                RSTransDetails("UnitID").value = IIf((.TextMatrix(RowNum, .ColIndex("UnitId")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("UnitId"))))
                
                  i = RowNum
'RSTransDetails("LineId").value = IIf(val(FG.TextMatrix(i, FG.ColIndex("OrderLineId"))) = 0, i, val(FG.TextMatrix(i, FG.ColIndex("LineId"))))
                RSTransDetails("LineId").value = IIf(val(FG.TextMatrix(i, FG.ColIndex("OldID"))) = 0, i, val(FG.TextMatrix(i, FG.ColIndex("OldID"))))
                RSTransDetails("L").value = IIf((FG.TextMatrix(i, FG.ColIndex("L")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("L"))))
                RSTransDetails("W").value = IIf((FG.TextMatrix(i, FG.ColIndex("W")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("W"))))
                RSTransDetails("H1").value = IIf((FG.TextMatrix(i, FG.ColIndex("H1")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("H1"))))
                RSTransDetails("H2").value = IIf((FG.TextMatrix(i, FG.ColIndex("H2")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("H2"))))
                RSTransDetails("NoCount").value = IIf((FG.TextMatrix(i, FG.ColIndex("NoCount")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("NoCount"))))
                RSTransDetails("Width").value = IIf((FG.TextMatrix(i, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Width"))))
                RSTransDetails("length").value = IIf((FG.TextMatrix(i, FG.ColIndex("length")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("length"))))
                RSTransDetails("Height").value = IIf((FG.TextMatrix(i, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Height"))))
                RSTransDetails("Area").value = IIf((FG.TextMatrix(i, FG.ColIndex("Area")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Area"))))
                RSTransDetails("OldID").value = val(FG.TextMatrix(i, FG.ColIndex("OldID")))
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(i, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(i, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(i, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(i, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(i, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(i, FG.ColIndex("ClassId"))))
                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(i, FG.ColIndex("Remarks")) = ""), "", Trim(FG.TextMatrix(i, FG.ColIndex("Remarks"))))
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(.TextMatrix(RowNum, .ColIndex("ItemID")))
            LngUnitID = val(.TextMatrix(RowNum, .ColIndex("UnitId")))  ' val(Fg1.Cell(flexcpData, RowNum, Fg1.ColIndex("unitid")))
            DblQty = val(.TextMatrix(RowNum, .ColIndex("NetQty")))
            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("Price").value = val(IIf((.TextMatrix(RowNum, .ColIndex("Cost")) = ""), Null, val(.TextMatrix(RowNum, .ColIndex("Cost"))))) / RSTransDetails("QtyBySmalltUnit").value
            End If
                RSTransDetails.update
            End If

        Next RowNum
        End With
      '  UpdateTransactionsCost CStr(Transaction_ID)
                
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = general_noteid ' CStr(new_id("Notes", "NoteID", "", True))
        TxtNoteID.text = general_noteid
        RsNotesGeneral("NoteDate").value = ReciveDate.value
        RsNotesGeneral("Branch_no").value = val(Me.Dcbranch.BoundText)
        RsNotesGeneral("NoteType").value = 250
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("remark").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(19) '«–‰ «÷«›…
        RsNotesGeneral("sanad_year").value = year(ReciveDate.value)
        RsNotesGeneral("sanad_month").value = Month(ReciveDate.value)
        RsNotesGeneral.update
       Cn.Execute " Update TblQuality set  Transaction_ID2 =" & Transaction_ID & " , Receive_Voucher_Serial2 = '" & TxtReceive_Voucher_Serial2.text & "'  where ID=" & val(TxtSerial1.text) & ""
        CREATE_VOUCHER_GE2 Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·”‰œ"
        Else
            MsgBox " Vouchers Created "
        End If
     End If
     RsSavRec.Resync
FiLLTXT
ErrTrap:
End Sub

Private Sub Command2_Click()
    Dim Transaction_ID As Double
    Transaction_ID = get_transaction_id(TxtReceive_Voucher_Serial2, 28, 28)
    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmInpoutWorkOrder.show
    FrmInpoutWorkOrder.Retrive (Transaction_ID)
End Sub

Private Sub Command3_Click()
    Dim NoteSerial As String
    NoteSerial = get_transaction_NoteSerial2(val(Text2.text))
    If val(NoteSerial) = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmAccEditJournal.show
    FrmAccEditJournal.Retrive (NoteSerial)
End Sub

Private Sub Command4_Click()
    Dim Transaction_ID As Double
    Transaction_ID = get_transaction_id(TxtresiveVoucher, 28, 28)
    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmInpoutWorkOrder.show
    FrmInpoutWorkOrder.Retrive (Transaction_ID)
End Sub

Private Sub Command5_Click()
CreateIssue
End Sub

Private Sub Command6_Click()
    Dim Transaction_ID As Double
    Transaction_ID = get_transaction_id(TxtIssueSerial.text, 27, 27)
    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmOutProductionOrder.show
    FrmOutProductionOrder.Retrive (Transaction_ID)
End Sub

Private Sub Command7_Click()
    Dim NoteSerial As String
    NoteSerial = get_transaction_NoteSerial2(val(Text1.text))
    If val(NoteSerial) = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmAccEditJournal.show
    FrmAccEditJournal.Retrive (NoteSerial)
End Sub

Private Sub Command8_Click()
    Dim NoteSerial As String
    NoteSerial = get_transaction_NoteSerial2(val(Txtnots2.text))
    If val(NoteSerial) = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmAccEditJournal.show
    FrmAccEditJournal.Retrive (NoteSerial)
End Sub

Private Sub DCboStoreName_Change()
DCboStoreName_Click (0)
End Sub

Private Sub DCboStoreName_Click(Area As Integer)
TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
End Sub

Private Sub DCboStoreName2_Change()
DCboStoreName2_Click (0)
End Sub

Private Sub DCboStoreName2_Click(Area As Integer)
TxtStoreID2.text = getStoreCoding(val(DCboStoreName2.BoundText))
End Sub

Private Sub DcbStoreFinish_Change()
DcbStoreFinish_Click (0)
End Sub

Private Sub DcbStoreFinish_Click(Area As Integer)
Text3.text = getStoreCoding(val(DcbStoreFinish.BoundText))
End Sub

Private Sub FG_AfterEdit(ByVal row As Long, ByVal Col As Long)

'RelainGrid
'Exit Sub
With FG
RelainGrid
Select Case .ColKey(Col)
Case "RejectionQty", "NetQty"
If val(.TextMatrix(row, .ColIndex("RemainQty"))) < 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬂ„Ì… «·„—›Ê÷…Ê«·ﬂ„Ì… «·„” ·„… «ﬂ»— „‰ «·ﬂ„Ì… «·„ »ﬁÌ…"
Else
MsgBox "The rejected quantity can not be greater than the quantity produced"
End If
.TextMatrix(row, .ColIndex("RejectionQty")) = 0
If val(.TextMatrix(row, .ColIndex("OriginalQty"))) <> 0 Then
.TextMatrix(row, .ColIndex("PercentReject")) = val(.TextMatrix(row, .ColIndex("RejectionQty"))) / val(.TextMatrix(row, .ColIndex("OriginalQty"))) * 100
End If
Exit Sub
End If
Case "NetQty"
If val(.TextMatrix(row, .ColIndex("RemainQty"))) < 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬂ„Ì… «·„—›Ê÷…Ê«·ﬂ„Ì… «·„” ·„… «ﬂ»— „‰ «·ﬂ„Ì… «·„ »ﬁÌ…"
Else
MsgBox "The rejected quantity can not be greater than the quantity produced"
End If
.TextMatrix(row, .ColIndex("NetQty")) = 0
If val(.TextMatrix(row, .ColIndex("OriginalQty"))) <> 0 Then
.TextMatrix(row, .ColIndex("PercentReceive")) = val(.TextMatrix(row, .ColIndex("PercentReceive"))) / val(.TextMatrix(row, .ColIndex("OriginalQty"))) * 100
End If
Exit Sub
End If
 RelainGrid
 
End Select
End With
End Sub

Private Sub FG_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
If Me.TxtModFlg.text = "R" Then Exit Sub
With FG
Select Case .ColKey(Col)
Case "OriginalQty"
Cancel = True
Case "PercentReject"
Cancel = True
Case "PercentReceive"
Cancel = True
Case "ItemCode"
Cancel = True
Case "ItemName"
Cancel = True
Case "UnitName"
Cancel = True
Case "Cost"
Cancel = True
Case "ItemQty"
Cancel = True
Case "Amount"
Cancel = True
Case "RejectionQty"
.ComboList = ""
Case "AmountRejection"
Cancel = True
Case "NetQty"
.ComboList = ""
Case "NetAmount"
Cancel = True
Case "Remarks"
.ComboList = ""
Case "RemainQty"
Cancel = True
Case "RemainValue"
Cancel = True

End Select
End With

End Sub


Sub RelainGrid()
Dim i As Integer
Dim SumItemQty As Double
Dim SumAmount As Double
Dim SumRejectionQty As Double
Dim SumAmountRejection As Double
Dim SumNetQty As Double
Dim SumNetAmount As Double

SumItemQty = 0
SumAmount = 0
SumRejectionQty = 0
SumAmountRejection = 0
SumNetQty = 0
SumNetAmount = 0
With FG
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
.TextMatrix(i, .ColIndex("RemainQty")) = val(.TextMatrix(i, .ColIndex("OriginalQty"))) - val(.TextMatrix(i, .ColIndex("ItemQty"))) - (val(.TextMatrix(i, .ColIndex("NetQty"))) + val(.TextMatrix(i, .ColIndex("RejectionQty"))))
.TextMatrix(i, .ColIndex("Amount")) = val(.TextMatrix(i, .ColIndex("ItemQty"))) * val(.TextMatrix(i, .ColIndex("Cost")))
.TextMatrix(i, .ColIndex("AmountRejection")) = val(.TextMatrix(i, .ColIndex("RejectionQty"))) * val(.TextMatrix(i, .ColIndex("Cost")))
.TextMatrix(i, .ColIndex("NetAmount")) = val(.TextMatrix(i, .ColIndex("NetQty"))) * val(.TextMatrix(i, .ColIndex("Cost")))
.TextMatrix(i, .ColIndex("RemainValue")) = val(.TextMatrix(i, .ColIndex("RemainQty"))) * val(.TextMatrix(i, .ColIndex("Cost")))
SumItemQty = SumItemQty + val(.TextMatrix(i, .ColIndex("ItemQty")))
SumAmount = SumAmount + val(.TextMatrix(i, .ColIndex("Amount")))
SumRejectionQty = SumRejectionQty + val(.TextMatrix(i, .ColIndex("RejectionQty")))
SumAmountRejection = SumAmountRejection + val(.TextMatrix(i, .ColIndex("AmountRejection")))
SumNetQty = SumNetQty + val(.TextMatrix(i, .ColIndex("NetQty")))
SumNetAmount = SumNetAmount + val(.TextMatrix(i, .ColIndex("NetAmount")))
End If
Next i
End With
lbl(23).Caption = SumItemQty
lbl(24).Caption = SumAmount
lbl(25).Caption = SumRejectionQty
lbl(26).Caption = SumAmountRejection
lbl(27).Caption = SumNetQty
lbl(28).Caption = SumNetAmount
'show_parts
End Sub
Function GetQty(Optional ItemID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(isnull(dbo.TblQualityDet.RejectionQty,0) + isnull(dbo.TblQualityDet.NetQty,0)) AS Qty"
sql = sql & " FROM         dbo.TblQuality LEFT OUTER JOIN"
sql = sql & "                      dbo.TblQualityDet ON dbo.TblQuality.ID = dbo.TblQualityDet.QualityID"
sql = sql & " WHERE     (dbo.TblQualityDet.ItemID = " & ItemID & ") AND (dbo.TblQuality.OderNo = N'" & TxtOderNo.text & "') AND (dbo.TblQuality.ID <> " & val(TxtSerial1.text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetQty = IIf(IsNull(rs2("Qty").value), 0, rs2("Qty").value)
Else
GetQty = 0
End If
End Function

Private Sub FG_CellButtonClick(ByVal row As Long, ByVal Col As Long)
With FG
Select Case .ColKey(Col)
Case "Show"
If val(.TextMatrix(row, .ColIndex("NoSerial"))) <> 0 Then
Load FrmTestCertificate
FrmTestCertificate.FindRec (val(.TextMatrix(row, .ColIndex("NoSerial"))))
FrmTestCertificate.show
Else
FrmTestCertificate.btnNew_Click
FrmTestCertificate.TxtRefNo.text = GetRefNo()
FrmTestCertificate.TxtQulityID.text = TxtSerial1.text
FrmTestCertificate.FillGrid

FrmTestCertificate.show
End If
End Select
End With
End Sub

Private Sub fg_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
Dim StrSQL As String
Dim StrList
Dim RsNote As New ADODB.Recordset
With FG
Select Case .ColKey(Col)

Case "ItemSize"
  StrSQL = "Select * From TblItemsSizes Order by SizeName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "SizeName", "SizeId")

        If StrList <> "" Then
            .ColComboList(.ColIndex("ItemSize")) = "|" & StrList
        End If
    
  Case "ColorID"
        StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "ColorName", "ColorID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("ColorID")) = "|" & StrList
        End If



Case "Show"
.ColComboList(.ColIndex("Show")) = "..."
End Select
End With
End Sub
Sub HidBtoon()
    If SystemOptions.AllowCraeJLQuality = True Then
    CmdResiveVoucher.Enabled = True
    Command1.Enabled = True
    Else
    CmdResiveVoucher.Enabled = False
    Command1.Enabled = False
    End If
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
HidBtoon
    If SystemOptions.UserInterface = ArabicInterface Then
With CBoBasedON
.Clear
.AddItem "«„— «‰ «Ã"
End With
With DcbTypeProcess
.Clear
.AddItem " —›÷"
.AddItem "›—“"
End With
Else

With CBoBasedON
.Clear
.AddItem "Product Order"
End With

With DcbTypeProcess

.Clear
.Clear
.AddItem " Rejection"
.AddItem "Sort"
End With
'fg_StartEdit 1,
End If
    conection = "select * from TblQuality order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetStores Me.DCboStoreName2
    Dcombos.GetStores Me.DcbStoreFinish
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
Function GetRefNo() As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select RefNo from TblBatchSheet where BatchNo='" & TxtBatchNo.text & "' "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetRefNo = IIf(IsNull(rs2("RefNo").value), "", rs2("RefNo").value)
Else
GetRefNo = ""
End If
End Function
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.text = "E" Then
                 StrSQL = "Delete From TblQualityDet Where QualityID =" & val(TxtSerial1.text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
              End If

    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("TypeProcess").value = val(DcbTypeProcess.ListIndex)
    RsSavRec.Fields("BasedOn").value = val(CBoBasedON.ListIndex)
    RsSavRec.Fields("OderNo").value = TxtOderNo.text
    RsSavRec.Fields("StoreID").value = val(Me.DCboStoreName.BoundText)
    RsSavRec.Fields("StoreID2").value = val(Me.DCboStoreName2.BoundText)
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("TotalItemQty").value = val(lbl(23).Caption)
    RsSavRec.Fields("TotalAmount").value = val(lbl(24).Caption)
    RsSavRec.Fields("TotalRejectionQty").value = val(lbl(25).Caption)
    RsSavRec.Fields("TotalAmountRejection").value = val(lbl(26).Caption)
    RsSavRec.Fields("TotalNetQty").value = val(lbl(27).Caption)
    RsSavRec.Fields("TotalNetAmount").value = val(lbl(28).Caption)
    RsSavRec.Fields("BatchNo").value = TxtBatchNo.text
    RsSavRec.Fields("StoreFInishID").value = val(Me.DcbStoreFinish.BoundText)
    RsSavRec.Fields("StoreID").value = val(Me.DCboStoreName.BoundText)
    RsSavRec.update
  
''//////////////////////////
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblQualityDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With FG
       For i = .FixedRows To .rows - 1
       If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
       RsDevsub.AddNew
                 RsDevsub("QualityID").value = val(Me.TxtSerial1.text)
                 RsDevsub("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                 RsDevsub("UnitId").value = IIf((.TextMatrix(i, .ColIndex("UnitId"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitId"))))
                 RsDevsub("Cost").value = IIf((.TextMatrix(i, .ColIndex("Cost"))) = "", 0, val(.TextMatrix(i, .ColIndex("Cost"))))
                 RsDevsub("ItemQty").value = IIf((.TextMatrix(i, .ColIndex("ItemQty"))) = "", 0, val(.TextMatrix(i, .ColIndex("ItemQty"))))
                 RsDevsub("RejectionQty").value = IIf((.TextMatrix(i, .ColIndex("RejectionQty"))) = "", 0, val((.TextMatrix(i, .ColIndex("RejectionQty")))))
                 RsDevsub("NetQty").value = IIf((.TextMatrix(i, .ColIndex("NetQty"))) = "", 0, val(.TextMatrix(i, .ColIndex("NetQty"))))
                 RsDevsub("Amount").value = IIf((.TextMatrix(i, .ColIndex("Amount"))) = "", 0, val(.TextMatrix(i, .ColIndex("Amount"))))
                 RsDevsub("AmountRejection").value = IIf((.TextMatrix(i, .ColIndex("AmountRejection"))) = "", 0, val(.TextMatrix(i, .ColIndex("AmountRejection"))))
                 RsDevsub("NetAmount").value = IIf((.TextMatrix(i, .ColIndex("NetAmount"))) = "", 0, val(.TextMatrix(i, .ColIndex("NetAmount"))))
                 
                 RsDevsub("RemainQty").value = IIf((.TextMatrix(i, .ColIndex("RemainQty"))) = "", 0, val(.TextMatrix(i, .ColIndex("RemainQty"))))
                 RsDevsub("RemainValue").value = IIf((.TextMatrix(i, .ColIndex("RemainValue"))) = "", 0, val(.TextMatrix(i, .ColIndex("RemainValue"))))
                 RsDevsub("OriginalQty").value = IIf((.TextMatrix(i, .ColIndex("OriginalQty"))) = "", 0, val(.TextMatrix(i, .ColIndex("OriginalQty"))))
                 RsDevsub("PercentReject").value = IIf((.TextMatrix(i, .ColIndex("PercentReject"))) = "", 0, val(.TextMatrix(i, .ColIndex("PercentReject"))))
                 RsDevsub("PercentReceive").value = IIf((.TextMatrix(i, .ColIndex("PercentReceive"))) = "", 0, val(.TextMatrix(i, .ColIndex("PercentReceive"))))
                 
                 
        
                       
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(i, .ColIndex("Remarks"))))
                RsDevsub("LineId").value = i
                RsDevsub("L").value = IIf((FG.TextMatrix(i, FG.ColIndex("L")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("L"))))
                RsDevsub("W").value = IIf((FG.TextMatrix(i, FG.ColIndex("W")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("W"))))
                RsDevsub("H1").value = IIf((FG.TextMatrix(i, FG.ColIndex("H1")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("H1"))))
                RsDevsub("H2").value = IIf((FG.TextMatrix(i, FG.ColIndex("H2")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("H2"))))
                RsDevsub("NoCount").value = IIf((FG.TextMatrix(i, FG.ColIndex("NoCount")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("NoCount"))))
                RsDevsub("Width").value = IIf((FG.TextMatrix(i, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Width"))))
                RsDevsub("length").value = IIf((FG.TextMatrix(i, FG.ColIndex("length")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("length"))))
                RsDevsub("Height").value = IIf((FG.TextMatrix(i, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Height"))))
                RsDevsub("Area").value = IIf((FG.TextMatrix(i, FG.ColIndex("Area")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Area"))))
                RsDevsub("OldID").value = val(FG.TextMatrix(i, FG.ColIndex("OldID")))
                RsDevsub("ColorID").value = IIf((FG.TextMatrix(i, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(i, FG.ColIndex("ColorID"))))
                RsDevsub("ItemSize").value = IIf((FG.TextMatrix(i, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(i, FG.ColIndex("ItemSize"))))
                RsDevsub("ClassId").value = IIf((FG.TextMatrix(i, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(i, FG.ColIndex("ClassId"))))

 
        
       RsDevsub.update
      End If
     Next i
    End With

  ''///////////////////
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
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
    TxtBatchNo.text = IIf(IsNull(RsSavRec.Fields("BatchNo").value), "", RsSavRec.Fields("BatchNo").value)
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    DcbTypeProcess.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeProcess").value), -1, RsSavRec.Fields("TypeProcess").value)
    Me.CBoBasedON.ListIndex = IIf(IsNull(RsSavRec.Fields("BasedOn").value), -1, RsSavRec.Fields("BasedOn").value)
    TxtOderNo.text = IIf(IsNull(RsSavRec.Fields("OderNo").value), "", RsSavRec.Fields("OderNo").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID").value), 0, RsSavRec.Fields("StoreID").value)   ': ProgressBar1.value = 90
    Me.DCboStoreName2.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID2").value), "", RsSavRec.Fields("StoreID2").value)
    Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)  ': ProgressBar1.value = 10
    TxtresiveVoucher.text = IIf(IsNull(RsSavRec.Fields("Receive_Voucher_Serial").value), "", RsSavRec.Fields("Receive_Voucher_Serial").value)
    TxtReceive_Voucher_Serial2.text = IIf(IsNull(RsSavRec.Fields("Receive_Voucher_Serial2").value), "", RsSavRec.Fields("Receive_Voucher_Serial2").value)
    lbl(23).Caption = IIf(IsNull(RsSavRec.Fields("TotalItemQty").value), "", RsSavRec.Fields("TotalItemQty").value)
    lbl(24).Caption = IIf(IsNull(RsSavRec.Fields("TotalAmount").value), "", RsSavRec.Fields("TotalAmount").value)
    lbl(25).Caption = IIf(IsNull(RsSavRec.Fields("TotalRejectionQty").value), "", RsSavRec.Fields("TotalRejectionQty").value)
    lbl(26).Caption = IIf(IsNull(RsSavRec.Fields("TotalAmountRejection").value), "", RsSavRec.Fields("TotalAmountRejection").value)
    lbl(27).Caption = IIf(IsNull(RsSavRec.Fields("TotalNetQty").value), "", RsSavRec.Fields("TotalNetQty").value)
    lbl(28).Caption = IIf(IsNull(RsSavRec.Fields("TotalNetAmount").value), "", RsSavRec.Fields("TotalNetAmount").value)
    Text1.text = IIf(IsNull(RsSavRec.Fields("Transaction_ID").value), "", RsSavRec.Fields("Transaction_ID").value)
    Text2.text = IIf(IsNull(RsSavRec.Fields("Transaction_ID2").value), "", RsSavRec.Fields("Transaction_ID2").value)
    Txtnots2.text = IIf(IsNull(RsSavRec.Fields("Transaction_ID3").value), "", RsSavRec.Fields("Transaction_ID3").value)
    TxtresiveVoucher.text = IIf(IsNull(RsSavRec.Fields("Receive_Voucher_Serial").value), "", RsSavRec.Fields("Receive_Voucher_Serial").value)
    Me.DcbStoreFinish.BoundText = IIf(IsNull(RsSavRec.Fields("StoreFInishID").value), "", RsSavRec.Fields("StoreFInishID").value)
    ReciveDate.value = IIf(IsNull(RsSavRec.Fields("ReciveDate").value), Date, RsSavRec.Fields("ReciveDate").value)
    DateIssu.value = IIf(IsNull(RsSavRec.Fields("DateIssu").value), Date, RsSavRec.Fields("DateIssu").value)
    TxtIssueSerial.text = IIf(IsNull(RsSavRec.Fields("IssueSerial").value), "", RsSavRec.Fields("IssueSerial").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
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
      If Dcbranch.text = "" And val(Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            Dcbranch.SetFocus
            Exit Sub
     End If
    If DcbTypeProcess.text = "" And val(DcbTypeProcess.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·⁄„·Ì…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Type Process ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            DcbTypeProcess.SetFocus
            Exit Sub
     End If
     
    If CBoBasedON.text = "" And val(CBoBasedON.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— »‰«¡ ⁄·Ï  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Based On ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            CBoBasedON.SetFocus
            Exit Sub
     End If
         If TxtOderNo.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡   «œŒ«· —ﬁ„ «·«„—  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Order No. ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            TxtOderNo.SetFocus
            Exit Sub
     End If
     
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text
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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblQuality", "ID", "")
    Me.TxtSerial1.text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 1
sql = " SELECT     dbo.TblQualityDet.ID, dbo.TblQualityDet.QualityID, dbo.TblQualityDet.Cost, dbo.TblQualityDet.ItemQty, dbo.TblQualityDet.RejectionQty, dbo.TblQualityDet.NetQty,TblQualityDet.*, "
sql = sql + "                      dbo.TblQualityDet.Amount, dbo.TblQualityDet.AmountRejection, dbo.TblQualityDet.NetAmount, dbo.TblQualityDet.ItemID, dbo.TblItems.ItemName,"
sql = sql + "                      dbo.TblItems.ItemNamee , dbo.TblItems.fullcode, dbo.TblQualityDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee ,dbo.TblQualityDet.Remarks ,dbo.TblQualityDet.RemainQty ,dbo.TblQualityDet.RemainValue  ,dbo.TblQualityDet.NoSerial ,dbo.TblQualityDet.OriginalQty,dbo.TblQualityDet.PercentReject,dbo.TblQualityDet.PercentReceive"
sql = sql + " FROM         dbo.TblQualityDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblUnites ON dbo.TblQualityDet.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
sql = sql + "                      dbo.TblItems ON dbo.TblQualityDet.ItemID = dbo.TblItems.ItemID"
sql = sql + " Where (dbo.TblQualityDet.QualityID = " & val(TxtSerial1.text) & ")"
sql = sql + " Order By IsNull(TblQualityDet.LineID,TblQualityDet.Id)"
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With FG
     .rows = .FixedRows + Rs1.RecordCount
              For i = .FixedRows To Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("NoSerial")) = IIf(IsNull(Rs1("NoSerial").value), "", Rs1("NoSerial").value)
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   .TextMatrix(i, .ColIndex("Cost")) = IIf(IsNull(Rs1("Cost").value), "", Rs1("Cost").value)
                   .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("ItemQty")) = IIf(IsNull(Rs1("ItemQty").value), "", Rs1("ItemQty").value)
                   .TextMatrix(i, .ColIndex("RejectionQty")) = IIf(IsNull(Rs1("RejectionQty").value), "", Rs1("RejectionQty").value)
                   .TextMatrix(i, .ColIndex("NetQty")) = IIf(IsNull(Rs1("NetQty").value), "", Rs1("NetQty").value)
                   .TextMatrix(i, .ColIndex("Amount")) = IIf(IsNull(Rs1("Amount").value), "", Rs1("Amount").value)
                   .TextMatrix(i, .ColIndex("AmountRejection")) = IIf(IsNull(Rs1("AmountRejection").value), "", Rs1("AmountRejection").value)
                   .TextMatrix(i, .ColIndex("NetAmount")) = IIf(IsNull(Rs1("NetAmount").value), "", Rs1("NetAmount").value)
                   .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value)
                   .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(Rs1("UnitId").value), 0, Rs1("UnitId").value)
                   .TextMatrix(i, .ColIndex("RemainQty")) = IIf(IsNull(Rs1("RemainQty").value), 0, Rs1("RemainQty").value)
                   .TextMatrix(i, .ColIndex("RemainValue")) = IIf(IsNull(Rs1("RemainValue").value), 0, Rs1("RemainValue").value)
                   .TextMatrix(i, .ColIndex("OriginalQty")) = IIf(IsNull(Rs1("OriginalQty").value), "", Rs1("OriginalQty").value)
                   .TextMatrix(i, .ColIndex("PercentReject")) = IIf(IsNull(Rs1("PercentReject").value), "", Rs1("PercentReject").value)
                   .TextMatrix(i, .ColIndex("PercentReceive")) = IIf(IsNull(Rs1("PercentReceive").value), "", Rs1("PercentReceive").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", (Rs1("ItemName").value))
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", (Rs1("UnitName").value))
                   Else
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", (Rs1("ItemNamee").value))
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", (Rs1("UnitNamee").value))
                   End If
                   
                   fg_StartEdit CLng(i), FG.ColIndex("ItemSize"), False
            
               
                
             fg_StartEdit i, .ColIndex("ColorID"), False
           
            fg_StartEdit i, .ColIndex("ClassID"), False
                   
                    FG.TextMatrix(i, FG.ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks")), "", (Rs1("Remarks").value))
                    FG.TextMatrix(i, FG.ColIndex("L")) = IIf(IsNull(Rs1("L")), "", (Rs1("L").value))
                    FG.TextMatrix(i, FG.ColIndex("W")) = IIf(IsNull(Rs1("W")), "", (Rs1("W").value))
                    FG.TextMatrix(i, FG.ColIndex("OldID")) = IIf(IsNull(Rs1("OldID")), "", (Rs1("OldID").value))
                    FG.TextMatrix(i, FG.ColIndex("H1")) = IIf(IsNull(Rs1("H1")), "", (Rs1("H1").value))
                    FG.TextMatrix(i, FG.ColIndex("H2")) = IIf(IsNull(Rs1("H2")), "", (Rs1("H2").value))
                    FG.TextMatrix(i, FG.ColIndex("NoCount")) = IIf(IsNull(Rs1("NoCount")), "", (Rs1("NoCount").value))
                    FG.TextMatrix(i, FG.ColIndex("Area")) = IIf(IsNull(Rs1("Area")), "", (Rs1("Area").value))
                    FG.TextMatrix(i, FG.ColIndex("Height")) = IIf(IsNull(Rs1("Height")), "", (Rs1("Height").value))
                    FG.TextMatrix(i, FG.ColIndex("length")) = IIf(IsNull(Rs1("length")), "", (Rs1("length").value))
                    FG.TextMatrix(i, FG.ColIndex("Width")) = IIf(IsNull(Rs1("Width")), "", (Rs1("Width").value))
                    FG.TextMatrix(i, FG.ColIndex("ColorID")) = IIf(IsNull(Rs1("ColorID")), 1, (Rs1("ColorID").value))
                    FG.TextMatrix(i, FG.ColIndex("ItemSize")) = IIf(IsNull(Rs1("ItemSize")), 1, (Rs1("ItemSize").value))
                    FG.TextMatrix(i, FG.ColIndex("ClassID")) = IIf(IsNull(Rs1("ClassID")), 1, (Rs1("ClassID").value))


                   
                   
                   Rs1.MoveNext
             Next i
        End With
        RelainGrid
' ''/////////////////////////
        Exit Sub
ErrTrap:
    End Sub
    Private Sub RemoveGridRow2()
    With Me.FG
        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With
    ReLineGrid
End Sub


Private Sub ISButton5_Click()
print_report
End Sub

Public Function show_parts()
    Dim RowNum As Integer
    FG1.Clear flexClearScrollable, flexClearEverything
    FG1.rows = 2
    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("ItemID")) <> "" Then
            If add_part_item(val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID"))), val(FG.TextMatrix(RowNum, FG.ColIndex("NetQty"))) + val(FG.TextMatrix(RowNum, FG.ColIndex("RejectionQty")))) Then
        
            End If
        End If

    Next RowNum

End Function
Public Function add_part_item(LngItemID As Long, _
                              Optional Qty As Double) As Boolean
    '131315
    Dim StrSQL As String
    Dim RsParts As ADODB.Recordset
    Dim i As Integer
  
    StrSQL = "SELECT  dbo.TblItemsParts.Unitid,  dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.TableID   ,dbo.TblItems.ItemName, dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.ItemID, dbo.TblItems.ItemCode"
    StrSQL = StrSQL + " FROM         dbo.TblItems INNER JOIN "
    StrSQL = StrSQL + " dbo.TblItemsParts ON dbo.TblItems.ItemID = dbo.TblItemsParts.PartItemID"
    StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & LngItemID
    StrSQL = StrSQL + " Order By TableID"
    Dim item_cost As Variant
    Set RsParts = New ADODB.Recordset
    RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsParts.EOF Or RsParts.BOF) Then

        For i = 0 To RsParts.RecordCount - 1
               
            item_cost = ModItemCostPrice.GetCostItemPrice(RsParts("PartItemID").value, 0, , , SystemOptions.SysMainStockCostMethod, , , , , RsParts("Unitid").value)

            If add_item_to_parts_grid(val(RsParts("PartItemID").value), RsParts("ItemCode").value, RsParts("ItemName").value, item_cost, val(RsParts("PartItemQty").value), Qty, val(RsParts("Unitid").value)) = True Then
            End If
                  
            RsParts.MoveNext
        Next i

    End If

End Function
Public Function add_item_to_parts_grid(ItemID As Long, _
                                       itemcode As String, _
                                       ItemName As String, _
                                       cost As Variant, _
                                       Qty As Double, _
                                       productQty As Double, Optional UnitID As Integer)
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long
    Dim StrSQL As String
    LngNewRow = ModFgLib.SetFgForNewRow(FG1, FG1.ColIndex("Code"))

    StrSQL = "SELECT TblItemsUnits.JunckID, TblItemsUnits.ItemID, TblItemsUnits.UnitID," & "TblUnites.UnitName, TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder,TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice,TblItemsUnits.UnitPurPrice"
    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits ON TblUnites.UnitID =" & "TblItemsUnits.UnitID "
    StrSQL = StrSQL + " Where  TblUnites.UnitID=" & val(UnitID)
    Dim rs As New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    Dim UnitName As String

    If Not (rs.BOF Or rs.EOF) Then
        UnitID = IIf(IsNull(rs("UnitID").value), 0, rs("UnitID").value)
        UnitName = IIf(IsNull(rs("UnitName").value), 0, rs("UnitName").value)
    End If

    With Me.FG1
        .TextMatrix(LngNewRow, .ColIndex("id")) = ItemID
        .TextMatrix(LngNewRow, .ColIndex("code")) = itemcode
        .TextMatrix(LngNewRow, .ColIndex("Name")) = ItemName
        .TextMatrix(LngNewRow, .ColIndex("count")) = Qty
        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = UnitID
        .TextMatrix(LngNewRow, .ColIndex("Unitname")) = UnitName
        .TextMatrix(LngNewRow, .ColIndex("Cost")) = cost
        .TextMatrix(LngNewRow, .ColIndex("Valu")) = cost * Qty
        .TextMatrix(LngNewRow, .ColIndex("TotalQty")) = productQty * Qty
        .TextMatrix(LngNewRow, .ColIndex("Total")) = productQty * cost * Qty
        .AutoSize 0, .Cols - 1, False
        If .rows > 1 Then
            Me.TxtTotalMaterials.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .rows - 1, .ColIndex("Total"))
            Me.txtCount.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Count"), .rows - 1, .ColIndex("Count"))
            Me.TxtTotalQty.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalQty"), .rows - 1, .ColIndex("TotalQty"))
            
        Else
            Me.TxtTotalMaterials.text = 0
                 Me.txtCount.text = 0
                      Me.TxtTotalQty.text = 0
        End If

    End With

End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
 Dim StoreID As Integer
    If KeyAscii = vbKeyReturn Then
    StoreID = getStoreInformatin(Text3.text)
        Me.DcbStoreFinish.BoundText = StoreID
    End If
End Sub

Private Sub TxtOderNo_Change()
If Me.TxtModFlg.text <> "R" Then
RetriveOrderProd TxtOderNo.text
ReLineGrid
End If
End Sub

Private Sub TxtOderNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
      Order_no_search2.RetrunType = 40
            Order_no_search2.show vbModal
    End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
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
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
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
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
      Dim StrSQL As String
      DeleteTransactiomsVoucher val(Text2.text)
      DeleteTransactiomsVoucher val(Text1.text)
      DeleteTransactiomsVoucher val(Txtnots2.text)
      StrSQL = "delete from   TblQualityDet where QualityID =" & val(TxtSerial1.text) & ""
      Cn.Execute StrSQL
                RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                                          RsSavRec.delete
            FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 1
            lbl(23).Caption = "0"
            lbl(24).Caption = "0"
            lbl(25).Caption = "0"
            lbl(26).Caption = "0"
            lbl(27).Caption = "0"
            lbl(28).Caption = "0"
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
           'Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
        Select Case Me.TxtModFlg.text
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
    If TxtModFlg.text = "N" Then
    XPDtbTrans.Enabled = True
CmdResiveVoucher.Enabled = False
Command1.Enabled = False
Command4.Enabled = False
Command2.Enabled = False
Command7.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        
    ElseIf TxtModFlg.text = "R" Then
  HidBtoon
    Command4.Enabled = True
    Command2.Enabled = True
    Command7.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
        XPDtbTrans.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.text <> "" Then
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
   ElseIf TxtModFlg.text = "E" Then
 '  CmdResiveVoucher.Enabled = False
Command1.Enabled = False
Command4.Enabled = False
Command2.Enabled = False
Command7.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
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
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
    If TxtSerial1.text <> "" Then
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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

    TxtModFlg.text = "N"
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
    XPDtbTrans.value = Date
    ReciveDate.value = Date
    DcbTypeProcess.ListIndex = 0
    lbl(23).Caption = "0"
    lbl(24).Caption = "0"
    lbl(25).Caption = "0"
    lbl(26).Caption = "0"
    lbl(27).Caption = "0"
    lbl(28).Caption = "0"
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            Sendkeys "{TAB}"
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
   ''''''''''''''''''''////
 


   lbl(33).Caption = "ROM Store"
   C1Tab1.Caption = "Data|ROM Items"
   Label5.Caption = "Raw Of Material Item"
   Label10.Caption = "Total"
   Me.lbl(4).Caption = "ID"
   Me.lbl(2).Caption = "Date"
   lbl(7).Caption = "Branch"
   lbl(3).Caption = "Type"
   lbl(11).Caption = "Based On"
   lbl(0).Caption = "Finish Goods Store"
   lbl(1).Caption = "Rejection Store"
   Cmd(0).Caption = "Delete"
   Cmd(1).Caption = "Delete All"
   lbl(9).Caption = "Totals"
   Label6.Caption = "No"
   Label7.Caption = "Issue Date"
   Command6.Caption = "View VCHR"
   Command8.Caption = "View JE"
   Command5.Caption = "Create Issue VCHR"
   
   Command4.Caption = "View VCHR"
   Command2.Caption = "View VCHR"
   Command7.Caption = "View JE"
   Command3.Caption = "View JE"
   Label3.Caption = "No."
   Label16.Caption = "No."
   CmdResiveVoucher.Caption = "Create Receive Voucher Rejection"
   Command1.Caption = "Create Receive Voucher"
    Me.Caption = "Quality "
     Label1(2).Caption = Me.Caption
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
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
    Label4.Caption = "Recive Date"
    lbl(5).Caption = "Bach No."
  With Me.FG1
        .TextMatrix(0, .ColIndex("Code")) = "Item Code "
        .TextMatrix(0, .ColIndex("Name")) = "Item Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name "
        .TextMatrix(0, .ColIndex("Valu")) = " Value "
        .TextMatrix(0, .ColIndex("TotalQty")) = "TotalQty"
        .TextMatrix(0, .ColIndex("Count")) = "Qty"
        .TextMatrix(0, .ColIndex("Cost")) = "Cost "
        .TextMatrix(0, .ColIndex("Total")) = "Total"
  End With
  With Me.FG
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
  .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
  .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
  .TextMatrix(0, .ColIndex("Cost")) = "Cost"
  .TextMatrix(0, .ColIndex("ItemQty")) = "Prod. Qty"
  .TextMatrix(0, .ColIndex("Amount")) = "Total"
  .TextMatrix(0, .ColIndex("RejectionQty")) = "Rejection Qty"
  .TextMatrix(0, .ColIndex("AmountRejection")) = "Total"
  .TextMatrix(0, .ColIndex("NetQty")) = "Qty Received"
  .TextMatrix(0, .ColIndex("NetAmount")) = "Total"
  .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
  .TextMatrix(0, .ColIndex("RemainQty")) = "Remain Qty"
  .TextMatrix(0, .ColIndex("RemainValue")) = "Total"
  .TextMatrix(0, .ColIndex("Show")) = "Show Test"
  .TextMatrix(0, .ColIndex("NoSerial")) = "No.Test"
  .TextMatrix(0, .ColIndex("OriginalQty")) = "Original Qty"
  .TextMatrix(0, .ColIndex("PercentReject")) = "Percent.Reject"
  .TextMatrix(0, .ColIndex("PercentReceive")) = "Percent.Receive"
  End With

ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblQuality"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end

Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "  SELECT     dbo.TblQuality.RecordDate, dbo.TblQuality.ID, dbo.TblQuality.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
MySQL = MySQL & "                      dbo.TblQuality.TypeProcess, dbo.TblQuality.BasedOn, dbo.TblQuality.OderNo, dbo.TblQuality.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreAdress,"
MySQL = MySQL & "                      dbo.TblQuality.StoreID2, TblStore_1.StoreName AS StoreName2, dbo.TblQuality.Receive_Voucher_Serial, dbo.TblQuality.Receive_Voucher_Serial2,"
MySQL = MySQL & "                      dbo.TblQuality.TotalItemQty, dbo.TblQuality.TotalRejectionQty, dbo.TblQuality.TotalNetQty, dbo.TblQuality.TotalAmount, dbo.TblQuality.TotalAmountRejection,"
MySQL = MySQL & "                      dbo.TblQuality.TotalNetAmount, dbo.TblQuality.ReciveDate, dbo.TblQualityDet.Cost, dbo.TblQualityDet.ItemQty, dbo.TblQualityDet.RejectionQty,"
MySQL = MySQL & "                      dbo.TblQualityDet.NetQty, dbo.TblQualityDet.Amount, dbo.TblQualityDet.AmountRejection, dbo.TblQualityDet.NetAmount, dbo.TblQualityDet.ItemID,"
MySQL = MySQL & "                      dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblQualityDet.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
MySQL = MySQL & "                      dbo.TblStore.StoreNamee, TblStore_1.StoreNamee AS StoreNamee2"
MySQL = MySQL & " FROM         dbo.TblUnites RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblQualityDet ON dbo.TblUnites.UnitID = dbo.TblQualityDet.UnitId LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems ON dbo.TblQualityDet.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblQuality ON dbo.TblQualityDet.QualityID = dbo.TblQuality.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblStore TblStore_1 ON dbo.TblQuality.StoreID2 = TblStore_1.StoreID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblStore ON dbo.TblQuality.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblQuality.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblQuality.ID = " & val(TxtSerial1.text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepRejection.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepRejectione.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
      Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
     hide_logo = False
 End Function

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreID
    End If
End Sub

Private Sub TxtStoreID2_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID2)
        DCboStoreName2.BoundText = StoreID
    End If
End Sub
