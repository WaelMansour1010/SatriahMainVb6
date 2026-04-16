VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmBankDeposite3 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”‰œ ﬁ»÷  ‰ﬁ«ÿ «·»Ì⁄  "
   ClientHeight    =   9300
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   15870
   HelpContextID   =   580
   Icon            =   "FrmBankDeposite3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   15870
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
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15825
      _cx             =   27914
      _cy             =   16325
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmBankDeposite3.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8220
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15765
         _cx             =   27808
         _cy             =   14499
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
         Caption         =   ".|«·»Ì«‰« "
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
            Height          =   7800
            Left            =   16410
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   45
            Width           =   15675
            _cx             =   27649
            _cy             =   13758
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
            Begin VB.CommandButton Command1 
               Caption         =   "⁄—÷"
               Height          =   615
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Frame Frame4 
               Caption         =   " Õ·Ì·Ï «·„»Ì⁄« "
               Height          =   6015
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   0
               Width           =   8775
               Begin VSFlex8UCtl.VSFlexGrid Gridxxx 
                  Height          =   3765
                  Left            =   0
                  TabIndex        =   119
                  Top             =   480
                  Width           =   8685
                  _cx             =   15319
                  _cy             =   6641
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
                  Rows            =   50
                  Cols            =   14
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   400
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBankDeposite3.frx":0410
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
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000A&
                  ForeColor       =   &H000000FF&
                  Height          =   375
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   0
                  Width           =   1935
               End
               Begin VB.Label Labelx 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«Ã„«·Ì "
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   13
                  Left            =   6000
                  TabIndex        =   121
                  Top             =   4440
                  Width           =   1575
               End
               Begin VB.Label lblTotalTransaction 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000A&
                  ForeColor       =   &H000000FF&
                  Height          =   375
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   4440
                  Width           =   1935
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7800
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   15675
            _cx             =   27649
            _cy             =   13758
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
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   645
               Index           =   5
               Left            =   0
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   0
               Width           =   15645
               _cx             =   27596
               _cy             =   1138
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
               Picture         =   "FrmBankDeposite3.frx":0651
               Caption         =   "”‰œ ﬁ»÷  ‰ﬁ«ÿ «·»Ì⁄  "
               Align           =   0
               AutoSizeChildren=   0
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
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3840
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   375
                  Index           =   0
                  Left            =   1695
                  TabIndex        =   36
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBankDeposite3.frx":132B
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
                  Left            =   630
                  TabIndex        =   37
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBankDeposite3.frx":16C5
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
                  Left            =   2220
                  TabIndex        =   38
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBankDeposite3.frx":1A5F
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
                  Left            =   1155
                  TabIndex        =   39
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBankDeposite3.frx":1DF9
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
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   7635
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   18105
               _cx             =   31935
               _cy             =   13467
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
               Begin VB.CheckBox chkIsRecOnly 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁ»÷ ›ﬁÿ"
                  Height          =   225
                  Left            =   13650
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   3570
                  Value           =   1  'Checked
                  Width           =   1485
               End
               Begin VB.CheckBox chkIsRetOnly 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„—œÊœ«  ›ﬁÿ"
                  Height          =   255
                  Left            =   10350
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   3540
                  Width           =   1575
               End
               Begin VB.CheckBox chkIsSalesOnly 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„»Ì⁄«  ›ﬁÿ"
                  Height          =   225
                  Left            =   11970
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   3570
                  Width           =   1485
               End
               Begin VB.CheckBox Check17 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  Height          =   270
                  Left            =   10350
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   7800
                  Width           =   1185
               End
               Begin VB.CheckBox chkDue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ «·‘Ìﬂ«  «·„” Õﬁ…  ›ﬁÿ"
                  Height          =   195
                  Left            =   6090
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   7920
                  Width           =   3345
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
                  Height          =   405
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   7890
                  Width           =   2040
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
                  Height          =   405
                  Left            =   -630
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   7650
                  Width           =   2040
               End
               Begin VB.TextBox TxtBankName 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   19740
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   2700
               End
               Begin VB.TextBox TxtNoteSerial1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   12615
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   690
                  Width           =   1590
               End
               Begin VB.TextBox XXXX 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7650
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   8610
                  Width           =   2130
               End
               Begin VB.TextBox TxtTotalCheques 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   7800
                  Width           =   1890
               End
               Begin VB.TextBox TxtTotalCash 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2400
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   7650
                  Width           =   1890
               End
               Begin VB.TextBox txtchequeno 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   19890
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   1500
               End
               Begin VB.TextBox TxtValue1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3135
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   7620
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.TextBox TxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   6465
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   3960
                  Visible         =   0   'False
                  Width           =   1770
               End
               Begin VB.Frame Frame1 
                  Caption         =   "„⁄·Ê„« "
                  Height          =   2115
                  Left            =   22950
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   1050
                  Width           =   5535
                  Begin MSDataListLib.DataCombo xxx 
                     Height          =   315
                     Index           =   0
                     Left            =   120
                     TabIndex        =   58
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
                     TabIndex        =   59
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
                     TabIndex        =   62
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
                     TabIndex        =   61
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
                     TabIndex        =   60
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
                     TabIndex        =   57
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
                     TabIndex        =   56
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
                     TabIndex        =   55
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
                     TabIndex        =   54
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
                     TabIndex        =   53
                     Top             =   840
                     Width           =   1200
                  End
               End
               Begin VB.TextBox txtRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   885
                  Left            =   9885
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   47
                  Top             =   2550
                  Width           =   4350
               End
               Begin VB.CheckBox ChkLocked 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  Height          =   465
                  Left            =   18420
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   2220
                  Width           =   2880
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
                  Height          =   375
                  Left            =   19290
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   2790
                  Value           =   -1  'True
                  Width           =   1365
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
                  Height          =   375
                  Left            =   19245
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   2790
                  Width           =   2025
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   18330
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   4530
                  Width           =   1935
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
                  Height          =   525
                  Left            =   19725
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Text            =   "0"
                  Top             =   2220
                  Visible         =   0   'False
                  Width           =   555
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
                  Height          =   315
                  Left            =   20355
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   690
                  Visible         =   0   'False
                  Width           =   1485
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   255
                  Left            =   18330
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   4650
                  Width           =   2880
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
                  Left            =   -4665
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   12090
                  Width           =   2565
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
                  Left            =   20370
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   690
                  Visible         =   0   'False
                  Width           =   2550
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   1515
                  Left            =   15600
                  TabIndex        =   7
                  Top             =   -510
                  Visible         =   0   'False
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   2672
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
                  FormatString    =   $"FrmBankDeposite3.frx":2193
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
                  Height          =   285
                  Left            =   9930
                  TabIndex        =   12
                  Top             =   690
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   112525313
                  CurrentDate     =   41640
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   19530
                  TabIndex        =   13
                  Top             =   2670
                  Width           =   5295
                  _ExtentX        =   9340
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
                  Left            =   19365
                  TabIndex        =   15
                  Top             =   1980
                  Width           =   1920
                  _ExtentX        =   3387
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
                  Left            =   20325
                  TabIndex        =   30
                  Top             =   1050
                  Width           =   3795
                  _ExtentX        =   6694
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
                  Height          =   525
                  Left            =   20730
                  TabIndex        =   42
                  Top             =   2100
                  Width           =   2460
                  _ExtentX        =   4339
                  _ExtentY        =   926
                  _Version        =   393216
                  Format          =   112525313
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   20190
                  TabIndex        =   48
                  Top             =   2790
                  Width           =   795
                  _ExtentX        =   1402
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
                  ButtonImage     =   "FrmBankDeposite3.frx":246E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   19185
                  TabIndex        =   49
                  Top             =   2790
                  Width           =   765
                  _ExtentX        =   1349
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
                  ButtonImage     =   "FrmBankDeposite3.frx":2808
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DCEmp 
                  Height          =   315
                  Left            =   19770
                  TabIndex        =   50
                  Top             =   1740
                  Width           =   3135
                  _ExtentX        =   5530
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
                  Left            =   20070
                  TabIndex        =   51
                  Top             =   2790
                  Width           =   5145
                  _ExtentX        =   9075
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
                  Height          =   2115
                  Left            =   0
                  TabIndex        =   64
                  Top             =   8160
                  Width           =   11805
                  _cx             =   20823
                  _cy             =   3731
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
                  FormatString    =   $"FrmBankDeposite3.frx":2DA2
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
                  Left            =   15780
                  TabIndex        =   66
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   4905
                  _ExtentX        =   8652
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   9285
                  TabIndex        =   67
                  Top             =   1800
                  Width           =   4920
                  _ExtentX        =   8678
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo Dcbank1 
                  Height          =   315
                  Left            =   3510
                  TabIndex        =   72
                  Top             =   8730
                  Visible         =   0   'False
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   7
                  Left            =   9075
                  TabIndex        =   82
                  Top             =   3900
                  Width           =   795
                  _ExtentX        =   1402
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
                  ButtonImage     =   "FrmBankDeposite3.frx":3109
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   9
                  Left            =   2265
                  TabIndex        =   83
                  Top             =   8295
                  Width           =   795
                  _ExtentX        =   1402
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
                  ButtonImage     =   "FrmBankDeposite3.frx":34A3
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   10
                  Left            =   0
                  TabIndex        =   84
                  Top             =   8295
                  Visible         =   0   'False
                  Width           =   1440
                  _ExtentX        =   2540
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
                  ButtonImage     =   "FrmBankDeposite3.frx":383D
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo Dcbranch 
                  Height          =   315
                  Left            =   9285
                  TabIndex        =   87
                  Top             =   2130
                  Width           =   4920
                  _ExtentX        =   8678
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCChequeBox 
                  Height          =   315
                  Left            =   3210
                  TabIndex        =   90
                  Top             =   7950
                  Width           =   6765
                  _ExtentX        =   11933
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker FromDate 
                  Height          =   285
                  Left            =   12720
                  TabIndex        =   109
                  Top             =   3900
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   112525313
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DTPicker2 
                  Height          =   285
                  Left            =   0
                  TabIndex        =   110
                  Top             =   0
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   112525313
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   285
                  Left            =   9990
                  TabIndex        =   111
                  Top             =   3900
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   112525313
                  CurrentDate     =   41640
               End
               Begin MSDataListLib.DataCombo DcGeneralBox 
                  Height          =   315
                  Left            =   9285
                  TabIndex        =   113
                  Top             =   1440
                  Width           =   4920
                  _ExtentX        =   8678
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboUserName 
                  Height          =   360
                  Left            =   9270
                  TabIndex        =   122
                  Top             =   1080
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   635
                  _Version        =   393216
                  BackColor       =   16761024
                  ForeColor       =   0
                  Text            =   ""
                  RightToLeft     =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin C1SizerLibCtl.C1Tab TabMain 
                  Height          =   3405
                  Left            =   -30
                  TabIndex        =   126
                  Top             =   4200
                  Width           =   15660
                  _cx             =   27622
                  _cy             =   6006
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
                  BackColor       =   12648447
                  ForeColor       =   -2147483630
                  FrontTabColor   =   14871017
                  BackTabColor    =   12648447
                  TabOutlineColor =   -2147483632
                  FrontTabForeColor=   16711680
                  Caption         =   "»Ì«‰«  «·ﬁ»÷|”‰œ«  «·’—›|”‰œ«  «·«” ·«„"
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
                  Begin C1SizerLibCtl.C1Elastic ELe 
                     Height          =   3030
                     Index           =   0
                     Left            =   45
                     TabIndex        =   127
                     TabStop         =   0   'False
                     Top             =   45
                     Width           =   15570
                     _cx             =   27464
                     _cy             =   5345
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
                     Begin VSFlex8UCtl.VSFlexGrid FgItems 
                        Height          =   2970
                        Index           =   0
                        Left            =   21930
                        TabIndex        =   128
                        Top             =   435
                        Width           =   15420
                        _cx             =   27199
                        _cy             =   5239
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
                        RowHeightMin    =   300
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   -1  'True
                        FormatString    =   $"FrmBankDeposite3.frx":3DD7
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
                     Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
                        Height          =   1935
                        Left            =   75
                        TabIndex        =   131
                        Top             =   30
                        Width           =   15375
                        _cx             =   27120
                        _cy             =   3413
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
                        Cols            =   31
                        FixedRows       =   1
                        FixedCols       =   2
                        RowHeightMin    =   0
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   -1  'True
                        FormatString    =   $"FrmBankDeposite3.frx":3E97
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
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   195
                        Index           =   8
                        Left            =   13305
                        TabIndex        =   132
                        Top             =   2085
                        Visible         =   0   'False
                        Width           =   1680
                        _ExtentX        =   2963
                        _ExtentY        =   344
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
                        ButtonImage     =   "FrmBankDeposite3.frx":4333
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin VB.Label lblAccountBalance 
                        Alignment       =   1  'Right Justify
                        Caption         =   "—’Ìœ «·Õ”«» "
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H000000FF&
                        Height          =   450
                        Left            =   0
                        RightToLeft     =   -1  'True
                        TabIndex        =   133
                        Top             =   2085
                        Width           =   4605
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic ELe 
                     Height          =   3030
                     Index           =   3
                     Left            =   16305
                     TabIndex        =   129
                     TabStop         =   0   'False
                     Top             =   45
                     Width           =   15570
                     _cx             =   27464
                     _cy             =   5345
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
                     Begin VSFlex8UCtl.VSFlexGrid FgItems 
                        Height          =   2955
                        Index           =   1
                        Left            =   21690
                        TabIndex        =   130
                        Top             =   465
                        Width           =   15345
                        _cx             =   27067
                        _cy             =   5212
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
                        RowHeightMin    =   300
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   -1  'True
                        FormatString    =   $"FrmBankDeposite3.frx":48CD
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
                     Begin VSFlex8UCtl.VSFlexGrid grdMaster2 
                        Height          =   2100
                        Left            =   120
                        TabIndex        =   136
                        Top             =   30
                        Width           =   15435
                        _cx             =   27226
                        _cy             =   3704
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
                        SelectionMode   =   0
                        GridLines       =   1
                        GridLinesFixed  =   2
                        GridLineWidth   =   1
                        Rows            =   12
                        Cols            =   24
                        FixedRows       =   1
                        FixedCols       =   1
                        RowHeightMin    =   320
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   -1  'True
                        FormatString    =   $"FrmBankDeposite3.frx":498D
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
                        AccessibleName  =   "ReCostDet"
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   24
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic ELe 
                     Height          =   3030
                     Index           =   4
                     Left            =   16605
                     TabIndex        =   134
                     TabStop         =   0   'False
                     Top             =   45
                     Width           =   15570
                     _cx             =   27464
                     _cy             =   5345
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
                     Begin VSFlex8UCtl.VSFlexGrid FgItems 
                        Height          =   2985
                        Index           =   2
                        Left            =   21690
                        TabIndex        =   135
                        Top             =   330
                        Width           =   15345
                        _cx             =   27067
                        _cy             =   5265
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
                        RowHeightMin    =   300
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   -1  'True
                        FormatString    =   $"FrmBankDeposite3.frx":4D62
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
                     Begin VSFlex8UCtl.VSFlexGrid grdMaster3 
                        Height          =   2580
                        Left            =   0
                        TabIndex        =   137
                        Top             =   90
                        Width           =   15435
                        _cx             =   27226
                        _cy             =   4551
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
                        SelectionMode   =   0
                        GridLines       =   1
                        GridLinesFixed  =   2
                        GridLineWidth   =   1
                        Rows            =   12
                        Cols            =   24
                        FixedRows       =   1
                        FixedCols       =   1
                        RowHeightMin    =   320
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   -1  'True
                        FormatString    =   $"FrmBankDeposite3.frx":4E22
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
                        AccessibleName  =   "ReCostDet"
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   24
                     End
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid grdMaster 
                  Height          =   2640
                  Left            =   30
                  TabIndex        =   138
                  Top             =   900
                  Width           =   9075
                  _cx             =   16007
                  _cy             =   4657
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
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   12
                  Cols            =   24
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBankDeposite3.frx":51FA
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
                  AccessibleName  =   "ReCostDet"
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin VB.Label Labelx 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«”„ «·ﬂ«‘Ì—"
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Index           =   6
                  Left            =   14070
                  TabIndex        =   123
                  Top             =   1080
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì  «—ÌŒ"
                  Height          =   285
                  Index           =   23
                  Left            =   11190
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   3900
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   285
                  Index           =   22
                  Left            =   14430
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   3900
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·‘Ìﬂ«  «·„Õœœ…"
                  Height          =   285
                  Index           =   21
                  Left            =   5910
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   8610
                  Width           =   1620
               End
               Begin VB.Label TxtPaymentCounts 
                  Alignment       =   1  'Right Justify
                  Caption         =   "0"
                  Height          =   375
                  Left            =   3900
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   7980
                  Width           =   1605
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   285
                  Index           =   19
                  Left            =   10485
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   8610
                  Width           =   810
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ Õ«›Ÿ… «·‘Ìﬂ« "
                  Height          =   285
                  Index           =   18
                  Left            =   9300
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   6450
                  Width           =   1770
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   285
                  Index           =   17
                  Left            =   14805
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   2160
                  Width           =   825
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·‘Ìﬂ« "
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   2025
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   8610
                  Width           =   1500
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·‰ﬁœ"
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   1395
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   6240
                  Width           =   1200
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·‘Ìﬂ"
                  Height          =   255
                  Left            =   6450
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   8340
                  Visible         =   0   'False
                  Width           =   825
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁÌ„Â"
                  Height          =   255
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   7920
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»‰ﬂ"
                  Height          =   285
                  Index           =   16
                  Left            =   10485
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   7590
                  Visible         =   0   'False
                  Width           =   1065
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—’Ìœ"
                  Height          =   255
                  Left            =   8025
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   3960
                  Visible         =   0   'False
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’‰œÊﬁ «·—∆Ì”Ì"
                  Height          =   285
                  Index           =   15
                  Left            =   14145
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   1440
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’‰œÊﬁ «·›—⁄Ì"
                  Height          =   285
                  Index           =   14
                  Left            =   14430
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   1830
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìœ«⁄«  ‘Ìﬂ« "
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   13
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   6060
                  Width           =   2010
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìœ«⁄«  ‰ﬁœÌ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   12
                  Left            =   15660
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   1800
                  Width           =   2430
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   315
                  Index           =   3
                  Left            =   14835
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   2670
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   525
                  Index           =   2
                  Left            =   19200
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   2100
                  Width           =   390
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰œÊ»"
                  Height          =   315
                  Index           =   0
                  Left            =   19035
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   1740
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   285
                  Index           =   5
                  Left            =   11565
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   690
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  Height          =   270
                  Index           =   8
                  Left            =   19380
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   3480
                  Width           =   2130
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”‰œ"
                  Height          =   240
                  Index           =   7
                  Left            =   14715
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   690
                  Width           =   915
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
                  Height          =   585
                  Left            =   16380
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   1170
                  Width           =   1095
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   1
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   8265
         Width           =   15765
         _cx             =   27808
         _cy             =   1693
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
         AutoSizeChildren=   0
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
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   480
            Width           =   1590
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   13080
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   570
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmBankDeposite3.frx":55CC
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   225
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBankDeposite3.frx":5966
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
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
            ButtonImage     =   "FrmBankDeposite3.frx":5D00
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   9060
            TabIndex        =   23
            Top             =   30
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   873
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
            Height          =   495
            Index           =   1
            Left            =   8160
            TabIndex        =   24
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
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
            Height          =   495
            Index           =   2
            Left            =   7320
            TabIndex        =   25
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   3
            Left            =   6315
            TabIndex        =   26
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   4
            Left            =   5280
            TabIndex        =   27
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
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
            Height          =   495
            Index           =   6
            Left            =   2040
            TabIndex        =   28
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   5
            Left            =   4350
            TabIndex        =   29
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   375
            Left            =   11040
            TabIndex        =   34
            Tag             =   "Delete Row"
            Top             =   0
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "FrmBankDeposite3.frx":609A
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
            Left            =   7680
            TabIndex        =   93
            Top             =   525
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   12
            Left            =   2880
            TabIndex        =   105
            Top             =   30
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â «·”‰œ"
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
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   375
            Left            =   6240
            TabIndex        =   116
            Top             =   525
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
            Caption         =   "ﬁÌœ —ﬁ„"
            Height          =   315
            Index           =   24
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   600
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   20
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   37
            Left            =   780
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   600
            Width           =   615
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   1830
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   600
            Width           =   825
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
            Height          =   210
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   225
            Width           =   1740
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
            Height          =   210
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   1515
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
      ButtonImage     =   "FrmBankDeposite3.frx":60B6
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
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
      TabIndex        =   102
      Top             =   9360
      Width           =   7155
   End
End
Attribute VB_Name = "FrmBankDeposite3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long
 
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "„”·”· " & TxtNoteSerial1.Text & CHR(13) & "   «· «—ÌŒ " & dbRecordDate & CHR(13) & "   «·›—⁄ " & DcBranch & CHR(13) & "   «·»‰ﬂ «·„Êœ⁄ »Â  " & Dcbank & CHR(13) & "   „·«ÕŸ«  " & txtRemarks & CHR(13) & "   —ﬁ„ «·ﬁÌœ " & TxtNoteSerial & CHR(13) & "   «Ã„«·Ì «·‰ﬁœ " & TxtTotalCashView & CHR(13) & "   «Ã„«·Ì «·‘Ìﬂ«  " & TxtTotalChequesView
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Serial " & TxtNoteSerial1.Text & CHR(13) & "   Date " & dbRecordDate & CHR(13) & "   Branch " & DcBranch & CHR(13) & "Deposite Bank" & Dcbank & CHR(13) & "   Remarks " & txtRemarks & CHR(13) & " Ge NO" & TxtNoteSerial & CHR(13) & "  Total Cash " & TxtTotalCashView & CHR(13) & "  Total Cheques " & TxtTotalChequesView
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.Show
End Sub

Function check_previous_dev(year As String, month As String) As Boolean
 
 
End Function

Function check_previous_dev1(year As String, month As String) As Boolean
 
 
End Function

Function Create_dev()
 
End Function

Function Create_dev1()
 
End Function

Private Sub ALLButton2_Click()
    'dcbank.text = ""

    dcproject.Text = ""
    FillGridWithData

    DoEvents
    Create_dev
    CmdOk_Click
End Sub

Private Sub ALLButton3_Click()
 
End Sub

Private Sub CboPayMentType_Change()
 
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        get_all_employee
    Else

        With Me.Grid
            .Rows = 2
            .Clear flexClearScrollable
        End With

    End If

End Sub

Private Sub CmbMonth_Click()
    CmdOk_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()

End Sub

Function create_report_data()

End Function

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

Private Sub Combo1_Click()
 
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

     'On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
        If CheckAcconts = False Then
              Exit Sub
        End If
 
        If Trim(Me.DcGeneralBox.BoundText) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ» ≈Œ Ì«— «·’‰œÊﬁ «·⁄«„ ..!!"
                    Else
                        Msg = "Specify General Box.!!"
                    End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcGeneralBox.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
 
         If Trim(Me.DcboBox.BoundText) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ» ≈Œ Ì«— «·’‰œÊﬁ «·›—⁄Ì ..!!"
                    Else
                        Msg = "Specify Sub Box.!!"
                    End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboBox.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
        
    End If

    '-------------------------------------------------------------------------------------------
  
    If TxtNoteSerial.Text = "" Then
        If Notes_coding(val(my_branch), dbRecordDate.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If Notes_coding(val(my_branch), dbRecordDate.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                TxtNoteSerial.Text = Notes_coding(val(my_branch), dbRecordDate.value)
            End If
        End If
    End If
        
    If TxtNoteSerial1.Text = "" Then
        If Voucher_coding(val(my_branch), dbRecordDate.value, 59, 59) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «Ìœ«⁄  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If Voucher_coding(val(my_branch), dbRecordDate.value, 59, 59) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ    ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                TxtNoteSerial1.Text = Voucher_coding(val(my_branch), dbRecordDate.value, 59, 59)
            End If
        End If
    End If
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
        TxtNoteID.Text = CStr(new_id("Notes", "NoteID", "", True))
            Me.TxtlBanksDepositeId.Text = CStr(new_id("tblGeneralCashing", "id", "", True))
            
        Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
        
    ElseIf Me.TxtModFlg.Text = "E" Then
                 
        Cn.Execute "delete tblGeneralCashingdetails where tblGeneralCashingId=" & val(Me.TxtlBanksDepositeId.Text)
        StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
    End If
    
    rs("id").value = TxtlBanksDepositeId.Text
     
    rs("branch_no").value = IIf(Me.DcBranch.BoundText = "", Null, Me.DcBranch.BoundText)
    
    rs("GeneralBoxId").value = IIf(Me.DcGeneralBox.BoundText = "", Null, Me.DcGeneralBox.BoundText)
    rs("SubBoxId").value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
    rs("CashierId").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
    
    rs("FromDate").value = FromDate.value
    rs("ToDate").value = ToDate.value
   
    rs("RecordDate").value = dbRecordDate.value
    rs("Remarks").value = IIf(Me.txtRemarks.Text = "", "", Me.txtRemarks.Text)
 
    rs("NoteID").value = CStr(TxtNoteID.Text)
    rs("NoteSerial").value = Trim$(Me.TxtNoteSerial.Text) '„”·”· «·ﬁÌœ
    rs("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.Text) '„”·”· «–‰ «·’—›
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
 
    rs.update
    
    
    
    Dim i As Integer

recalgrid
    Set RsDev = New ADODB.Recordset
        
   ' RsDev.Open "tblGeneralCashingDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
       StrSQL = "SELECT     *  from dbo.tblGeneralCashingDetails Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        




 
  
    With Me.VSFlexGrid1

        For i = 1 To .Rows - 1
 
            If .TextMatrix(i, .ColIndex("PaymentID")) <> "" Then
         
                RsDev.AddNew
                RsDev("tblGeneralCashingId").value = Me.TxtlBanksDepositeId.Text
                RsDev("TransType").value = val(.TextMatrix(i, .ColIndex("PaymentID")))
                RsDev("value").value = val(.TextMatrix(i, .ColIndex("balance")))
                RsDev("CollectedValue").value = val(.TextMatrix(i, .ColIndex("CollectedValue")))
                RsDev("CommissionValue").value = val(.TextMatrix(i, .ColIndex("CommissionValue")))
                RsDev("Different").value = val(.TextMatrix(i, .ColIndex("Different")))
                 RsDev("Accountsus").value = .TextMatrix(i, .ColIndex("Accountsus"))
                 RsDev("Accountcom").value = .TextMatrix(i, .ColIndex("Accountcom"))
                 RsDev("Account_Code").value = .TextMatrix(i, .ColIndex("Account_Code"))
                 
                 RsDev("CommissionPercentage").value = val(.TextMatrix(i, .ColIndex("CommissionPercentage")))
                 RsDev("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
                 
                 RsDev("Remarks").value = .TextMatrix(i, .ColIndex("Remarks"))
                 
               'RsDev("TransType").value = val(.TextMatrix(i, .ColIndex("PaymentID")))
                      
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
 
    RsDev.Close
   ' RsDev.Open "tblGeneralCashingDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
 
 

    createVoucher
           updatePaymenttransactions 1, val(DcboBox.BoundText), FromDate.value, ToDate.value

    Cn.CommitTrans
    BeginTrans = False
 
    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·”‰œ" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = "Saved" & CHR(13)
                Msg = Msg + "Do you want enter another One"
            End If
   
            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

            '   Retrive
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

            lbl(27).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
            '  Fg_Journal.Enabled = False
    End Select

    Retrive
    TxtModFlg.Text = "R"
    'End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Function CheckAcconts() As Boolean
CheckAcconts = True
 Dim Account_Code_dynamic As String
 Dim Account_Code_dynamic1 As String
 If GetValueAddedAccount(dbRecordDate.value, , , 1, 21) = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                   MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„»Ì⁄« "
            Else
                  MsgBox "Value added account not specified"
            End If
            CheckAcconts = False
End If

If GetValueAddedAccount(dbRecordDate.value, , , 1, 9) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                     MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ·„—œÊœ«  «·„»Ì⁄« "
                Else
                       MsgBox "Value added account not specified"
                End If
CheckAcconts = False
End If




         Account_Code_dynamic = get_account_code_branch(2, val(Me.DcBranch.BoundText))
        
            If Account_Code_dynamic = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "Branch Not Created", vbCritical
                        End If

              CheckAcconts = False
            ElseIf Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Sales Account Not Defined in this Branch", vbCritical
                    End If
CheckAcconts = False
                
         
                End If
 
            
            
         Account_Code_dynamic1 = get_account_code_branch(3, val(Me.DcBranch.BoundText))
        
            If Account_Code_dynamic1 = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "Branch Not Created", vbCritical
                        End If

              CheckAcconts = False
            ElseIf Account_Code_dynamic1 = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» „ «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Sales Account Not DefinœÊœ«  ed in this Branch", vbCritical
                    End If
CheckAcconts = False
                
         
                End If
                
 
End Function

Function createVoucher()
    Dim bankDes As String
    Dim AccountCode As String
    Dim AccountCode1 As String
 
    Dim NoteID As String
    Dim sql As String
 
bankDes = "”‰œ ﬁ»÷ ⁄„Ê„Ì"
    '//////////////////////////////////////Notes////////////////////////////////////
    Dim line_no As Integer
    Dim RsNotes As New ADODB.Recordset
  '  RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (1 = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    If Me.TxtModFlg.Text = "E" Then
                  
        sql = "Delete notes where NoteID=" & val(Me.TxtNoteID.Text)
    End If

    RsNotes.AddNew
    NoteID = CStr(TxtNoteID.Text)
    RsNotes("NoteID").value = CStr(TxtNoteID.Text)
    RsNotes("NoteType").value = 59
    RsNotes("NoteDate").value = dbRecordDate.value
    RsNotes("UserID").value = user_id
    RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.Text) '„”·”· «·ﬁÌœ
    RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.Text) '„”·”· «–‰ «·’—›
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("numbering_type1").value = sand_numbering_type(59) '‰Ê⁄  —ﬁÌ„ ”‰œ «·«Ìœ«⁄
    RsNotes("sanad_year").value = year(dbRecordDate.value)
    RsNotes("sanad_month").value = month(dbRecordDate.value)
    RsNotes("note_value_by_characters").value = WriteNo(Format(val(TxtTotalCash.Text) + val(TxtTotalCheques.Text), "0.00"), 0, True, ".")
    RsNotes("remark").value = txtRemarks.Text & bankDes
    RsNotes("Branch_no").value = val(Me.DcBranch.BoundText)
                
    RsNotes.update
                
    line_no = 0
Dim i As Integer
  
'*********************************ﬁÌœ «À»«  «·„»Ì⁄«  **********************************************
        
  


Dim LngDevID  As Long
Dim debitorcredit As Integer
Dim Tvalue As Double
        With Me.Gridxxx
 
            For i = 1 To .Rows - 1

                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
              If val(.TextMatrix(i, .ColIndex("PaymentID"))) = 0 Then
                   AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
               Else
                    AccountCode = (.TextMatrix(i, .ColIndex("Accountsus")))
               End If
                    bankDes = "«Ìœ«⁄   „‰    " & (.TextMatrix(i, .ColIndex("PaymentName")))
                    line_no = line_no + 1
  
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
Tvalue = val(.TextMatrix(i, .ColIndex("Value")))
        If Tvalue > 0 Then
        debitorcredit = 0
        Else
        debitorcredit = 1
        Tvalue = Abs(Tvalue)
        End If

                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, Tvalue, debitorcredit, .TextMatrix(i, .ColIndex("PaymentName")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , Tvalue, , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    
                    End If
         
                End If

            Next i


Dim salesAccount As String
Dim ReturnsalesAccount As String

Dim VatsalesAccount As String
Dim VatReturnsalesAccount As String
Dim vaTAccount As String
Dim VATValue As Double
Dim X As Boolean
 salesAccount = get_account_code_branch(2, val(Me.DcBranch.BoundText))
 ReturnsalesAccount = get_account_code_branch(3, val(Me.DcBranch.BoundText))
 X = GetValueAddedAccount(dbRecordDate.value, , VatsalesAccount, 1, 21)
 X = GetValueAddedAccount(dbRecordDate.value, VatReturnsalesAccount, , 1, 9)


            For i = 1 To .Rows - 1

                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
               Tvalue = val(.TextMatrix(i, .ColIndex("beforeVat")))
               VATValue = val(.TextMatrix(i, .ColIndex("Vat")))
               If Tvalue > 0 Then ' „»Ì⁄« 
               AccountCode = salesAccount
              vaTAccount = VatsalesAccount
               debitorcredit = 1
                             
               Else
                AccountCode = ReturnsalesAccount
                vaTAccount = VatReturnsalesAccount
               Tvalue = Abs(Tvalue)
               VATValue = Abs(VATValue)
               debitorcredit = 0
               End If
            '        AccountCode = (.TextMatrix(i, .ColIndex("Accountcom")))
                    bankDes = "«Ìœ«⁄   „‰    " & (.TextMatrix(i, .ColIndex("PaymentName")))
                    line_no = line_no + 1
  
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, Tvalue, debitorcredit, .TextMatrix(i, .ColIndex("PaymentName")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , Tvalue, , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    
                    End If
         
                 line_no = line_no + 1
                 
                   If ModAccounts.AddNewDev(LngDevID, line_no, vaTAccount, VATValue, debitorcredit, .TextMatrix(i, .ColIndex("PaymentName")) & bankDes & " ﬁ „÷«›…", val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , Tvalue, , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    
                    End If
                End If

            Next i
            
             
        End With
    
     
     
'*******************************************
  
  With VSFlexGrid1
                   
    If VSFlexGrid1.Rows > 1 Then
If val(.TextMatrix(1, .ColIndex("CollectedValue"))) > 0 And (.TextMatrix(1, .ColIndex("PaymentID"))) = 0 Then
        Dim RsDev  As ADODB.Recordset
        Set RsDev = New ADODB.Recordset
     '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
           StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
 line_no = line_no + 1
bankDes = "«Ìœ«⁄ ‰ﬁœÌ „‰    " & DcboBox.Text
        '«·ÿ—› «·„œÌ‰      «·’‰œÊﬁ «·⁄„Ê„Ì ﬂ«‘
       ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
       AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcGeneralBox.BoundText))
       
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.DcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = val(.TextMatrix(1, .ColIndex("NetValue")))
        RsDev("Credit_Or_Debit").value = 0
                    
        RsDev("RecordDate").value = Me.dbRecordDate.value
        RsDev("Notes_ID").value = val(Me.TxtNoteID.Text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
    
    
 line_no = line_no + 1

        '«·ÿ—› «·œ«∆‰      «·’‰œÊﬁ   «·€—⁄Ì «·ﬂ«‘Ì—
       ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
       AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
       
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.DcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = val(.TextMatrix(1, .ColIndex("NetValue")))
        RsDev("Credit_Or_Debit").value = 1
                    
        RsDev("RecordDate").value = Me.dbRecordDate.value
        RsDev("Notes_ID").value = val(Me.TxtNoteID.Text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
    
    
    End If

End If


End With
 
    'ﬁÌœ «‰Ê«⁄ «·œ›⁄ «·«Œ—Ì
    If VSFlexGrid1.Rows > 2 Then
 
         
         

        With VSFlexGrid1
 
            For i = 2 To .Rows - 1

                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("NetValue"))) > 0 Then
               
                    AccountCode = (.TextMatrix(i, .ColIndex("Account_Code")))
                    bankDes = "«Ìœ«⁄   „‰    " & (.TextMatrix(i, .ColIndex("PaymentName")))
                    line_no = line_no + 1
  
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("NetValue"))), 0, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("NetValue")), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    
                    End If
         
                End If

            Next i



            For i = 2 To .Rows - 1

                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("CommissionValue"))) > 0 Then
               
                    AccountCode = (.TextMatrix(i, .ColIndex("Accountcom")))
                    bankDes = "«Ìœ«⁄   „‰    " & (.TextMatrix(i, .ColIndex("PaymentName")))
                    line_no = line_no + 1
  
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("CommissionValue"))), 0, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("CommissionValue")), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    
                    End If
         
                End If

            Next i
            
            
                     For i = 2 To .Rows - 1

                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("CollectedValue"))) > 0 Then
               
                    AccountCode = (.TextMatrix(i, .ColIndex("Accountsus")))
                    bankDes = "«Ìœ«⁄   „‰    " & (.TextMatrix(i, .ColIndex("PaymentName")))
                    line_no = line_no + 1
  
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("CollectedValue"))), 1, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("CollectedValue")), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    
                    End If
         
                End If

            Next i
            
        End With
    
    End If
  
  
    updateNotesValueAndNobytext (val(Me.TxtNoteID.Text))

ErrTrap:
End Function

Function checkSelectCheque() As Boolean
    checkSelectCheque = False
    Dim i As Integer

    With Me.Grid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("BoxId")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
              
                checkSelectCheque = True
                Exit Function
            End If

        Next i

    End With

End Function

Private Sub Check17_Click()
 

    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Grid1
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = True
            Next i

        End With

    Else

        With Me.Grid1

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = False
            Next i

        End With

    End If

     



        ReLineGrid
End Sub

Private Sub Cmd_Click(Index As Integer)
   
   '  On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me

       
            'dbRecordDate.SetFocus
  
         
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            VSFlexGrid1.Enabled = True
            Me.DcBranch.BoundText = Current_branch
         
        Case 1
        If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 
            TxtModFlg.Text = "E"
            'Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
         
            ' Grid1.Rows = Grid1.Rows + 1
            Grid1.Enabled = True
            CuurentLogdata
Command1_Click

        Case 2
         If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            Dim Msg As String

            If Trim(DcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.DcBranch.BoundText
         
            SaveData
           
        Case 3
            Undo

        Case 4
        If ChekClodePeriod(dbRecordDate.value) = True Then
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

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
'wael
    '        Load FrmCatchSearsh
    
    '    FrmCatchSearsh.Show vbModal

        Case 6
            Unload Me

        Case 7
            If chkIsRecOnly.value = vbChecked Then
                FillGridWithData (val(Me.DCboUserName.BoundText))
            End If
            If chkIsRecOnly.value = vbChecked Or chkIsSalesOnly.value = vbChecked Then
                FillGridWithData2 (val(Me.DCboUserName.BoundText))
            End If

    
     '       If val(TxtValue.text) < 0 Then
    '
    '            MsgBox "—’Ìœ «·Œ“Ì‰… œ«∆‰ ·« Ì„ﬂ‰ «·«Ìœ«⁄ ÊÌ„ﬂ‰ﬂ ﬂ «»… «·„»·€ «·„—«œ «Ìœ«⁄Â ÌœÊÌ«", vbInformation
    '            TxtValue.text = 0
    '            Exit Sub
    '        End If

            addrow1
recalgrid
        Case 8
            RemoveGridRow
    
            '   ViewDataList
        Case 9
            addrow1

        Case 10
            RemoveGridRow1

        Case 11

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.Text, , 200, val(TxtNoteID.Text)
        
        Case 12

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport (TxtlBanksDepositeId.Text)
    End Select

    Exit Sub
ErrTrap:

End Sub

Function PrintReport(ID As Integer)

    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
  '  MySQL = "SELECT     TOP 100 PERCENT dbo.tblGeneralCashingDetails.tblGeneralCashingId, dbo.tblGeneralCashingDetails.box_or_bank, dbo.tblGeneralCashingDetails.[value], "
  '  MySQL = MySQL & " dbo.tblGeneralCashingDetails.ChequeNo, dbo.tblGeneralCashingDetails.Remarks, dbo.tblGeneralCashingDetails.BoxID, dbo.TblBoxesData.BoxName,"
  '  MySQL = MySQL & "  dbo.TblBoxesData.BoxNameE, dbo.tblGeneralCashingDetails.BankName, dbo.tblGeneralCashingDetails.DueDate, dbo.tblGeneralCashing.NoteSerial1,"
  '  MySQL = MySQL & " dbo.tblGeneralCashing.NoteSerial, dbo.tblGeneralCashing.RecordDate, dbo.tblGeneralCashing.bankid, dbo.BanksData.BankName AS DepositeBankName,"
  '  MySQL = MySQL & " dbo.tblGeneralCashing.ID"
  '  MySQL = MySQL & " FROM         dbo.tblGeneralCashingDetails INNER JOIN"
  '  MySQL = MySQL & " dbo.TblBoxesData ON dbo.tblGeneralCashingDetails.BoxID = dbo.TblBoxesData.BoxID INNER JOIN"
  '  MySQL = MySQL & " dbo.tblGeneralCashing ON dbo.tblGeneralCashingDetails.tblGeneralCashingId = dbo.tblGeneralCashing.id LEFT OUTER JOIN"
  '  MySQL = MySQL & " dbo.BanksData ON dbo.tblGeneralCashing.bankid = dbo.BanksData.BankID"
  '  MySQL = MySQL & "  Where (1 = 1) and dbo.tblGeneralCashing.ID=" & id
  '  MySQL = MySQL & "  ORDER BY dbo.tblGeneralCashing.NoteSerial1"
MySQL = " SELECT     TOP 100 PERCENT dbo.tblGeneralCashing.id, dbo.tblGeneralCashing.RecordDate, dbo.tblGeneralCashing.NoteSerial1, dbo.tblGeneralCashing.NoteSerial,"
 MySQL = MySQL & "                      dbo.tblGeneralCashing.OldNoteSerial1, dbo.tblGeneralCashing.Remarks, dbo.tblGeneralCashing.ToDate, dbo.tblGeneralCashing.FromDate,"
 MySQL = MySQL & "                      dbo.tblGeneralCashing.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.tblGeneralCashing.GeneralBoxId,"
  MySQL = MySQL & "                     dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE, dbo.tblGeneralCashing.SubBoxId, TblBoxesData_1.BoxName AS SubDcboBox,"
 MySQL = MySQL & "                      TblBoxesData_1.BoxNameE AS SubDcboBoxE, dbo.tblGeneralCashing.NoteID, dbo.tblGeneralCashingDetails.[value], dbo.tblGeneralCashingDetails.TransType,"
 MySQL = MySQL & "                      dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.tblGeneralCashingDetails.CollectedValue,"
  MySQL = MySQL & "                     dbo.tblGeneralCashingDetails.CommissionValue, dbo.tblGeneralCashingDetails.Different, dbo.tblGeneralCashingDetails.Remarks AS RemarksDet,"
 MySQL = MySQL & "                      dbo.tblGeneralCashingDetails.NoteID AS NoteIDDet, dbo.tblGeneralCashingDetails.Accountsus, dbo.tblGeneralCashingDetails.Accountcom,"
 MySQL = MySQL & "                      dbo.tblGeneralCashingDetails.Account_Code , dbo.tblGeneralCashingDetails.CommissionPercentage, dbo.tblGeneralCashingDetails.netvalue"
 MySQL = MySQL & " FROM         dbo.TblPaymentType RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.tblGeneralCashingDetails ON dbo.TblPaymentType.PaymentID = dbo.tblGeneralCashingDetails.TransType RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.tblGeneralCashing LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBoxesData TblBoxesData_1 ON dbo.tblGeneralCashing.SubBoxId = TblBoxesData_1.BoxID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBoxesData ON dbo.tblGeneralCashing.GeneralBoxId = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.tblGeneralCashing.branch_no = dbo.TblBranchesData.branch_id ON"
  MySQL = MySQL & "                     dbo.tblGeneralCashingDetails.tblGeneralCashingId = dbo.tblGeneralCashing.id"
 MySQL = MySQL & " Where (dbo.tblGeneralCashing.id = " & ID & ")"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "RepCatchSupport.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "RepCatchSupport.rpt"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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

End Function

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim i As Integer
    On Error GoTo ErrTrap

    'check Cheque Not Payed

    With Me.Grid1

        For i = 1 To .Rows - 1
                 
            If .TextMatrix(i, .ColIndex("NoteID")) <> "" Then
                If ChequeBoxCollect(val(.TextMatrix(i, .ColIndex("NoteID")))) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«   Õ’Ì·  "
                    Msg = Msg & CHR(13) & " ··‘Ìﬂ —ﬁ„ " & .TextMatrix(i, .ColIndex("ChequeNo"))
                    Msg = Msg & CHR(13) & "»ﬁÌ„… " & .TextMatrix(i, .ColIndex("Value"))
                    Msg = Msg & CHR(13) & " ⁄·Ï »‰ﬂ " & .TextMatrix(i, .ColIndex("BankName"))
                                          
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                          
                    Exit Sub
                End If
                                        
            End If
                                  
        Next i

    End With
 
    If TxtlBanksDepositeId.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial1.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
      
            StrSQL = "Delete From notes Where NoteID=" & val(TxtNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

     
 
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
       updatePaymenttransactions -1, val(DcboBox.BoundText), FromDate.value, ToDate.value
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                
     
               
                      VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            
            
                    TxtModFlg_Change
           
                Else
                    Retrive
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
    
        With Me.Grid

            If Me.TxtModFlg <> "E" Then Exit Sub
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
            LogTextA = "  Õ–› «·Œ“Ì‰…   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("BoxName")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
            LogTexte = "  Delete  Box   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("BoxName")) & " With Value " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                                                         
            AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtNoteSerial, TxtNoteSerial1
        End With
  
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow1()

    With Me.Grid1

        If .Row <= 0 Then Exit Sub
    
        Cn.Execute "update  TblChecqueBoxContent set Deposited=0 where NoteID=" & val(.TextMatrix(.Row, .ColIndex("NoteID")))
                                                        
        .RemoveItem .Row

    End With

    ReLineGrid
End Sub

Function addrow()

    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
    On Error Resume Next

    If Trim(Me.DcGeneralBox.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·’‰œÊﬁ «·⁄«„..!!"
        Else
            Msg = "Specify General Box.!!"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcGeneralBox.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
 
 
    If Trim(Me.DcboBox.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·’‰œÊﬁ «·›—⁄Ì..!!"
        Else
            Msg = "Specify  Sub Box.!!"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboBox.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
    
 
    Me.Grid.Rows = Me.Grid.Rows + 1
    LngRow = Me.Grid.Rows - 1
 
    With Me.Grid
  
        .TextMatrix(LngRow, .ColIndex("BoxId")) = val(DcboBox.BoundText)
    
        .TextMatrix(LngRow, .ColIndex("BoxName")) = DcboBox.Text
    
        .TextMatrix(LngRow, .ColIndex("Value")) = val(TxtValue.Text)
    
        .TextMatrix(LngRow, .ColIndex("Remarks")) = ""
     
        If Me.TxtModFlg = "E" Then
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                           
            LogTextA = "  «Ìœ«⁄ „‰ «·Œ“Ì‰…  " & DcboBox & " »ﬁÌ„… " & val(TxtValue.Text)
            LogTexte = "Deposite From Box  " & DcboBox & " With Value " & val(TxtValue.Text)
                    
            AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtNoteSerial, TxtNoteSerial1
        End If
                                                     
        .AutoSize 0, .Cols - 1, False
    End With
 
    Me.TxtValue.Text = ""
    DcboBox.BoundText = ""
    ReLineGrid

End Function

Function addrow1()

    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
    On Error Resume Next

    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
        If Trim(Me.DcGeneralBox.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·’‰œÊﬁ «·⁄«„..!!"
        Else
            Msg = "Specify General Box.!!"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcGeneralBox.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
 
 
    If Trim(Me.DcboBox.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·’‰œÊﬁ «·›—⁄Ì..!!"
        Else
            Msg = "Specify  Sub Box.!!"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboBox.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
    


    StrSQL = "SELECT     dbo.TblPaymentType.PaymentID, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.TblPaymentType.commision, "
StrSQL = StrSQL & "   dbo.TblPaymentType.Accountsus , dbo.TblPaymentType.Accountcom, dbo.BanksData.Account_Code"
StrSQL = StrSQL & " FROM         dbo.TblPaymentType LEFT OUTER JOIN"
StrSQL = StrSQL & " dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
 
 
 
 '  If chkDue.value = vbChecked Then
 '          StrSQL = StrSQL + " and (DueDate <=" & SQLDate(dbRecordDate.value, True) & ")"
 ' End If
 '
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2


        With Me.VSFlexGrid1
 LngRow = 1
            .TextMatrix(LngRow, .ColIndex("PaymentID")) = 0
     If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(LngRow, .ColIndex("PaymentName")) = "‰ﬁœÌ"
      Else
      .TextMatrix(LngRow, .ColIndex("PaymentName")) = "Cash"
      End If
      
      
          Dim AccountCode As String
    Dim Balance As Double
    Dim balancetype As Integer
    Dim FirstPeriodDateInthisYear  As Date

 
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
   
    Balance = GetActualAccountBalance(AccountCode, , FirstPeriodDateInthisYear, dbRecordDate.value)
  
    
            .TextMatrix(LngRow, .ColIndex("Balance")) = Balance 'IIf(IsNull(rs("Balance").value), "", rs("Balance").value)
    
            .TextMatrix(LngRow, .ColIndex("CommissionPercentage")) = 0
         .TextMatrix(LngRow, .ColIndex("CollectedValue")) = GetPointSAles1(val(DcboBox.BoundText), val(.TextMatrix(LngRow, .ColIndex("PaymentID"))), FromDate.value, ToDate.value, val(DCboUserName.BoundText))
    '.TextMatrix(LngRow, .ColIndex("CollectedValue")) = .TextMatrix(LngRow, .ColIndex("Balance"))
 
        End With

'Exit Function

    For i = 1 To rs.RecordCount
        Me.VSFlexGrid1.Rows = Me.VSFlexGrid1.Rows + 1
        LngRow = Me.VSFlexGrid1.Rows - 1
   
 
 
 
        With Me.VSFlexGrid1
 
            .TextMatrix(LngRow, .ColIndex("PaymentID")) = IIf(IsNull(rs("PaymentID").value), 0, rs("PaymentID").value)
     If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(LngRow, .ColIndex("PaymentName")) = IIf(IsNull(rs("PaymentName").value), "", rs("PaymentName").value)
      Else
      .TextMatrix(LngRow, .ColIndex("PaymentName")) = IIf(IsNull(rs("PaymentNamee").value), "", rs("PaymentNamee").value)
      End If
            .TextMatrix(LngRow, .ColIndex("Balance")) = GetPointSAles1(val(DcboBox.BoundText), val(.TextMatrix(LngRow, .ColIndex("PaymentID"))), FromDate.value, ToDate.value, val(DCboUserName.BoundText))
    .TextMatrix(LngRow, .ColIndex("CollectedValue")) = .TextMatrix(LngRow, .ColIndex("Balance"))
            .TextMatrix(LngRow, .ColIndex("CommissionPercentage")) = IIf(IsNull(rs("commision").value), 0, rs("commision").value)
    
            .TextMatrix(LngRow, .ColIndex("Account_Code")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            .TextMatrix(LngRow, .ColIndex("Accountcom")) = IIf(IsNull(rs("Accountcom").value), "", rs("Accountcom").value)
            .TextMatrix(LngRow, .ColIndex("Accountsus")) = IIf(IsNull(rs("Accountsus").value), "", rs("Accountsus").value)
            
            .TextMatrix(LngRow, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
  
            .AutoSize 0, .Cols - 1, False
        End With

        rs.MoveNext
    Next i
 
    'Me.TxtValue.text = ""
    'txtchequeno.text = ""
    'Dcbank1.BoundText = ""
    'TxtValue1.text = ""

    ReLineGrid

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
                     
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
          
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Dcdep_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcedara_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub CmdAttach_Click()
     On Error Resume Next
           If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtNoteSerial1, "0712201403"

End Sub

Private Sub Command1_Click()
FillGridWithData (val(Me.DCboUserName.BoundText))
End Sub

Private Sub dbRecordDate_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
End Sub

Private Sub dcbank_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
End Sub

Function SHow_grig_col()
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If Rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If Rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If Rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If Rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If Rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If Rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If Rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If Rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If Rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If Rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If Rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If Rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If Rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If Rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If Rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If Rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If Rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If Rs2("s18").value = True Then
            .ColHidden(.ColIndex("sgn")) = False
        Else
            .ColHidden(.ColIndex("sgn")) = True
        End If
     
    End With

End Function

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

Private Sub DcboBox_Change()
    Dim AccountCode As String
    Dim Balance As Double
    Dim balancetype As Integer
    Dim FirstPeriodDateInthisYear  As Date

    If val(DcboBox.BoundText) = 0 Then TxtValue.Text = 0: Exit Sub

    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
    'get_balanceFromGlNew Accountcode, , , , FirstPeriodDateInthisYear, Date, , , Balance, Val(Me.DcBranch.BoundText)

    Balance = GetActualAccountBalance(AccountCode, val(Me.DcBranch.BoundText), FirstPeriodDateInthisYear, dbRecordDate.value)
    'getBalanceWithOpeningBalance Accountcode, Val(dcBranch.BoundText), Date, balance, balanceType

    TxtValue.Text = Balance
    
    
   
            
            
End Sub

Private Sub DcboBox_Click(Area As Integer)
 
    Dim AccountCode As String
AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
   
    lblAccountBalance.Caption = GetbalanceBar(AccountCode)
            
End Sub

Public Sub FillGridWithData(Optional Emp_id As Integer)
Dim Total As Double
    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
 
My_SQL = " SELECT         SUM(dbo.TblTransactionPayments.value * dbo.TblTransactionPayments.Effect) AS TotalValue, dbo.TblTransactionPayments.PaymentID, dbo.TblTransactionPayments.Effect, "
My_SQL = My_SQL & "                         dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee,"
My_SQL = My_SQL & "                         dbo.TblPaymentType.branch_no , dbo.TblPaymentType.maxvalue, dbo.TblPaymentType.TypTran"
My_SQL = My_SQL & "   FROM            dbo.TblTransactionPayments LEFT OUTER JOIN"
My_SQL = My_SQL & "                         dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
My_SQL = My_SQL & "  WHERE        (dbo.TblTransactionPayments.Transaction_ID IN"
My_SQL = My_SQL & "                             (SELECT        Transaction_ID"
My_SQL = My_SQL & "                                From dbo.transactions"
My_SQL = My_SQL & "                                WHERE        (POSBillType = 1) "
'My_SQL = My_SQL & " AND (Transaction_Date = CONVERT(DATETIME, '2019-02-14 00:00:00', 102)) "

 My_SQL = My_SQL & "  AND  (Transaction_Date >='" & SQLDate(FromDate) & "'"
 My_SQL = My_SQL & "  AND   Transaction_Date <='" & SQLDate(ToDate) & "')"
 


My_SQL = My_SQL & "  AND (Emp_ID = " & Emp_id & ") "
My_SQL = My_SQL & "  AND (Transaction_Type = 21 OR"
My_SQL = My_SQL & "                                                         Transaction_Type = 9)))"
My_SQL = My_SQL & "  GROUP BY dbo.TblTransactionPayments.PaymentID, dbo.TblTransactionPayments.Effect, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom,"
My_SQL = My_SQL & "                         dbo.TblPaymentType.commision , dbo.TblPaymentType.PaymentNamee, dbo.TblPaymentType.branch_no, dbo.TblPaymentType.maxvalue, dbo.TblPaymentType.TypTran"
My_SQL = My_SQL & " ORDER BY dbo.TblTransactionPayments.PaymentID"


 

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
 Total = 0
    With Me.Gridxxx
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "", rs.Fields("PaymentName").value)
               .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(rs.Fields("TotalValue").value), 0, rs.Fields("TotalValue").value)
               
               Dim beforeVat As Double
               Dim Vat As Double
               beforeVat = val(.TextMatrix(i, .ColIndex("value"))) / 1.05
               Vat = beforeVat * 0.05
                
                .TextMatrix(i, .ColIndex("beforeVat")) = Round(beforeVat, 2)
                .TextMatrix(i, .ColIndex("Vat")) = Round(Vat, 2)
                
               Total = Total + val(.TextMatrix(i, .ColIndex("value")))
              .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), IIf(IsNull(rs.Fields("TotalValue").value), 0, rs.Fields("TotalValue").value), rs.Fields("PaymentID").value)
              .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
              
              
               If SystemOptions.UserInterface = ArabicInterface Then
                     If .TextMatrix(i, .ColIndex("PaymentName")) = "" Then
                                If val(.TextMatrix(i, .ColIndex("value"))) >= 0 Then
                                       .TextMatrix(i, .ColIndex("PaymentName")) = "‰ﬁœÌ „»Ì⁄« "
                                ElseIf val(.TextMatrix(i, .ColIndex("value"))) < 0 Then
                                       .TextMatrix(i, .ColIndex("PaymentName")) = "‰ﬁœÌ „—œÊœ«  "
                                End If
                     
                     End If
                     
               Else
                
                               If val(.TextMatrix(i, .ColIndex("value"))) >= 0 Then
                                       .TextMatrix(i, .ColIndex("PaymentName")) = "Cash Sales "
                                ElseIf val(.TextMatrix(i, .ColIndex("value"))) < 0 Then
                                       .TextMatrix(i, .ColIndex("PaymentName")) = "Cash Return "
                                End If
                                
               End If
               
                
                 rs.MoveNext
            Next
   
            rs.Close
        End If
lblTotalTransaction.Caption = Total
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub



Public Sub FillGridWithData2(Optional Emp_id As Integer)
Dim Total As Double
    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim sql As String

    Set rs = New ADODB.Recordset
 
 
Dim mTime As Date
Dim mTime2 As Date
mTime = Time
Dim MinDate As Date


If SystemOptions.CostStarting = True Then
     Dim FirstPeriodDateInthisYear  As Date
     getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                               
    MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
Else
    MinDate = "1-1-1900"
End If

  grdMaster.Clear flexClearScrollable, flexClearEverything
  grdMaster.Rows = 1
  'grdDet.Clear flexClearScrollable, flexClearEverything
  'grdDet.Rows = 1
Dim Rs2 As ADODB.Recordset
Dim RsDet  As New ADODB.Recordset
Set Rs2 = New ADODB.Recordset
 
If chkIsSalesOnly.value = vbUnchecked And chkIsRetOnly.value = vbUnchecked Then Exit Sub

sql = " SELECT     DISTINCT dbo.Transactions.Transaction_Date,Transactions.FixesAssetsID,Transactions.Emp_ID,Transactions.DepartementID, dbo.Transactions.CusID"
sql = sql & " ,dbo.Transactions.Doctype, dbo.Transactions.Transaction_ID,Transactions.NoteID, Transactions.NoteSerial,Transactions.BranchId,  dbo.Transactions.NoteSerial1, dbo.Transactions.StoreID,"
sql = sql & "                       dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.Transaction_Type"
',Notes.Note_Value,"
sql = sql & "           ,tc.CusName,tc.CusNamee"
sql = sql & " ,Note_Value = (Select Sum(IsNull(ShowQty,0) *  IsNull(showPrice,0) ) From Transaction_Details DD Where DD.Transaction_ID = dbo.Transactions.Transaction_ID )"
'sql = sql & " ,Note_Value2 = (Select Sum(IsNull(ShowQty,0) *  IsNull(OldshowPrice,0) ) From Transaction_Details DD Where DD.Transaction_ID = dbo.Transactions.Transaction_ID )"
sql = sql & " ,T2.StoreID StoreID2,ts.StoreName StoreName2 "

sql = sql & "  FROM            Transactions LEFT OUTER JOIN"
sql = sql & "                           TblStore ON Transactions.StoreID = TblStore.StoreID LEFT OUTER JOIN"
sql = sql & "                           Notes ON Notes.NoteID = Transactions.NoteId LEFT OUTER JOIN"
sql = sql & "                           TblCustemers AS tc ON Transactions.CusID = tc.CusID LEFT OUTER JOIN"
sql = sql & "                           Transactions AS T2 ON Transactions.Transaction_ID = T2.ReturnID LEFT OUTER JOIN                          TblStore AS ts ON ts.StoreID = T2.StoreID"

'If DcboItemID1.Text <> "" Then
sql = sql & "               Left Outer join            Transaction_Details  ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID "



sql = sql & "    WHERE       "

sql = sql & "               Transactions.POSBillType = 1 "
'My_SQL = My_SQL & " AND (Transaction_Date = CONVERT(DATETIME, '2019-02-14 00:00:00', 102)) "

 sql = sql & "  AND  (Transactions.Transaction_Date >='" & SQLDate(FromDate) & "'"
 sql = sql & "   AND   Transactions.Transaction_Date <='" & SQLDate(ToDate) & "')"
 


sql = sql & "   AND (Transactions.Emp_ID = " & Emp_id & ") "
sql = sql & "   AND "
If chkIsSalesOnly.value = vbChecked And chkIsRetOnly.value = vbChecked Then
    sql = sql & " (Transactions.Transaction_Type = 21 OR Transactions.Transaction_Type = 9 )"
ElseIf chkIsSalesOnly.value = vbChecked And chkIsRetOnly.value = vbUnchecked Then
    sql = sql & " (Transactions.Transaction_Type = 21 )"
ElseIf chkIsSalesOnly.value = vbUnchecked And chkIsRetOnly.value = vbChecked Then
    sql = sql & " (Transactions.Transaction_Type = 9 )"
End If

sql = sql & "   Order By transactions.Transaction_Date,transactions.Transaction_ID,transactions.NoteSerial1"

 

Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
' Set cProgress = New ClsProgress
'        BolFrmLoaded = True
'        cProgress.ProgressType = Waiting
'        cProgress.StartProgress
With grdMaster
If Rs2.RecordCount > 0 Then
    .Rows = .Rows + Rs2.RecordCount
    Rs2.MoveFirst

    For i = 1 To .Rows - 1
        If .Rows <= i Then Exit Sub
        .TextMatrix(i, .ColIndex(("Ser"))) = i
        '.TextMatrix(I, .ColIndex(("IDRef"))) = IIf(IsNull(Rs2("ID").value), "", Rs2("ID").value)
        .TextMatrix(i, .ColIndex(("Transaction_ID"))) = IIf(IsNull(Rs2("Transaction_ID").value), "", Rs2("Transaction_ID").value)
        .TextMatrix(i, .ColIndex(("Transaction_Type"))) = IIf(IsNull(Rs2("Transaction_Type").value), "", Rs2("Transaction_Type").value)
        .TextMatrix(i, .ColIndex(("CusID"))) = IIf(IsNull(Rs2("CusID").value), "", Rs2("CusID").value)
        
        .TextMatrix(i, .ColIndex(("NoteSerial"))) = IIf(IsNull(Rs2("NoteSerial").value), "", Rs2("NoteSerial").value)
        .TextMatrix(i, .ColIndex(("NoteID"))) = IIf(IsNull(Rs2("NoteID").value), "", Rs2("NoteID").value)
        .TextMatrix(i, .ColIndex(("BranchId"))) = IIf(IsNull(Rs2("BranchId").value), "", Rs2("BranchId").value)
        
    
        
        If Rs2("Transaction_Type").value = 19 Then
            .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«–‰ ’—›"
        ElseIf Rs2("Transaction_Type").value = 992 Or Rs2("Transaction_Type").value = 10 Then
                .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = " ÕÊÌ·«  »Ì‰ «·„Œ«“‰"
        ElseIf Rs2("Transaction_Type").value = 992 Or Rs2("Transaction_Type").value = 11 Then
                                .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«” ·«„ „‰ „Œ“‰"
        ElseIf Rs2("Transaction_Type").value = 21 Then
            .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "„»Ì⁄« "
        ElseIf Rs2("Transaction_Type").value = 9 Then
            .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "„—œÊœ« "
        End If
    
        .TextMatrix(i, .ColIndex(("Doctype"))) = IIf(IsNull(Rs2("docType").value), "", Rs2("docType").value)
        .TextMatrix(i, .ColIndex(("NoteSerial1"))) = IIf(IsNull(Rs2("NoteSerial1").value), "", Rs2("NoteSerial1").value)
        .TextMatrix(i, .ColIndex(("Transaction_Date"))) = IIf(IsNull(Rs2("Transaction_Date").value), "", Rs2("Transaction_Date").value)
        .TextMatrix(i, .ColIndex(("StoreID"))) = IIf(IsNull(Rs2("StoreID").value), "", Rs2("StoreID").value)
        .TextMatrix(i, .ColIndex(("Note_Value"))) = IIf(IsNull(Rs2("Note_Value").value), "", Rs2("Note_Value").value)
        
        .TextMatrix(i, .ColIndex(("CusID"))) = IIf(IsNull(Rs2("CusID").value), "", Rs2("CusID").value)
        .TextMatrix(i, .ColIndex(("FixesAssetsID"))) = IIf(IsNull(Rs2("FixesAssetsID").value), "", Rs2("FixesAssetsID").value)
        .TextMatrix(i, .ColIndex(("Emp_ID"))) = IIf(IsNull(Rs2("Emp_ID").value), "", Rs2("Emp_ID").value)
        .TextMatrix(i, .ColIndex(("DepartementID"))) = IIf(IsNull(Rs2("DepartementID").value), "", Rs2("DepartementID").value)
        
        .TextMatrix(i, .ColIndex(("CusID"))) = IIf(IsNull(Rs2("CusID").value), "", Rs2("CusID").value)
        .TextMatrix(i, .ColIndex(("StoreID2"))) = IIf(IsNull(Rs2("StoreID2").value), "", Rs2("StoreID2").value)
        .TextMatrix(i, .ColIndex(("StoreName2"))) = IIf(IsNull(Rs2("StoreName2").value), "", Rs2("StoreName2").value)
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex(("StoreName"))) = IIf(IsNull(Rs2("StoreName").value), "", Rs2("StoreName").value)
            .TextMatrix(i, .ColIndex(("CusName"))) = IIf(IsNull(Rs2("CusName").value), "", Rs2("CusName").value)
        Else
            .TextMatrix(i, .ColIndex(("StoreName"))) = IIf(IsNull(Rs2("StoreNamee").value), "", Rs2("StoreNamee").value)
            .TextMatrix(i, .ColIndex(("CusName"))) = IIf(IsNull(Rs2("CusNamee").value), "", Rs2("CusNamee").value)
        End If
        
    
    
        Rs2.MoveNext
        DoEvents
    Next i
End If
End With

MsgBox " „ «·«œ—«Ã"





ErrTrap:
End Sub



Private Sub DCboUserName_Change()
 Dim PettyId As Long
 Dim BoxID As Long
     
    getCashireData val(DCboUserName.BoundText), 0, 0, 0, PettyId, 0, BoxID, val(Me.DCboUserName.BoundText)
DcboBox.BoundText = BoxID
'Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", PettyId)
'cashierBocaccount = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", BoxID)
             

End Sub

Private Sub DCboUserName_Click(Area As Integer)
DCboUserName_Change
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcChequeBox_Click(Area As Integer)
'addrow1
End Sub

Private Sub DcGeneralBox_Click(Area As Integer)
Dim AccountCode As String
AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcGeneralBox.BoundText))
   
    lblAccountBalance.Caption = GetbalanceBar(AccountCode)
            
End Sub

Private Sub dcproject_Click(Area As Integer)

    If dcproject.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(dcproject.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub Dcterm_Click(Area As Integer)

    If Dcterm.BoundText = "" Then Exit Sub

    My_SQL = " select  fullcode,name from terms_operations where term_fullcode='" & Dcterm.BoundText & "'"
    fill_combo dcopr, My_SQL
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ScreenNameArabic = "”‰œ ﬁ»÷ «·’‰œÊﬁ «·⁄«„ "
    ScreenNameEnglish = "General Cashing Voucher"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 20
 
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(12).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With

    With Me.Grid1
        Set .WallPaper = GrdBack.Picture
     
    End With

    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL
If SystemOptions.UserInterface = ArabicInterface Then
  My_SQL = "SELECT        dbo.cachierData.EmpID,  dbo.TblEmployee.Emp_Name "
My_SQL = My_SQL & " FROM            dbo.cachierData INNER JOIN"
My_SQL = My_SQL & "                          dbo.TblEmployee ON dbo.cachierData.EmpID = dbo.TblEmployee.Emp_ID"
My_SQL = My_SQL & "  Where (dbo.cachierData.Ctype = 0)"
  
  Else
  My_SQL = "SELECT        dbo.cachierData.EmpID,  dbo.TblEmployee.Emp_Namee "
My_SQL = My_SQL & " FROM            dbo.cachierData INNER JOIN"
My_SQL = My_SQL & "                          dbo.TblEmployee ON dbo.cachierData.EmpID = dbo.TblEmployee.Emp_ID"
My_SQL = My_SQL & "  Where (dbo.cachierData.Ctype = 0)"

  End If
  

    fill_combo DCboUserName, My_SQL




    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic

    Dcombos.GetBanks Me.Dcbank
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBoxes Me.DcGeneralBox
    Dcombos.GetBranches Me.DcBranch
    
    Dcombos.GetBanks Me.Dcbank1
    Dcombos.GetChequeBox Me.DCChequeBox

    

    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From tblGeneralCashing   WHERE 1=1"
    StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
            If SystemOptions.usertype <> UserAdminAll Then
        'StrSQL = StrSQL & " where   branch_no=" & Current_branch
    End If
    StrSQL = StrSQL & " order by noteserial1"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
Cmd(12).Caption = "Print"
Check17.Caption = "Select All"
CmdAttach.Caption = "Attachments"
lbl(22).Caption = "From Date"
lbl(23).Caption = "To Date"
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"
chkDue.Caption = "Show Due only"
    Cmd(11).Caption = "JE Print"
    Label4.Caption = "Total Cash"
    Label6.Caption = "Total Cheque"
    lbl(19).Caption = "JE NO"
    lbl(21).Caption = "Cheques Sel."
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
lbl(24).Caption = "Reg.No"
    Me.Caption = "   CASH RECEIPTS"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(5).Caption = " Date"
    'Ele(3).Caption = "Select Interval"
    'lbl(2).Caption = "Year"
    lbl(17).Caption = "Branch"

    lbl(15).Caption = "Main Box"
    lbl(3).Caption = "Remarks"
    lbl(12).Caption = "Cash Deposite"
    lbl(14).Caption = "Sub Box "
    Label1.Caption = "Value"
    Cmd(7).Caption = "Add"
    Cmd(8).Caption = "Remove"

    lbl(13).Caption = "Cheques "
    lbl(18).Caption = "Cheques  Box"
    lbl(16).Caption = " From Bank"
    Label3.Caption = "Chq. NO"
    Label2.Caption = "Value"
    Cmd(9).Caption = "Add"
    Cmd(10).Caption = "Remove"
   With VSFlexGrid1
   .TextMatrix(0, .ColIndex("Ser")) = "Serial"
   .TextMatrix(0, .ColIndex("PaymentName")) = "Payment"
   .TextMatrix(0, .ColIndex("balance")) = "Balance"
   .TextMatrix(0, .ColIndex("CollectedValue")) = "CollectedValue"
   .TextMatrix(0, .ColIndex("CommissionPercentage")) = "CommissionPercentage"
   .TextMatrix(0, .ColIndex("CommissionValue")) = "CommissionValue"
   .TextMatrix(0, .ColIndex("different")) = "Different"
    .TextMatrix(0, .ColIndex("NetValue")) = "NetValue"
   .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
   End With
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("BoxName")) = "BoxId"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With

    With Me.Grid1
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("Select")) = "Select"

        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        .TextMatrix(0, .ColIndex("ChequeNO")) = "Cheque NO"

        .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"

    End With

    lbl(20).Caption = "Curr Rec."
    lbl(37).Caption = "Total Rec."
End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .Rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                       
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

  
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

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
     
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 20
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
 
        ReLineGrid
    End With

End Sub

Private Sub ReLineGrid()
    On Error Resume Next
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("BoxId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i

        Me.TxtTotalCash.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    
    End With
                 
    IntCounter = 0

    With Me.Grid1

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("BoxId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i

        Me.TxtTotalCheques.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With

    CalCulateParts
End Sub

Private Sub CalCulateParts()
    Dim i As Integer
    Dim IntCount As Integer

    Dim SngTotal As Single

    With Me.Grid1
        SngTotal = 0
        IntCount = 0

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                IntCount = IntCount + 1
                SngTotal = SngTotal + val(.TextMatrix(i, .ColIndex("Value")))
            End If

        Next i

    End With

    Me.TxtPaymentCounts.Caption = IntCount
    Me.TxtTotalCheques.Text = SngTotal
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, Optional NoteID As Long = 0)
    'Exit Sub
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
          
    Grid1.Clear flexClearScrollable, flexClearEverything
    Grid1.Rows = 1
          
    Gridxxx.Clear flexClearScrollable, flexClearEverything
    Gridxxx.Rows = 1
          
               
    TxtTotalCash.Text = 0
    TxtTotalCheques.Text = 0
DCChequeBox.Text = ""
DcboBox.Text = ""

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
 
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If

    If Lngid <> 0 Then
        rs.Find "id=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    If NoteID <> 0 Then
        rs.Find "NoteID=" & NoteID, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    
    Me.TxtNoteID.Text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(27).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    DcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
 DcGeneralBox.BoundText = IIf(IsNull(rs("GeneralBoxId").value), "", rs("GeneralBoxId").value)
 DcboBox.BoundText = IIf(IsNull(rs("SubBoxId").value), "", rs("SubBoxId").value)
 DCboUserName.BoundText = IIf(IsNull(rs("CashierId").value), "", rs("CashierId").value)
 
 
    Me.TxtlBanksDepositeId.Text = IIf(IsNull(rs("id").value), "", rs("id").value)
 
 FromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
 ToDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
 
    dbRecordDate.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)

'    Dcbank.BoundText = IIf(IsNull(rs("bankid").value), "", rs("bankid").value)

    txtRemarks.Text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
 

  
    
    'StrSQL = " SELECT   * FROM         dbo.TbllBanksDepositeDetails  "
    'StrSQL = StrSQL & "  where box_or_bank=0 and  TbllBanksDepositeId=" & Val(Me.TxtlBanksDepositeId.text)
  
    StrSQL = " SELECT     dbo.tblGeneralCashing.id, dbo.tblGeneralCashingDetails.TransType, dbo.tblGeneralCashingDetails.[value], dbo.tblGeneralCashingDetails.CollectedValue, "
    StrSQL = StrSQL & "     dbo.tblGeneralCashingDetails.CommissionValue, dbo.tblGeneralCashingDetails.Different, dbo.tblGeneralCashingDetails.Remarks,"
    StrSQL = StrSQL & "   dbo.tblGeneralCashingDetails.NoteID, dbo.tblGeneralCashingDetails.Accountsus, dbo.tblGeneralCashingDetails.Accountcom,"
    StrSQL = StrSQL & "    dbo.tblGeneralCashingDetails.Account_Code , dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee"
    StrSQL = StrSQL & " ,     dbo.tblGeneralCashingDetails.CommissionPercentage, dbo.tblGeneralCashingDetails.NetValue FROM         dbo.tblGeneralCashing INNER JOIN"
    StrSQL = StrSQL & "   dbo.tblGeneralCashingDetails ON dbo.tblGeneralCashing.id = dbo.tblGeneralCashingDetails.tblGeneralCashingId LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblPaymentType ON dbo.tblGeneralCashingDetails.TransType = dbo.TblPaymentType.PaymentID"
 
    
    
    StrSQL = StrSQL & "  Where (dbo.tblGeneralCashing.id = " & val(Me.TxtlBanksDepositeId.Text) & ")"

    
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
    
            .Rows = .FixedRows + RsDev.RecordCount



            For i = .FixedRows To .Rows - 1
  
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(RsDev("TransType").value), 0, val(RsDev("TransType").value))
            If .TextMatrix(i, .ColIndex("PaymentID")) = 0 Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    .TextMatrix(i, .ColIndex("PaymentName")) = "‰ﬁœÌ"
                                Else
                                    .TextMatrix(i, .ColIndex("PaymentName")) = "Cash"
                                End If
            Else
                                  If SystemOptions.UserInterface = ArabicInterface Then
                                    .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(RsDev("PaymentName").value), "", RsDev("PaymentName").value)
                                Else
                                    .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(RsDev("PaymentNamee").value), "", RsDev("PaymentNamee").value)
                                End If
            End If
 
                .TextMatrix(i, .ColIndex("Balance")) = IIf(IsNull(RsDev("value").value), 0, val(RsDev("value").value))
            
            .TextMatrix(i, .ColIndex("CollectedValue")) = IIf(IsNull(RsDev("CollectedValue").value), 0, val(RsDev("CollectedValue").value))
             .TextMatrix(i, .ColIndex("CommissionPercentage")) = IIf(IsNull(RsDev("CommissionPercentage").value), 0, val(RsDev("CommissionPercentage").value))
            .TextMatrix(i, .ColIndex("CommissionValue")) = IIf(IsNull(RsDev("CommissionValue").value), 0, val(RsDev("CommissionValue").value))
            .TextMatrix(i, .ColIndex("different")) = IIf(IsNull(RsDev("different").value), 0, val(RsDev("different").value))
           .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDev("NetValue").value), 0, val(RsDev("NetValue").value))
          
              
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDev("Remarks").value), "", RsDev("Remarks").value)
                         


                            .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                                            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(RsDev("Accountcom").value), "", RsDev("Accountcom").value)
                           .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(RsDev("Accountsus").value), "", RsDev("Accountsus").value)
                                                            
                RsDev.MoveNext
            Next i
 
        End With

    End If
  

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    If Col <> Grid.ColIndex("Remarks") Then
        Cancel = True
    End If

End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid1

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select

        ReLineGrid
    
        If Me.TxtModFlg = "E" Then

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
            If .Cell(flexcpChecked, Row, .ColIndex("Select")) = flexChecked Then
                LogTextA = "   ÕœÌœ «·‘Ìﬂ —ﬁ„   " & .Cell(flexcpTextDisplay, Row, .ColIndex("ChequeNo")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "⁄·Ï »‰ﬂ " & .Cell(flexcpTextDisplay, Row, .ColIndex("BankName"))
                LogTexte = "Select Cheque No  " & .Cell(flexcpTextDisplay, Row, .ColIndex("ChequeNo")) & " With Value " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "On Bank " & .Cell(flexcpTextDisplay, Row, .ColIndex("BankName"))
                                                         
            Else
                                                          
                LogTextA = "«·€«¡    ÕœÌœ «·‘Ìﬂ —ﬁ„   " & .Cell(flexcpTextDisplay, Row, .ColIndex("ChequeNo")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "⁄·Ï »‰ﬂ " & .Cell(flexcpTextDisplay, Row, .ColIndex("BankName"))
                LogTexte = "DeSelect Cheque No  " & .Cell(flexcpTextDisplay, Row, .ColIndex("ChequeNo")) & " With Value " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "On Bank " & .Cell(flexcpTextDisplay, Row, .ColIndex("BankName"))
                                                         
            End If
                                                         
            AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtNoteSerial, TxtNoteSerial1
        End If
                                                     
    End With

End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)
    Dim Msg As String

    With Grid1
 
        Select Case .ColKey(Col)

            Case "Remarks"
                Cancel = False
                Exit Sub

            Case "Select"
     
                If .TextMatrix(.Row, .ColIndex("NoteID")) <> "" Then
                    If ChequeBoxCollect(val(.TextMatrix(.Row, .ColIndex("NoteID")))) = False Then
                        Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                        Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«   Õ’Ì· "
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Cancel = True
                        Undo
                        '     .Cell(flexcpChecked, .Row, .ColIndex("Select")) = flexChecked
           
                        Exit Sub
                    End If
                End If
    
                Cancel = False
                Exit Sub
        End Select

        Cancel = True
    End With

End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(9).Enabled = True
    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        Cmd(9).Enabled = False
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False
        Cmd(9).Enabled = False
        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub TxtTotalCash_Change()
    TxtTotalCashView.Text = Format(val(TxtTotalCash.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtTotalCheques_Change()
    TxtTotalChequesView.Text = Format(val(TxtTotalCheques.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtValue.Text, 0)
End Sub

Private Sub TxtValue1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtValue1.Text, 0)
End Sub

Function recalgrid()
   Dim i As Integer
  With VSFlexGrid1
                    For i = .FixedRows To .Rows - 1
            
                        If .TextMatrix(i, .ColIndex("PaymentID")) <> "" Then
                            '  If .TextMatrix(i, .ColIndex("PaymentID")) <> 0 Then
                              
                               '  .TextMatrix(i, .ColIndex("CollectedValue")) = Round(val(.TextMatrix(i, .ColIndex("balance"))), 2)
                            '      .TextMatrix(i, .ColIndex("different")) = 0
                                  
                            '     Else
                            '       .TextMatrix(i, .ColIndex("different")) = Round(val(.TextMatrix(i, .ColIndex("balance"))) - val(.TextMatrix(i, .ColIndex("CollectedValue"))), 2)
                            '         .TextMatrix(i, .ColIndex("NetValue")) = Round(val(.TextMatrix(i, .ColIndex("CollectedValue"))), 2)
                            '  End If
                              
                                  .TextMatrix(i, .ColIndex("different")) = Round(val(.TextMatrix(i, .ColIndex("balance"))) - val(.TextMatrix(i, .ColIndex("CollectedValue"))), 2)
                                     .TextMatrix(i, .ColIndex("NetValue")) = Round(val(.TextMatrix(i, .ColIndex("CollectedValue"))), 2)
                                     
                           .TextMatrix(i, .ColIndex("CommissionValue")) = Round(val(.TextMatrix(i, .ColIndex("CommissionPercentage"))) / 100 * val(.TextMatrix(i, .ColIndex("CollectedValue"))), 2)
                             .TextMatrix(i, .ColIndex("NetValue")) = Round(val(.TextMatrix(i, .ColIndex("CollectedValue"))) - val(.TextMatrix(i, .ColIndex("CommissionValue"))), 2)
                             
                    '        If SystemOptions.UserInterface = ArabicInterface Then
                  '  '            bankDes = bankDes & " „‰  " & .TextMatrix(i, .ColIndex("PaymentName")) & Chr(13)
                  '          Else
                  '              bankDes = bankDes & " From  " & .TextMatrix(i, .ColIndex("PaymentName")) & Chr(13)
                  '          End If
                  '
                    End If
            Next i
    
End With
 
End Function

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
recalgrid
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With VSFlexGrid1

If .ColKey(Col) <> "CollectedValue" And .ColKey(Col) <> "Remarks" Then
  'Cancel = True

End If

If .ColKey(Col) = "CollectedValue" And .Row <> 1 Then
  'Cancel = True

End If

  

    End With
End Sub

 

Private Sub VSFlexGrid1_Click()
   Static lNoteRow&, lNoteCol&, r&, c&

    With VSFlexGrid1
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
       ' r = Fg_Journal.Row
        'c = Fg_Journal.Col

        If .TextMatrix(r, .ColIndex("Account_Code")) <> "" And val(.TextMatrix(r, .ColIndex("Ser"))) = 1 Then
            '        ALLButton1_Click
            lblAccountBalance.Caption = GetbalanceBar(.TextMatrix(.Row, .ColIndex("Account_Code")))
            Else
            
            lblAccountBalance.Caption = GetbalanceBar(.TextMatrix(.Row, .ColIndex("Accountsus")))
            
        End If
    
    End With
    
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
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
