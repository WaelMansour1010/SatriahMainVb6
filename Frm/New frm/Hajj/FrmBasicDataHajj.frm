VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBasicDataHajj 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   Icon            =   "FrmBasicDataHajj.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   13410
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9780
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13410
      _cx             =   23654
      _cy             =   17251
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
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   552
         Left            =   -120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   13536
         _cx             =   23865
         _cy             =   979
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
         Caption         =   "   «·»Ì«‰«  «·«”«”Ì…    "
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
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   3
            Top             =   120
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBasicDataHajj.frx":038A
            ColorButton     =   -2147483634
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
            Left            =   90
            TabIndex        =   4
            Top             =   120
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBasicDataHajj.frx":0724
            ColorButton     =   -2147483634
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
            Left            =   1680
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBasicDataHajj.frx":0ABE
            ColorButton     =   -2147483634
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
            Left            =   615
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBasicDataHajj.frx":0E58
            ColorButton     =   -2147483634
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8880
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   13215
         _cx             =   23310
         _cy             =   15663
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
         Caption         =   "»Ì«‰«  «”«”Ì…|»Ì«‰«  «”«”Ì…|»Ì«‰«  «”«”Ì…|»Ì‰«  «”«”Ì…"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   8460
            Left            =   14160
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   45
            Width           =   13125
            _cx             =   23151
            _cy             =   14923
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic11 
               Height          =   3975
               Left            =   0
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   0
               Width           =   6315
               _cx             =   11139
               _cy             =   7011
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
               Caption         =   "«·„Ê«ﬁ⁄"
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
               Begin VB.TextBox txtNameE10 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1740
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   3525
                  Width           =   3435
               End
               Begin VB.TextBox txtName10 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   1740
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   3150
                  Width           =   3435
               End
               Begin VB.TextBox txtCode10 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3375
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   2835
                  Width           =   1800
               End
               Begin VSFlex8Ctl.VSFlexGrid fg10 
                  Height          =   2310
                  Left            =   0
                  TabIndex        =   103
                  Top             =   330
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   4075
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":11F2
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   330
                  Index           =   16
                  Left            =   360
                  TabIndex        =   104
                  Top             =   2910
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   582
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":12A9
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   17
                  Left            =   360
                  TabIndex        =   105
                  Top             =   3525
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":1643
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   8
                  Left            =   360
                  TabIndex        =   106
                  Top             =   3210
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":7EA5
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   28
                  Left            =   5175
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   3525
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   225
                  Index           =   27
                  Left            =   5115
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   2850
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   26
                  Left            =   5175
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   3135
                  Width           =   1020
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic12 
               Height          =   3975
               Left            =   6435
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   0
               Width           =   6570
               _cx             =   11589
               _cy             =   7011
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
               Caption         =   "«‰Ê«⁄ «·—Õ·« "
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
               Begin VB.TextBox txtNameE9 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1755
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   3525
                  Width           =   3675
               End
               Begin VB.TextBox txtName9 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   1755
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   3150
                  Width           =   3675
               End
               Begin VB.TextBox txtCode9 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   252
                  Left            =   3555
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   2835
                  Width           =   1875
               End
               Begin VSFlex8Ctl.VSFlexGrid fg9 
                  Height          =   2310
                  Left            =   0
                  TabIndex        =   114
                  Top             =   330
                  Width           =   6570
                  _cx             =   11589
                  _cy             =   4075
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":823F
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   330
                  Index           =   18
                  Left            =   360
                  TabIndex        =   115
                  Top             =   2910
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   582
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":82F6
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   19
                  Left            =   360
                  TabIndex        =   116
                  Top             =   3525
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":8690
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   9
                  Left            =   360
                  TabIndex        =   117
                  Top             =   3210
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":EEF2
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   31
                  Left            =   5430
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   3525
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   225
                  Index           =   30
                  Left            =   5430
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   2850
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   29
                  Left            =   5430
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   3135
                  Width           =   1020
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   4140
               Left            =   6435
               TabIndex        =   128
               TabStop         =   0   'False
               Top             =   4080
               Width           =   6570
               _cx             =   11589
               _cy             =   7303
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
               Caption         =   "«·„Ê«”„"
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
               Begin XtremeSuiteControls.RadioButton Omra_Hajj 
                  Height          =   255
                  Index           =   0
                  Left            =   4200
                  TabIndex        =   183
                  Top             =   3840
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "⁄„—Â"
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChCurrYear 
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   182
                  Top             =   2760
                  Width           =   1335
                  _Version        =   786432
                  _ExtentX        =   2355
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "«·„Ê”„ «·Õ«·Ì"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox txtNameE1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1875
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   3540
                  Width           =   3495
               End
               Begin VB.TextBox txtName1 
                  Alignment       =   1  'Right Justify
                  Height          =   270
                  Left            =   1875
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   3165
                  Width           =   3495
               End
               Begin VB.TextBox txtCode1 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   288
                  Left            =   3435
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   2820
                  Width           =   1935
               End
               Begin VSFlex8Ctl.VSFlexGrid fg1 
                  Height          =   2310
                  Left            =   0
                  TabIndex        =   132
                  Top             =   330
                  Width           =   6510
                  _cx             =   11483
                  _cy             =   4075
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
                  FormatString    =   $"FrmBasicDataHajj.frx":F28C
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   345
                  Index           =   0
                  Left            =   420
                  TabIndex        =   133
                  Top             =   2910
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   609
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":F38F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   345
                  Index           =   1
                  Left            =   420
                  TabIndex        =   134
                  Top             =   3540
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   609
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":F729
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   3
                  Left            =   420
                  TabIndex        =   135
                  Top             =   3210
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":15F8B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin XtremeSuiteControls.RadioButton Omra_Hajj 
                  Height          =   255
                  Index           =   1
                  Left            =   2880
                  TabIndex        =   184
                  Top             =   3840
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ÕÃ"
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·„Ê”„"
                  Height          =   285
                  Index           =   51
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   185
                  Top             =   3840
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   285
                  Index           =   7
                  Left            =   5370
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   3540
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   270
                  Index           =   0
                  Left            =   5310
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   2835
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   3
                  Left            =   5370
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   3150
                  Width           =   1080
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic13 
               Height          =   3975
               Left            =   0
               TabIndex        =   139
               TabStop         =   0   'False
               Top             =   4080
               Width           =   6315
               _cx             =   11139
               _cy             =   7011
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
               Caption         =   "«‰Ê«⁄ «·«⁄ „«œ"
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
               Begin VB.TextBox txtCode11 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3255
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   2835
                  Width           =   1800
               End
               Begin VB.TextBox txtName11 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   1620
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   3150
                  Width           =   3435
               End
               Begin VB.TextBox txtNameE11 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1620
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   3525
                  Width           =   3435
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg11 
                  Height          =   2310
                  Left            =   0
                  TabIndex        =   143
                  Top             =   330
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   4075
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":16325
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   330
                  Index           =   20
                  Left            =   360
                  TabIndex        =   144
                  Top             =   2910
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   582
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":163DC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   21
                  Left            =   360
                  TabIndex        =   145
                  Top             =   3525
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":16776
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   10
                  Left            =   360
                  TabIndex        =   146
                  Top             =   3210
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":1CFD8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   41
                  Left            =   5175
                  RightToLeft     =   -1  'True
                  TabIndex        =   149
                  Top             =   3135
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   225
                  Index           =   40
                  Left            =   5115
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   2850
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   39
                  Left            =   5175
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   3525
                  Width           =   1020
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8460
            Index           =   2
            Left            =   45
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   45
            Width           =   13125
            _cx             =   23151
            _cy             =   14923
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   4035
               Left            =   120
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   105
               Width           =   6255
               _cx             =   11033
               _cy             =   7117
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
               Caption         =   "«·ŒÿÊÿ «·ÃÊÌ…"
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
               Begin VB.TextBox txtCode6 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   288
                  Left            =   2940
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   2985
                  Width           =   1935
               End
               Begin VB.TextBox txtName6 
                  Alignment       =   1  'Right Justify
                  Height          =   276
                  Left            =   1380
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   3330
                  Width           =   3495
               End
               Begin VB.TextBox txtNameE6 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1380
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   3615
                  Width           =   3495
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
                  Height          =   2460
                  Left            =   120
                  TabIndex        =   14
                  Top             =   330
                  Width           =   6135
                  _cx             =   10821
                  _cy             =   4339
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":1D372
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   360
                  Index           =   10
                  Left            =   120
                  TabIndex        =   15
                  Top             =   3060
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   635
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":1D429
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   11
                  Left            =   120
                  TabIndex        =   16
                  Top             =   3615
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":1D7C3
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   270
                  Index           =   1
                  Left            =   120
                  TabIndex        =   17
                  Top             =   3390
                  Width           =   840
                  _ExtentX        =   1482
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":24025
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   18
                  Left            =   4935
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   3330
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   270
                  Index           =   19
                  Left            =   4935
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   3000
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   20
                  Left            =   4935
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   3615
                  Width           =   1080
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   4035
               Left            =   6750
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   105
               Width           =   6255
               _cx             =   11033
               _cy             =   7117
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
               Caption         =   "«·„ÿ«—« "
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
               Begin VB.TextBox txtNameE5 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1380
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   3615
                  Width           =   3495
               End
               Begin VB.TextBox txtName5 
                  Alignment       =   1  'Right Justify
                  Height          =   276
                  Left            =   1380
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   3330
                  Width           =   3495
               End
               Begin VB.TextBox txtCode5 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   288
                  Left            =   2940
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   2985
                  Width           =   1935
               End
               Begin VSFlex8Ctl.VSFlexGrid fg5 
                  Height          =   2460
                  Left            =   0
                  TabIndex        =   25
                  Top             =   330
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   4339
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":243BF
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   8
                  Left            =   480
                  TabIndex        =   26
                  Top             =   2955
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":24476
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   9
                  Left            =   480
                  TabIndex        =   27
                  Top             =   3615
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":24810
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   0
                  Left            =   480
                  TabIndex        =   28
                  Top             =   3285
                  Width           =   840
                  _ExtentX        =   1482
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":2B072
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   11
                  Left            =   5115
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   3615
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   270
                  Index           =   16
                  Left            =   5055
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   3000
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   17
                  Left            =   5115
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   3330
                  Width           =   1080
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   3900
               Left            =   120
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   4230
               Width           =   6255
               _cx             =   11033
               _cy             =   6879
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
               Caption         =   "‰Ê⁄ «·«—ﬂ«»"
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
               Begin VB.TextBox txtCode7 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   3195
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   2895
                  Width           =   1920
               End
               Begin VB.TextBox txtName7 
                  Alignment       =   1  'Right Justify
                  Height          =   276
                  Left            =   1680
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   3195
                  Width           =   3435
               End
               Begin VB.TextBox txtNameE7 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1680
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   3480
                  Width           =   3435
               End
               Begin VSFlex8Ctl.VSFlexGrid fg7 
                  Height          =   2355
                  Left            =   0
                  TabIndex        =   36
                  Top             =   330
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   4154
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":2B40C
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   330
                  Index           =   12
                  Left            =   240
                  TabIndex        =   37
                  Top             =   2850
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   582
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":2B4C3
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   13
                  Left            =   240
                  TabIndex        =   38
                  Top             =   3480
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":2B85D
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   6
                  Left            =   240
                  TabIndex        =   39
                  Top             =   3150
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":320BF
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   9
                  Left            =   5115
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   3195
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   225
                  Index           =   10
                  Left            =   5055
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   2895
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   22
                  Left            =   5115
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   3480
                  Width           =   1020
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   4215
               Left            =   6750
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   4230
               Width           =   6255
               _cx             =   11033
               _cy             =   7435
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
               Caption         =   "√‰Ê«⁄ «·»—«„Ã"
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
               Begin VB.TextBox TxtValuee1 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1065
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   178
                  Top             =   2880
                  Width           =   1395
               End
               Begin VB.TextBox txtNameE3 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1065
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   3480
                  Width           =   4035
               End
               Begin VB.TextBox txtName3 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1065
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   3195
                  Width           =   4035
               End
               Begin VB.TextBox txtCode3 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   3450
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   2895
                  Width           =   1650
               End
               Begin VSFlex8Ctl.VSFlexGrid fg3 
                  Height          =   2355
                  Left            =   135
                  TabIndex        =   47
                  Top             =   330
                  Width           =   6135
                  _cx             =   10821
                  _cy             =   4154
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
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":32459
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   330
                  Index           =   4
                  Left            =   135
                  TabIndex        =   48
                  Top             =   2850
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   582
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":32590
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   5
                  Left            =   135
                  TabIndex        =   49
                  Top             =   3480
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":3292A
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   2
                  Left            =   135
                  TabIndex        =   50
                  Top             =   3150
                  Width           =   810
                  _ExtentX        =   1429
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":3918C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo VehicleType 
                  Height          =   315
                  Left            =   1065
                  TabIndex        =   180
                  Top             =   3840
                  Width           =   4035
                  _ExtentX        =   7117
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·Õ«›·…"
                  Height          =   285
                  Index           =   50
                  Left            =   5115
                  RightToLeft     =   -1  'True
                  TabIndex        =   181
                  Top             =   3840
                  Visible         =   0   'False
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·»—‰«„Ã"
                  Height          =   225
                  Index           =   49
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   179
                  Top             =   2880
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   5
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   3480
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   225
                  Index           =   6
                  Left            =   5055
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   2895
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   8
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   3195
                  Width           =   1035
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Frame1 
            Height          =   8460
            Left            =   13860
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   45
            Width           =   13125
            _cx             =   23151
            _cy             =   14923
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
            Caption         =   "”Ì«—«  «·„ ⁄«ÂœÌ‰"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   4140
               Left            =   6690
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   15
               Width           =   6375
               _cx             =   11245
               _cy             =   7303
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
               Caption         =   "«‰Ê«⁄ «·„”«—« "
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
               Begin VB.TextBox txtCode4 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   276
                  Left            =   3315
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   2640
                  Width           =   1920
               End
               Begin VB.TextBox txtName4 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1800
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   2970
                  Width           =   3435
               End
               Begin VB.TextBox txtNameE4 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   1800
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   3270
                  Width           =   3435
               End
               Begin VSFlex8Ctl.VSFlexGrid fg4 
                  Height          =   2190
                  Left            =   120
                  TabIndex        =   58
                  Top             =   345
                  Width           =   6195
                  _cx             =   10927
                  _cy             =   3863
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":39526
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   6
                  Left            =   480
                  TabIndex        =   59
                  Top             =   3060
                  Width           =   840
                  _ExtentX        =   1482
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":395FD
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   7
                  Left            =   480
                  TabIndex        =   60
                  Top             =   3735
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":39997
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo dcCity 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   61
                  Top             =   3615
                  Visible         =   0   'False
                  Width           =   3435
                  _ExtentX        =   6059
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   5
                  Left            =   480
                  TabIndex        =   62
                  Top             =   3405
                  Width           =   840
                  _ExtentX        =   1482
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":401F9
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Õ«›Ÿ… "
                  Height          =   285
                  Index           =   12
                  Left            =   5235
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   3615
                  Visible         =   0   'False
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   300
                  Index           =   13
                  Left            =   5235
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   2970
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   270
                  Index           =   14
                  Left            =   5175
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   2640
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   285
                  Index           =   15
                  Left            =   5235
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   3270
                  Width           =   1020
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   4140
               Left            =   0
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   15
               Width           =   6570
               _cx             =   11589
               _cy             =   7303
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
               Caption         =   "«·„ƒ””«  "
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
               Begin VB.TextBox txtNameE2 
                  Alignment       =   1  'Right Justify
                  Height          =   264
                  Left            =   1755
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   3495
                  Width           =   3495
               End
               Begin VB.TextBox txtName2 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   1755
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   3165
                  Width           =   3495
               End
               Begin VB.TextBox txtCode2 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   252
                  Left            =   3315
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   2880
                  Width           =   1935
               End
               Begin VSFlex8Ctl.VSFlexGrid fg2 
                  Height          =   2400
                  Left            =   120
                  TabIndex        =   71
                  Top             =   360
                  Width           =   6390
                  _cx             =   11271
                  _cy             =   4233
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
                  FormatString    =   $"FrmBasicDataHajj.frx":40593
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   360
                  Index           =   2
                  Left            =   540
                  TabIndex        =   72
                  Top             =   3060
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   635
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":40692
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   3
                  Left            =   540
                  TabIndex        =   73
                  Top             =   3705
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":40A2C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   345
                  Index           =   4
                  Left            =   540
                  TabIndex        =   74
                  Top             =   3360
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   609
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":4728E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo dcGroup 
                  Height          =   315
                  Left            =   1755
                  TabIndex        =   75
                  Top             =   3810
                  Visible         =   0   'False
                  Width           =   3495
                  _ExtentX        =   6165
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   240
                  Index           =   2
                  Left            =   5190
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   2895
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   330
                  Index           =   4
                  Left            =   5250
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   3165
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   270
                  Index           =   21
                  Left            =   5250
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   3495
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ã„Ê⁄…"
                  Height          =   315
                  Index           =   23
                  Left            =   5250
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   3810
                  Visible         =   0   'False
                  Width           =   1020
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic9 
               Height          =   4020
               Left            =   0
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   4200
               Width           =   13005
               _cx             =   22939
               _cy             =   7091
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
               Caption         =   "»Ì«‰«  «·„”«—« "
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
               Begin VB.TextBox TxtKMNo 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1440
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   3240
                  Width           =   1395
               End
               Begin VB.ComboBox DcbPeriodType 
                  Appearance      =   0  'Flat
                  Height          =   315
                  ItemData        =   "FrmBasicDataHajj.frx":47628
                  Left            =   4335
                  List            =   "FrmBasicDataHajj.frx":47632
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   3225
                  Width           =   960
               End
               Begin VB.ComboBox DcbPathType 
                  Appearance      =   0  'Flat
                  Height          =   315
                  ItemData        =   "FrmBasicDataHajj.frx":47642
                  Left            =   7650
                  List            =   "FrmBasicDataHajj.frx":4764C
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   2880
                  Width           =   1320
               End
               Begin VB.TextBox TxtDriverCost 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1440
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   2880
                  Width           =   1395
               End
               Begin VB.TextBox TxtDisalCost 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4335
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   2880
                  Width           =   1740
               End
               Begin VB.TextBox TxtPeriod 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5355
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   3225
                  Width           =   720
               End
               Begin VB.TextBox TxtRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1440
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   91
                  Top             =   3615
                  Width           =   4635
               End
               Begin VB.TextBox txtNameE8 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   7650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   3630
                  Width           =   3795
               End
               Begin VB.TextBox txtName8 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   7650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   3240
                  Width           =   3795
               End
               Begin VB.TextBox txtCode8 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   10050
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   2895
                  Width           =   1395
               End
               Begin VSFlex8Ctl.VSFlexGrid fg8 
                  Height          =   2370
                  Left            =   0
                  TabIndex        =   83
                  Top             =   360
                  Width           =   12945
                  _cx             =   22834
                  _cy             =   4180
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
                  Cols            =   14
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":4765C
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   14
                  Left            =   0
                  TabIndex        =   92
                  Top             =   2970
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":47882
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   15
                  Left            =   0
                  TabIndex        =   93
                  Top             =   3630
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":47C1C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   375
                  Index           =   7
                  Left            =   0
                  TabIndex        =   94
                  Top             =   3300
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":4E47E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ  ﬂ„ «·„ Êﬁ⁄…"
                  Height          =   270
                  Index           =   38
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   3240
                  Width           =   1260
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„œ… «·„ Êﬁ⁄…"
                  Height          =   270
                  Index           =   37
                  Left            =   6255
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   3240
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬂ·›… «·”«∆ﬁ"
                  Height          =   270
                  Index           =   36
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   2880
                  Width           =   1260
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œÌ“·"
                  Height          =   270
                  Index           =   35
                  Left            =   5535
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   2850
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬂ·›… «·„”«—"
                  Height          =   270
                  Index           =   34
                  Left            =   6510
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   2850
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·„”«—"
                  Height          =   270
                  Index           =   33
                  Left            =   8850
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   2895
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   270
                  Index           =   32
                  Left            =   6255
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   3615
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   270
                  Index           =   25
                  Left            =   11625
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   3630
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   270
                  Index           =   24
                  Left            =   11625
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   2910
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   1
                  Left            =   11625
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   3240
                  Width           =   1140
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic14 
            Height          =   8460
            Left            =   14460
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   45
            Width           =   13125
            _cx             =   23151
            _cy             =   14923
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic15 
               Height          =   3975
               Left            =   0
               TabIndex        =   151
               TabStop         =   0   'False
               Top             =   0
               Width           =   6315
               _cx             =   11139
               _cy             =   7011
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
               Caption         =   "«‰Ê«⁄ «·„ÿ«·»« "
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
               Begin VB.TextBox txtCode13 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3375
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   2835
                  Width           =   1800
               End
               Begin VB.TextBox txtName13 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   1740
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   153
                  Top             =   3150
                  Width           =   3435
               End
               Begin VB.TextBox txtNameE13 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1740
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   3525
                  Width           =   3435
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg13 
                  Height          =   2310
                  Left            =   0
                  TabIndex        =   155
                  Top             =   330
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   4075
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":4E818
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   330
                  Index           =   22
                  Left            =   360
                  TabIndex        =   156
                  Top             =   2910
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   582
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":4E8CF
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   23
                  Left            =   360
                  TabIndex        =   157
                  Top             =   3525
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":4EC69
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   11
                  Left            =   360
                  TabIndex        =   158
                  Top             =   3210
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":554CB
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   44
                  Left            =   5175
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   3135
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   225
                  Index           =   43
                  Left            =   5115
                  RightToLeft     =   -1  'True
                  TabIndex        =   160
                  Top             =   2850
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   42
                  Left            =   5175
                  RightToLeft     =   -1  'True
                  TabIndex        =   159
                  Top             =   3525
                  Width           =   1020
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic16 
               Height          =   4575
               Left            =   6375
               TabIndex        =   162
               TabStop         =   0   'False
               Top             =   0
               Width           =   6510
               _cx             =   11483
               _cy             =   8070
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
               Caption         =   "«‰Ê«⁄ «·Õ”„Ì« "
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
               Begin VB.ComboBox DcbDiscount 
                  Height          =   315
                  Left            =   4395
                  TabIndex        =   176
                  Top             =   4200
                  Width           =   1035
               End
               Begin VB.TextBox TxtValuee 
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
                  Left            =   3435
                  MaxLength       =   50
                  TabIndex        =   175
                  Top             =   4200
                  Width           =   900
               End
               Begin VB.TextBox txtCode12 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   3495
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
                  Top             =   2835
                  Width           =   1875
               End
               Begin VB.TextBox txtName12 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1755
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   3150
                  Width           =   3615
               End
               Begin VB.TextBox txtNameE12 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1755
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   163
                  Top             =   3525
                  Width           =   3615
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg12 
                  Height          =   2310
                  Left            =   0
                  TabIndex        =   166
                  Top             =   330
                  Width           =   6510
                  _cx             =   11483
                  _cy             =   4075
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
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBasicDataHajj.frx":55865
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   330
                  Index           =   24
                  Left            =   360
                  TabIndex        =   167
                  Top             =   2910
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   582
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":55983
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   25
                  Left            =   360
                  TabIndex        =   168
                  Top             =   3525
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":55D1D
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   315
                  Index           =   12
                  Left            =   360
                  TabIndex        =   169
                  Top             =   3210
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmBasicDataHajj.frx":5C57F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcbAccount 
                  Height          =   315
                  Left            =   1755
                  TabIndex        =   173
                  Top             =   3840
                  Width           =   3615
                  _ExtentX        =   6376
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·Õ”„"
                  Height          =   195
                  Index           =   5
                  Left            =   5370
                  TabIndex        =   177
                  Top             =   4200
                  Width           =   1260
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Õ”«»"
                  Height          =   300
                  Index           =   48
                  Left            =   5490
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   3840
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   47
                  Left            =   5370
                  RightToLeft     =   -1  'True
                  TabIndex        =   172
                  Top             =   3135
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   225
                  Index           =   46
                  Left            =   5370
                  RightToLeft     =   -1  'True
                  TabIndex        =   171
                  Top             =   2850
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   300
                  Index           =   45
                  Left            =   5370
                  RightToLeft     =   -1  'True
                  TabIndex        =   170
                  Top             =   3525
                  Width           =   1020
               End
            End
         End
      End
   End
End
Attribute VB_Name = "FrmBasicDataHajj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip

Dim rs_CopaniesGroup As ADODB.Recordset
Dim rs_TourismCompanies As ADODB.Recordset
Dim rs_ProgrammTypes  As ADODB.Recordset
Dim rs_Hotels  As ADODB.Recordset
Dim rs_AirPort As ADODB.Recordset
Dim rs_AirLines As ADODB.Recordset
Dim rsDuration As ADODB.Recordset
Dim rs_VehicleType As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim rs_Shrines As ADODB.Recordset
Dim Rs_Trips As ADODB.Recordset
Dim rs_Locations  As ADODB.Recordset
Dim rs_TypeDependence As ADODB.Recordset
Dim rs_TypeDeduction As ADODB.Recordset
Dim rs_TypeClaim As ADODB.Recordset
Dim FromDate_ As Date
Dim ToDate_ As Date
Dim FromDateH_ As String
Dim ToDateH_ As String


Private Sub btnModify_Click(Index As Integer)
Select Case Index
Case 0
Update_AirPort
Case 1
Update_AirLines
Case 2
Update_ProgrammTypes
Case 3
Update_CompaniesGroup
Case 4
Update_TourismCompanies
Case 5
Update_Hotels
Case 6
    Update_VehicleType
Case 7
    Update_Shrines

Case 8
    Update_Locations
Case 9
    Update_Trips
Case 10
Update_TypeDependencep
Case 11
Update_TypeClaim
Case 12
Update_TypeDeduction
End Select
End Sub


Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            Add_CompaniesGroup
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_CompaniesGroup
        Case 2
                Add_TourismCompanies
        Case 3
             If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_TourismCompanies
        Case 4
                Add_ProgrammTypes
        Case 5
                 If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_ProgrammTypes
        Case 6
                Add_Hotels
         Case 7
           If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_Hotels
        
        Case 8
              Add_AirPort
        Case 9
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_AirPort
         Case 10
    
     
              Add_AirLines
                Case 11
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_AirLines
            
           Case 12
                Add_VehicleType
                
            Case 13
                Del_VehicleType
            Case 14
                    Add_Shrines
            Case 15
                Del_Shrines
                
            Case 16
                    Add_Locations
            Case 17
                    Del_Locations
            Case 18
                    Add_Trips
           Case 19
                Del_Trips
           Case 20
           Add_TypeDependence
           Case 21
           Del_TypeDependencep
           Case 22
           Add_TypeClaim
          Case 23
          Del_TypeClaim
          Case 24
           Add_TypeDeduction
          Case 25
          Del_TypeDeduction
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub Add_CompaniesGroup()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
 If TxtName1.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName1.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_CopaniesGroup = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblCompaniesGroup  "
    rs_CopaniesGroup.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_CopaniesGroup.AddNew
    txtCode1.Text = CStr(new_id("TblCompaniesGroup", "ID", "", True))
    rs_CopaniesGroup("ID") = IIf(txtCode1.Text = "", Null, txtCode1.Text)
    rs_CopaniesGroup("Name") = IIf(TxtName1.Text = "", Null, TxtName1.Text)
    rs_CopaniesGroup("NameE") = IIf(txtNameE1.Text = "", Null, txtNameE1.Text)
    rs_CopaniesGroup("creationUserid") = user_id
    rs_CopaniesGroup("creationDate") = Date
    If ChCurrYear.value = vbChecked Then
    rs_CopaniesGroup("CurrYear") = 1
    Else
    rs_CopaniesGroup("CurrYear") = 0
    End If
    If Omra_Hajj(1).value = True Then
    rs_CopaniesGroup("Omra_Hajj") = 1
    Else
    rs_CopaniesGroup("Omra_Hajj") = 0
    End If
    
    rs_CopaniesGroup.update
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_CompaniesGroup
    Clear_CompaniesGroup
    fill_Group
Exit Sub
errortrap:

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
Private Sub Add_TypeDependence()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
 If txtName11.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName11.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_TypeDependence = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblTypeDependence  "
    rs_TypeDependence.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_TypeDependence.AddNew
    txtCode11.Text = CStr(new_id("TblTypeDependence", "ID", "", True))
    rs_TypeDependence("ID") = IIf(txtCode11.Text = "", Null, txtCode11.Text)
    rs_TypeDependence("Name") = IIf(txtName11.Text = "", Null, txtName11.Text)
    rs_TypeDependence("NameE") = IIf(txtNameE11.Text = "", Null, txtNameE11.Text)
    rs_TypeDependence.update
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_TypeDependence
    Clear_TypeDependencep
Exit Sub
errortrap:

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
Private Sub Add_TypeDeduction()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
 If txtName12.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName12.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_TypeDeduction = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblTypeDeduction  "
    rs_TypeDeduction.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_TypeDeduction.AddNew
    txtCode12.Text = CStr(new_id("TblTypeDeduction", "ID", "", True))
    rs_TypeDeduction("ID") = IIf(txtCode12.Text = "", Null, txtCode12.Text)
    rs_TypeDeduction("Name") = IIf(txtName12.Text = "", Null, txtName12.Text)
    rs_TypeDeduction("NameE") = IIf(txtNameE12.Text = "", Null, txtNameE12.Text)
    rs_TypeDeduction("AccountCode") = IIf(DcbAccount.Text = "", Null, DcbAccount.BoundText)
    rs_TypeDeduction("Typ") = IIf(val(DcbDiscount.ListIndex) = -1, Null, val(DcbDiscount.ListIndex))
    rs_TypeDeduction("Valuee") = IIf(val(TxtValuee.Text) = 0, Null, val(TxtValuee.Text))
    rs_TypeDeduction.update
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_TypeDeduction
    Clear_TypeDeduction
Exit Sub
errortrap:

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
Private Sub Add_TypeClaim()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
 If txtName13.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName13.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_TypeClaim = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblTypeClaim  "
    rs_TypeClaim.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_TypeClaim.AddNew
    txtCode13.Text = CStr(new_id("TblTypeClaim", "ID", "", True))
    rs_TypeClaim("ID") = IIf(txtCode13.Text = "", Null, txtCode13.Text)
    rs_TypeClaim("Name") = IIf(txtName13.Text = "", Null, txtName13.Text)
    rs_TypeClaim("NameE") = IIf(txtNameE13.Text = "", Null, txtNameE13.Text)
    rs_TypeClaim.update
    Cn.CommitTrans
    BeginTrans = False
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_TypeClaim
    Clear_TypeClaim
Exit Sub
errortrap:

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



Private Sub Add_TourismCompanies()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  

  
 If TxtName2.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName2.SetFocus
 Exit Sub
 End If
 
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_TourismCompanies = New ADODB.Recordset
    StrSQL = "SELECT  *  From tbltourismcompanies  "
    rs_TourismCompanies.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_TourismCompanies.AddNew
    TxtCode2.Text = CStr(new_id("tbltourismcompanies", "ID", "", True))
    rs_TourismCompanies("ID") = IIf(TxtCode2.Text = "", Null, TxtCode2.Text)
    rs_TourismCompanies("Name") = IIf(TxtName2.Text = "", Null, TxtName2.Text)
    rs_TourismCompanies("NameE") = IIf(txtNameE2.Text = "", Null, txtNameE2.Text)
    'rs_TourismCompanies("GroupID") = IIf(dcGroup.BoundText = "", Null, dcGroup.BoundText)
    rs_TourismCompanies("creationUserid") = user_id
    rs_TourismCompanies("creationDate") = Date
    rs_TourismCompanies.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_TourismCompanies
    Clear_TourismCompanies
Exit Sub
errortrap:

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


Private Sub Add_ProgrammTypes()
  Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
        If val(VehicleType.BoundText) = 0 Or VehicleType.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·Õ«›·…"
        Else
        MsgBox "Please Select The Type Of Vehicle"
        End If
        VehicleType.SetFocus
        Exit Sub
        End If
  
 If TxtName3.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName3.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_ProgrammTypes = New ADODB.Recordset
    StrSQL = "SELECT  *  From tblprogrammtypes  "
    rs_ProgrammTypes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    rs_ProgrammTypes.AddNew
    txtCode3.Text = CStr(new_id("tblprogrammtypes", "ID", "", True))
    rs_ProgrammTypes("VehicleType") = IIf(val(VehicleType.BoundText) = 0, Null, val(VehicleType.BoundText))
    rs_ProgrammTypes("ID") = IIf(txtCode3.Text = "", Null, txtCode3.Text)
    rs_ProgrammTypes("Name") = IIf(TxtName3.Text = "", Null, TxtName3.Text)
    rs_ProgrammTypes("NameE") = IIf(txtNameE3.Text = "", Null, txtNameE3.Text)
    rs_ProgrammTypes("Valuee") = val(TxtValuee1.Text)
    rs_ProgrammTypes("creationUserid") = user_id
    rs_ProgrammTypes("creationDate") = Date
    rs_ProgrammTypes.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_ProgrammTypes
    Clear_ProgrammTypes
Exit Sub
errortrap:

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



Private Sub Add_Hotels()
  Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  

  
 If txtName4.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName4.SetFocus
 Exit Sub
 End If
  

  
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_Hotels = New ADODB.Recordset
    StrSQL = "SELECT  *  From tblhotels  "
    rs_Hotels.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_Hotels.AddNew
    
    txtCode4.Text = CStr(new_id("tblhotels", "ID", "", True))
    rs_Hotels("ID") = IIf(txtCode4.Text = "", Null, txtCode4.Text)
    rs_Hotels("Name") = IIf(txtName4.Text = "", Null, txtName4.Text)
    rs_Hotels("NameE") = IIf(txtNameE4.Text = "", Null, txtNameE4.Text)
  '  rs_Hotels("cityID") = IIf(DcCity.BoundText = "", Null, DcCity.BoundText)
    rs_Hotels("creationUserid") = user_id
    rs_Hotels("creationDate") = Date
    rs_Hotels.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Hotels
    Clear_Hotels
Exit Sub
errortrap:

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
Private Sub Add_AirLines()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  

  
 If txtName6.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName6.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_AirLines = New ADODB.Recordset
    StrSQL = "SELECT  *  From tblairlines  "
    rs_AirLines.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    rs_AirLines.AddNew
    txtCode6.Text = CStr(new_id("tblairlines", "ID", "", True))
    
    rs_AirLines("ID") = IIf(txtCode6.Text = "", Null, txtCode6.Text)
    rs_AirLines("Name") = IIf(txtName6.Text = "", Null, txtName6.Text)
    rs_AirLines("NameE") = IIf(txtNameE6.Text = "", Null, txtNameE6.Text)
    rs_AirLines("creationUserid") = user_id
    rs_AirLines("creationDate") = Date
    rs_AirLines.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_AirLines
    
    
    txtCode6.Text = ""
    txtName6.Text = ""
    txtNameE6.Text = ""

Exit Sub
errortrap:

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
Private Sub Update_AirLines()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If txtName6.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName6.SetFocus
 Exit Sub
 End If
 
        str = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("id"))
        sr = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    
    BeginTrans = True
     StrSQL = "Update tblairlines Set  NameE= '" & txtNameE6.Text & "' ,Name = '" & txtName6.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Retrive_AirLines
    Me.txtCode6.Text = ""
    Me.txtName6.Text = ""
    Me.txtNameE6.Text = ""

Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Update_TypeDependencep()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If txtName11.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName11.SetFocus
 Exit Sub
 End If
        str = Fg11.TextMatrix(Fg11.Row, Fg11.ColIndex("id"))
        sr = Fg11.TextMatrix(Fg11.Row, Fg11.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    
    BeginTrans = True
     StrSQL = "Update TblTypeDependence Set  NameE= '" & txtNameE11.Text & "' ,Name = '" & txtName11.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Clear_TypeDependencep
    Retrive_TypeDependence
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Update_TypeDeduction()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If txtName12.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName12.SetFocus
 Exit Sub
 End If
        str = Fg12.TextMatrix(Fg12.Row, Fg12.ColIndex("id"))
        sr = Fg12.TextMatrix(Fg12.Row, Fg12.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    
    BeginTrans = True
     StrSQL = "Update TblTypeDeduction Set  NameE= '" & txtNameE12.Text & "' ,Name = '" & txtName12.Text & "'  "
     StrSQL = StrSQL & ",AccountCode='" & Me.DcbAccount.BoundText & "' ,Typ=" & val(Me.DcbDiscount.ListIndex) & " ,Valuee=" & val(Me.TxtValuee.Text) & "  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
 Retrive_TypeDeduction
Clear_TypeDeduction
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Update_TypeClaim()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If txtName13.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName13.SetFocus
 Exit Sub
 End If
        str = Fg13.TextMatrix(Fg13.Row, Fg13.ColIndex("id"))
        sr = Fg13.TextMatrix(Fg13.Row, Fg13.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    
    BeginTrans = True
     StrSQL = "Update TblTypeClaim Set  NameE= '" & txtNameE13.Text & "' ,Name = '" & txtName13.Text & "'  "
     StrSQL = StrSQL & " Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
 Retrive_TypeClaim
Clear_TypeClaim
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Update_AirPort()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If txtName5.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName5.SetFocus
 Exit Sub
 End If
 
    str = Fg5.TextMatrix(Fg5.Row, Fg5.ColIndex("id"))
    sr = Fg5.TextMatrix(Fg5.Row, Fg5.ColIndex("serial"))
    If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
    StrSQL = "Update tblairport Set  NameE='" & txtNameE5.Text & "',Name='" & txtName5.Text & "'  Where ID=" & val(str)
    Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Retrive_AirPort
    Clear_AirPort
    
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Update_Hotels()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If txtName4.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName4.SetFocus
 Exit Sub
 End If
 

  
        str = FG4.TextMatrix(FG4.Row, FG4.ColIndex("id"))
        sr = FG4.TextMatrix(FG4.Row, FG4.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update tblhotels Set   NameE='" & txtNameE4.Text & "',Name='" & txtName4.Text & "'   Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Hotels
    Clear_Hotels
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Update_ProgrammTypes()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
        If val(VehicleType.BoundText) = 0 Or VehicleType.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·Õ«›·…"
        Else
        MsgBox "Please Select The Type Of Vehicle"
        End If
        VehicleType.SetFocus
        Exit Sub
        End If
        
 If TxtName3.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName3.SetFocus
 Exit Sub
 End If
 
        str = FG3.TextMatrix(FG3.Row, FG3.ColIndex("id"))
        sr = FG3.TextMatrix(FG3.Row, FG3.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update tblprogrammtypes Set VehicleType=" & val(VehicleType.BoundText) & " , Valuee=" & val(TxtValuee1.Text) & ", NameE='" & txtNameE3.Text & "',Name='" & TxtName3.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Retrive_ProgrammTypes
    Clear_ProgrammTypes
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Update_TourismCompanies()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If TxtName2.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName2.SetFocus
 Exit Sub
 End If
 
    If DCGroup.BoundText = "" Then
           If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox ("«Œ — «·„Ã„Ê⁄… «Ê·«")
           Else
                  MsgBox ("Select Groub ")
           End If
           DCGroup.SetFocus
           Exit Sub
    End If
 
 
    str = FG2.TextMatrix(FG2.Row, FG2.ColIndex("id"))
    sr = FG2.TextMatrix(FG2.Row, FG2.ColIndex("serial"))
    If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
    StrSQL = "Update tbltourismcompanies set NameE='" & txtNameE2.Text & "',Name='" & TxtName2.Text & "'  and  GroupID = " & val(DCGroup.BoundText) & "  Where ID=" & val(str)
    Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
   Retrive_TourismCompanies
 Clear_TourismCompanies
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Update_CompaniesGroup()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If TxtName1.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName1.SetFocus
 Exit Sub
 End If
 Dim CurrYear As Integer
 Dim Omra_Hajj1 As Integer
 If ChCurrYear.value = vbChecked Then
 CurrYear = 1
 Else
 CurrYear = 0
 End If
  If Omra_Hajj(1).value = True Then
 Omra_Hajj1 = 1
 Else
 Omra_Hajj1 = 0
 End If
        str = FG1.TextMatrix(FG1.Row, FG1.ColIndex("id"))
        sr = FG1.TextMatrix(FG1.Row, FG1.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblCompaniesGroup Set  NameE='" & txtNameE1.Text & "',Name='" & TxtName1.Text & "',CurrYear=" & CurrYear & ",Omra_Hajj=" & Omra_Hajj1 & "  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
   Retrive_CompaniesGroup
  Clear_CompaniesGroup
Exit Sub
errortrap:

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
   End If
End Sub

Private Sub Add_AirPort()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  

  
 If txtName5.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName5.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_AirPort = New ADODB.Recordset
    StrSQL = "SELECT  *  From tblairport  "
    rs_AirPort.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    rs_AirPort.AddNew
    txtCode5.Text = CStr(new_id("tblairport", "ID", "", True))
    rs_AirPort("ID") = IIf(txtCode5.Text = "", Null, txtCode5.Text)
    rs_AirPort("Name") = IIf(txtName5.Text = "", Null, txtName5.Text)
    rs_AirPort("NameE") = IIf(txtNameE5.Text = "", Null, txtNameE5.Text)
    rs_AirPort("creationUserid") = user_id
    rs_AirPort("creationDate") = Date
    rs_AirPort.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_AirPort
    Clear_AirPort
Exit Sub
errortrap:

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

Private Sub Del_CompaniesGroup()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = FG1.TextMatrix(FG1.Row, FG1.ColIndex("id"))
        sr = FG1.TextMatrix(FG1.Row, FG1.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_CopaniesGroup.RecordCount < 1 Then
                   
                   Dim mm As String
                   mm = "  select * from TblTourismCompanies where GroupID = " & val(str)
                  Set Rs_Temp = New ADODB.Recordset
                  Rs_Temp.Open mm, Cn, adOpenStatic, adLockOptimistic, adCmdText
                  If Rs_Temp.RecordCount > 0 Then
                                MsgBox ("·« Ì„ﬂ‰ Õ–› «·„Ã„Ê⁄… ° »—Ã«¡ Õ–› «·‘—ﬂ«  «·„÷«›… ⁄·ÌÂ« «Ê·« ")
                                Exit Sub
                  End If
                   
                   StrSQL = "delete From TblCompaniesGroup  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblCompaniesGroup"
                   Set rs_CopaniesGroup = New ADODB.Recordset
                   rs_CopaniesGroup.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   fill_Group
                   Clear_CompaniesGroup
                   
                If rs_CopaniesGroup.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_CompaniesGroup
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_CompaniesGroup
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_CopaniesGroup.CancelUpdate
    'End If
End Sub



Private Sub Del_TourismCompanies()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = FG2.TextMatrix(FG2.Row, FG2.ColIndex("id"))
        sr = FG2.TextMatrix(FG2.Row, FG2.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_TourismCompanies.RecordCount < 1 Then
                StrSQL = "delete From tbltourismcompanies  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From tbltourismcompanies"
                   Set rs_TourismCompanies = New ADODB.Recordset
                   rs_TourismCompanies.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    Clear_TourismCompanies
                If rs_TourismCompanies.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_TourismCompanies
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_TourismCompanies
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_TourismCompanies.CancelUpdate
    'End If
End Sub


Private Sub Del_ProgrammTypes()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = FG3.TextMatrix(FG3.Row, FG3.ColIndex("id"))
        sr = FG3.TextMatrix(FG3.Row, FG3.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_ProgrammTypes.RecordCount < 1 Then
                StrSQL = "delete From tblprogrammtypes  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From tblprogrammtypes"
                   Set rs_ProgrammTypes = New ADODB.Recordset
                   rs_ProgrammTypes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          Clear_ProgrammTypes
                If rs_ProgrammTypes.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_ProgrammTypes
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_ProgrammTypes
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_ProgrammTypes.CancelUpdate
    'End If
End Sub

Private Sub Del_Hotels()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = FG4.TextMatrix(FG4.Row, FG4.ColIndex("id"))
        sr = FG4.TextMatrix(FG4.Row, FG4.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_Hotels.RecordCount < 1 Then
                StrSQL = "delete From tblhotels  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From tblhotels"
                   Set rs_Hotels = New ADODB.Recordset
                   rs_Hotels.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    
                    Clear_Hotels
                If rs_Hotels.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_Hotels
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_Hotels
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_Hotels.CancelUpdate
    'End If
End Sub
   
Private Sub Del_AirLines()

    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("id"))
        sr = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
               
                If Not rs_AirLines.RecordCount < 1 Then
                StrSQL = "delete From tblairlines  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From tblairlines"
                   Set rs_AirLines = New ADODB.Recordset
                   rs_AirLines.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                    txtCode6.Text = ""
                    txtName6.Text = ""
                    txtNameE6.Text = ""
                    
          
                If rs_AirLines.RecordCount < 1 Then
                                Else
                   Retrive_AirLines
                End If
            End If
        End If

    Else
        
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      
        Exit Sub
    End If
 Retrive_AirLines
  
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_AirLines.CancelUpdate
    'End If
End Sub

Private Sub Del_AirPort()

    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = Fg5.TextMatrix(Fg5.Row, Fg5.ColIndex("id"))
        sr = Fg5.TextMatrix(Fg5.Row, Fg5.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_AirPort.RecordCount < 1 Then
                StrSQL = "delete From tblairport  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From tblairport"
                   Set rs_AirPort = New ADODB.Recordset
                   rs_AirPort.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    Clear_AirPort
                If rs_AirPort.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_AirPort
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_AirPort
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_AirPort.CancelUpdate
    'End If
End Sub
   Private Sub Retrive_TypeDependence()
Dim i As Integer
     Set rs_TypeDependence = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblTypeDependence  "
    rs_TypeDependence.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Fg11.Rows = 1
    If rs_TypeDependence.RecordCount > 0 Then
        rs_TypeDependence.MoveFirst
        With Fg11
        .Rows = rs_TypeDependence.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_TypeDependence("id").value), "", rs_TypeDependence("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_TypeDependence("name").value), "", rs_TypeDependence("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_TypeDependence("namee").value), "", rs_TypeDependence("namee").value)
          rs_TypeDependence.MoveNext
         Next
         End With
    End If
End Sub
   Private Sub Retrive_TypeDeduction()
Dim i As Integer
     Set rs_TypeDeduction = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblTypeDeduction  "
    rs_TypeDeduction.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Fg12.Rows = 1
    If rs_TypeDeduction.RecordCount > 0 Then
        rs_TypeDeduction.MoveFirst
        With Fg12
        .Rows = rs_TypeDeduction.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_TypeDeduction("id").value), "", rs_TypeDeduction("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_TypeDeduction("name").value), "", rs_TypeDeduction("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_TypeDeduction("namee").value), "", rs_TypeDeduction("namee").value)
         .TextMatrix(i, .ColIndex("Typ")) = IIf(IsNull(rs_TypeDeduction("Typ").value), -1, rs_TypeDeduction("Typ").value)
         .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs_TypeDeduction("AccountCode").value), "", rs_TypeDeduction("AccountCode").value)
         .TextMatrix(i, .ColIndex("Valuee")) = IIf(IsNull(rs_TypeDeduction("Valuee").value), "", rs_TypeDeduction("Valuee").value)
          rs_TypeDeduction.MoveNext
         Next
         End With
    End If
End Sub
Private Sub Retrive_TypeClaim()
Dim i As Integer
     Set rs_TypeClaim = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblTypeClaim  "
    rs_TypeClaim.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Fg13.Rows = 1
    If rs_TypeClaim.RecordCount > 0 Then
        rs_TypeClaim.MoveFirst
        With Fg13
        .Rows = rs_TypeClaim.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_TypeClaim("id").value), "", rs_TypeClaim("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_TypeClaim("name").value), "", rs_TypeClaim("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_TypeClaim("namee").value), "", rs_TypeClaim("namee").value)
          rs_TypeClaim.MoveNext
         Next
         End With
    End If
End Sub
Private Sub Retrive_CompaniesGroup()
Dim i As Integer
     Set rs_CopaniesGroup = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblCompaniesGroup  "
    rs_CopaniesGroup.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    FG1.Rows = 1
    If rs_CopaniesGroup.RecordCount > 0 Then
        rs_CopaniesGroup.MoveFirst
        With FG1
        .Rows = rs_CopaniesGroup.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_CopaniesGroup("id").value), "", rs_CopaniesGroup("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_CopaniesGroup("name").value), "", rs_CopaniesGroup("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_CopaniesGroup("namee").value), "", rs_CopaniesGroup("namee").value)
         .TextMatrix(i, .ColIndex("CurrYear")) = IIf(IsNull(rs_CopaniesGroup("CurrYear").value), 0, rs_CopaniesGroup("CurrYear").value)
         .TextMatrix(i, .ColIndex("Omra_Hajj")) = IIf(IsNull(rs_CopaniesGroup("Omra_Hajj").value), 0, rs_CopaniesGroup("Omra_Hajj").value)
         
          rs_CopaniesGroup.MoveNext
         Next
         End With
    End If
End Sub


Private Sub Retrive_TourismCompanies()
Dim i As Integer
     Set rs_TourismCompanies = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From tbltourismcompanies  "
    rs_TourismCompanies.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    FG2.Rows = 1
    If rs_TourismCompanies.RecordCount > 0 Then
        rs_TourismCompanies.MoveFirst
        With FG2
        .Rows = rs_TourismCompanies.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_TourismCompanies("id").value), "", rs_TourismCompanies("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_TourismCompanies("name").value), "", rs_TourismCompanies("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_TourismCompanies("namee").value), "", rs_TourismCompanies("namee").value)
     ' .TextMatrix(i, .ColIndex("groupid")) = IIf(IsNull(rs_TourismCompanies("groupid").value), "", rs_TourismCompanies("groupid").value)
          rs_TourismCompanies.MoveNext
         Next
         End With
    End If
End Sub

Private Sub Retrive_ProgrammTypes()
Dim i As Integer
     Set rs_ProgrammTypes = New ADODB.Recordset
    Dim StrSQL As String
   ' StrSQL = "SELECT  *  From tblprogrammtypes  "
  StrSQL = " SELECT     dbo.TBLCarTypes.name AS Vehname, dbo.TBLCarTypes.namee AS VehnameE, dbo.TblProgrammTypes.*"
  StrSQL = StrSQL & "   FROM         dbo.TblProgrammTypes LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TBLCarTypes ON dbo.TblProgrammTypes.VehicleType = dbo.TBLCarTypes.id"
    rs_ProgrammTypes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    FG3.Rows = 1
    If rs_ProgrammTypes.RecordCount > 0 Then
        rs_ProgrammTypes.MoveFirst
        With FG3
        .Rows = rs_ProgrammTypes.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_ProgrammTypes("id").value), "", rs_ProgrammTypes("id").value)
         .TextMatrix(i, .ColIndex("VehicleType")) = IIf(IsNull(rs_ProgrammTypes("VehicleType").value), "", rs_ProgrammTypes("VehicleType").value)
         .TextMatrix(i, .ColIndex("Valuee")) = IIf(IsNull(rs_ProgrammTypes("Valuee").value), "", rs_ProgrammTypes("Valuee").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_ProgrammTypes("name").value), "", rs_ProgrammTypes("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_ProgrammTypes("namee").value), "", rs_ProgrammTypes("namee").value)
         If SystemOptions.UserInterface = ArabicInterface Then
          .TextMatrix(i, .ColIndex("Vehname")) = IIf(IsNull(rs_ProgrammTypes("Vehname").value), "", rs_ProgrammTypes("Vehname").value)
         Else
          .TextMatrix(i, .ColIndex("Vehname")) = IIf(IsNull(rs_ProgrammTypes("VehnameE").value), "", rs_ProgrammTypes("VehnameE").value)
         End If
          rs_ProgrammTypes.MoveNext
         Next
         End With
    End If
End Sub
Private Sub Clear_TypeClaim()
txtCode13.Text = ""
txtName13.Text = ""
txtNameE13.Text = ""
End Sub
Private Sub Clear_TypeDeduction()
txtCode11.Text = ""
txtName11.Text = ""
txtNameE11.Text = ""
DcbAccount.BoundText = ""
DcbDiscount.ListIndex = -1
TxtValuee.Text = ""
End Sub

Private Sub Clear_TypeDependencep()
txtCode12.Text = ""
txtName12.Text = ""
txtNameE12.Text = ""
End Sub

Private Sub Clear_CompaniesGroup()
txtCode1.Text = ""
TxtName1.Text = ""
txtNameE1.Text = ""
Me.ChCurrYear.value = vbUnchecked
End Sub

Private Sub Clear_TourismCompanies()
TxtCode2.Text = ""
TxtName2.Text = ""
txtNameE2.Text = ""
DCGroup.BoundText = ""
End Sub

Private Sub Clear_ProgrammTypes()
txtCode3.Text = ""
TxtName3.Text = ""
txtNameE3.Text = ""
TxtValuee1.Text = ""
VehicleType.BoundText = 0
End Sub

Private Sub Clear_Hotels()
txtCode4.Text = ""
txtName4.Text = ""
txtNameE4.Text = ""

End Sub

Private Sub Clear_AirPort()
txtCode5.Text = ""
txtName5.Text = ""
txtNameE5.Text = ""
'DcBranch5.BoundText = ""
End Sub


Private Sub Retrive_Hotels()
Dim i As Integer
     Set rs_Hotels = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From tblhotels  "
    rs_Hotels.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    FG4.Rows = 1
    If rs_Hotels.RecordCount > 0 Then
        rs_Hotels.MoveFirst
        With FG4
        .Rows = rs_Hotels.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_Hotels("id").value), "", rs_Hotels("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_Hotels("name").value), "", rs_Hotels("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_Hotels("namee").value), "", rs_Hotels("namee").value)
          .TextMatrix(i, .ColIndex("cityID")) = IIf(IsNull(rs_Hotels("cityID").value), "", rs_Hotels("cityID").value)
          rs_Hotels.MoveNext
         Next
         End With
    End If
End Sub
Private Sub Retrive_AirLines()
Dim i As Integer
     Set rs_AirLines = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From tblairlines  "
    rs_AirLines.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    VSFlexGrid1.Rows = 1
    If rs_AirLines.RecordCount > 0 Then
        rs_AirLines.MoveFirst
        With VSFlexGrid1
        .Rows = rs_AirLines.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_AirLines("id").value), "", rs_AirLines("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_AirLines("name").value), "", rs_AirLines("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_AirLines("namee").value), "", rs_AirLines("namee").value)

          rs_AirLines.MoveNext
         Next
         End With
    End If

End Sub
Private Sub Retrive_AirPort()
Dim i As Integer
     Set rs_AirPort = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From tblairport  "
    rs_AirPort.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Fg5.Rows = 1
    If rs_AirPort.RecordCount > 0 Then
        rs_AirPort.MoveFirst
        With Fg5
        .Rows = rs_AirPort.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_AirPort("id").value), "", rs_AirPort("id").value)

         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_AirPort("name").value), "", rs_AirPort("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_AirPort("namee").value), "", rs_AirPort("namee").value)
   
          rs_AirPort.MoveNext
         Next
         End With
    End If
End Sub



Private Sub Fg1_Click()
With Me.FG1
If .Row > 0 Then
txtCode1.Text = .TextMatrix(.Row, .ColIndex("id"))
TxtName1.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE1.Text = .TextMatrix(.Row, .ColIndex("NameE"))
If val(.TextMatrix(.Row, .ColIndex("CurrYear"))) = 1 Then
ChCurrYear.value = vbChecked
Else
ChCurrYear.value = vbUnchecked
End If
If val(.TextMatrix(.Row, .ColIndex("Omra_Hajj"))) = 1 Then
Omra_Hajj(1).value = True
Else
Omra_Hajj(0).value = True
End If
End If
End With
End Sub

Private Sub Fg11_Click()
    With Me.Fg11
            If .Row > 0 Then
                    txtCode11.Text = .TextMatrix(.Row, .ColIndex("id"))
                    txtName11.Text = .TextMatrix(.Row, .ColIndex("Name"))
                    Me.txtNameE11.Text = .TextMatrix(.Row, .ColIndex("NameE"))
            End If
    End With
End Sub

Private Sub Fg12_Click()
With Me.Fg12
If .Row > 0 Then
txtCode12.Text = .TextMatrix(.Row, .ColIndex("id"))
txtName12.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE12.Text = .TextMatrix(.Row, .ColIndex("NameE"))
Me.TxtValuee.Text = .TextMatrix(.Row, .ColIndex("Valuee"))
Me.DcbDiscount.ListIndex = val(.TextMatrix(.Row, .ColIndex("Typ")))
Me.DcbAccount.BoundText = .TextMatrix(.Row, .ColIndex("AccountCode"))
End If
End With
End Sub

Private Sub Fg13_Click()
With Me.Fg13
If .Row > 0 Then
txtCode13.Text = .TextMatrix(.Row, .ColIndex("id"))
txtName13.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE13.Text = .TextMatrix(.Row, .ColIndex("NameE"))
End If
End With
End Sub

Private Sub Fg2_Click()
With Me.FG2
If .Row > 0 Then
TxtCode2.Text = .TextMatrix(.Row, .ColIndex("id"))
TxtName2.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE2.Text = .TextMatrix(.Row, .ColIndex("NameE"))
'dcGroup.BoundText = .TextMatrix(.Row, .ColIndex("groupid"))
End If
End With
End Sub

Private Sub fg3_Click()
With Me.FG3
If .Row > 0 Then
txtCode3.Text = .TextMatrix(.Row, .ColIndex("id"))
TxtName3.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE3.Text = .TextMatrix(.Row, .ColIndex("NameE"))
Me.TxtValuee1.Text = .TextMatrix(.Row, .ColIndex("Valuee"))
Me.VehicleType.BoundText = val(.TextMatrix(.Row, .ColIndex("VehicleType")))
End If
End With
End Sub

Private Sub fg4_Click()
With Me.FG4
If .Row > 0 Then
txtCode4.Text = .TextMatrix(.Row, .ColIndex("id"))
txtName4.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE4.Text = .TextMatrix(.Row, .ColIndex("NameE"))
'DcCity.BoundText = .TextMatrix(.Row, .ColIndex("cityID"))
End If
End With
End Sub

Private Sub fg5_Click()
With Me.Fg5
If .Row > 0 Then
txtCode5.Text = .TextMatrix(.Row, .ColIndex("id"))
txtName5.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE5.Text = .TextMatrix(.Row, .ColIndex("NameE"))
End If
End With
End Sub

Private Sub fg7_Click()
With Me.fg7
    If .Row > 0 Then
            TXTCode7.Text = .TextMatrix(.Row, .ColIndex("id"))
            TxtName7.Text = .TextMatrix(.Row, .ColIndex("Name"))
            Me.TxtNameE7.Text = .TextMatrix(.Row, .ColIndex("NameE"))
    End If
End With
End Sub

Private Sub fg8_Click()
    With Me.fg8
            If .Row > 0 Then
                    txtCode8.Text = .TextMatrix(.Row, .ColIndex("id"))
                    txtName8.Text = .TextMatrix(.Row, .ColIndex("Name"))
                    Me.txtNameE8.Text = .TextMatrix(.Row, .ColIndex("NameE"))
                    TxtDisalCost.Text = .TextMatrix(.Row, .ColIndex("DisalCost"))
                    TxtDriverCost.Text = .TextMatrix(.Row, .ColIndex("DriverCost"))
                    TxtRemarks.Text = .TextMatrix(.Row, .ColIndex("Remarks"))
                    Me.TxtKMNo.Text = .TextMatrix(.Row, .ColIndex("KMNo"))
                    Me.txtPeriod.Text = .TextMatrix(.Row, .ColIndex("Period"))
                    Me.DcbPeriodType.ListIndex = val(.TextMatrix(.Row, .ColIndex("PeriodType")))
                    Me.DcbPathType.ListIndex = val(.TextMatrix(.Row, .ColIndex("PathType")))
            End If
    End With
End Sub

Private Sub fg9_Click()
    With Me.fg9
            If .Row > 0 Then
                    txtCode9.Text = .TextMatrix(.Row, .ColIndex("id"))
                    txtName9.Text = .TextMatrix(.Row, .ColIndex("Name"))
                    Me.txtNameE9.Text = .TextMatrix(.Row, .ColIndex("NameE"))
            End If
    End With
End Sub

Private Sub fg10_Click()
    With Me.fg10
            If .Row > 0 Then
                    txtCode10.Text = .TextMatrix(.Row, .ColIndex("id"))
                    txtName10.Text = .TextMatrix(.Row, .ColIndex("Name"))
                    Me.txtNameE10.Text = .TextMatrix(.Row, .ColIndex("NameE"))
            End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
   
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
    On Error GoTo ErrTrap
    C1Tab1.CurrTab = 0
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
   ' Dcombos.getCountriesGovernments DcCity
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
    Dcombos.GetTblCarsDataGroup VehicleType, 1, True
   fill_Group
       
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
    With DcbPeriodType
    .Clear
    .AddItem "œﬁÌﬁ…"
    .AddItem "”«⁄…"
    End With
    With DcbPathType
    .Clear
    .AddItem "⁄„—…"
    .AddItem "ÕÃ"
    End With
      With DcbDiscount
    .Clear
    .AddItem "ﬁÌ„Â"
    .AddItem "‰”»…"
    End With
    Else
        With DcbPeriodType
    .Clear
    .AddItem "Minute"
    .AddItem "Hour"
    End With
    With DcbDiscount
    .Clear
    .AddItem "Value"
    .AddItem "Percentage"
    End With
    With DcbPathType
    .Clear
    .AddItem "Amra"
    .AddItem "Hajj"
    End With
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «‰Ê«⁄ «·„Œ«·›«   "
    LogTexte = " Open Window " & "  Violation Types "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
    Resize_Form Me
    Retrive_TypeClaim
    Retrive_TypeDeduction
    Retrive_AirLines
    Retrive_VehicleType
    Retrive_ProgrammTypes
    Retrive_CompaniesGroup
    Retrive_AirPort
    Retrive_Hotels
    Retrive_Shrines
    Retrive_TourismCompanies
    Retrive_Trips
    Retrive_Locations
    Retrive_TypeDependence
    Exit Sub
ErrTrap:
End Sub
'

Private Sub fill_Group()

Dim str As String
str = "select id , name from tblCompaniesGroup"
fill_combo DCGroup, str
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
     
  lbl(17).Caption = "Name Ar"
  lbl(18).Caption = "Name Ar"
  lbl(8).Caption = "Name Ar"
  lbl(3).Caption = "Name Ar"
  lbl(4).Caption = "Name Ar"
  lbl(13).Caption = "Name Ar"
   lbl(11).Caption = "Name Eng"
   lbl(20).Caption = "Name Eng"
   lbl(5).Caption = "Name Eng"
   lbl(7).Caption = "Name Eng"
   lbl(21).Caption = "Name Eng"
   lbl(15).Caption = "Name Eng"
   lbl(16).Caption = "Code"
   lbl(19).Caption = "Code"
   lbl(6).Caption = "Code"
   lbl(0).Caption = "Code"
   lbl(2).Caption = "Code"
   lbl(14).Caption = "Code"
Cmd(8).Caption = "Add"
Cmd(10).Caption = "Add"
Cmd(4).Caption = "Add"
Cmd(6).Caption = "Add"
Cmd(2).Caption = "Add"
Cmd(0).Caption = "Add"
Cmd(9).Caption = "Delete"
Cmd(1).Caption = "Delete"
Cmd(11).Caption = "Delete"
Cmd(3).Caption = "Delete"
Cmd(5).Caption = "Delete"
Cmd(7).Caption = "Delete"
lbl(9).Caption = "Com.Entr"
lbl(10).Caption = "Com.Out"
lbl(12).Caption = "Branch"
  btnModify(0).Caption = "Update"
  btnModify(1).Caption = "Update"
  btnModify(2).Caption = "Update"
  btnModify(3).Caption = "Update"
  btnModify(4).Caption = "Update"
  btnModify(5).Caption = "Update"
    Me.Caption = "  Baisc Data"
    EleHeader.Caption = Me.Caption
    
    C1Elastic5.Caption = "Group of Investors"
    C1Elastic3.Caption = "Type of Investors"
    C1Elastic4.Caption = "Type Contributions"
    C1Elastic2.Caption = "Group Contributions"
    C1Elastic6.Caption = "Type Development"
    C1Elastic7.Caption = "Type Separation"
  With Fg5
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With VSFlexGrid1
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With FG3
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With FG1
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With FG2
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
    .TextMatrix(0, .ColIndex("EntryFees")) = "Com.Ent"
  .TextMatrix(0, .ColIndex("ExitFees")) = "Com.Out"
  End With
    With FG4
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «‰Ê«⁄ «·„Œ«·›«   "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtDisalCost_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDisalCost.Text, 0)
End Sub

Private Sub TxtDriverCost_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDriverCost.Text, 0)
End Sub

Private Sub TxtKMNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtKMNo.Text, 0)
End Sub

Private Sub txtName1_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtName10_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtName2_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub


Private Sub txtName3_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtName4_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtName5_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub


Private Sub txtName6_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub


Private Sub txtName8_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub


Private Sub txtName9_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE1_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub



Private Sub txtNameE10_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtNameE2_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub


Private Sub txtNameE3_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtNameE4_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub


Private Sub txtNameE5_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtNameE6_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub


Private Sub txtNameE8_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub


Private Sub txtNameE9_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtPeriod_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.txtPeriod.Text, 0)
End Sub

Private Sub TxtValuee1_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtValuee1.Text, 0)
End Sub

Private Sub VSFlexGrid1_Click()
With Me.VSFlexGrid1
    If .Row > 0 Then
            txtCode6.Text = .TextMatrix(.Row, .ColIndex("id"))
            txtName6.Text = .TextMatrix(.Row, .ColIndex("Name"))
            Me.txtNameE6.Text = .TextMatrix(.Row, .ColIndex("NameE"))
    End If
End With
End Sub



Private Sub Add_VehicleType()
  Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
 
  
 If TxtName7.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName7.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_VehicleType = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblvehicleType  "
    rs_VehicleType.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    rs_VehicleType.AddNew
    
    TXTCode7.Text = CStr(new_id("TblvehicleType", "ID", "", True))
    rs_VehicleType("ID") = IIf(TXTCode7.Text = "", Null, TXTCode7.Text)
    rs_VehicleType("Name") = IIf(TxtName7.Text = "", Null, TxtName7.Text)
    rs_VehicleType("NameE") = IIf(TxtNameE7.Text = "", Null, TxtNameE7.Text)
    rs_VehicleType("creationUserid") = user_id
    rs_VehicleType("creationDate") = Date
    rs_VehicleType.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_VehicleType
    Clear_VehicleType
Exit Sub
errortrap:

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

Private Sub Update_VehicleType()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If TxtName7.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName7.SetFocus
 Exit Sub
 End If
 
        str = fg7.TextMatrix(fg7.Row, fg7.ColIndex("id"))
        sr = fg7.TextMatrix(fg7.Row, fg7.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblvehicleType Set  NameE='" & TxtNameE7.Text & "',Name='" & TxtName7.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Retrive_VehicleType
    Clear_VehicleType
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Del_VehicleType()
    
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = fg7.TextMatrix(fg7.Row, fg7.ColIndex("id"))
        sr = fg7.TextMatrix(fg7.Row, fg7.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
               
                If Not rs_VehicleType.RecordCount < 1 Then
                StrSQL = "delete From TblvehicleType  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblvehicleType"
                   Set rs_VehicleType = New ADODB.Recordset
                   rs_VehicleType.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   Clear_VehicleType
                If rs_VehicleType.RecordCount < 1 Then
                                
                Else
                   Retrive_VehicleType
                End If
            End If
        End If

    Else
        
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      
        Exit Sub
    End If
 Retrive_VehicleType
  
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_VehicleType.CancelUpdate
    'End If
End Sub



Private Sub Retrive_VehicleType()
Dim i As Integer
     Set rs_VehicleType = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblvehicleType  "
    rs_VehicleType.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    fg7.Rows = 1
    If rs_VehicleType.RecordCount > 0 Then
        rs_VehicleType.MoveFirst
        With fg7
        .Rows = rs_VehicleType.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_VehicleType("id").value), "", rs_VehicleType("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_VehicleType("name").value), "", rs_VehicleType("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_VehicleType("namee").value), "", rs_VehicleType("namee").value)
          rs_VehicleType.MoveNext
         Next
         End With
    End If
End Sub

Private Sub Clear_VehicleType()
TXTCode7.Text = ""
TxtName7.Text = ""
TxtNameE7.Text = ""

End Sub


Private Sub Add_Shrines()
  Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
 
  
 If txtName8.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName8.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_Shrines = New ADODB.Recordset
    StrSQL = "SELECT  *  From  TblShrines  "
    rs_Shrines.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    rs_Shrines.AddNew
    
    txtCode8.Text = CStr(new_id("TblShrines", "ID", "", True))
    rs_Shrines("ID") = IIf(txtCode8.Text = "", Null, txtCode8.Text)
    rs_Shrines("Name") = IIf(txtName8.Text = "", Null, txtName8.Text)
    rs_Shrines("NameE") = IIf(txtNameE8.Text = "", Null, txtNameE8.Text)
    rs_Shrines("creationUserid") = user_id
    rs_Shrines("creationDate") = Date
    rs_Shrines("PathType") = IIf(val(DcbPathType.ListIndex) = -1, Null, val(DcbPathType.ListIndex))
    rs_Shrines("DisalCost") = IIf(TxtDisalCost.Text = "", Null, val(TxtDisalCost.Text))
    rs_Shrines("DriverCost") = IIf(TxtDriverCost.Text = "", Null, val(TxtDriverCost.Text))
    rs_Shrines("Remarks") = IIf(TxtRemarks.Text = "", Null, TxtRemarks.Text)
    rs_Shrines("Period") = IIf(txtPeriod.Text = "", Null, val(txtPeriod.Text))
    rs_Shrines("KMNo") = IIf(TxtKMNo.Text = "", Null, val(TxtKMNo.Text))
    rs_Shrines("PeriodType") = IIf(val(DcbPeriodType.ListIndex) = -1, Null, val(DcbPeriodType.ListIndex))
    rs_Shrines.update
    
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Shrines
    Clear_Shrines
Exit Sub
errortrap:

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



Private Sub Update_Shrines()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    On Error GoTo errortrap
    Dim str As String, sr As String
    
    If txtName8.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("«œŒ· «·«”„ «Ê·«")
        Else
            MsgBox ("Enter Name ")
        End If
        txtName8.SetFocus
        Exit Sub
    End If
 
    str = fg8.TextMatrix(fg8.Row, fg8.ColIndex("id"))
    sr = fg8.TextMatrix(fg8.Row, fg8.ColIndex("serial"))
    
    If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblShrines Set  NameE='" & txtNameE8.Text & "',Name='" & txtName8.Text & "' "
     StrSQL = StrSQL & " , PathType=" & val(Me.DcbPathType.ListIndex) & ",Remarks='" & TxtRemarks.Text & "'"
     StrSQL = StrSQL & " , PeriodType=" & val(Me.DcbPeriodType.ListIndex) & ",Period=" & val(txtPeriod.Text) & ""
     StrSQL = StrSQL & " , DisalCost=" & val(TxtDisalCost.Text) & ",DriverCost=" & val(TxtDriverCost.Text) & " ,KMNo=" & val(TxtKMNo.Text) & "  Where ID=" & val(str)
     Cn.Execute StrSQL, , adExecuteNoRecords
     Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Else
    MsgBox ("Save Successfully")
    End If
    Retrive_Shrines
    Clear_Shrines
Exit Sub
errortrap:

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
   End If
End Sub


Private Sub Del_Shrines()
    
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = fg8.TextMatrix(fg8.Row, fg8.ColIndex("id"))
        sr = fg8.TextMatrix(fg8.Row, fg8.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
               
                If Not rs_Shrines.RecordCount < 1 Then
                   StrSQL = "delete From TblShrines  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblShrines "
                   Set rs_Shrines = New ADODB.Recordset
                   rs_Shrines.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   Clear_Shrines
                If rs_Shrines.RecordCount < 1 Then
                                
                Else
                   Retrive_Shrines
                End If
            End If
        End If

    Else
        
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      
        Exit Sub
    End If
 Retrive_Shrines
  
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_Shrines.CancelUpdate
    'End If
End Sub


Private Sub Retrive_Shrines()
Dim i As Integer
     Set rs_Shrines = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblShrines  "
    rs_Shrines.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    fg8.Rows = 1
    If rs_Shrines.RecordCount > 0 Then
        rs_Shrines.MoveFirst
        With fg8
        .Rows = rs_Shrines.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_Shrines("id").value), "", rs_Shrines("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_Shrines("name").value), "", rs_Shrines("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_Shrines("namee").value), "", rs_Shrines("namee").value)
.TextMatrix(i, .ColIndex("DisalCost")) = IIf(IsNull(rs_Shrines("DisalCost").value), 0, rs_Shrines("DisalCost").value)
.TextMatrix(i, .ColIndex("DriverCost")) = IIf(IsNull(rs_Shrines("DriverCost").value), 0, rs_Shrines("DriverCost").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs_Shrines("Remarks").value), "", rs_Shrines("Remarks").value)
.TextMatrix(i, .ColIndex("KMNo")) = IIf(IsNull(rs_Shrines("KMNo").value), 0, rs_Shrines("KMNo").value)
.TextMatrix(i, .ColIndex("PeriodType")) = IIf(IsNull(rs_Shrines("PeriodType").value), -1, rs_Shrines("PeriodType").value)
.TextMatrix(i, .ColIndex("PathType")) = IIf(IsNull(rs_Shrines("PathType").value), -1, rs_Shrines("PathType").value)
.TextMatrix(i, .ColIndex("Period")) = IIf(IsNull(rs_Shrines("Period").value), 0, rs_Shrines("Period").value)

If Not (IsNull(rs_Shrines("PeriodType").value)) Then
If rs_Shrines("PeriodType").value = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("PeriodTp")) = rs_Shrines("Period").value & " " & "œﬁÌﬁ…"
Else
.TextMatrix(i, .ColIndex("PeriodTp")) = "Minute" & " " & rs_Shrines("Period").value
End If
Else

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("PeriodTp")) = rs_Shrines("Period").value & " " & "”«⁄…"
Else
.TextMatrix(i, .ColIndex("PeriodTp")) = "Hur" & " " & rs_Shrines("Period").value
End If
End If
End If

If Not (IsNull(rs_Shrines("PathType").value)) Then
If rs_Shrines("PathType").value = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("PathTypeT")) = "⁄„—…"
Else
.TextMatrix(i, .ColIndex("PathTypeT")) = "Amra"
End If
Else

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("PathTypeT")) = "ÕÃ"
Else
.TextMatrix(i, .ColIndex("PathTypeT")) = "Hajj"
End If
End If
End If

          rs_Shrines.MoveNext
         Next
         End With
    End If

End Sub


Private Sub Clear_Shrines()
txtCode8.Text = ""
txtName8.Text = ""
txtNameE8.Text = ""
'DcBranch5.BoundText = ""
End Sub


Private Sub Add_Trips()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  

 If txtName9.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName9.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set Rs_Trips = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblTrips  "
    Rs_Trips.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Rs_Trips.AddNew
    txtCode9.Text = CStr(new_id("TblTrips", "ID", "", True))
    Rs_Trips("ID") = IIf(txtCode9.Text = "", Null, txtCode9.Text)
    Rs_Trips("Name") = IIf(txtName9.Text = "", Null, txtName9.Text)
    Rs_Trips("NameE") = IIf(txtNameE9.Text = "", Null, txtNameE9.Text)
    Rs_Trips("creationUserid") = user_id
    Rs_Trips("creationDate") = Date
    Rs_Trips.update
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Trips
    Clear_Trips
    'fill_Trips
Exit Sub
errortrap:

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


Private Sub Add_Locations()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  

 If txtName10.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName10.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_Locations = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblLocations  "
    rs_Locations.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs_Locations.AddNew
    txtCode10.Text = CStr(new_id("TblLocations", "ID", "", True))
    rs_Locations("ID") = IIf(txtCode10.Text = "", Null, txtCode10.Text)
    rs_Locations("Name") = IIf(txtName10.Text = "", Null, txtName10.Text)
    rs_Locations("NameE") = IIf(txtNameE10.Text = "", Null, txtNameE10.Text)
    rs_Locations("creationUserid") = user_id
    rs_Locations("creationDate") = Date
    rs_Locations.update
    Cn.CommitTrans
    BeginTrans = False
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Locations
    Clear_Locations
    'fill_Locations
Exit Sub
errortrap:

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

Private Sub Update_Trips()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If txtName9.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName9.SetFocus
 Exit Sub
 End If
 
        str = fg9.TextMatrix(fg9.Row, fg9.ColIndex("id"))
        sr = fg9.TextMatrix(fg9.Row, fg9.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblTrips Set  NameE='" & txtNameE9.Text & "',Name='" & txtName9.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Trips
    Clear_Trips
Exit Sub
errortrap:

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
   End If
End Sub


Private Sub Update_Locations()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If txtName10.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 txtName10.SetFocus
 Exit Sub
 End If
 
        str = fg10.TextMatrix(fg10.Row, fg10.ColIndex("id"))
        sr = fg10.TextMatrix(fg10.Row, fg10.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblLocations Set  NameE='" & txtNameE10.Text & "',Name='" & txtName10.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Locations
    Clear_Locations
Exit Sub
errortrap:

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
   End If
End Sub
Private Sub Del_Trips()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = fg9.TextMatrix(fg9.Row, fg9.ColIndex("id"))
        sr = fg9.TextMatrix(fg9.Row, fg9.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs_Trips.RecordCount < 1 Then
                StrSQL = "delete From TblTrips  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblTrips"
                   Set Rs_Trips = New ADODB.Recordset
                   Rs_Trips.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   Clear_Trips
                If Rs_Trips.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_Trips
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_Trips
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    Rs_Trips.CancelUpdate
    'End If
End Sub
Private Sub Del_TypeDeduction()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = Fg12.TextMatrix(Fg12.Row, Fg12.ColIndex("id"))
        sr = Fg12.TextMatrix(Fg12.Row, Fg12.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_TypeDeduction.RecordCount < 1 Then
                StrSQL = "delete From TblTypeDeduction  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblTypeDeduction"
                   Set rs_TypeDeduction = New ADODB.Recordset
                   rs_TypeDeduction.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   Clear_TypeDeduction
                If rs_TypeDeduction.RecordCount < 1 Then
              Clear_TypeDeduction
                Else
                Retrive_TypeDeduction
                End If
            End If
        End If

    Else
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_TypeDeduction
    
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_TypeDeduction.CancelUpdate
    'End If
End Sub

Private Sub Del_TypeClaim()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = Fg13.TextMatrix(Fg13.Row, Fg13.ColIndex("id"))
        sr = Fg13.TextMatrix(Fg13.Row, Fg13.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_TypeClaim.RecordCount < 1 Then
                StrSQL = "delete From TblTypeClaim  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblTypeDeduction"
                   Set rs_TypeClaim = New ADODB.Recordset
                   rs_TypeClaim.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   Clear_TypeClaim
                If rs_TypeClaim.RecordCount < 1 Then
              Clear_TypeClaim
                Else
                Retrive_TypeClaim
                End If
            End If
        End If

    Else
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_TypeClaim
    
    Exit Sub
ErrTrap:
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_TypeClaim.CancelUpdate
    'End If
End Sub

Private Sub Del_TypeDependencep()
    Dim Msg As String
    Dim StrSQL As String
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = Fg11.TextMatrix(Fg11.Row, Fg11.ColIndex("id"))
        sr = Fg11.TextMatrix(Fg11.Row, Fg11.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_TypeDependence.RecordCount < 1 Then
                StrSQL = "delete From TblTypeDependence  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblTypeDependence"
                   Set rs_TypeDependence = New ADODB.Recordset
                   rs_TypeDependence.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   Clear_TypeDependencep
                If rs_TypeDependence.RecordCount < 1 Then
                   Clear_TypeDependencep
                Else
                   Retrive_TypeDependence
                End If
            End If
        End If

    Else
       
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_TypeDependence
    
    Exit Sub
ErrTrap:
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_TypeDependence.CancelUpdate
End Sub


Private Sub Del_Locations()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = fg10.TextMatrix(fg10.Row, fg10.ColIndex("id"))
        sr = fg10.TextMatrix(fg10.Row, fg10.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_Locations.RecordCount < 1 Then
                StrSQL = "delete From TblLocations  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblLocations"
                   Set rs_Locations = New ADODB.Recordset
                   rs_Locations.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   Clear_Locations
                If rs_Locations.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_Locations
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_Locations
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_Locations.CancelUpdate
    'End If
End Sub

Private Sub Retrive_Trips()
Dim i As Integer
     Set Rs_Trips = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = " SELECT  *  From TblTrips  "
    Rs_Trips.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    fg9.Rows = 1
    If Rs_Trips.RecordCount > 0 Then
        Rs_Trips.MoveFirst
        With fg9
        .Rows = Rs_Trips.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_Trips("id").value), "", Rs_Trips("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs_Trips("name").value), "", Rs_Trips("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs_Trips("namee").value), "", Rs_Trips("namee").value)
       
          Rs_Trips.MoveNext
         Next
         End With
    End If
End Sub


Private Sub Retrive_Locations()
Dim i As Integer
     Set rs_Locations = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = " SELECT  *  From TblLocations "
    rs_Locations.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    fg10.Rows = 1
    If rs_Locations.RecordCount > 0 Then
        rs_Locations.MoveFirst
        With fg10
        .Rows = rs_Locations.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_Locations("id").value), "", rs_Locations("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs_Locations("name").value), "", rs_Locations("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs_Locations("namee").value), "", rs_Locations("namee").value)

          rs_Locations.MoveNext
         Next
         End With
    End If
End Sub

Private Sub Clear_Trips()
txtCode9.Text = ""
txtName9.Text = ""
txtNameE9.Text = ""
End Sub

Private Sub Clear_Locations()
txtCode10.Text = ""
txtName10.Text = ""
txtNameE10.Text = ""
End Sub

