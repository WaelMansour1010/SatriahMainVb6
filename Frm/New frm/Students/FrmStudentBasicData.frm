VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmStudentBasicData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18270
   Icon            =   "FrmStudentBasicData.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   18270
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9945
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18270
      _cx             =   32226
      _cy             =   17542
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
         Height          =   585
         Left            =   -105
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   19050
         _cx             =   33602
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
            ButtonImage     =   "FrmStudentBasicData.frx":038A
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
            ButtonImage     =   "FrmStudentBasicData.frx":0724
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
            ButtonImage     =   "FrmStudentBasicData.frx":0ABE
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
            ButtonImage     =   "FrmStudentBasicData.frx":0E58
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
         Height          =   9300
         Left            =   -30
         TabIndex        =   7
         Top             =   600
         Width           =   18480
         _cx             =   32597
         _cy             =   16404
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
         Caption         =   "»Ì«‰«  «”«”Ì…|»Ì«‰«  «”«”Ì…|«·„—›ﬁ« "
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   8880
            Left            =   45
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   45
            Width           =   18390
            _cx             =   32438
            _cy             =   15663
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   4575
               Left            =   6120
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   4320
               Width           =   6030
               _cx             =   10636
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
               Caption         =   "«·„Ê«œ «·œ—«”Ì…"
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
               Begin VB.TextBox txtName3 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1500
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   3165
                  Width           =   3120
               End
               Begin VB.TextBox txtCode3 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2880
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   2790
                  Width           =   1740
               End
               Begin VB.TextBox TxtNoHour 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3480
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   4200
                  Width           =   1140
               End
               Begin VB.TextBox txtNameE3 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1500
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   3480
                  Width           =   3120
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1500
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   4200
                  Width           =   1140
               End
               Begin VSFlex8Ctl.VSFlexGrid fg3 
                  Height          =   2355
                  Left            =   120
                  TabIndex        =   15
                  Top             =   360
                  Width           =   5760
                  _cx             =   10160
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
                  FormatString    =   $"FrmStudentBasicData.frx":11F2
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
                  Index           =   5
                  Left            =   480
                  TabIndex        =   16
                  Top             =   3240
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":1312
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   6
                  Left            =   480
                  TabIndex        =   17
                  Top             =   3960
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":16AC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   330
                  Index           =   2
                  Left            =   480
                  TabIndex        =   18
                  Top             =   3600
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":7F0E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcbCurs 
                  Bindings        =   "FrmStudentBasicData.frx":82A8
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   19
                  Top             =   3840
                  Width           =   3120
                  _ExtentX        =   5503
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   5
                  Left            =   4710
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   3480
                  Width           =   930
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   6
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   2820
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   8
                  Left            =   4710
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   3165
                  Width           =   930
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·”«⁄« "
                  Height          =   285
                  Index           =   13
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   4230
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   285
                  Index           =   14
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   4230
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì »⁄ "
                  Height          =   285
                  Index           =   15
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   3840
                  Width           =   960
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   4335
               Left            =   12240
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   0
               Width           =   6030
               _cx             =   10636
               _cy             =   7646
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
               Caption         =   "«‰Ê«⁄ «·⁄ﬁÊœ"
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
               Begin VB.TextBox txtNameE1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   3960
                  Width           =   3210
               End
               Begin VB.TextBox txtName1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   3645
                  Width           =   3210
               End
               Begin VB.TextBox txtCode1 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3120
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   3270
                  Width           =   1770
               End
               Begin VSFlex8Ctl.VSFlexGrid fg1 
                  Height          =   2715
                  Left            =   0
                  TabIndex        =   30
                  Top             =   360
                  Width           =   6000
                  _cx             =   10583
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmStudentBasicData.frx":82BD
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
                  Index           =   1
                  Left            =   240
                  TabIndex        =   31
                  Top             =   3120
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":8375
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   2
                  Left            =   240
                  TabIndex        =   32
                  Top             =   3840
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":870F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   330
                  Index           =   0
                  Left            =   240
                  TabIndex        =   33
                  Top             =   3480
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
                  ButtonImage     =   "FrmStudentBasicData.frx":EF71
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   11
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   3960
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   16
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   3285
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   17
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   3645
                  Width           =   990
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   4335
               Left            =   6120
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   0
               Width           =   6030
               _cx             =   10636
               _cy             =   7646
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
               Caption         =   "«‰Ê«⁄ «·„ƒÂ·«  «·œ—«”Ì…"
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
               Begin VB.TextBox txtCode2 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3060
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   3150
                  Width           =   1770
               End
               Begin VB.TextBox txtName2 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   3525
                  Width           =   3180
               End
               Begin VB.TextBox txtNameE2 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   3840
                  Width           =   3180
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg2 
                  Height          =   2715
                  Left            =   150
                  TabIndex        =   41
                  Top             =   360
                  Width           =   5850
                  _cx             =   10319
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmStudentBasicData.frx":F30B
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
                  Index           =   3
                  Left            =   150
                  TabIndex        =   42
                  Top             =   3120
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":F3C2
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   4
                  Left            =   150
                  TabIndex        =   43
                  Top             =   3720
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":F75C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   285
                  Index           =   1
                  Left            =   150
                  TabIndex        =   44
                  Top             =   3480
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmStudentBasicData.frx":15FBE
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   18
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   3525
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   19
                  Left            =   4770
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   3165
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   20
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   3840
                  Width           =   960
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   4335
               Left            =   0
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   0
               Width           =   6030
               _cx             =   10636
               _cy             =   7646
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
               Caption         =   " Œ’’«  «·„œ—»Ì‰"
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
               Begin VB.TextBox txtNameE4 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   3840
                  Width           =   3180
               End
               Begin VB.TextBox txtName4 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   3525
                  Width           =   3180
               End
               Begin VB.TextBox txtCode4 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3060
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   3150
                  Width           =   1770
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg4 
                  Height          =   2715
                  Left            =   150
                  TabIndex        =   52
                  Top             =   360
                  Width           =   5850
                  _cx             =   10319
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmStudentBasicData.frx":16358
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
                  Index           =   0
                  Left            =   120
                  TabIndex        =   53
                  Top             =   3120
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":1640F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   7
                  Left            =   150
                  TabIndex        =   54
                  Top             =   3720
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":167A9
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   285
                  Index           =   3
                  Left            =   150
                  TabIndex        =   55
                  Top             =   3480
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmStudentBasicData.frx":1D00B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   0
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   3840
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   1
                  Left            =   4770
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   3165
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   2
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   3525
                  Width           =   960
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   4575
               Left            =   12240
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   4320
               Width           =   6030
               _cx             =   10636
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
               Caption         =   "«‰Ê«⁄ «·œÊ—«  «·œ—«”Ì…"
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
               Begin VB.TextBox txtCode5 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3060
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   3150
                  Width           =   1770
               End
               Begin VB.TextBox txtName5 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   3525
                  Width           =   3180
               End
               Begin VB.TextBox txtNameE5 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   3840
                  Width           =   3180
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg5 
                  Height          =   2715
                  Left            =   150
                  TabIndex        =   63
                  Top             =   360
                  Width           =   5850
                  _cx             =   10319
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmStudentBasicData.frx":1D3A5
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
                  Index           =   8
                  Left            =   150
                  TabIndex        =   64
                  Top             =   3360
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":1D45C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   9
                  Left            =   150
                  TabIndex        =   65
                  Top             =   3960
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":1D7F6
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   285
                  Index           =   4
                  Left            =   150
                  TabIndex        =   66
                  Top             =   3720
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmStudentBasicData.frx":24058
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   3
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   3525
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   4
                  Left            =   4770
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   3165
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   7
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   3840
                  Width           =   960
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   4455
               Left            =   0
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   4320
               Width           =   6030
               _cx             =   10636
               _cy             =   7858
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
               Caption         =   "«·ﬁ«⁄«  «·œ—«”Ì…"
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
                  Height          =   315
                  Left            =   3060
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   3150
                  Width           =   1770
               End
               Begin VB.TextBox txtName6 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   3525
                  Width           =   3180
               End
               Begin VB.TextBox txtNameE6 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   3840
                  Width           =   3180
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg6 
                  Height          =   2715
                  Left            =   150
                  TabIndex        =   74
                  Top             =   360
                  Width           =   5850
                  _cx             =   10319
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmStudentBasicData.frx":243F2
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
                  Index           =   10
                  Left            =   150
                  TabIndex        =   75
                  Top             =   3240
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":244A9
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   11
                  Left            =   150
                  TabIndex        =   76
                  Top             =   3840
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":24843
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   285
                  Index           =   5
                  Left            =   150
                  TabIndex        =   77
                  Top             =   3600
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmStudentBasicData.frx":2B0A5
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   9
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   3525
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   10
                  Left            =   4770
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   3165
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   12
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   3840
                  Width           =   960
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   8880
            Left            =   19125
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   45
            Width           =   18390
            _cx             =   32438
            _cy             =   15663
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
               Height          =   4575
               Left            =   11760
               TabIndex        =   82
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
               Caption         =   "«‰Ê«⁄ «· œ—Ì»"
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
               Begin VB.TextBox TxtCode7 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3480
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   3270
                  Width           =   1770
               End
               Begin VB.TextBox TxtName7 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2040
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   3645
                  Width           =   3210
               End
               Begin VB.TextBox TxtNameE7 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2040
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   3960
                  Width           =   3210
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
                  Height          =   2715
                  Left            =   120
                  TabIndex        =   86
                  Top             =   360
                  Width           =   6360
                  _cx             =   11218
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmStudentBasicData.frx":2B43F
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
                  Index           =   14
                  Left            =   240
                  TabIndex        =   87
                  Top             =   3240
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":2B512
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   15
                  Left            =   240
                  TabIndex        =   88
                  Top             =   3960
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":2B8AC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   330
                  Index           =   7
                  Left            =   240
                  TabIndex        =   89
                  Top             =   3600
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
                  ButtonImage     =   "FrmStudentBasicData.frx":3210E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin XtremeSuiteControls.RadioButton TypeTrain 
                  Height          =   255
                  Index           =   0
                  Left            =   2625
                  TabIndex        =   104
                  Top             =   3240
                  Width           =   810
                  _Version        =   786432
                  _ExtentX        =   1429
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "›—œÌ"
                  BackColor       =   14737632
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton TypeTrain 
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   105
                  Top             =   3240
                  Width           =   1425
                  _Version        =   786432
                  _ExtentX        =   2514
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„‰ ÂÌ »«· ÊŸÌ›"
                  BackColor       =   14737632
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   29
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   3645
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   28
                  Left            =   5370
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   3285
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   27
                  Left            =   5370
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   3960
                  Width           =   990
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic14 
               Height          =   4575
               Left            =   12240
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   4680
               Visible         =   0   'False
               Width           =   6030
               _cx             =   10636
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
               Caption         =   "«‰Ê«⁄ «·œÊ—«  «·œ—«”Ì…"
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
               Begin VB.TextBox Text17 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   3840
                  Width           =   3180
               End
               Begin VB.TextBox Text16 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   3525
                  Width           =   3180
               End
               Begin VB.TextBox Text15 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3060
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   3150
                  Width           =   1770
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid5 
                  Height          =   2715
                  Left            =   150
                  TabIndex        =   97
                  Top             =   360
                  Width           =   5850
                  _cx             =   10319
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmStudentBasicData.frx":324A8
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
                  Index           =   20
                  Left            =   150
                  TabIndex        =   98
                  Top             =   3360
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":3255F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   150
                  TabIndex        =   99
                  Top             =   3960
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmStudentBasicData.frx":328F9
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   285
                  Index           =   10
                  Left            =   150
                  TabIndex        =   100
                  Top             =   3720
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmStudentBasicData.frx":3915B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   315
                  Index           =   38
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   3840
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ"
                  Height          =   285
                  Index           =   37
                  Left            =   4770
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   3165
                  Width           =   990
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   315
                  Index           =   36
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   3525
                  Width           =   960
               End
            End
         End
      End
   End
End
Attribute VB_Name = "FrmStudentBasicData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim Rs_StuConr As ADODB.Recordset
Dim Rs_StudentQuli As ADODB.Recordset
Dim Rs_STudentCurs As ADODB.Recordset
Dim Rs_StuTeach As ADODB.Recordset
Dim Rs_StuClRooms As ADODB.Recordset
Dim Rs_StTypeCurs As ADODB.Recordset
Dim Rs_TypeTrining As ADODB.Recordset


Private Sub btnModify_Click(Index As Integer)
Select Case Index
Case 0
Update_StudentContr
Case 1
Update_StudentQualification
Case 2
Update_StudentCurs
Case 3
Update_StudentTeachers
Case 4
Update_StudentTypeCurs
Case 5
Update_ClassRoom
Case 7
  If TypeTrain(0).value = False And TypeTrain(1).value = False Then
  If SystemOptions.UserInterface = ArabicInterface Then
  MsgBox "Ì—ÃÏ «Œ Ì«— «·‰Ê⁄"
  Else
  MsgBox "Please Select Type"
  End If
  Exit Sub
  End If
Update_TypeTrining
End Select
End Sub

Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
    Case 14
    Add_TypeTrining
    Case 15
        If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_TypetTrining
            
        Case 1
            Add_StudentContract
        Case 2

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_StudentContr
        Case 3
                Add_StudentQualification
        Case 4
             If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_StudentQualification
        Case 5
                Add_StudentCurs
        Case 6
                 If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_StudentCurs
      Case 0
      Add_StudentTeachers
     Case 7
     Del_StudentTeachers
     Case 8
     Add_StudentTypeCurs
    Case 9
     Del_StudentTypeCurs
    Case 10
     Add_StudentClassRooms
    Case 11
     Del_StudentClassRooms
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub Add_StudentClassRooms()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
   Dim msg As String
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
    Set Rs_StuClRooms = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblStudentClassRooms  "
    Rs_StuClRooms.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs_StuClRooms.AddNew
    txtCode6.Text = CStr(new_id("TblStudentClassRooms", "ID", "", True))
    Rs_StuClRooms("ID") = IIf(txtCode6.Text = "", Null, val(txtCode6.Text))
    Rs_StuClRooms("Name") = IIf(txtName6.Text = "", Null, txtName6.Text)
    Rs_StuClRooms("NameE") = IIf(txtNameE6.Text = "", Null, txtNameE6.Text)
    Rs_StuClRooms.update
    Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
    MsgBox "Save Successfully"
    End If
    Retrive_StudentClassRooms
    Clear_StudentClassRooms
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Add_StudentContract()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
   Dim msg As String
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
    Set Rs_StuConr = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblStudentContract  "
    Rs_StuConr.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs_StuConr.AddNew
    txtCode1.Text = CStr(new_id("TblStudentContract", "ID", "", True))
    Rs_StuConr("ID") = IIf(txtCode1.Text = "", Null, txtCode1.Text)
    Rs_StuConr("Name") = IIf(TxtName1.Text = "", Null, TxtName1.Text)
    Rs_StuConr("NameE") = IIf(txtNameE1.Text = "", Null, txtNameE1.Text)
    Rs_StuConr.update
    Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
    MsgBox "Save Successfully"
    End If
    Retrive_StudentContr
    Clear_StudenContrac
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Add_TypeTrining()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
   Dim msg As String
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
  If TypeTrain(0).value = False And TypeTrain(1).value = False Then
  If SystemOptions.UserInterface = ArabicInterface Then
  MsgBox "Ì—ÃÏ «Œ Ì«— «·‰Ê⁄"
  Else
  MsgBox "Please Select Type"
  End If
  Exit Sub
  End If
    Cn.BeginTrans
    BeginTrans = True
    Set Rs_TypeTrining = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblStudentTypeTrinng  "
    Rs_TypeTrining.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs_TypeTrining.AddNew
    TxtCode7.Text = CStr(new_id("TblStudentTypeTrinng", "ID", "", True))
    Rs_TypeTrining("ID") = IIf(TxtCode7.Text = "", Null, TxtCode7.Text)
    Rs_TypeTrining("Name") = IIf(TxtName7.Text = "", Null, TxtName7.Text)
    Rs_TypeTrining("NameE") = IIf(TxtNameE7.Text = "", Null, TxtNameE7.Text)
    If TypeTrain(0).value = True Then
    Rs_TypeTrining("typ") = 0
    ElseIf TypeTrain(1).value = True Then
    Rs_TypeTrining("typ") = 1
    End If
    Rs_TypeTrining.update
    Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
    MsgBox "Save Successfully"
    End If
    Retrive_TypeTrining
    Clear_TypeTrining
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Add_StudentTypeCurs()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
   Dim msg As String
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
    Set Rs_StTypeCurs = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblStudentTypeCurs  "
    Rs_StTypeCurs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs_StTypeCurs.AddNew
    txtCode5.Text = CStr(new_id("TblStudentTypeCurs", "ID", "", True))
    Rs_StTypeCurs("ID") = IIf(txtCode5.Text = "", Null, txtCode5.Text)
    Rs_StTypeCurs("Name") = IIf(txtName5.Text = "", Null, txtName5.Text)
    Rs_StTypeCurs("NameE") = IIf(txtNameE5.Text = "", Null, txtNameE5.Text)
    Rs_StTypeCurs.update
    Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
    MsgBox "Save Successfully"
    End If
    Retrive_StudentTypeCurs
    Clear_StudentTypeCurs
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Add_StudentTeachers()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
   Dim msg As String
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
    Set Rs_StuTeach = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblStudentTeachers  "
    Rs_StuTeach.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs_StuTeach.AddNew
    txtCode4.Text = CStr(new_id("TblStudentTeachers", "ID", "", True))
    Rs_StuTeach("ID") = IIf(txtCode4.Text = "", Null, txtCode4.Text)
    Rs_StuTeach("Name") = IIf(txtName4.Text = "", Null, txtName4.Text)
    Rs_StuTeach("NameE") = IIf(txtNameE4.Text = "", Null, txtNameE4.Text)
    Rs_StuTeach.update
    Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
    MsgBox "Save Successfully"
    End If
    Retrive_StudentTeachers
    Clear_StudentTeachers
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub



Private Sub Add_StudentQualification()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
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
    Set Rs_StudentQuli = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblStudentQualification  "
    Rs_StudentQuli.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs_StudentQuli.AddNew
    TxtCode2.Text = CStr(new_id("TblStudentQualification", "ID", "", True))
    Rs_StudentQuli("ID") = IIf(TxtCode2.Text = "", Null, TxtCode2.Text)
    Rs_StudentQuli("Name") = IIf(TxtName2.Text = "", Null, TxtName2.Text)
    Rs_StudentQuli("NameE") = IIf(txtNameE2.Text = "", Null, txtNameE2.Text)
    Rs_StudentQuli.update
    Cn.CommitTrans
    BeginTrans = False
   If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
   Else
   MsgBox "Save Successfully"
  End If
    Retrive_StudentQualification
    Clear_StudentQualification
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub


Private Sub Add_StudentCurs()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
  On Error GoTo errortrap
  
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
    
    Set Rs_STudentCurs = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblStudentCurs  "
    Rs_STudentCurs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Rs_STudentCurs.AddNew
    txtCode3.Text = CStr(new_id("TblStudentCurs", "ID", "", True))
    Rs_STudentCurs("ID") = IIf(txtCode3.Text = "", Null, val(txtCode3.Text))
    Rs_STudentCurs("Name") = IIf(TxtName3.Text = "", Null, TxtName3.Text)
    Rs_STudentCurs("NameE") = IIf(txtNameE3.Text = "", Null, txtNameE3.Text)
    Rs_STudentCurs("NoHour") = IIf(TxtNoHour.Text = "", Null, val(TxtNoHour.Text))
    Rs_STudentCurs("Price") = IIf(TxtPrice.Text = "", Null, val(TxtPrice.Text))
    Rs_STudentCurs("CursTypeID") = val(DcbCurs.BoundText)
    
    Rs_STudentCurs.update
    
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
  Else
   MsgBox "Save Successfully"
 End If
    Retrive_StudentCurs
    Clear_StudentCurs
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Update_StudentQualification()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
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
 
        str = Fg2.TextMatrix(Fg2.Row, Fg2.ColIndex("id"))
        sr = Fg2.TextMatrix(Fg2.Row, Fg2.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblStudentQualification Set  NameE='" & txtNameE2.Text & "',Name='" & TxtName2.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
  Else
   MsgBox "Save Successfully"
 End If
    Retrive_StudentContr
    Clear_StudenContrac
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  End If
End Sub
Private Sub Update_StudentContr()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
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
 
        str = Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("id"))
        sr = Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblStudentContract Set  NameE='" & txtNameE1.Text & "',Name='" & TxtName1.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
  Else
   MsgBox "Save Successfully"
 End If
    Retrive_StudentContr
    Clear_StudenContrac
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
End Sub
''////
Private Sub Update_TypeTrining()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
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
 
        str = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("id"))
        sr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
    If TypeTrain(0).value = True Then
     StrSQL = "Update TblStudentTypeTrinng Set typ=0, NameE='" & TxtNameE7.Text & "',Name='" & TxtName7.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
         Else
         StrSQL = "Update TblStudentTypeTrinng Set typ=1, NameE='" & TxtNameE7.Text & "',Name='" & TxtName7.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
         End If
           
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
  Else
   MsgBox "Save Successfully"
 End If
    Retrive_TypeTrining
    Clear_TypeTrining
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
End Sub
Private Sub Update_ClassRoom()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
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
 
        str = Fg6.TextMatrix(Fg6.Row, Fg6.ColIndex("id"))
        sr = Fg6.TextMatrix(Fg6.Row, Fg6.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblStudentClassRooms Set  NameE='" & txtNameE6.Text & "',Name='" & txtName6.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
  Else
   MsgBox "Save Successfully"
 End If
    Retrive_StudentClassRooms
    Clear_StudentClassRooms
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
End Sub

Private Sub Update_StudentTeachers()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
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
 
        str = Fg4.TextMatrix(Fg4.Row, Fg4.ColIndex("id"))
        sr = Fg4.TextMatrix(Fg4.Row, Fg4.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblStudentTeachers Set  NameE='" & txtNameE4.Text & "',Name='" & txtName4.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
  Else
   MsgBox "Save Successfully"
 End If
    Retrive_StudentTeachers
    Clear_StudentTeachers
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  End If
End Sub
Private Sub Update_StudentTypeCurs()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
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
     StrSQL = "Update TblStudentTypeCurs Set  NameE='" & txtNameE5.Text & "',Name='" & txtName5.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
  Else
   MsgBox "Save Successfully"
 End If
    Retrive_StudentTypeCurs
    Clear_StudentTypeCurs
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  End If
End Sub
Private Sub Update_StudentClassRooms()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
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
 
        str = Fg6.TextMatrix(Fg6.Row, Fg6.ColIndex("id"))
        sr = Fg6.TextMatrix(Fg6.Row, Fg6.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblStudentClassRooms Set  NameE='" & txtNameE6.Text & "',Name='" & txtName6.Text & "'  Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
                     If SystemOptions.UserInterface = ArabicInterface Then
                       MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
                     Else
                      MsgBox "Save Successfully"
                    End If
    Retrive_StudentClassRooms
    Clear_StudentClassRooms
    End If
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

                If Err.Number = -2147217900 Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
                                msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
                                msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
                               Else
                               End If
                    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Update_StudentCurs()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If TxtName3.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·«”„ «Ê·«")
 Else
        MsgBox ("Enter Name ")
 End If
 TxtName3.SetFocus
 Exit Sub
 End If
        str = fg3.TextMatrix(fg3.Row, fg3.ColIndex("id"))
        sr = fg3.TextMatrix(fg3.Row, fg3.ColIndex("serial"))
        If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
     StrSQL = "Update TblStudentCurs Set  NameE='" & txtNameE3.Text & "',Name='" & TxtName3.Text & "',NoHour=" & val(TxtNoHour.Text) & " ,Price=" & val(TxtPrice.Text) & ",CursTypeID=" & val(Me.DcbCurs.BoundText) & "   Where ID=" & val(str)
           Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
  Else
   MsgBox "Save Successfully"
 End If
    Retrive_StudentCurs
    Clear_StudentCurs
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
       End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
 msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
 Else
 msg = "Sorry error douring Save " & Chr(13)
End If
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
End Sub
Private Sub Del_StudentContr()
    Dim msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("id"))
        sr = Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("serial"))
        
        If str <> "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & Chr(13)
        msg = msg + (sr) & Chr(13)
        msg = msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
msg = "Confirm Delete"
End If
        If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not Rs_StuConr.RecordCount < 1 Then
                StrSQL = "delete From TblStudentContract  where  ID =" & val(str)
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblStudentContract"
                   Set Rs_StuConr = New ADODB.Recordset
                   Rs_StuConr.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If Rs_StuConr.RecordCount < 1 Then
                Else
                   Retrive_StudentContr
                End If
            End If
        End If

    Else
       If SystemOptions.UserInterface = ArabicInterface Then
        msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        msg = "This process is not available to the lack of records "
        End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_StudentContr
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Else
     msg = "You can not delete this record " & Chr(13) & "There are linked to this record data "
    End If
    msg = msg & Chr(13) & Err.description
    MsgBox msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub Del_StudentQualification()
    Dim msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = Fg2.TextMatrix(Fg2.Row, Fg2.ColIndex("id"))
        sr = Fg2.TextMatrix(Fg2.Row, Fg2.ColIndex("serial"))
        
        If str <> "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & Chr(13)
        msg = msg + (sr) & Chr(13)
        msg = msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
msg = "Confirm Delete"
End If
        If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs_StudentQuli.RecordCount < 1 Then
                StrSQL = "delete From TblStudentQualification  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblStudentQualification"
                   Set Rs_StudentQuli = New ADODB.Recordset
                   Rs_StudentQuli.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If Rs_StudentQuli.RecordCount < 1 Then
                Else
                   Retrive_StudentQualification
                End If
            End If
        End If

    Else
       If SystemOptions.UserInterface = ArabicInterface Then
        msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        msg = "This process is not available to the lack of records "
        End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_StudentQualification
    
    Exit Sub
ErrTrap:
   If SystemOptions.UserInterface = ArabicInterface Then
    msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Else
     msg = "You can not delete this record " & Chr(13) & "There are linked to this record data "
    End If
    msg = msg & Chr(13) & Err.description
    MsgBox msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    Rs_StudentQuli.CancelUpdate
    'End If
End Sub
Private Sub Del_TypetTrining()
    Dim msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("id"))
        sr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("serial"))
        
        If str <> "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & Chr(13)
        msg = msg + (sr) & Chr(13)
        msg = msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
msg = "Confirm Delete"
End If
        If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs_TypeTrining.RecordCount < 1 Then
                StrSQL = "delete From TblStudentTypeTrinng  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblStudentTypeTrinng"
                   Set Rs_TypeTrining = New ADODB.Recordset
                   Rs_TypeTrining.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If Rs_TypeTrining.RecordCount < 1 Then
                Else
                   Retrive_TypeTrining
                End If
            End If
        End If

    Else
       If SystemOptions.UserInterface = ArabicInterface Then
        msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        msg = "This process is not available to the lack of records "
        End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_TypeTrining
    
    Exit Sub
ErrTrap:
   If SystemOptions.UserInterface = ArabicInterface Then
    msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Else
     msg = "You can not delete this record " & Chr(13) & "There are linked to this record data "
    End If
    msg = msg & Chr(13) & Err.description
    MsgBox msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    Rs_TypeTrining.CancelUpdate
    'End If
End Sub
Private Sub Del_StudentCurs()
    Dim msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = fg3.TextMatrix(fg3.Row, fg3.ColIndex("id"))
        sr = fg3.TextMatrix(fg3.Row, fg3.ColIndex("serial"))
        
        If str <> "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & Chr(13)
        msg = msg + (sr) & Chr(13)
        msg = msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
msg = "Confirm Delete"
End If
        If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs_STudentCurs.RecordCount < 1 Then
                StrSQL = "delete From TblStudentCurs  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblStudentCurs"
                   Set Rs_STudentCurs = New ADODB.Recordset
                   Rs_STudentCurs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If Rs_STudentCurs.RecordCount < 1 Then
            
                Else
                   Retrive_StudentCurs
                End If
            End If
        End If

    Else
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        msg = "This process is not available to the lack of records "
        End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_StudentCurs

    Exit Sub
ErrTrap:
  If SystemOptions.UserInterface = ArabicInterface Then
    msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Else
     msg = "You can not delete this record " & Chr(13) & "There are linked to this record data "
    End If
    msg = msg & Chr(13) & Err.description
    MsgBox msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    Rs_STudentCurs.CancelUpdate
    'End If
End Sub
Private Sub Del_StudentTeachers()
    Dim msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = Fg4.TextMatrix(Fg4.Row, Fg4.ColIndex("id"))
        sr = Fg4.TextMatrix(Fg4.Row, Fg4.ColIndex("serial"))
        
        If str <> "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & Chr(13)
        msg = msg + (sr) & Chr(13)
        msg = msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
msg = "Confirm Delete"
End If
        If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs_StuTeach.RecordCount < 1 Then
                StrSQL = "delete From TblStudentTeachers  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblStudentTeachers"
                   Set Rs_StuTeach = New ADODB.Recordset
                   Rs_StuTeach.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If Rs_StuTeach.RecordCount < 1 Then
            
                Else
                   Retrive_StudentTeachers
                End If
            End If
        End If

    Else
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        msg = "This process is not available to the lack of records "
        End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_StudentTeachers

    Exit Sub
ErrTrap:
  If SystemOptions.UserInterface = ArabicInterface Then
    msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Else
     msg = "You can not delete this record " & Chr(13) & "There are linked to this record data "
    End If
    msg = msg & Chr(13) & Err.description
    MsgBox msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    Rs_StuTeach.CancelUpdate
End Sub
Private Sub Del_StudentTypeCurs()
    Dim msg As String
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
 If SystemOptions.UserInterface = ArabicInterface Then
        msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & Chr(13)
        msg = msg + (sr) & Chr(13)
        msg = msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
msg = "Confirm Delete"
End If
        If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs_StTypeCurs.RecordCount < 1 Then
                StrSQL = "delete From TblStudentTypeCurs  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblStudentTypeCurs"
                   Set Rs_StTypeCurs = New ADODB.Recordset
                   Rs_StTypeCurs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If Rs_StTypeCurs.RecordCount < 1 Then
            
                Else
                   Retrive_StudentTypeCurs
                End If
            End If
        End If

    Else
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        msg = "This process is not available to the lack of records "
        End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_StudentTypeCurs

    Exit Sub
ErrTrap:
  If SystemOptions.UserInterface = ArabicInterface Then
    msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Else
     msg = "You can not delete this record " & Chr(13) & "There are linked to this record data "
    End If
    msg = msg & Chr(13) & Err.description
    MsgBox msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    Rs_StTypeCurs.CancelUpdate
End Sub

Private Sub Del_StudentClassRooms()
    Dim msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = Fg6.TextMatrix(Fg6.Row, Fg6.ColIndex("id"))
        sr = Fg6.TextMatrix(Fg6.Row, Fg6.ColIndex("serial"))
        
        If str <> "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & Chr(13)
        msg = msg + (sr) & Chr(13)
        msg = msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
msg = "Confirm Delete"
End If
        If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs_StuClRooms.RecordCount < 1 Then
                StrSQL = "delete From TblStudentClassRooms  where  ID =" & val(str)
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblStudentClassRooms"
                   Set Rs_StuClRooms = New ADODB.Recordset
                   Rs_StuClRooms.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If Rs_StuClRooms.RecordCount < 1 Then
            
                Else
                   Retrive_StudentTypeCurs
                End If
            End If
        End If

    Else
    If SystemOptions.UserInterface = ArabicInterface Then
        msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        msg = "This process is not available to the lack of records "
        End If
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_StudentClassRooms
    Exit Sub
ErrTrap:
  If SystemOptions.UserInterface = ArabicInterface Then
    msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Else
     msg = "You can not delete this record " & Chr(13) & "There are linked to this record data "
    End If
    msg = msg & Chr(13) & Err.description
    MsgBox msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    Rs_StuClRooms.CancelUpdate
End Sub
Private Sub Retrive_StudentTeachers()
Dim i As Integer
     Set Rs_StuTeach = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblStudentTeachers  "
    Rs_StuTeach.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Fg4.Rows = 1
    If Rs_StuTeach.RecordCount > 0 Then
        Rs_StuTeach.MoveFirst
        With Fg4
        .Rows = Rs_StuTeach.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_StuTeach("id").value), "", Rs_StuTeach("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs_StuTeach("name").value), "", Rs_StuTeach("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs_StuTeach("namee").value), "", Rs_StuTeach("namee").value)
          Rs_StuTeach.MoveNext
         Next
         End With
    End If
End Sub
Private Sub Retrive_StudentTypeCurs()
Dim i As Integer
     Set Rs_StTypeCurs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblStudentTypeCurs  "
    Rs_StTypeCurs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Fg5.Rows = 1
    If Rs_StTypeCurs.RecordCount > 0 Then
        Rs_StTypeCurs.MoveFirst
        With Fg5
        .Rows = Rs_StTypeCurs.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_StTypeCurs("id").value), "", Rs_StTypeCurs("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs_StTypeCurs("name").value), "", Rs_StTypeCurs("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs_StTypeCurs("namee").value), "", Rs_StTypeCurs("namee").value)
          Rs_StTypeCurs.MoveNext
         Next
         End With
    End If
End Sub
Private Sub Retrive_StudentClassRooms()
Dim i As Integer
     Set Rs_StuClRooms = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblStudentClassRooms  "
    Rs_StuClRooms.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Fg6.Rows = 1
    If Rs_StuClRooms.RecordCount > 0 Then
        Rs_StuClRooms.MoveFirst
        With Fg6
        .Rows = Rs_StuClRooms.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_StuClRooms("id").value), "", Rs_StuClRooms("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs_StuClRooms("name").value), "", Rs_StuClRooms("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs_StuClRooms("namee").value), "", Rs_StuClRooms("namee").value)
          Rs_StuClRooms.MoveNext
         Next
         End With
    End If
End Sub
Private Sub Retrive_TypeTrining()
Dim i As Integer
     Set Rs_TypeTrining = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblStudentTypeTrinng  "
    Rs_TypeTrining.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    VSFlexGrid2.Rows = 1
    If Rs_TypeTrining.RecordCount > 0 Then
        Rs_TypeTrining.MoveFirst
        With VSFlexGrid2
        .Rows = Rs_TypeTrining.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("Typ")) = IIf(IsNull(Rs_TypeTrining("Typ").value), "", Rs_TypeTrining("Typ").value)
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_TypeTrining("id").value), "", Rs_TypeTrining("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs_TypeTrining("name").value), "", Rs_TypeTrining("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs_TypeTrining("namee").value), "", Rs_TypeTrining("namee").value)
          Rs_TypeTrining.MoveNext
         Next
         End With
    End If
End Sub

Private Sub Retrive_StudentContr()
Dim i As Integer
     Set Rs_StuConr = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblStudentContract  "
    Rs_StuConr.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Fg1.Rows = 1
    If Rs_StuConr.RecordCount > 0 Then
        Rs_StuConr.MoveFirst
        With Fg1
        .Rows = Rs_StuConr.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_StuConr("id").value), "", Rs_StuConr("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs_StuConr("name").value), "", Rs_StuConr("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs_StuConr("namee").value), "", Rs_StuConr("namee").value)
          Rs_StuConr.MoveNext
         Next
         End With
    End If
End Sub


Private Sub Retrive_StudentQualification()
Dim i As Integer
     Set Rs_StudentQuli = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblStudentQualification  "
    Rs_StudentQuli.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Fg2.Rows = 1
    If Rs_StudentQuli.RecordCount > 0 Then
        Rs_StudentQuli.MoveFirst
        With Fg2
        .Rows = Rs_StudentQuli.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_StudentQuli("id").value), "", Rs_StudentQuli("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs_StudentQuli("name").value), "", Rs_StudentQuli("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs_StudentQuli("namee").value), "", Rs_StudentQuli("namee").value)
          Rs_StudentQuli.MoveNext
         Next
         End With
    End If
End Sub

Private Sub Retrive_StudentCurs()
Dim i As Integer
     Set Rs_STudentCurs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblStudentCurs  "
    Rs_STudentCurs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    fg3.Rows = 1
    If Rs_STudentCurs.RecordCount > 0 Then
        Rs_STudentCurs.MoveFirst
        With fg3
        .Rows = Rs_STudentCurs.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_STudentCurs("id").value), "", Rs_STudentCurs("id").value)
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs_STudentCurs("name").value), "", Rs_STudentCurs("name").value)
         .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs_STudentCurs("namee").value), "", Rs_STudentCurs("namee").value)
         .TextMatrix(i, .ColIndex("NoHour")) = IIf(IsNull(Rs_STudentCurs("NoHour").value), "", Rs_STudentCurs("NoHour").value)
         .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rs_STudentCurs("Price").value), "", Rs_STudentCurs("Price").value)
         .TextMatrix(i, .ColIndex("CursTypeID")) = IIf(IsNull(Rs_STudentCurs("CursTypeID").value), "", Rs_STudentCurs("CursTypeID").value)
          Rs_STudentCurs.MoveNext
         Next
         End With
    End If
End Sub

Private Sub Clear_StudentTypeCurs()
txtCode5.Text = ""
txtName5.Text = ""
txtNameE5.Text = ""
End Sub
Private Sub Clear_StudentClassRooms()
txtCode6.Text = ""
txtName6.Text = ""
txtNameE6.Text = ""
End Sub
Private Sub Clear_TypeTrining()
TxtCode7.Text = ""
TxtName7.Text = ""
TxtNameE7.Text = ""
End Sub

Private Sub Clear_StudenContrac()
txtCode1.Text = ""
TxtName1.Text = ""
txtNameE1.Text = ""
End Sub
Private Sub Clear_StudentTeachers()
txtCode4.Text = ""
txtName4.Text = ""
txtNameE4.Text = ""
End Sub
Private Sub Clear_StudentQualification()
TxtCode2.Text = ""
TxtName2.Text = ""
txtNameE2.Text = ""
End Sub

Private Sub Clear_StudentCurs()
txtCode3.Text = ""
TxtName3.Text = ""
txtNameE3.Text = ""
Me.TxtPrice.Text = 0
Me.TxtNoHour.Text = 0
Me.DcbCurs.BoundText = 0
End Sub

Private Sub DcbCurs_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyF3 Then
Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
   Dcombos.GetStudentTypeCurs DcbCurs
End If
End Sub

Private Sub Fg1_Click()
With Me.Fg1
If .Row > 0 Then
txtCode1.Text = .TextMatrix(.Row, .ColIndex("id"))
TxtName1.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE1.Text = .TextMatrix(.Row, .ColIndex("NameE"))
End If
End With
End Sub



Private Sub fg2_Click()
With Me.Fg2
If .Row > 0 Then
TxtCode2.Text = .TextMatrix(.Row, .ColIndex("id"))
TxtName2.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE2.Text = .TextMatrix(.Row, .ColIndex("NameE"))
End If
End With

End Sub

Private Sub fg3_Click()
With Me.fg3
If .Row > 0 Then
txtCode3.Text = .TextMatrix(.Row, .ColIndex("id"))
TxtName3.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE3.Text = .TextMatrix(.Row, .ColIndex("NameE"))
TxtNoHour.Text = val(.TextMatrix(.Row, .ColIndex("NoHour")))
TxtPrice.Text = val(.TextMatrix(.Row, .ColIndex("Price")))
DcbCurs.BoundText = val(.TextMatrix(.Row, .ColIndex("CursTypeID")))
End If
End With
End Sub





Private Sub fg4_Click()
With Me.Fg4
If .Row > 0 Then
txtCode4.Text = .TextMatrix(.Row, .ColIndex("id"))
txtName4.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE4.Text = .TextMatrix(.Row, .ColIndex("NameE"))
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

Private Sub Fg6_Click()
With Me.Fg6
If .Row > 0 Then
txtCode6.Text = .TextMatrix(.Row, .ColIndex("id"))
txtName6.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.txtNameE6.Text = .TextMatrix(.Row, .ColIndex("NameE"))
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
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «·»Ì«‰«  «·«”«”Ì… ··ÿ·«»  "
    LogTextE = " Open Window " & "  Violation Types "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "O", "", ""

    Dim My_SQL As String
    Resize_Form Me
    
   Retrive_StudentContr
   Retrive_StudentQualification
   Retrive_StudentCurs
   Retrive_StudentTypeCurs
   Retrive_StudentTeachers
   Retrive_StudentClassRooms
   Retrive_TypeTrining
    Exit Sub

ErrTrap:
End Sub
'
Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
     Me.Caption = "Data"
     lbl(29).Caption = "Name Ar"
     lbl(28).Caption = "Code"
     lbl(27).Caption = "Name Eng"
     Cmd(14).Caption = "Add"
     Cmd(15).Caption = "Delete"
     TypeTrain(0).RightToLeft = False
TypeTrain(1).RightToLeft = False
TypeTrain(0).Caption = "Personal"
TypeTrain(1).Caption = "Employment"
    btnModify(7).Caption = "Update"
     C1Elastic11.Caption = "Type Trining"
  lbl(17).Caption = "Name Ar"
  lbl(18).Caption = "Name Ar"
  lbl(2).Caption = "Name Ar"
  lbl(3).Caption = "Name Ar"
  lbl(8).Caption = "Name Ar"
  lbl(9).Caption = "Name Ar"
   lbl(11).Caption = "Name Eng"
   lbl(20).Caption = "Name Eng"
   lbl(0).Caption = "Name Eng"
   lbl(7).Caption = "Name Eng"
   lbl(5).Caption = "Name Eng"
   lbl(12).Caption = "Name Eng"
   lbl(16).Caption = "Code"
   lbl(19).Caption = "Code"
   lbl(1).Caption = "Code"
   lbl(4).Caption = "Code"
   lbl(6).Caption = "Code"
   lbl(10).Caption = "Code"
Cmd(1).Caption = "Add"
Cmd(3).Caption = "Add"
Cmd(0).Caption = "Add"
Cmd(8).Caption = "Add"
Cmd(5).Caption = "Add"
Cmd(10).Caption = "Add"
Cmd(4).Caption = "Delete"
Cmd(7).Caption = "Delete"
Cmd(11).Caption = "Delete"
Cmd(6).Caption = "Delete"
Cmd(2).Caption = "Delete"
Cmd(9).Caption = "Delete"
C1Elastic6.Caption = "Types of Contracts"
C1Elastic2.Caption = "Specialties Instructors"
C1Elastic7.Caption = "Type Qualifications"
C1Elastic5.Caption = "ClassRooms"
C1Elastic4.Caption = "Subjects"
C1Elastic3.Caption = "Type of Subjects"
lbl(15).Caption = "Type Curs"
lbl(14).Caption = "Value"
lbl(13).Caption = "No.Hours"

  btnModify(0).Caption = "Update"
  btnModify(1).Caption = "Update"
  btnModify(2).Caption = "Update"
  btnModify(3).Caption = "Update"
  btnModify(4).Caption = "Update"
  btnModify(5).Caption = "Update"
    Me.Caption = "  Baisc Data"
    C1Tab1.TabCaption(0) = "Baisc Data"
    C1Tab1.TabCaption(1) = "Baisc Data"
    EleHeader.Caption = Me.Caption
     With VSFlexGrid2
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
  
  With Fg5
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With Fg6
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With fg3
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With Fg1
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With Fg2
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With
    With Fg4
  .TextMatrix(0, .ColIndex("id")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Name Ar"
  .TextMatrix(0, .ColIndex("NameE")) = "Name Eng"
  End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «·ÿ·«» «·«”«”Ì…  "
    LogTextE = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "O", "", ""

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

Private Sub txtName1_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtName2_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtName3_Change()
Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
   Dcombos.GetStudentTypeCurs DcbCurs
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

Private Sub TxtName7_GotFocus()

SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE1_GotFocus()
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

Private Sub TxtNameE7_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtNoHour_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNoHour.Text, 0)
End Sub

Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPrice.Text, 0)
End Sub

Private Sub VSFlexGrid2_Click()
With Me.VSFlexGrid2
If .Row > 0 Then
TxtCode7.Text = .TextMatrix(.Row, .ColIndex("id"))
TxtName7.Text = .TextMatrix(.Row, .ColIndex("Name"))
Me.TxtNameE7.Text = .TextMatrix(.Row, .ColIndex("NameE"))
TypeTrain(val(.TextMatrix(.Row, .ColIndex("typ")))).value = True
End If
End With
End Sub
