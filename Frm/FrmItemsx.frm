VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E1BFA30F-D929-4F80-AEDD-76FC2BDF5E23}#1.0#0"; "ciaXPPopUp30.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmItems 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·√’‰«›"
   ClientHeight    =   8550
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   16020
   HelpContextID   =   210
   Icon            =   "FrmItems.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   16020
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
      Height          =   8550
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   16020
      _cx             =   28258
      _cy             =   15081
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
      Begin C1SizerLibCtl.C1Elastic EleFooter 
         Height          =   600
         Left            =   15
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   7875
         Width           =   15990
         _cx             =   28205
         _cy             =   1058
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
            Height          =   375
            Index           =   0
            Left            =   14205
            TabIndex        =   16
            Top             =   90
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   12735
            TabIndex        =   17
            Top             =   90
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   11415
            TabIndex        =   18
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   3
            Left            =   9855
            TabIndex        =   19
            Top             =   90
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   4
            Left            =   8310
            TabIndex        =   20
            Top             =   90
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   5
            Left            =   6765
            TabIndex        =   21
            Top             =   90
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   6
            Left            =   75
            TabIndex        =   22
            Top             =   90
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   7
            Left            =   5250
            TabIndex        =   23
            Top             =   90
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   375
            Left            =   1785
            TabIndex        =   24
            Top             =   90
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   30
            Left            =   3360
            TabIndex        =   297
            Top             =   90
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
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
      End
      Begin C1SizerLibCtl.C1Elastic EleMiddle 
         Height          =   7185
         Left            =   15
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   675
         Width           =   15870
         _cx             =   27993
         _cy             =   12674
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
         BorderWidth     =   1
         ChildSpacing    =   2
         Splitter        =   -1  'True
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
            Height          =   7035
            Left            =   15
            TabIndex        =   26
            Top             =   15
            Width           =   12390
            _cx             =   21855
            _cy             =   12409
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
            FrontTabForeColor=   -2147483630
            Caption         =   $"FrmItems.frx":038A
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
            Picture(0)      =   "FrmItems.frx":0425
            Picture(1)      =   "FrmItems.frx":07BF
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   6570
               Left            =   15735
               TabIndex        =   306
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
                  Height          =   5970
                  Left            =   0
                  TabIndex        =   307
                  Top             =   120
                  Width           =   12060
                  _cx             =   21272
                  _cy             =   10530
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
                  BackColorBkg    =   16777215
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmItems.frx":0B59
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
                  ExplorerBar     =   5
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
               Height          =   6570
               Index           =   2
               Left            =   13035
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
               Begin VB.TextBox txtTypenew 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   6120
                  RightToLeft     =   -1  'True
                  TabIndex        =   316
                  Top             =   75
                  Width           =   780
               End
               Begin VB.TextBox TxtSource 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   8280
                  RightToLeft     =   -1  'True
                  TabIndex        =   314
                  Top             =   75
                  Width           =   780
               End
               Begin VB.TextBox txtDippre 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   312
                  Top             =   75
                  Width           =   1020
               End
               Begin VB.TextBox txtContent 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   310
                  Top             =   75
                  Width           =   1020
               End
               Begin VB.TextBox TxtWight 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   308
                  Top             =   75
                  Width           =   780
               End
               Begin VB.TextBox TxtOverHead 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   12900
                  TabIndex        =   231
                  Top             =   -165
                  Visible         =   0   'False
                  Width           =   540
               End
               Begin VB.CheckBox ChkAssplied 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’‰› „Ã„⁄"
                  Height          =   225
                  Left            =   10410
                  TabIndex        =   73
                  Top             =   330
                  Width           =   1365
               End
               Begin VB.CheckBox chkItemMaking 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ ‰”» «·«’‰«› ··«‰ «Ã"
                  Height          =   225
                  Left            =   -690
                  TabIndex        =   57
                  Top             =   330
                  Visible         =   0   'False
                  Width           =   2310
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   2070
                  Index           =   7
                  Left            =   0
                  TabIndex        =   55
                  TabStop         =   0   'False
                  Top             =   3285
                  Width           =   11895
                  _cx             =   20981
                  _cy             =   3651
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
                  Begin VSFlex8UCtl.VSFlexGrid FgAttachs 
                     Height          =   1335
                     Left            =   0
                     TabIndex        =   59
                     Top             =   15
                     Width           =   11895
                     _cx             =   20981
                     _cy             =   2355
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
                     Cols            =   6
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmItems.frx":0C3D
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
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   825
                     Index           =   5
                     Left            =   -540
                     TabIndex        =   195
                     TabStop         =   0   'False
                     Top             =   1455
                     Width           =   12975
                     _cx             =   22886
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
                     Begin VB.TextBox TxtAttachedItemCode 
                        Alignment       =   1  'Right Justify
                        Height          =   360
                        Left            =   10920
                        TabIndex        =   198
                        Top             =   300
                        Width           =   1515
                     End
                     Begin VB.TextBox TxtItemQty 
                        Alignment       =   1  'Right Justify
                        Height          =   360
                        Index           =   1
                        Left            =   4785
                        MaxLength       =   5
                        TabIndex        =   197
                        Top             =   300
                        Width           =   945
                     End
                     Begin VB.TextBox TxtItemPrice 
                        Alignment       =   1  'Right Justify
                        Height          =   360
                        Index           =   1
                        Left            =   2865
                        MaxLength       =   5
                        TabIndex        =   196
                        Top             =   300
                        Width           =   1785
                     End
                     Begin MSDataListLib.DataCombo DcboItemID1 
                        Height          =   315
                        Left            =   7245
                        TabIndex        =   199
                        Top             =   300
                        Width           =   3540
                        _ExtentX        =   6244
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   315
                        Index           =   10
                        Left            =   1905
                        TabIndex        =   200
                        Top             =   285
                        Width           =   825
                        _ExtentX        =   1455
                        _ExtentY        =   556
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
                        ColorButton     =   14871017
                     End
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   315
                        Index           =   11
                        Left            =   540
                        TabIndex        =   201
                        Top             =   285
                        Width           =   960
                        _ExtentX        =   1693
                        _ExtentY        =   556
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
                     End
                     Begin MSDataListLib.DataCombo DataCombo6 
                        Height          =   315
                        Left            =   5730
                        TabIndex        =   259
                        Top             =   300
                        Visible         =   0   'False
                        Width           =   1515
                        _ExtentX        =   2672
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·ÊÕœ…"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H000000C0&
                        Height          =   345
                        Index           =   49
                        Left            =   6285
                        TabIndex        =   260
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   810
                     End
                     Begin VB.Label lblLabel2 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«”„ «·’‰›"
                        ForeColor       =   &H000000C0&
                        Height          =   240
                        Left            =   8475
                        TabIndex        =   207
                        Top             =   0
                        Width           =   2040
                     End
                     Begin VB.Label lblLabel1 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "ﬂÊœ «·’‰›"
                        ForeColor       =   &H000000C0&
                        Height          =   240
                        Left            =   11070
                        TabIndex        =   206
                        Top             =   0
                        Width           =   1080
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "0"
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
                        Height          =   240
                        Index           =   28
                        Left            =   690
                        TabIndex        =   205
                        ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
                        Top             =   45
                        Width           =   135
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "⁄œœ «·√’‰«›"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H000000C0&
                        Height          =   240
                        Index           =   27
                        Left            =   1635
                        TabIndex        =   204
                        ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
                        Top             =   45
                        Width           =   960
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·ﬂ„Ì…"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H000000C0&
                        Height          =   255
                        Index           =   25
                        Left            =   4650
                        TabIndex        =   203
                        Top             =   45
                        Width           =   1080
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«· ﬂ·›…"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H000000C0&
                        Height          =   240
                        Index           =   26
                        Left            =   2460
                        TabIndex        =   202
                        Top             =   45
                        Width           =   1500
                     End
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·’‰›"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   795
                     Index           =   24
                     Left            =   0
                     TabIndex        =   58
                     Top             =   15
                     Width           =   13665
                  End
               End
               Begin VB.CheckBox ChkRelated 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·Â ’‰› „·Õﬁ"
                  Height          =   240
                  Left            =   10380
                  TabIndex        =   54
                  Top             =   3045
                  Width           =   1515
               End
               Begin VB.CheckBox ChkItemMakingNew 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’‰› Ì „ «‰ «Ã…"
                  Height          =   225
                  Left            =   10380
                  TabIndex        =   12
                  Top             =   90
                  Width           =   1380
               End
               Begin VB.TextBox TxtItemComment 
                  Alignment       =   1  'Right Justify
                  Height          =   750
                  Left            =   135
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   13
                  Top             =   5760
                  Width           =   11895
               End
               Begin VSFlex8UCtl.VSFlexGrid FG 
                  Height          =   1740
                  Left            =   0
                  TabIndex        =   71
                  Top             =   585
                  Width           =   11895
                  _cx             =   20981
                  _cy             =   3069
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
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmItems.frx":0D21
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   720
                  Index           =   3
                  Left            =   0
                  TabIndex        =   78
                  TabStop         =   0   'False
                  Top             =   2310
                  Width           =   12030
                  _cx             =   21220
                  _cy             =   1270
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
                  Begin VB.TextBox TxtItemPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Index           =   0
                     Left            =   2865
                     MaxLength       =   5
                     TabIndex        =   81
                     Top             =   375
                     Width           =   1515
                  End
                  Begin VB.TextBox TxtItemQty 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Index           =   0
                     Left            =   4380
                     MaxLength       =   5
                     TabIndex        =   80
                     Top             =   375
                     Width           =   945
                  End
                  Begin VB.TextBox TxtItemCode 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   9840
                     TabIndex        =   79
                     Top             =   375
                     Width           =   1785
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   300
                     Index           =   8
                     Left            =   1500
                     TabIndex        =   82
                     Top             =   285
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   529
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
                     ColorButton     =   14871017
                  End
                  Begin MSDataListLib.DataCombo DcboItems 
                     Height          =   315
                     Left            =   6555
                     TabIndex        =   83
                     Top             =   375
                     Width           =   3285
                     _ExtentX        =   5794
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   300
                     Index           =   9
                     Left            =   135
                     TabIndex        =   84
                     Top             =   285
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   529
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
                  End
                  Begin MSDataListLib.DataCombo dcItemunit 
                     Height          =   315
                     Left            =   5325
                     TabIndex        =   102
                     Top             =   405
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·ÊÕœ…"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   375
                     Index           =   36
                     Left            =   5610
                     TabIndex        =   103
                     Top             =   0
                     Width           =   675
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· ﬂ·›…"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   315
                     Index           =   17
                     Left            =   2730
                     TabIndex        =   90
                     Top             =   45
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬂ„Ì…"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   390
                     Index           =   18
                     Left            =   4380
                     TabIndex        =   89
                     Top             =   45
                     Width           =   675
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·’‰›"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   315
                     Index           =   19
                     Left            =   6975
                     TabIndex        =   88
                     Top             =   45
                     Width           =   2595
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂÊœ «·’‰›"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   315
                     Index           =   20
                     Left            =   10110
                     TabIndex        =   87
                     Top             =   45
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "0"
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
                     Height          =   315
                     Index           =   21
                     Left            =   270
                     TabIndex        =   86
                     ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
                     Top             =   45
                     Width           =   270
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·√’‰«›"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   315
                     Index           =   22
                     Left            =   1095
                     TabIndex        =   85
                     ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
                     Top             =   45
                     Width           =   960
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‰Ê⁄"
                  Height          =   195
                  Index           =   58
                  Left            =   6630
                  RightToLeft     =   -1  'True
                  TabIndex        =   317
                  Top             =   75
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„’œ—"
                  Height          =   195
                  Index           =   57
                  Left            =   8790
                  RightToLeft     =   -1  'True
                  TabIndex        =   315
                  Top             =   75
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«ÃÂ«œ"
                  Height          =   195
                  Index           =   56
                  Left            =   750
                  RightToLeft     =   -1  'True
                  TabIndex        =   313
                  Top             =   75
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Õ ÊÏ"
                  Height          =   195
                  Index           =   55
                  Left            =   2550
                  RightToLeft     =   -1  'True
                  TabIndex        =   311
                  Top             =   75
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ê“‰ «·„⁄Ì«—Ì"
                  Height          =   195
                  Index           =   54
                  Left            =   4590
                  RightToLeft     =   -1  'True
                  TabIndex        =   309
                  Top             =   75
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  Height          =   195
                  Index           =   48
                  Left            =   13875
                  RightToLeft     =   -1  'True
                  TabIndex        =   233
                  Top             =   165
                  Width           =   270
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”»… «· Õ„Ì·"
                  Height          =   195
                  Index           =   47
                  Left            =   11130
                  RightToLeft     =   -1  'True
                  TabIndex        =   232
                  Top             =   -45
                  Visible         =   0   'False
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ«  ⁄·Ï «·’‰›:"
                  ForeColor       =   &H00000080&
                  Height          =   240
                  Index           =   16
                  Left            =   8880
                  TabIndex        =   53
                  Top             =   5460
                  Width           =   3015
               End
            End
            Begin C1SizerLibCtl.C1Elastic NO 
               Height          =   6570
               Left            =   45
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
               Begin VB.TextBox TxtbarCodeNO 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   6840
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   172
                  Top             =   1320
                  Width           =   2310
               End
               Begin VB.TextBox TxtBinLocation 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   9705
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   169
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.TextBox TxtFactoryNO 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   0
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   510
                  Width           =   1500
               End
               Begin VB.TextBox TxtCatlogNO 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   4380
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   870
                  Width           =   1365
               End
               Begin VB.TextBox XPTxtNamee 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2595
                  MaxLength       =   255
                  TabIndex        =   92
                  Top             =   495
                  Width           =   3150
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Create Prices"
                  Height          =   165
                  Left            =   270
                  TabIndex        =   91
                  Top             =   3495
                  Visible         =   0   'False
                  Width           =   1365
               End
               Begin VB.TextBox TxtPartNo 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   0
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   120
                  Width           =   1230
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   4650
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   120
                  Width           =   1095
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Create Units"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   72
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   1365
               End
               Begin VB.Frame Frame2 
                  Caption         =   "„Êﬁ› «·ÿ·»Ì« "
                  Height          =   2490
                  Left            =   6555
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   3870
                  Width           =   5745
                  Begin VB.TextBox TxtRequired 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   960
                     MaxLength       =   10
                     TabIndex        =   104
                     Top             =   1320
                     Width           =   2805
                  End
                  Begin VB.TextBox Text5 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   960
                     TabIndex        =   70
                     Top             =   2040
                     Width           =   2775
                  End
                  Begin VB.TextBox Text4 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   960
                     TabIndex        =   69
                     Top             =   1680
                     Width           =   2775
                  End
                  Begin VB.TextBox TxtMaxValueqty 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   960
                     TabIndex        =   68
                     Top             =   960
                     Width           =   2775
                  End
                  Begin VB.TextBox Txtminvalueqty 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   960
                     TabIndex        =   67
                     Top             =   600
                     Width           =   2775
                  End
                  Begin VB.TextBox TxtAvilableqty 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   960
                     TabIndex        =   66
                     Top             =   240
                     Width           =   2775
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Õœ ≈⁄«œ… «·ÿ·»"
                     Height          =   270
                     Index           =   8
                     Left            =   3765
                     TabIndex        =   105
                     Top             =   1335
                     Width           =   1635
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·ﬂ„Ì… «·„ÕÃÊ“…"
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   65
                     Top             =   2040
                     Width           =   1335
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ﬂ„Ì… «·ÿ·»Ì…"
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   64
                     Top             =   1680
                     Width           =   1335
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·Õœ  «·«ﬁ’Ï"
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   63
                     Top             =   960
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·Õœ «·«œ‰Ï"
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   62
                     Top             =   600
                     Width           =   1335
                  End
                  Begin VB.Label Text1 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·„ «Õ"
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   61
                     Top             =   240
                     Width           =   1335
                  End
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   4230
                  Style           =   2  'Dropdown List
                  TabIndex        =   4
                  Top             =   1365
                  Width           =   1380
               End
               Begin VB.ComboBox CboItemType 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   5
                  Top             =   960
                  Width           =   3960
               End
               Begin VB.CheckBox ChkAr 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„· „⁄ «·’‰›"
                  Height          =   420
                  Left            =   4230
                  TabIndex        =   8
                  Top             =   3150
                  Width           =   1920
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   360
                  Index           =   0
                  Left            =   1920
                  TabIndex        =   50
                  TabStop         =   0   'False
                  Top             =   3555
                  Width           =   3135
                  _cx             =   5530
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
                  Appearance      =   5
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
                  Begin VB.OptionButton OptGaurType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÌÊ„"
                     Height          =   225
                     Index           =   1
                     Left            =   90
                     TabIndex        =   11
                     Top             =   60
                     Width           =   780
                  End
                  Begin VB.OptionButton OptGaurType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Â—"
                     Height          =   225
                     Index           =   0
                     Left            =   930
                     TabIndex        =   10
                     Top             =   60
                     Value           =   -1  'True
                     Width           =   855
                  End
               End
               Begin VB.TextBox TxtGuarValue 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   5610
                  MaxLength       =   2
                  TabIndex        =   9
                  Top             =   3555
                  Width           =   540
               End
               Begin VB.CheckBox ChkGuar 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·Â ÷„«‰"
                  Height          =   360
                  Left            =   9150
                  TabIndex        =   7
                  Top             =   3555
                  Width           =   1785
               End
               Begin VB.TextBox XPTxtCode 
                  Height          =   315
                  Left            =   4095
                  MaxLength       =   50
                  TabIndex        =   1
                  Top             =   1620
                  Visible         =   0   'False
                  Width           =   2055
               End
               Begin VB.TextBox XPTxtName 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6840
                  MaxLength       =   255
                  RightToLeft     =   -1  'True
                  TabIndex        =   2
                  Text            =   " "
                  Top             =   450
                  Width           =   3960
               End
               Begin VB.TextBox XPTxtID 
                  Alignment       =   1  'Right Justify
                  Height          =   450
                  Left            =   12570
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   0
                  Top             =   825
                  Visible         =   0   'False
                  Width           =   1785
               End
               Begin VB.CheckBox XPChkSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ·Â ”Ì—Ì«·"
                  Height          =   435
                  Left            =   9015
                  TabIndex        =   6
                  Top             =   3150
                  Width           =   1920
               End
               Begin MSDataListLib.DataCombo XPCboGroup 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   3
                  Top             =   120
                  Width           =   3960
                  _ExtentX        =   6985
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCPreFix 
                  Height          =   315
                  Left            =   2730
                  TabIndex        =   75
                  Top             =   120
                  Width           =   1920
                  _ExtentX        =   3387
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
               End
               Begin MSDataListLib.DataCombo DBCboClientName 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   94
                  Top             =   1365
                  Width           =   2865
                  _ExtentX        =   5054
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton SearchCashCustomer 
                  Height          =   330
                  Left            =   1095
                  TabIndex        =   173
                  TabStop         =   0   'False
                  ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
                  Top             =   120
                  Width           =   540
                  _ExtentX        =   953
                  _ExtentY        =   582
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
                  ButtonImage     =   "FrmItems.frx":0E4E
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1455
                  Index           =   4
                  Left            =   0
                  TabIndex        =   182
                  TabStop         =   0   'False
                  Top             =   1680
                  Width           =   12165
                  _cx             =   21458
                  _cy             =   2566
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
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
                  ForeColor       =   192
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "√”⁄«— «·’‰›"
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
                  Begin VB.TextBox lastorderPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   3465
                     MaxLength       =   10
                     TabIndex        =   303
                     Top             =   960
                     Width           =   1605
                  End
                  Begin VB.TextBox lstorderdate 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   480
                     MaxLength       =   10
                     TabIndex        =   302
                     Top             =   960
                     Width           =   1605
                  End
                  Begin VB.TextBox TxtItemMaxDiscount 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   6645
                     MaxLength       =   10
                     TabIndex        =   228
                     Top             =   780
                     Width           =   1260
                  End
                  Begin VB.TextBox TxtDealerPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   450
                     MaxLength       =   10
                     TabIndex        =   187
                     Top             =   570
                     Width           =   1605
                  End
                  Begin VB.TextBox TxtCusPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   3435
                     MaxLength       =   10
                     TabIndex        =   186
                     Top             =   570
                     Width           =   1605
                  End
                  Begin VB.TextBox XPTxtSall 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   6645
                     MaxLength       =   10
                     TabIndex        =   185
                     Top             =   480
                     Width           =   1260
                  End
                  Begin VB.TextBox XPTxtPurchase 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   6645
                     MaxLength       =   10
                     TabIndex        =   184
                     Top             =   150
                     Width           =   1260
                  End
                  Begin VB.TextBox TxtFreeQty 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   450
                     MaxLength       =   10
                     TabIndex        =   183
                     Top             =   225
                     Width           =   1605
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Œ— ”⁄— ›Ì «·⁄—÷"
                     Height          =   195
                     Index           =   53
                     Left            =   5130
                     TabIndex        =   305
                     Top             =   975
                     Width           =   1320
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «Œ— ”⁄— ⁄—÷"
                     Height          =   195
                     Index           =   52
                     Left            =   2085
                     TabIndex        =   304
                     Top             =   975
                     Width           =   1380
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«ﬁ’Ì Œ’„"
                     Height          =   180
                     Index           =   44
                     Left            =   8805
                     TabIndex        =   227
                     Top             =   870
                     Width           =   930
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄(œÌ·—)"
                     Height          =   375
                     Index           =   11
                     Left            =   2280
                     TabIndex        =   194
                     Top             =   585
                     Width           =   1035
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄(⁄„Ì·)"
                     Height          =   270
                     Index           =   10
                     Left            =   5160
                     TabIndex        =   193
                     Top             =   585
                     Width           =   1140
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Œ— ”⁄— ‘—«¡"
                     Height          =   195
                     Index           =   5
                     Left            =   8250
                     TabIndex        =   192
                     Top             =   255
                     Width           =   1485
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄(„” Â·ﬂ)"
                     Height          =   210
                     Index           =   7
                     Left            =   8250
                     TabIndex        =   191
                     Top             =   585
                     Width           =   1485
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„ Ê”ÿ  «· ﬂ·›…"
                     Height          =   195
                     Index           =   30
                     Left            =   5265
                     TabIndex        =   190
                     ToolTipText     =   "ÌŸÂ— »⁄œ œŒÊ· «Ê· ⁄„·Ì… ‘—«¡ Ê—’Ìœ «›  «ÕÏ"
                     Top             =   255
                     Width           =   1035
                  End
                  Begin VB.Label LblCostPrice 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E2E9E9&
                     BorderStyle     =   1  'Fixed Single
                     ForeColor       =   &H80000008&
                     Height          =   285
                     Left            =   3435
                     TabIndex        =   189
                     Top             =   255
                     Width           =   1605
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”Ì… «·’‰› «·„Ã«‰Ì"
                     Height          =   180
                     Index           =   45
                     Left            =   2280
                     TabIndex        =   188
                     Top             =   225
                     Width           =   1035
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   2220
                  Index           =   6
                  Left            =   0
                  TabIndex        =   221
                  TabStop         =   0   'False
                  Top             =   3960
                  Width           =   6555
                  _cx             =   11562
                  _cy             =   3916
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
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
                  ForeColor       =   192
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "’Ê—… «·’‰›"
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
                  Begin Dynamic_Byte.NewViewBox ImgPic 
                     Height          =   1470
                     Left            =   0
                     TabIndex        =   222
                     ToolTipText     =   "≈÷€ÿ ⁄·Ï «·’Ê—… „— Ì‰ ·· ﬂ»Ì—"
                     Top             =   150
                     Width           =   4350
                     _ExtentX        =   7673
                     _ExtentY        =   2593
                  End
                  Begin ImpulseButton.ISButton CmdPic 
                     Height          =   480
                     Index           =   0
                     Left            =   4350
                     TabIndex        =   223
                     Top             =   150
                     Width           =   1020
                     _ExtentX        =   1799
                     _ExtentY        =   847
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "≈÷«›… ’Ê—…"
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
                     ButtonImage     =   "FrmItems.frx":124B
                     ColorButton     =   14871017
                     Alignment       =   1
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton CmdPic 
                     Height          =   405
                     Index           =   1
                     Left            =   4350
                     TabIndex        =   224
                     Top             =   630
                     Width           =   1020
                     _ExtentX        =   1799
                     _ExtentY        =   714
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ–› «·’Ê—…"
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
                     ButtonImage     =   "FrmItems.frx":15E5
                     ColorButton     =   14871017
                     Alignment       =   1
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton CmdAttach 
                     Height          =   285
                     Left            =   4455
                     TabIndex        =   230
                     Top             =   990
                     Width           =   915
                     _ExtentX        =   1614
                     _ExtentY        =   503
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
               End
               Begin MSDataListLib.DataCombo DcTemplate 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   226
                  Top             =   960
                  Width           =   3420
                  _ExtentX        =   6033
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁ«·»"
                  Height          =   195
                  Index           =   43
                  Left            =   3420
                  TabIndex        =   225
                  Top             =   960
                  Width           =   810
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»«—ﬂÊœ"
                  Height          =   315
                  Index           =   46
                  Left            =   9015
                  TabIndex        =   171
                  Top             =   1440
                  Width           =   555
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Êﬁ⁄"
                  Height          =   270
                  Index           =   40
                  Left            =   10800
                  TabIndex        =   170
                  Top             =   1440
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·„’‰⁄"
                  Height          =   315
                  Index           =   35
                  Left            =   1770
                  TabIndex        =   100
                  Top             =   630
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬂ «·ÊÃ"
                  Height          =   195
                  Index           =   34
                  Left            =   4920
                  TabIndex        =   98
                  Top             =   990
                  Width           =   1770
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ê—œ «·«› —«÷Ì"
                  Height          =   315
                  Index           =   32
                  Left            =   2865
                  TabIndex        =   95
                  Top             =   1365
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   195
                  Index           =   31
                  Left            =   4920
                  TabIndex        =   93
                  Top             =   630
                  Width           =   1770
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁÿ⁄Â/«·„ÊœÌ·"
                  Height          =   435
                  Index           =   0
                  Left            =   1500
                  TabIndex        =   76
                  Top             =   120
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Õ«·… «·€«·»… ··’‰›"
                  Height          =   435
                  Index           =   29
                  Left            =   5460
                  TabIndex        =   56
                  Top             =   1395
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·’‰›"
                  Height          =   270
                  Index           =   15
                  Left            =   10800
                  TabIndex        =   52
                  Top             =   1095
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ÿ«„ «· ⁄«„·"
                  Height          =   285
                  Index           =   14
                  Left            =   6420
                  TabIndex        =   51
                  Top             =   3285
                  Width           =   1650
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ÿ«„ «·÷„«‰"
                  Height          =   300
                  Index           =   13
                  Left            =   10800
                  TabIndex        =   49
                  Top             =   3615
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·÷„«‰ «·√› —«÷Ì…"
                  Height          =   270
                  Index           =   12
                  Left            =   1230
                  TabIndex        =   48
                  Top             =   3615
                  Width           =   6975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·’‰›"
                  Height          =   450
                  Index           =   6
                  Left            =   12300
                  TabIndex        =   37
                  Top             =   945
                  Visible         =   0   'False
                  Width           =   960
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   195
                  Index           =   23
                  Left            =   4785
                  TabIndex        =   36
                  Top             =   120
                  Width           =   1770
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   270
                  Index           =   3
                  Left            =   10800
                  TabIndex        =   35
                  Top             =   570
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„Ã„Ê⁄…"
                  Height          =   270
                  Index           =   4
                  Left            =   10800
                  TabIndex        =   34
                  Top             =   180
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ã· «·Õ«·Ì:"
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
                  Height          =   330
                  Index           =   1
                  Left            =   3960
                  TabIndex        =   33
                  Top             =   6225
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·”Ã·« :"
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
                  Height          =   330
                  Index           =   2
                  Left            =   690
                  TabIndex        =   32
                  Top             =   6195
                  Width           =   1635
               End
               Begin VB.Label XPTxtCount 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   210
                  Left            =   135
                  TabIndex        =   31
                  Top             =   6195
                  Width           =   405
               End
               Begin VB.Label XPTxtCurrent 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   330
                  Left            =   2595
                  TabIndex        =   30
                  Top             =   6195
                  Width           =   690
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ÿ«„ «·”Ì—Ì«·"
                  Height          =   435
                  Index           =   9
                  Left            =   11070
                  TabIndex        =   29
                  Top             =   3285
                  Width           =   1095
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6570
               Index           =   1
               Left            =   13335
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
               Begin VB.TextBox TxtUnitFactor 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   6285
                  MaxLength       =   6
                  TabIndex        =   240
                  Top             =   3510
                  Width           =   1230
               End
               Begin VB.CheckBox ChkDef 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÊÕœ… ≈› —«÷Ì…"
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   10935
                  TabIndex        =   239
                  Top             =   3510
                  Width           =   1365
               End
               Begin VB.TextBox TxtUnitSalesPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4515
                  MaxLength       =   6
                  TabIndex        =   238
                  Top             =   3510
                  Width           =   945
               End
               Begin VB.TextBox TxtUnitPurPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5610
                  MaxLength       =   6
                  TabIndex        =   237
                  Top             =   3510
                  Width           =   675
               End
               Begin VB.Frame Frame1 
                  Enabled         =   0   'False
                  Height          =   1275
                  Left            =   6420
                  TabIndex        =   111
                  Top             =   6240
                  Visible         =   0   'False
                  Width           =   10260
                  Begin VB.TextBox TxtRowNumber 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   120
                     TabIndex        =   112
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   645
                  End
                  Begin MSDataListLib.DataCombo DcboItems1 
                     Height          =   315
                     Left            =   4920
                     TabIndex        =   113
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   3675
                     _ExtentX        =   6482
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin C1SizerLibCtl.C1Elastic EltCont 
                     Height          =   510
                     Left            =   1350
                     TabIndex        =   114
                     TabStop         =   0   'False
                     Top             =   4740
                     Width           =   1650
                     _cx             =   2910
                     _cy             =   900
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
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   420
                        Index           =   23
                        Left            =   900
                        TabIndex        =   115
                        Top             =   90
                        Visible         =   0   'False
                        Width           =   720
                        _ExtentX        =   1270
                        _ExtentY        =   741
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
                        ButtonImage     =   "FrmItems.frx":197F
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   420
                        Index           =   22
                        Left            =   180
                        TabIndex        =   116
                        Top             =   60
                        Width           =   690
                        _ExtentX        =   1217
                        _ExtentY        =   741
                        ButtonStyle     =   1
                        ButtonPositionImage=   1
                        Caption         =   "≈·€«¡"
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
                        ButtonImage     =   "FrmItems.frx":1D19
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISB»ÕÀ 
                        Height          =   330
                        Left            =   5880
                        TabIndex        =   117
                        TabStop         =   0   'False
                        ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                        Top             =   90
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
                        ButtonImage     =   "FrmItems.frx":20B3
                        ColorButton     =   14737632
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISB ÕœÌÀ 
                        Height          =   330
                        Left            =   6045
                        TabIndex        =   118
                        TabStop         =   0   'False
                        ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
                        Top             =   105
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
                        ButtonImage     =   "FrmItems.frx":244D
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                  End
                  Begin VB.Label itemnamex 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·’‰›"
                     Height          =   315
                     Index           =   2
                     Left            =   9360
                     TabIndex        =   119
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1095
                  End
               End
               Begin MSDataListLib.DataCombo DcboUnits 
                  Height          =   315
                  Left            =   7650
                  TabIndex        =   241
                  Top             =   3510
                  Width           =   3285
                  _ExtentX        =   5794
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VSFlex8Ctl.VSFlexGrid FgUnites 
                  Height          =   3000
                  Left            =   135
                  TabIndex        =   242
                  Top             =   135
                  Width           =   12165
                  _cx             =   21458
                  _cy             =   5292
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
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmItems.frx":27E7
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   225
                  Index           =   20
                  Left            =   3825
                  TabIndex        =   243
                  Top             =   3435
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmItems.frx":29CF
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   225
                  Index           =   21
                  Left            =   2865
                  TabIndex        =   244
                  Top             =   3435
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmItems.frx":2D69
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Notes : To Define Units According to Small unit"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   600
                  Index           =   37
                  Left            =   540
                  RightToLeft     =   -1  'True
                  TabIndex        =   245
                  Top             =   4245
                  Visible         =   0   'False
                  Width           =   4650
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   $"FrmItems.frx":3303
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   1515
                  Index           =   33
                  Left            =   540
                  RightToLeft     =   -1  'True
                  TabIndex        =   251
                  Top             =   4875
                  Width           =   5340
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·ÊÕœ…"
                  Height          =   270
                  Index           =   0
                  Left            =   8745
                  TabIndex        =   250
                  Top             =   3210
                  Width           =   1785
               End
               Begin VB.Label lbl«·⁄·«ﬁ…„⁄ 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄·«ﬁ… „⁄ «·ÊÕœ… «·”«»ﬁ…"
                  Height          =   270
                  Index           =   1
                  Left            =   6150
                  TabIndex        =   249
                  Top             =   3210
                  Width           =   1920
               End
               Begin VB.Label lblÊÕœ…≈› —«÷Ì… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÊÕœ… ≈› —«÷Ì…"
                  Height          =   270
                  Index           =   3
                  Left            =   10800
                  TabIndex        =   248
                  Top             =   3210
                  Width           =   1635
               End
               Begin VB.Label lbl”⁄—«·»Ì⁄ 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”⁄— «·»Ì⁄"
                  Height          =   270
                  Index           =   4
                  Left            =   4650
                  TabIndex        =   247
                  Top             =   3210
                  Width           =   810
               End
               Begin VB.Label lbl”⁄—«·‘—«¡ 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”⁄— «·‘—«¡"
                  Height          =   270
                  Index           =   5
                  Left            =   5325
                  TabIndex        =   246
                  Top             =   3210
                  Width           =   960
               End
               Begin VB.Shape Shape1 
                  BorderWidth     =   2
                  FillColor       =   &H00C0FFFF&
                  FillStyle       =   0  'Solid
                  Height          =   2445
                  Left            =   0
                  Top             =   4110
                  Width           =   6840
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6570
               Index           =   8
               Left            =   13635
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   240
                  Index           =   0
                  Left            =   9570
                  MaxLength       =   6
                  TabIndex        =   274
                  Top             =   5640
                  Width           =   540
               End
               Begin VB.TextBox TxtDiscount 
                  Height          =   240
                  Index           =   0
                  Left            =   8880
                  TabIndex        =   273
                  Top             =   5640
                  Width           =   690
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   240
                  Index           =   1
                  Left            =   8340
                  MaxLength       =   6
                  TabIndex        =   272
                  Top             =   5640
                  Width           =   540
               End
               Begin VB.TextBox TxtDiscount 
                  Height          =   240
                  Index           =   1
                  Left            =   7785
                  TabIndex        =   271
                  Top             =   5640
                  Width           =   555
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   240
                  Index           =   2
                  Left            =   7110
                  MaxLength       =   6
                  TabIndex        =   270
                  Top             =   5640
                  Width           =   540
               End
               Begin VB.TextBox TxtDiscount 
                  Height          =   240
                  Index           =   2
                  Left            =   6555
                  TabIndex        =   269
                  Top             =   5640
                  Width           =   555
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   240
                  Index           =   3
                  Left            =   5880
                  MaxLength       =   6
                  TabIndex        =   268
                  Top             =   5640
                  Width           =   675
               End
               Begin VB.TextBox TxtDiscount 
                  Height          =   240
                  Index           =   3
                  Left            =   5325
                  TabIndex        =   267
                  Top             =   5640
                  Width           =   555
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   240
                  Index           =   4
                  Left            =   4785
                  MaxLength       =   6
                  TabIndex        =   266
                  Top             =   5640
                  Width           =   540
               End
               Begin VB.TextBox TxtDiscount 
                  Height          =   240
                  Index           =   4
                  Left            =   4095
                  TabIndex        =   265
                  Top             =   5640
                  Width           =   555
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   240
                  Index           =   5
                  Left            =   3555
                  MaxLength       =   6
                  TabIndex        =   264
                  Top             =   5640
                  Width           =   540
               End
               Begin VB.TextBox TxtDiscount 
                  Height          =   240
                  Index           =   5
                  Left            =   2865
                  TabIndex        =   263
                  Top             =   5640
                  Width           =   555
               End
               Begin VB.OptionButton optBranch 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·ﬂ· «·›—Ê⁄"
                  Height          =   150
                  Index           =   0
                  Left            =   10380
                  RightToLeft     =   -1  'True
                  TabIndex        =   262
                  Top             =   4995
                  Value           =   -1  'True
                  Width           =   1095
               End
               Begin VB.OptionButton optBranch 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "›—⁄ „Õœœ"
                  Height          =   150
                  Index           =   1
                  Left            =   9300
                  RightToLeft     =   -1  'True
                  TabIndex        =   261
                  Top             =   4995
                  Width           =   810
               End
               Begin VB.Frame Frame3 
                  Caption         =   "«”⁄«— »Ì⁄ «·’‰›"
                  Enabled         =   0   'False
                  Height          =   225
                  Left            =   405
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   6360
                  Visible         =   0   'False
                  Width           =   12165
                  Begin VB.TextBox TxtPriceDes 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   3000
                     MaxLength       =   6
                     TabIndex        =   127
                     Top             =   5520
                     Visible         =   0   'False
                     Width           =   2235
                  End
                  Begin VB.TextBox TxtPriceName 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   6480
                     MaxLength       =   50
                     TabIndex        =   126
                     Top             =   5520
                     Visible         =   0   'False
                     Width           =   2505
                  End
                  Begin VB.TextBox Text9 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   4590
                     MaxLength       =   6
                     TabIndex        =   125
                     Top             =   5400
                     Visible         =   0   'False
                     Width           =   1785
                  End
                  Begin VB.CheckBox ChkDefSalePrice 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «› —«÷Ì"
                     Height          =   315
                     Left            =   9090
                     TabIndex        =   124
                     Top             =   5520
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.TextBox TxtSalesPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   5340
                     MaxLength       =   6
                     TabIndex        =   123
                     Top             =   5520
                     Visible         =   0   'False
                     Width           =   1155
                  End
                  Begin VB.TextBox Text7 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   2940
                     MaxLength       =   6
                     TabIndex        =   122
                     Top             =   5640
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.TextBox Text6 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   120
                     TabIndex        =   121
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   645
                  End
                  Begin MSDataListLib.DataCombo DataCombo1 
                     Height          =   315
                     Left            =   6420
                     TabIndex        =   128
                     Top             =   5400
                     Visible         =   0   'False
                     Width           =   2625
                     _ExtentX        =   4630
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VSFlex8Ctl.VSFlexGrid FgPrices 
                     Height          =   1245
                     Left            =   10470
                     TabIndex        =   129
                     Top             =   6240
                     Visible         =   0   'False
                     Width           =   8955
                     _cx             =   15796
                     _cy             =   2196
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
                     Cols            =   12
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmItems.frx":3454
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
                  Begin MSDataListLib.DataCombo DataCombo2 
                     Height          =   315
                     Left            =   4920
                     TabIndex        =   130
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   3675
                     _ExtentX        =   6482
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic1 
                     Height          =   510
                     Left            =   1350
                     TabIndex        =   131
                     TabStop         =   0   'False
                     Top             =   5340
                     Visible         =   0   'False
                     Width           =   1650
                     _cx             =   2910
                     _cy             =   900
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
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   420
                        Index           =   12
                        Left            =   900
                        TabIndex        =   132
                        Top             =   570
                        Visible         =   0   'False
                        Width           =   720
                        _ExtentX        =   1270
                        _ExtentY        =   741
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
                        ButtonImage     =   "FrmItems.frx":365A
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   420
                        Index           =   13
                        Left            =   180
                        TabIndex        =   133
                        Top             =   540
                        Visible         =   0   'False
                        Width           =   690
                        _ExtentX        =   1217
                        _ExtentY        =   741
                        ButtonStyle     =   1
                        ButtonPositionImage=   1
                        Caption         =   "≈·€«¡"
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
                        ButtonImage     =   "FrmItems.frx":39F4
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton1 
                        Height          =   330
                        Left            =   5880
                        TabIndex        =   134
                        TabStop         =   0   'False
                        ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                        Top             =   90
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
                        ButtonImage     =   "FrmItems.frx":3D8E
                        ColorButton     =   14737632
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton2 
                        Height          =   330
                        Left            =   6045
                        TabIndex        =   135
                        TabStop         =   0   'False
                        ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
                        Top             =   105
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
                        ButtonImage     =   "FrmItems.frx":4128
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                  End
                  Begin VB.Label lbl”⁄—«·»Ì⁄ 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„·«ÕŸ« "
                     Height          =   255
                     Index           =   1
                     Left            =   3720
                     TabIndex        =   142
                     Top             =   5280
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.Label itemnamex 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«”„ «·’‰›"
                     Height          =   315
                     Index           =   0
                     Left            =   9360
                     TabIndex        =   141
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1095
                  End
                  Begin VB.Label lbl«”„«·ÊÕœ… 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·”⁄—"
                     Height          =   255
                     Index           =   1
                     Left            =   6420
                     TabIndex        =   140
                     Top             =   5220
                     Visible         =   0   'False
                     Width           =   2625
                  End
                  Begin VB.Label lbl«·⁄·«ﬁ…„⁄ 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄·«ﬁ… „⁄ «·ÊÕœ… «·”«»ﬁ…"
                     Height          =   255
                     Index           =   0
                     Left            =   4620
                     TabIndex        =   139
                     Top             =   5580
                     Visible         =   0   'False
                     Width           =   1755
                  End
                  Begin VB.Label lblÊÕœ…≈› —«÷Ì… 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «› —«÷Ì"
                     Height          =   255
                     Index           =   0
                     Left            =   9090
                     TabIndex        =   138
                     Top             =   5220
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label lbl”⁄—«·»Ì⁄ 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄"
                     Height          =   255
                     Index           =   0
                     Left            =   5460
                     TabIndex        =   137
                     Top             =   5220
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.Label lbl”⁄—«·‘—«¡ 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·‘—«¡"
                     Height          =   255
                     Index           =   0
                     Left            =   1980
                     TabIndex        =   136
                     Top             =   5700
                     Visible         =   0   'False
                     Width           =   795
                  End
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   300
                  Index           =   14
                  Left            =   2190
                  TabIndex        =   275
                  Top             =   5580
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   529
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
                  ButtonImage     =   "FrmItems.frx":44C2
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   300
                  Index           =   15
                  Left            =   1500
                  TabIndex        =   276
                  Top             =   5580
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   529
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
                  ButtonImage     =   "FrmItems.frx":485C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid FgSalePrice 
                  Height          =   4695
                  Left            =   2055
                  TabIndex        =   277
                  Top             =   90
                  Width           =   10110
                  _cx             =   17833
                  _cy             =   8281
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
                  Cols            =   26
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmItems.frx":4DF6
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
               Begin MSDataListLib.DataCombo DcUnit 
                  Height          =   315
                  Left            =   10110
                  TabIndex        =   278
                  Top             =   5640
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcBranch 
                  Height          =   315
                  Left            =   6150
                  TabIndex        =   279
                  Top             =   4995
                  Width           =   3000
                  _ExtentX        =   5292
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÊÕœÂ"
                  Height          =   285
                  Index           =   3
                  Left            =   10380
                  TabIndex        =   292
                  Top             =   5370
                  Width           =   555
               End
               Begin VB.Label lblPrice 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— 1"
                  Height          =   285
                  Index           =   0
                  Left            =   9435
                  TabIndex        =   291
                  Top             =   5370
                  Width           =   540
               End
               Begin VB.Label lblDiscount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ 1"
                  Height          =   285
                  Index           =   0
                  Left            =   8880
                  TabIndex        =   290
                  Top             =   5370
                  Width           =   690
               End
               Begin VB.Label lblPrice 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— 2"
                  Height          =   285
                  Index           =   1
                  Left            =   8340
                  TabIndex        =   289
                  Top             =   5370
                  Width           =   540
               End
               Begin VB.Label lblDiscount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ 2"
                  Height          =   285
                  Index           =   1
                  Left            =   7785
                  TabIndex        =   288
                  Top             =   5370
                  Width           =   555
               End
               Begin VB.Label lblPrice 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— 3"
                  Height          =   285
                  Index           =   2
                  Left            =   7110
                  TabIndex        =   287
                  Top             =   5370
                  Width           =   675
               End
               Begin VB.Label lblDiscount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ 3"
                  Height          =   285
                  Index           =   2
                  Left            =   6555
                  TabIndex        =   286
                  Top             =   5370
                  Width           =   555
               End
               Begin VB.Label lblPrice 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— 4"
                  Height          =   285
                  Index           =   3
                  Left            =   5880
                  TabIndex        =   285
                  Top             =   5370
                  Width           =   675
               End
               Begin VB.Label lblDiscount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ 4"
                  Height          =   285
                  Index           =   3
                  Left            =   5325
                  TabIndex        =   284
                  Top             =   5370
                  Width           =   555
               End
               Begin VB.Label lblPrice 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— 5"
                  Height          =   285
                  Index           =   4
                  Left            =   4785
                  TabIndex        =   283
                  Top             =   5370
                  Width           =   540
               End
               Begin VB.Label lblDiscount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ 5"
                  Height          =   285
                  Index           =   4
                  Left            =   4095
                  TabIndex        =   282
                  Top             =   5370
                  Width           =   555
               End
               Begin VB.Label lblPrice 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— 6"
                  Height          =   285
                  Index           =   5
                  Left            =   3555
                  TabIndex        =   281
                  Top             =   5370
                  Width           =   540
               End
               Begin VB.Label lblDiscount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„6"
                  Height          =   285
                  Index           =   5
                  Left            =   2865
                  TabIndex        =   280
                  Top             =   5370
                  Width           =   690
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6570
               Index           =   9
               Left            =   13935
               TabIndex        =   108
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
                  Caption         =   "«”⁄«— «·‘—«¡ „‰ «·„Ê—œÌ‰"
                  Enabled         =   0   'False
                  Height          =   5340
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   4920
                  Visible         =   0   'False
                  Width           =   11895
                  Begin VB.TextBox Text14 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   120
                     TabIndex        =   150
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   645
                  End
                  Begin VB.TextBox Text13 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   2940
                     MaxLength       =   6
                     TabIndex        =   149
                     Top             =   5640
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.TextBox TxtSalesPrice1 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   5340
                     MaxLength       =   6
                     TabIndex        =   148
                     Top             =   4440
                     Visible         =   0   'False
                     Width           =   1155
                  End
                  Begin VB.CheckBox ChkDefSalePrice1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «› —«÷Ì"
                     Height          =   315
                     Left            =   9090
                     TabIndex        =   147
                     Top             =   4440
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.TextBox Text11 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   4590
                     MaxLength       =   6
                     TabIndex        =   146
                     Top             =   5400
                     Visible         =   0   'False
                     Width           =   1785
                  End
                  Begin VB.TextBox TxtPriceName1 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   6480
                     MaxLength       =   50
                     TabIndex        =   145
                     Top             =   4440
                     Visible         =   0   'False
                     Width           =   2505
                  End
                  Begin VB.TextBox TxtPriceDes1 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   3000
                     MaxLength       =   6
                     TabIndex        =   144
                     Top             =   4440
                     Visible         =   0   'False
                     Width           =   2235
                  End
                  Begin MSDataListLib.DataCombo DataCombo3 
                     Height          =   315
                     Left            =   6420
                     TabIndex        =   151
                     Top             =   5400
                     Visible         =   0   'False
                     Width           =   2625
                     _ExtentX        =   4630
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VSFlex8Ctl.VSFlexGrid FgPrices1 
                     Height          =   1005
                     Left            =   1350
                     TabIndex        =   152
                     Top             =   5640
                     Visible         =   0   'False
                     Width           =   8955
                     _cx             =   15796
                     _cy             =   1773
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
                     Cols            =   12
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmItems.frx":51DC
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
                  Begin MSDataListLib.DataCombo DataCombo4 
                     Height          =   315
                     Left            =   4920
                     TabIndex        =   153
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   3675
                     _ExtentX        =   6482
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                     Height          =   510
                     Left            =   1350
                     TabIndex        =   154
                     TabStop         =   0   'False
                     Top             =   4740
                     Width           =   1650
                     _cx             =   2910
                     _cy             =   900
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
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   420
                        Index           =   16
                        Left            =   900
                        TabIndex        =   155
                        Top             =   90
                        Visible         =   0   'False
                        Width           =   720
                        _ExtentX        =   1270
                        _ExtentY        =   741
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
                        ButtonImage     =   "FrmItems.frx":53E3
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   420
                        Index           =   17
                        Left            =   180
                        TabIndex        =   156
                        Top             =   60
                        Visible         =   0   'False
                        Width           =   690
                        _ExtentX        =   1217
                        _ExtentY        =   741
                        ButtonStyle     =   1
                        ButtonPositionImage=   1
                        Caption         =   "≈·€«¡"
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
                        ButtonImage     =   "FrmItems.frx":577D
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton3 
                        Height          =   330
                        Left            =   5880
                        TabIndex        =   157
                        TabStop         =   0   'False
                        ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                        Top             =   90
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
                        ButtonImage     =   "FrmItems.frx":5B17
                        ColorButton     =   14737632
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton4 
                        Height          =   330
                        Left            =   6045
                        TabIndex        =   158
                        TabStop         =   0   'False
                        ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
                        Top             =   105
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
                        ButtonImage     =   "FrmItems.frx":5EB1
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   390
                     Index           =   18
                     Left            =   2160
                     TabIndex        =   159
                     Top             =   4320
                     Visible         =   0   'False
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
                     ButtonImage     =   "FrmItems.frx":624B
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   390
                     Index           =   19
                     Left            =   1440
                     TabIndex        =   160
                     Top             =   4320
                     Visible         =   0   'False
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
                     ButtonImage     =   "FrmItems.frx":65E5
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl”⁄—«·‘—«¡ 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·‘—«¡"
                     Height          =   255
                     Index           =   1
                     Left            =   1980
                     TabIndex        =   167
                     Top             =   5700
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.Label lbl”⁄—«·»Ì⁄ 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·‘—«¡"
                     Height          =   255
                     Index           =   3
                     Left            =   5460
                     TabIndex        =   166
                     Top             =   4140
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.Label lblÊÕœ…≈› —«÷Ì… 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «› —«÷Ì"
                     Height          =   255
                     Index           =   1
                     Left            =   9090
                     TabIndex        =   165
                     Top             =   4140
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label lbl«·⁄·«ﬁ…„⁄ 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄·«ﬁ… „⁄ «·ÊÕœ… «·”«»ﬁ…"
                     Height          =   255
                     Index           =   2
                     Left            =   4620
                     TabIndex        =   164
                     Top             =   5100
                     Visible         =   0   'False
                     Width           =   1755
                  End
                  Begin VB.Label lbl«”„«·ÊÕœ… 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·”⁄—"
                     Height          =   255
                     Index           =   2
                     Left            =   6420
                     TabIndex        =   163
                     Top             =   4140
                     Visible         =   0   'False
                     Width           =   2625
                  End
                  Begin VB.Label itemnamex 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«”„ «·’‰›"
                     Height          =   315
                     Index           =   1
                     Left            =   9360
                     TabIndex        =   162
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1095
                  End
                  Begin VB.Label lbl”⁄—«·»Ì⁄ 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„·«ÕŸ« "
                     Height          =   255
                     Index           =   2
                     Left            =   3720
                     TabIndex        =   161
                     Top             =   4080
                     Visible         =   0   'False
                     Width           =   795
                  End
               End
               Begin VSFlex8Ctl.VSFlexGrid FgVendorPrice 
                  Height          =   6315
                  Left            =   135
                  TabIndex        =   258
                  Top             =   240
                  Width           =   12030
                  _cx             =   21220
                  _cy             =   11139
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
                  BackColorBkg    =   16777215
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
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmItems.frx":6B7F
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
               Height          =   6570
               Index           =   10
               Left            =   14235
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
                  Height          =   6510
                  Index           =   12
                  Left            =   0
                  TabIndex        =   234
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   12165
                  _cx             =   21458
                  _cy             =   11483
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
                  Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
                     Height          =   7260
                     Left            =   0
                     TabIndex        =   301
                     Top             =   240
                     Width           =   12075
                     _cx             =   21299
                     _cy             =   12806
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
                     BackColorBkg    =   16777215
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
                     Cols            =   23
                     FixedRows       =   1
                     FixedCols       =   2
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmItems.frx":6E6B
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
                     ExplorerBar     =   5
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
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6570
               Index           =   11
               Left            =   14535
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   1140
                  Left            =   270
                  TabIndex        =   168
                  Top             =   270
                  Width           =   12030
                  _cx             =   21220
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
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmItems.frx":71D2
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   705
                  Index           =   13
                  Left            =   135
                  TabIndex        =   208
                  TabStop         =   0   'False
                  Top             =   1635
                  Width           =   12030
                  _cx             =   21220
                  _cy             =   1244
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
                  Begin VB.TextBox TxtRemark 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   293
                     Top             =   270
                     Width           =   2760
                  End
                  Begin VB.TextBox TxtItemPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Index           =   2
                     Left            =   4965
                     MaxLength       =   5
                     TabIndex        =   211
                     Top             =   270
                     Width           =   960
                  End
                  Begin VB.TextBox TxtItemQty 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Index           =   2
                     Left            =   6000
                     MaxLength       =   5
                     TabIndex        =   210
                     Top             =   270
                     Width           =   945
                  End
                  Begin VB.TextBox TxtCodeAother 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   10830
                     TabIndex        =   209
                     Top             =   270
                     Width           =   1065
                  End
                  Begin MSDataListLib.DataCombo Dcbiteem 
                     Height          =   315
                     Left            =   8280
                     TabIndex        =   212
                     Top             =   270
                     Width           =   2535
                     _ExtentX        =   4471
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   315
                     Index           =   24
                     Left            =   1230
                     TabIndex        =   213
                     Top             =   210
                     Width           =   540
                     _ExtentX        =   953
                     _ExtentY        =   556
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
                     ColorButton     =   14871017
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   315
                     Index           =   25
                     Left            =   135
                     TabIndex        =   214
                     Top             =   210
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   556
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
                  End
                  Begin MSDataListLib.DataCombo Dcbuniit 
                     Height          =   315
                     Left            =   7080
                     TabIndex        =   295
                     Top             =   270
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊÕœÂ"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   210
                     Index           =   51
                     Left            =   6720
                     TabIndex        =   296
                     Top             =   60
                     Width           =   1230
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„·«ÕŸ« "
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   195
                     Index           =   50
                     Left            =   3120
                     TabIndex        =   294
                     Top             =   30
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·”⁄—"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   195
                     Index           =   42
                     Left            =   4710
                     TabIndex        =   220
                     Top             =   30
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬂ„Ì…"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   210
                     Index           =   41
                     Left            =   5445
                     TabIndex        =   219
                     Top             =   60
                     Width           =   1230
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·√’‰«›"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   195
                     Index           =   39
                     Left            =   960
                     TabIndex        =   218
                     ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
                     Top             =   30
                     Width           =   1095
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "0"
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
                     Height          =   195
                     Index           =   38
                     Left            =   135
                     TabIndex        =   217
                     ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
                     Top             =   30
                     Width           =   270
                  End
                  Begin VB.Label Label8 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "ﬂÊœ «·’‰›"
                     ForeColor       =   &H000000C0&
                     Height          =   195
                     Left            =   10755
                     TabIndex        =   216
                     Top             =   0
                     Width           =   945
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«”„ «·’‰›"
                     ForeColor       =   &H000000C0&
                     Height          =   195
                     Left            =   8985
                     TabIndex        =   215
                     Top             =   0
                     Width           =   960
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6570
               Index           =   14
               Left            =   14835
               TabIndex        =   174
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
               Begin VB.Frame Frame6 
                  Height          =   6420
                  Left            =   0
                  TabIndex        =   175
                  Top             =   -7185
                  Width           =   12300
                  Begin VB.Frame Frame7 
                     BackColor       =   &H00E2E9E9&
                     Height          =   3090
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   179
                     Top             =   4920
                     Width           =   12240
                     Begin VB.Label Label6 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Height          =   300
                        Left            =   2640
                        RightToLeft     =   -1  'True
                        TabIndex        =   181
                        Top             =   2640
                        Width           =   1095
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ï ﬁÌ„… «·œ›⁄«  ÌœÊÌ«ı"
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
                        Height          =   255
                        Left            =   60
                        TabIndex        =   180
                        Top             =   2280
                        Visible         =   0   'False
                        Width           =   2595
                     End
                  End
                  Begin VB.Frame lblExt 
                     BackColor       =   &H00E2E9E9&
                     Height          =   3450
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   176
                     Top             =   0
                     Width           =   12240
                     Begin VB.Label Label12 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ï ﬁÌ„… «·œ›⁄«  ÌœÊÌ«ı"
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
                        Height          =   255
                        Left            =   60
                        TabIndex        =   178
                        Top             =   2280
                        Visible         =   0   'False
                        Width           =   2595
                     End
                     Begin VB.Label LbToTalExtra 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Height          =   300
                        Left            =   2640
                        RightToLeft     =   -1  'True
                        TabIndex        =   177
                        Top             =   2640
                        Width           =   1095
                     End
                  End
               End
               Begin VSFlex8Ctl.VSFlexGrid fgDiamonds 
                  Height          =   2865
                  Left            =   135
                  TabIndex        =   252
                  Top             =   90
                  Width           =   12030
                  _cx             =   21220
                  _cy             =   5054
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
                  Rows            =   1
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmItems.frx":7320
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
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   210
                  Index           =   27
                  Left            =   11610
                  TabIndex        =   253
                  Top             =   3180
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   370
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
                  ButtonImage     =   "FrmItems.frx":743D
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   210
                  Index           =   29
                  Left            =   10530
                  TabIndex        =   254
                  Top             =   3180
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   370
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ·"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmItems.frx":79D7
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid fgCameo 
                  Height          =   2775
                  Left            =   0
                  TabIndex        =   255
                  Top             =   3450
                  Width           =   12165
                  _cx             =   21458
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmItems.frx":7F71
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
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   210
                  Index           =   26
                  Left            =   11610
                  TabIndex        =   256
                  Top             =   6270
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   370
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
                  ButtonImage     =   "FrmItems.frx":8021
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   210
                  Index           =   28
                  Left            =   10110
                  TabIndex        =   257
                  Top             =   6270
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   370
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ·"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmItems.frx":85BB
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6570
               Index           =   15
               Left            =   15135
               TabIndex        =   229
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
                  Height          =   6540
                  Index           =   16
                  Left            =   0
                  TabIndex        =   235
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   12165
                  _cx             =   21458
                  _cy             =   11536
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
                  Begin VSFlex8UCtl.VSFlexGrid GridItemsDetails2 
                     Height          =   8520
                     Left            =   0
                     TabIndex        =   236
                     Top             =   60
                     Width           =   12090
                     _cx             =   21325
                     _cy             =   15028
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
                     Rows            =   15
                     Cols            =   14
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmItems.frx":8B55
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
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6570
               Index           =   17
               Left            =   15435
               TabIndex        =   298
               TabStop         =   0   'False
               Top             =   45
               Width           =   12300
               _cx             =   21696
               _cy             =   11589
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
               Begin VSFlex8UCtl.VSFlexGrid FgSum 
                  Height          =   3780
                  Left            =   120
                  TabIndex        =   299
                  Top             =   4275
                  Width           =   12060
                  _cx             =   21272
                  _cy             =   6667
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
                  Rows            =   15
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmItems.frx":8D87
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
               Begin VSFlex8UCtl.VSFlexGrid Fg1 
                  Height          =   4260
                  Left            =   -2520
                  TabIndex        =   300
                  Top             =   0
                  Width           =   14700
                  _cx             =   25929
                  _cy             =   7514
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
                  Rows            =   15
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmItems.frx":8E0F
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
         End
         Begin C1SizerLibCtl.C1Elastic EleRight 
            Height          =   7155
            Left            =   12435
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   15
            Width           =   3420
            _cx             =   6033
            _cy             =   12621
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
            Begin MSComctlLib.TreeView TreeItems 
               Height          =   6975
               Left            =   -240
               TabIndex        =   39
               Top             =   0
               Width           =   3285
               _ExtentX        =   5794
               _ExtentY        =   12303
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   441
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               Appearance      =   1
               Enabled         =   0   'False
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   645
         Left            =   15
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   15
         Width           =   15990
         _cx             =   28205
         _cy             =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   20.25
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
         Caption         =   "»Ì«‰«  «·√’‰«›"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   10320
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox TxtCutKey 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   7575
            TabIndex        =   43
            Top             =   120
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox TxtMenuState 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   6135
            TabIndex        =   42
            Text            =   "N"
            Top             =   180
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5115
            TabIndex        =   41
            Top             =   210
            Visible         =   0   'False
            Width           =   945
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   405
            Index           =   3
            Left            =   1155
            TabIndex        =   44
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
            ButtonImage     =   "FrmItems.frx":8F89
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
            Height          =   405
            Index           =   1
            Left            =   3630
            TabIndex        =   45
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   714
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
            ButtonImage     =   "FrmItems.frx":9323
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
            Height          =   405
            Index           =   0
            Left            =   2220
            TabIndex        =   46
            Top             =   90
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   714
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
            ButtonImage     =   "FrmItems.frx":96BD
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
            Height          =   405
            Index           =   2
            Left            =   75
            TabIndex        =   47
            Top             =   90
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   714
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
            ButtonImage     =   "FrmItems.frx":9A57
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin MSComDlg.CommonDialog cdg 
            Left            =   6330
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ciaXPPopMenu30.XPPopUp30 XPPopUp 
            Left            =   5550
            Top             =   60
            _ExtentX        =   900
            _ExtentY        =   873
            VisualStyle     =   0
            BeginProperty DefaultMenuItemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuItemSpacing =   0
         End
         Begin VB.Label LblItemName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   4680
            TabIndex        =   97
            Top             =   120
            Width           =   5655
         End
         Begin VB.Label LblItemCode 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   10920
            TabIndex        =   96
            Top             =   120
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "FrmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  
Dim rs As ADODB.Recordset
Dim Rsqty As ADODB.Recordset
  Private m_DealingForm As GridTransType
Dim TTP As clstooltip
Dim ItemReport As ClsItemsReport
Dim cDboSearch(2) As clsDCboSearch
Dim cSearch(1) As clsDCboSearch
Dim first_run As Boolean
Dim FirstPeriodDateInthisYear  As Date
Public CALLEDFPRM As Boolean
Public rowbarcod As Integer
Public namebarcod As String

Public Property Get DealingForm() As GridTransType
    DealingForm = m_DealingForm
End Property

Public Property Let DealingForm(ByVal vNewValue As GridTransType)
    'If vNewValue = OpeningBalance Or vNewValue = PurchaseTransaction Or vNewValue = InvoiceTransaction Then
    m_DealingForm = vNewValue
    'End If
End Property

Private Sub DataPassing()
    Dim StrSQL As String
    Dim StrList As String
      If FrmItems.CALLEDFPRM = False Then Exit Sub
    Dim RsNote As New ADODB.Recordset
    'On Error GoTo ErrTrap
    StrSQL = "select * From TblItems"
    RsNote.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    Select Case Me.DealingForm

        Case PurchaseTransaction

            With FrmBillBuy
                StrList = .Fg.BuildComboList(RsNote, "ItemName", "ItemID")

                If StrList <> "" Then
                    .Fg.ColComboList(.Fg.ColIndex("Name")) = "|" & StrList
                End If

                StrList = .Fg.BuildComboList(RsNote, "barCodeNO", "ItemID")

                If StrList <> "" Then
                    .Fg.ColComboList(.Fg.ColIndex("Code")) = "|" & StrList
                End If

                .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Code")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
                
                .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Name")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
            
            .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Price")) = 0
            .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Count")) = 1
            
            
            Dim RsUnitData As New ADODB.Recordset
            StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
            StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
            StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & XPTxtID.text
            StrSQL = StrSQL + " AND DefaultUnit=1"
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
         '       If DefalutUnitID = 0 Then
                    .Fg.Cell(flexcpData, .Fg.Row, .Fg.ColIndex("UnitID")) = RsUnitData("UnitID").value
                    .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("UnitID")) = RsUnitData("UnitName").value
         '       Else
         '           .Cell(flexcpData, LngRow, .ColIndex("UnitID")) = DefalutUnitID
         '           .TextMatrix(LngRow, .ColIndex("UnitID")) = DefalutUnitName
         '       End If
        
            End If

            RsUnitData.Close
            Set RsUnitData = Nothing
            
            Unload Me
            End With


        
        Case INVENTORYIN

            With FrmInpout
                StrList = .Fg.BuildComboList(RsNote, "ItemName", "ItemID")

                If StrList <> "" Then
                    .Fg.ColComboList(.Fg.ColIndex("Name")) = "|" & StrList
                End If

                StrList = .Fg.BuildComboList(RsNote, "barCodeNO", "ItemID")

                If StrList <> "" Then
                    .Fg.ColComboList(.Fg.ColIndex("Code")) = "|" & StrList
                End If

                .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Code")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
                
                .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Name")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
            .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Price")) = 0
            
            
             
            StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
            StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
            StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & XPTxtID.text
            StrSQL = StrSQL + " AND DefaultUnit=1"
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
         '       If DefalutUnitID = 0 Then
                    .Fg.Cell(flexcpData, .Fg.Row, .Fg.ColIndex("UnitID")) = RsUnitData("UnitID").value
                    .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("UnitID")) = RsUnitData("UnitName").value
         '       Else
         '           .Cell(flexcpData, LngRow, .ColIndex("UnitID")) = DefalutUnitID
         '           .TextMatrix(LngRow, .ColIndex("UnitID")) = DefalutUnitName
         '       End If
        
            End If

            RsUnitData.Close
            Set RsUnitData = Nothing
            
            Unload Me
            End With


        Case ShowPrice
            StrList = frmsalebill.Fg.BuildComboList(RsNote, "ItemName", "ItemID")

            If StrList <> "" Then
                frmsalebill.Fg.ColComboList(2) = "|" & StrList
            End If

            StrList = frmsalebill.Fg.BuildComboList(RsNote, "ItemCode", "ItemID")

            If StrList <> "" Then
                frmsalebill.Fg.ColComboList(1) = "|" & StrList
            End If

            frmsalebill.Fg.TextMatrix(frmsalebill.Fg.Row, 2) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))

        Case Maintenance

            With FrmMaintenence
                StrList = .Fg.BuildComboList(RsNote, "ItemName", "ItemID")

                If StrList <> "" Then
                    .Fg.ColComboList(.Fg.ColIndex("Name")) = "|" & StrList
                End If

                StrList = .Fg.BuildComboList(RsNote, "ItemCode", "ItemID")

                If StrList <> "" Then
                    .Fg.ColComboList(.Fg.ColIndex("Code")) = "|" & StrList
                End If

                .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Code")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
                .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Name")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
            End With

            '«·—’Ìœ «·«›  «ÕÌ
        Case OpeningBalance

            With FrmOpeningBalance
                StrList = .Fg.BuildComboList(RsNote, "ItemName", "ItemID")

                If StrList <> "" Then
                    .Fg.ColComboList(.Fg.ColIndex("Name")) = "|" & StrList
                End If

                StrList = .Fg.BuildComboList(RsNote, "ItemCode", "ItemID")

                If StrList <> "" Then
                    .Fg.ColComboList(.Fg.ColIndex("Code")) = "|" & StrList
                End If

                .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Code")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
                .Fg.TextMatrix(.Fg.Row, .Fg.ColIndex("Name")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
            End With

    End Select
 CALLEDFPRM = False
   Unload Me
    Exit Sub
ErrTrap:
End Sub

Private Sub RetriveQTY1(ItemID As String)

    Dim StrSQL As String
    Dim Num As Integer
    Dim Rsqty As ADODB.Recordset
    Dim RowNum As Long
    Dim ItemTransInfo As LastItemTransInfo
  
    On Error GoTo ErrTrap
    GridItemsDetails2.Clear flexClearScrollable, flexClearEverything
 
     Set Rsqty = New ADODB.Recordset
  
StrSQL = " SELECT     dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.[Count] * dbo.TransactionTypes.StockEffect AS countsactual, "
StrSQL = StrSQL & "  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & " dbo.TblUnites.UnitName , dbo.TblUnites.UnitNamee "
StrSQL = StrSQL & " FROM         dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL & "     dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "     dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
StrSQL = StrSQL & " Where (dbo.ItemsDetails.ItemID = " & ItemID & ")"
StrSQL = StrSQL & " GROUP BY dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.[Count] * dbo.TransactionTypes.StockEffect, dbo.ItemsDetails.ColorID,"
StrSQL = StrSQL & "  dbo.ItemsDetails.UnitID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ClassId, dbo.TransactionTypes.StockEffect, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
StrSQL = StrSQL & "  dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName,"
StrSQL = StrSQL & "   dbo.TblUnites.UnitNamee "




StrSQL = "SELECT     dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, SUM(dbo.ItemsDetails.[Count] * dbo.TransactionTypes.StockEffect) AS countsactual, "
StrSQL = StrSQL & "  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "  dbo.TblUnites.Unitname , dbo.TblUnites.UnitNamee"
StrSQL = StrSQL & "  FROM         dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL & "  dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
StrSQL = StrSQL & "  Where (dbo.ItemsDetails.ItemID = " & ItemID & ")"
StrSQL = StrSQL & "  GROUP BY dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.UnitID, dbo.ItemsDetails.SizeID,"
StrSQL = StrSQL & "  dbo.ItemsDetails.ClassId, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName,"
StrSQL = StrSQL & "  dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"







  Rsqty.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 GridItemsDetails2.Clear flexClearScrollable, flexClearEverything
    If Rsqty.RecordCount < 1 Then
     
    
        Exit Sub
     
         
    End If
    
   
    
        GridItemsDetails2.Rows = Rsqty.RecordCount + 1

        For Num = 1 To Rsqty.RecordCount

            With GridItemsDetails2
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num

    
            .TextMatrix(Num, .ColIndex("Quantity")) = IIf(IsNull(Rsqty("countsactual").value), 0, (Rsqty("countsactual").value))
            
              If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(Rsqty("StoreName").value), "", (Rsqty("StoreName").value))
               Else
                 .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(Rsqty("StoreNamee").value), "", (Rsqty("StoreNamee").value))
               End If
               
                .TextMatrix(Num, .ColIndex("ColorName")) = IIf(IsNull(Rsqty("ColorName").value), "", (Rsqty("ColorName").value))
                .TextMatrix(Num, .ColIndex("ItemSize")) = IIf(IsNull(Rsqty("SizeName").value), "", (Rsqty("SizeName").value))
                .TextMatrix(Num, .ColIndex("ClassName")) = IIf(IsNull(Rsqty("cclASS NAME").value), "", (Rsqty("cclASS NAME").value))
                .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(Rsqty("UnitName").value), "", (Rsqty("UnitName").value))
                .TextMatrix(Num, .ColIndex("ItemDetailedCode")) = IIf(IsNull(Rsqty("ItemDetailedCode").value), "", (Rsqty("ItemDetailedCode").value))
            .TextMatrix(Num, .ColIndex("ParrtNoCode")) = IIf(IsNull(Rsqty("ParrtNoCode").value), "", (Rsqty("ParrtNoCode").value))
            
         '  .TextMatrix(Num, .ColIndex("ProductionDate")) = IIf(IsNull(Rsqty("ProductionDate").value), "", (Rsqty("ProductionDate").value))
            '.TextMatrix(Num, .ColIndex("ExpireDate")) = IIf(IsNull(Rsqty("ExpireDate").value), "", (Rsqty("ExpireDate").value))
         
         
            End With

            Rsqty.MoveNext
        Next Num

        GridItemsDetails2.AutoSize 0, GridItemsDetails2.Cols - 1, False
 
 
    Exit Sub
ErrTrap:
End Sub

Private Sub RetriveQTY()

    Dim StrSQL As String
    Dim Num As Integer
    Dim RsData As ADODB.Recordset
    Dim RowNum As Long
    Dim ItemTransInfo As LastItemTransInfo
    Dim RsSumQty As ADODB.Recordset

   ' On Error GoTo ErrTrap
    FG1.Clear flexClearScrollable, flexClearEverything
    FgSum.Clear flexClearScrollable, flexClearEverything

    'GetItemData 0, Trim(Me.XPTxtCode.text)

    If Not (Rsqty.EOF Or Rsqty.BOF) Then
        If True Then
            If False = True Then
            
                '    LblHaveSerial.Visible = True
            Else
                '    LblHaveSerial.Visible = True
            End If
        End If
    
        FG1.Rows = Rsqty.RecordCount + 1

        For Num = 1 To Rsqty.RecordCount

            With FG1
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num

                '    .TextMatrix(Num, .ColIndex("Serial")) = IIf(IsNull(rs("ItemSerial").value), "·«ÌÊÃœ", (rs("ItemSerial").value))
                If Not (IsNull(Rsqty("SUMQTY").value)) Then
                    .TextMatrix(Num, .ColIndex("Quantity")) = Rsqty("SUMQTY").value
                Else
                    .TextMatrix(Num, .ColIndex("Quantity")) = 0
                End If
            
                .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(Rsqty("StoreName").value), "", (Rsqty("StoreName").value))
                .TextMatrix(Num, .ColIndex("ColorName")) = IIf(IsNull(Rsqty("ColorName").value), "", (Rsqty("ColorName").value))
                .TextMatrix(Num, .ColIndex("ItemSize")) = IIf(IsNull(Rsqty("SizeName").value), "", (Rsqty("SizeName").value))
                .TextMatrix(Num, .ColIndex("ClassName")) = IIf(IsNull(Rsqty("ClassName").value), "", (Rsqty("ClassName").value))
                .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(Rsqty("UnitName").value), "", (Rsqty("UnitName").value))
                .TextMatrix(Num, .ColIndex("serial")) = IIf(IsNull(Rsqty("ItemSerial").value), "", (Rsqty("ItemSerial").value))
            
            End With

            Rsqty.MoveNext
        Next Num

        FG1.AutoSize 0, FG1.Cols - 1, False

        '  Exit Sub
        If SystemOptions.UserInterface = ArabicInterface Then
            '    Me.Lbl(2).Caption = "≈Ã„«·Ï «·ﬂ„Ì«  «·„ÊÃÊœ… : " & FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Quantity"), FG.Rows - 1, FG.ColIndex("Quantity"))
        Else
            '    Me.Lbl(2).Caption = "Total Item Stock: " & FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Quantity"), FG.Rows - 1, FG.ColIndex("Quantity"))
        End If
    
        Set RsSumQty = New ADODB.Recordset

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
 
            StrSQL = "SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.TblStore.StoreName"
            StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
            StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
            StrSQL = StrSQL + "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
            StrSQL = StrSQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
            StrSQL = StrSQL + "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
            StrSQL = StrSQL + "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
            StrSQL = StrSQL + "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
            StrSQL = StrSQL + "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
            StrSQL = StrSQL + "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColoRs.ColorID"

            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
 
            StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FirstPeriodDateInthisYear, True) & ""
            StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Date, True) & ""
            StrSQL = StrSQL + " and Item_ID =" & val(XPTxtID.text)

            StrSQL = StrSQL + " GROUP BY dbo.TblStore.StoreName "
            StrSQL = StrSQL + " HAVING      (SUM(dbo.Transaction_Details.ShowQty * dbo.TransactionTypes.StockEffect) <> 0)"
        
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
     
        End If

        RsSumQty.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsSumQty.BOF Or RsSumQty.EOF) Then

            With Me.FgSum
                RsSumQty.MoveFirst
                .Rows = .FixedRows + RsSumQty.RecordCount

                For Num = .FixedRows To .Rows - 1
                    .TextMatrix(Num, .ColIndex("NumIndex")) = Num

                    If Not (IsNull(RsSumQty("SumQty").value)) Then
                        .TextMatrix(Num, .ColIndex("Quantity")) = Round(RsSumQty("SumQty").value, SystemOptions.SysDefCurrencyForamt)
                    Else
                        .TextMatrix(Num, .ColIndex("Quantity")) = ""
                    End If

                    .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(RsSumQty("StoreName").value), "", (RsSumQty("StoreName").value))
                    RsSumQty.MoveNext
                Next Num

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        RsSumQty.Close
        Set RsSumQty = Nothing
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            '      Me.Lbl(2).Caption = "·« ÊÃœ «Ì… ﬂ„Ì«  „‰ «·’‰›"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            '      Me.Lbl(2).Caption = "There Is NO Item Stock"
        End If
    End If

    'If Me.DCboItemsName.BoundText <> "" Then
    '    StrSQL = "Select * From TblItems Where ItemID=" & Me.XPTxtID.text & ""
    '    Set RsData = New ADODB.Recordset
    '    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '    If Not (RsData.BOF Or RsData.EOF) Then
    '        Lbl(8).Caption = IIf(IsNull(RsData("SallingPrice").value), "", RsData("SallingPrice").value)
    '        Lbl(9).Caption = IIf(IsNull(RsData("CustomerPrice").value), "", RsData("CustomerPrice").value)
    '        Lbl(10).Caption = IIf(IsNull(RsData("DealerPrice").value), "", RsData("DealerPrice").value)
    '    End If
    
    '    Set RsData = New ADODB.Recordset
    '    StrSQL = "select * From ItemsPrice where Item_ID=" & Me.DCboItemsName.BoundText
    '    RsData.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '    If Not (RsData.EOF Or RsData.BOF) Then
    '        FgItemPriceList.Rows = RsData.RecordCount + 1
    '        For RowNum = 1 To RsData.RecordCount
    '            With FgItemPriceList
    '                .TextMatrix(RowNum, .ColIndex("NumIndex")) = RowNum
    '                .TextMatrix(RowNum, .ColIndex("Form")) = _
    '                IIf(IsNull(RsData("From").value), "", Trim(RsData("From").value))
    '                .TextMatrix(RowNum, .ColIndex("To")) = _
    '                IIf(IsNull(RsData("To").value), "", Trim(RsData("To").value))
    '                .TextMatrix(RowNum, .ColIndex("Price")) = _
    '                IIf(IsNull(RsData("Price").value), "", Trim(RsData("Price").value))
    '            End With
    '            RsData.MoveNext
    '        Next RowNum
    '        FgItemPriceList.AutoSize 0, FgItemPriceList.Cols - 1, False
    '    End If
    '    ItemTransInfo = GetLastItemTrans(Val(Me.DCboItemsName.BoundText))
    '    Me.Lbl(16).Caption = ItemTransInfo.TransactionSerial
    
    '    If ItemTransInfo.TransactionDate <> "" Then
    '        Me.Lbl(17).Caption = DisplayDate(CDate(ItemTransInfo.TransactionDate))
    '    End If
    '    Me.Lbl(18).Caption = ItemTransInfo.StrCustomerName
    '    Me.Lbl(19).Caption = ItemTransInfo.SngItemPrice
 
    Exit Sub
ErrTrap:
End Sub

Private Function Get_DefalutUnitFactor(IntBegineRow As Integer, _
                                       IntDefalutRow As Integer) As Double
    'Aim:
    'Argument:
    '
    Dim DblRes As Double
    Dim i As Integer
    Dim BolCalAsc As Boolean
    Dim IntForStep As Integer

    If IntBegineRow < IntDefalutRow Then
        BolCalAsc = True
        IntForStep = 1
    ElseIf IntBegineRow > IntDefalutRow Then
        BolCalAsc = False
        IntForStep = -1
    ElseIf IntBegineRow = IntDefalutRow Then
        Get_DefalutUnitFactor = 1
        Exit Function
    End If

    DblRes = 1

    With Me.FgUnites

        If BolCalAsc = True Then

            For i = IntBegineRow + 1 To IntDefalutRow Step IntForStep

                If .TextMatrix(i, .ColIndex("UnitID")) <> "" Then
                    DblRes = (DblRes * IIf(val(.TextMatrix(i, .ColIndex("UnitFactor"))) = 0, 1, val(.TextMatrix(i, .ColIndex("UnitFactor")))))
                Else
                    Exit For
                End If

            Next i

        Else

            For i = IntBegineRow To IntDefalutRow + 1 Step IntForStep

                If .TextMatrix(i, .ColIndex("UnitID")) <> "" Then
                    DblRes = (DblRes * IIf(val(.TextMatrix(i, .ColIndex("UnitFactor"))) = 0, 1, val(.TextMatrix(i, .ColIndex("UnitFactor")))))
                Else
                    Exit For
                End If

            Next i

        End If

    End With

    If BolCalAsc = True Then
        Get_DefalutUnitFactor = DblRes
    Else
        Get_DefalutUnitFactor = (1 / DblRes)
    End If

End Function

Private Function Get_SmallUnitFactor(IntBegineRow As Integer) As Double
    Dim DblRes As Double
    Dim i As Integer

    DblRes = 1

    With Me.FgUnites

        For i = IntBegineRow + 1 To .Rows - 1 Step 1

            If .TextMatrix(i, .ColIndex("UnitID")) <> "" Then
                DblRes = (DblRes * IIf(val(.TextMatrix(i, .ColIndex("UnitFactor"))) = 0, 1, val(.TextMatrix(i, .ColIndex("UnitFactor")))))
            Else
                Exit For
            End If

        Next i

    End With

    Get_SmallUnitFactor = DblRes
End Function

Private Sub SaveData_Prices()

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim Msg As String
    Dim lngCount As Long
    Dim IntDefUnitRow As Integer

    StrSQL = "Delete  From TblSalesPrices Where ItemID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords

    Set rs = New ADODB.Recordset
  '  rs.Open "TblSalesPrices", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT     *  from dbo.TblSalesPrices Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    With Me.FgSalePrice

        If Me.FgSalePrice.Rows <> 1 Then

            For i = Me.FgSalePrice.FixedRows To Me.FgSalePrice.Rows - 1

                If val(.TextMatrix(i, .ColIndex("BranchId"))) > 0 Then
                    rs.AddNew
                    rs("ItemID").value = val(Me.XPTxtID.text) 'Val(Me.DcboItems1.BoundText)
                    rs("BranchId").value = val(.TextMatrix(i, .ColIndex("BranchId")))
                    rs("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                    rs("Price1").value = val(.TextMatrix(i, .ColIndex("Price1")))
                    rs("Price2").value = val(.TextMatrix(i, .ColIndex("Price2")))
                    rs("Price3").value = val(.TextMatrix(i, .ColIndex("Price3")))
                    rs("Price4").value = val(.TextMatrix(i, .ColIndex("Price4")))
                    rs("Price5").value = val(.TextMatrix(i, .ColIndex("Price5")))
                    rs("Price6").value = val(.TextMatrix(i, .ColIndex("Price6")))
              
                    rs("Discount1").value = val(.TextMatrix(i, .ColIndex("Discount1")))
                    rs("Discount2").value = val(.TextMatrix(i, .ColIndex("Discount2")))
                    rs("Discount3").value = val(.TextMatrix(i, .ColIndex("Discount3")))
                    rs("Discount4").value = val(.TextMatrix(i, .ColIndex("Discount4")))
                    rs("Discount5").value = val(.TextMatrix(i, .ColIndex("Discount5")))
                    rs("Discount6").value = val(.TextMatrix(i, .ColIndex("Discount6")))
             
                    rs.update
                End If

            Next i

        Else
 
        End If
 
    End With

    rs.Close
    Set rs = Nothing
 
End Sub

Private Sub SaveData_Pricesold()

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim Msg As String
    Dim lngCount As Long
    Dim IntDefUnitRow As Integer
  
    For i = Me.FgPrices.FixedRows To Me.FgPrices.Rows

        If Me.FgPrices.Rows <> 1 Then
            If FgPrices.Cell(flexcpChecked, i, FgPrices.ColIndex("DefaultUnit")) = flexChecked Then
                IntDefUnitRow = i
                Exit For
            End If
        End If

    Next i

    StrSQL = "Delete  From TblItemsPrices Where ItemID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords

    Set rs = New ADODB.Recordset
    rs.Open "TblItemsPrices", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Me.FgPrices

        If Me.FgPrices.Rows <> 1 Then

            For i = Me.FgPrices.FixedRows To Me.FgPrices.Rows - 1
                rs.AddNew
                rs("ItemID").value = val(Me.XPTxtID.text) 'Val(Me.DcboItems1.BoundText)
                rs("PriceId").value = i
                rs("PriceName").value = .TextMatrix(i, .ColIndex("PriceName"))
                rs("Pricevalue").value = val(.TextMatrix(i, .ColIndex("Pricevalue")))
                rs("des").value = .TextMatrix(i, .ColIndex("des"))
                rs("CustomerOrVendor").value = 0

                If .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexChecked Then
                    rs("DefaultUnit").value = 1
                Else
                    rs("DefaultUnit").value = 0
                End If
            
                rs.update
            Next i

        Else
 
        End If
 
    End With

    rs.Close
    Set rs = Nothing
 
    For i = Me.FgPrices1.FixedRows To Me.FgPrices1.Rows

        If Me.FgPrices1.Rows <> 1 Then
            If FgPrices1.Cell(flexcpChecked, i, FgPrices1.ColIndex("DefaultUnit")) = flexChecked Then
                IntDefUnitRow = i
                Exit For
            End If
        End If

    Next i
 
    Set rs = New ADODB.Recordset
    rs.Open "TblItemsPrices", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Me.FgPrices1

        If Me.FgPrices1.Rows <> 1 Then

            For i = Me.FgPrices1.FixedRows To Me.FgPrices1.Rows - 1
                rs.AddNew
                rs("ItemID").value = val(Me.XPTxtID.text) 'Val(Me.DcboItems1.BoundText)
                rs("PriceId").value = i
                rs("PriceName").value = .TextMatrix(i, .ColIndex("PriceName"))
                rs("Pricevalue").value = val(.TextMatrix(i, .ColIndex("Pricevalue")))
                rs("des").value = .TextMatrix(i, .ColIndex("des"))
                rs("CustomerOrVendor").value = 1

                If .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexChecked Then
                    rs("DefaultUnit").value = 1
                Else
                    rs("DefaultUnit").value = 0
                End If
            
                rs.update
            Next i

        Else
 
        End If
 
    End With

    rs.Close
    Set rs = Nothing

    If SystemOptions.UserInterface = ArabicInterface Then
        'Msg = " „  ⁄„·Ì… «·Õ›Ÿ...!!!"
    Else
        'Msg = "Saved........ !"
    End If

    'MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

End Sub
Private Function ItemsInGrid() As Long
    Dim i As Long
    Dim BolTemp As Boolean
    On Error GoTo ErrTrap

    With FgUnites

        If Trim(.TextMatrix(.FixedRows, FgUnites.ColIndex("UnitID"))) = "" Then
            ItemsInGrid = -1
        Else
            ItemsInGrid = 1
        End If

    End With

    Exit Function
ErrTrap:
    ItemsInGrid = -1
End Function


Private Function GetFgCheckCount() As Long

    Dim i As Long
    Dim IntCount As Long

    With FgUnites

        For i = .FixedRows To .Rows - 1

            If .Cell(flexcpChecked, i, FgUnites.ColIndex("DefaultUnit")) = flexChecked Then
                IntCount = IntCount + 1
            End If

        Next i

    End With

    GetFgCheckCount = IntCount
End Function
Private Sub SaveData_Unites()

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim Msg As String
    Dim lngCount As Long
    Dim IntDefUnitRow As Integer
    'If Val(Me.DcboItems1.BoundText) = 0 Then
    '    Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰› ...!!!"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If
    lngCount = ItemsInGrid()
    If lngCount = 0 Then
        Msg = "ÌÃ» ≈œŒ«· ÊÕœ… ⁄·Ï «·√ﬁ· ....!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
ElseIf Me.FgUnites.FixedRows + 1 = Me.FgUnites.Rows Then
        With Me.FgUnites
           .Cell(flexcpChecked, 1, .ColIndex("DefaultUnit")) = flexChecked
       End With
    Else
        If GetFgCheckCount() = 0 Then
        Msg = "ÌÃ»  ÕœÌœ ÊÕœ… ≈› —«÷Ì… ··’‰› ....!!!"
           MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    For i = Me.FgUnites.FixedRows To Me.FgUnites.Rows - 1

        If Me.FgUnites.Rows <> 1 Then
            If FgUnites.Cell(flexcpChecked, i, FgUnites.ColIndex("DefaultUnit")) = flexChecked Then
                IntDefUnitRow = i
                Exit For
            End If
        End If

    Next i

    StrSQL = "Delete  From TblItemsUnits Where ItemID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    Set rs = New ADODB.Recordset
   ' rs.Open "TblItemsUnits", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT     *  from dbo.TblItemsUnits Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    With FgUnites

        If Me.FgUnites.Rows <> 1 Then

            For i = Me.FgUnites.FixedRows To Me.FgUnites.Rows - 1
                rs.AddNew
                rs("ItemID").value = val(Me.XPTxtID.text) 'Val(Me.DcboItems1.BoundText)
                rs("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                rs("UnitFactor").value = val(.TextMatrix(i, .ColIndex("UnitFactor")))
                rs("UnitSalesPrice").value = val(.TextMatrix(i, .ColIndex("UnitSalesPrice")))
                rs("UnitPurPrice").value = val(.TextMatrix(i, .ColIndex("UnitPurPrice")))

                If .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexChecked Then
                    rs("DefaultUnit").value = 1
                Else
                    rs("DefaultUnit").value = 0
                End If

                rs("SecOrder").value = val(.TextMatrix(i, .ColIndex("SecOrder")))
                .TextMatrix(i, .ColIndex("FactorByDefaultUnit")) = Format(Get_DefalutUnitFactor(CInt(i), IntDefUnitRow), "0.000")
                rs("FactorByDefaultUnit").value = val(.TextMatrix(i, .ColIndex("FactorByDefaultUnit")))
            
                .TextMatrix(i, .ColIndex("FactorBySmallUnit")) = Format(Get_SmallUnitFactor(CInt(i)), "0.000")
                rs("FactorBySmallUnit").value = val(.TextMatrix(i, .ColIndex("FactorBySmallUnit")))
            
                rs.update
            Next i

        Else

            If Not Me.TxtModFlg.text = "E" Then
                rs.AddNew
                rs("ItemID").value = val(Me.XPTxtID.text)           'Val(Me.XPTxtID.text) 'Val(Me.DcboItems1.BoundText)
                rs("UnitID").value = 1
                rs("UnitFactor").value = 1
                rs("UnitSalesPrice").value = val(XPTxtSall.text)
                rs("UnitPurPrice").value = val(XPTxtPurchase.text)
                rs("DefaultUnit").value = 1
                rs("SecOrder").value = 1
                ' .TextMatrix(I, .ColIndex("FactorByDefaultUnit")) = 1
                rs("FactorByDefaultUnit").value = 1
            
                '.TextMatrix(I, .ColIndex("FactorBySmallUnit")) = 1
                rs("FactorBySmallUnit").value = 1
            
                rs.update
            End If
        End If

    End With

    rs.Close
    Set rs = Nothing

    If SystemOptions.UserInterface = ArabicInterface Then
        'Msg = " „  ⁄„·Ì… «·Õ›Ÿ...!!!"
    Else
        'Msg = "Saved........ !"
    End If

    'MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub RemoveFgRow2()

    With Me.FgPrices1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub

Private Sub RemoveFgRow6()

    With Me.fgCameo

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub
Private Sub RemoveFgRow7()

    With Me.fgDiamonds

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub
Private Sub RemoveFgRow1()

    With Me.FgSalePrice

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub

Private Sub RemoveFgRow1old()

    With Me.FgPrices

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub

Private Sub RemoveFgRow()

    With Me.FgUnites

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub

Private Sub CboItemType_Change()

    If CboItemType.ListIndex = -1 Then
        Exit Sub
    ElseIf CboItemType.ListIndex = 0 Then
        lbl(8).Enabled = True
        lbl(5).Enabled = True
        lbl(7).Enabled = True
        lbl(10).Enabled = True
        lbl(11).Enabled = True
        TxtRequired.Enabled = True
        XPTxtPurchase.Enabled = True
        XPTxtSall.Enabled = True
        TxtCusPrice.Enabled = True
        TxtDealerPrice.Enabled = True
    
    ElseIf CboItemType.ListIndex = 1 Then
        lbl(8).Enabled = False
        lbl(5).Enabled = False
        lbl(7).Enabled = True
        lbl(10).Enabled = False
        lbl(11).Enabled = False
        TxtRequired.Enabled = False
        XPTxtPurchase.Enabled = False
        XPTxtSall.Enabled = True
        TxtCusPrice.Enabled = False
        TxtDealerPrice.Enabled = False
    End If

End Sub

Private Sub CboItemType_Click()
    CboItemType_Change
End Sub

Sub activateass()

    If ChkAssplied.value = vbChecked Then
        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            Me.lbl(17).Enabled = True
            Me.lbl(18).Enabled = True
            Me.lbl(19).Enabled = True
            Me.lbl(20).Enabled = True
        
            Me.TxtItemCode.Enabled = True
            Me.DcboItems.Enabled = True
            Me.TxtItemQty(0).Enabled = True
            Me.TxtItemPrice(0).Enabled = True
            Me.Cmd(8).Enabled = True
            Me.Cmd(9).Enabled = True
            Fg.Visible = True
            '        Ele(1).Visible = True
            '     Ele(1).Width = Ele(7).Width
        End If

    Else
    
        Me.lbl(17).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(19).Enabled = False
        Me.lbl(20).Enabled = False
        Me.TxtItemCode.Enabled = False
        Me.DcboItems.Enabled = False
        Me.TxtItemQty(0).Enabled = False
        Me.TxtItemPrice(0).Enabled = False
        Me.Cmd(8).Enabled = False
        Me.Cmd(9).Enabled = False
    End If

End Sub

Private Sub ChkAssplied_Click()
    activateass
End Sub

Private Sub chkItemMaking_Click()
    'If chkItemMaking.value = vbChecked Then
    'FG.Visible = True
    ''Ele(1).Visible = True
 
    'Ele(1).Width = Ele(7).Width
    'End If

    If chkItemMaking.value = vbChecked Then
        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            Me.lbl(17).Enabled = True
            Me.lbl(18).Enabled = True
            Me.lbl(19).Enabled = True
            Me.lbl(20).Enabled = True
        
            Me.TxtItemCode.Enabled = True
            Me.DcboItems.Enabled = True
            Me.TxtItemQty(0).Enabled = True
            Me.TxtItemPrice(0).Enabled = True
            Me.Cmd(8).Enabled = True
            Me.Cmd(9).Enabled = True
            Fg.Visible = True
            '        Ele(1).Visible = True
            '     Ele(1).Width = Ele(7).Width
        End If

    Else
    
        Me.lbl(17).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(19).Enabled = False
        Me.lbl(20).Enabled = False
        Me.TxtItemCode.Enabled = False
        Me.DcboItems.Enabled = False
        Me.TxtItemQty(0).Enabled = False
        Me.TxtItemPrice(0).Enabled = False
        Me.Cmd(8).Enabled = False
        Me.Cmd(9).Enabled = False
    End If

End Sub

Private Sub ChkItemMakingNew_Click()

    If ChkItemMakingNew.value = vbChecked Then
        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            Me.lbl(17).Enabled = True
            Me.lbl(18).Enabled = True
            Me.lbl(19).Enabled = True
            Me.lbl(20).Enabled = True
        
            Me.TxtItemCode.Enabled = True
            Me.DcboItems.Enabled = True
            Me.TxtItemQty(0).Enabled = True
            Me.TxtItemPrice(0).Enabled = True
            Me.Cmd(8).Enabled = True
            Me.Cmd(9).Enabled = True
            Fg.Visible = True
            '        Ele(1).Visible = True
            '     Ele(1).Width = Ele(7).Width
        End If

    Else
    
        Me.lbl(17).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(19).Enabled = False
        Me.lbl(20).Enabled = False
        Me.TxtItemCode.Enabled = False
        Me.DcboItems.Enabled = False
        Me.TxtItemQty(0).Enabled = False
        Me.TxtItemPrice(0).Enabled = False
        Me.Cmd(8).Enabled = False
        Me.Cmd(9).Enabled = False
    End If

End Sub

Private Sub ChkRelated_Click()

    If ChkRelated.value = vbChecked Then
        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            Me.lbl(23).Enabled = True
            Me.lbl(24).Enabled = True
            Me.lbl(25).Enabled = True
            Me.lbl(26).Enabled = True
        
            Me.TxtAttachedItemCode.Enabled = True
            Me.DcboItemID1.Enabled = True
            Me.TxtItemQty(1).Enabled = True
            Me.TxtItemPrice(1).Enabled = True
            Me.Cmd(10).Enabled = True
            Me.Cmd(11).Enabled = True
        End If

    Else
    
        '   Me.Lbl(23).Enabled = False
        Me.lbl(24).Enabled = False
        Me.lbl(25).Enabled = False
        Me.lbl(26).Enabled = False
    
        Me.TxtAttachedItemCode.Enabled = False
        Me.DcboItemID1.Enabled = False
        Me.TxtItemQty(1).Enabled = False
        Me.TxtItemPrice(1).Enabled = False
        Me.Cmd(10).Enabled = False
        Me.Cmd(11).Enabled = False
    End If

End Sub

Public Sub Cmd_Click(Index As Integer)
'   On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    Select Case Index





        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

    TxtModFlg.text = "N"
     With VSFlexGrid2
          .Clear flexClearScrollable, flexClearEverything
     .Rows = 2
  End With
      With fgCameo
          .Clear flexClearScrollable, flexClearEverything
     .Rows = .FixedRows
  End With
  
            With VSFlexGrid1
          .Clear flexClearScrollable, flexClearEverything
     .Rows = .FixedRows
  End With
     With fgDiamonds
          .Clear flexClearScrollable, flexClearEverything
     .Rows = .FixedRows
  End With
             
             Me.fgCameo.Enabled = True
            Me.fgCameo.Rows = 2
            
            Me.fgDiamonds.Enabled = True
            Me.fgDiamonds.Rows = 2
            
            
            SetMeForNew
            'XPTxtID.text = CStr(new_id("TblItems", "ItemID", "", True))
          '  Set RsTemp = New ADODB.Recordset
          '  RsTemp.Open "TblItems", Cn, adOpenStatic, adLockOptimistic, adCmdTable

          '  If Not (RsTemp.EOF Or RsTemp.BOF) Then
          '      RsTemp.MoveLast
          '      XPTxtName.text = IIf(IsNull(RsTemp("ItemName").value), "", RsTemp("ItemName").value)
          '      RsTemp.Close
          '  Else
          '      RsTemp.Close
          '  End If

            'XPTxtCode.SetFocus
            Frame1.Enabled = True
            XPTxtName.text = ""
            OptGaurType(0).value = True
ChkDef.value = vbChecked
TxtUnitFactor.text = 1
DcboUnits.BoundText = 1
LblItemCode.Caption = ""
LblItemName.Caption = ""

        Case 1
    
            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
              Me.VSFlexGrid2.Enabled = True
            Me.VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
            
              Me.fgCameo.Enabled = True
            Me.fgCameo.Rows = fgCameo.Rows + 1
            Me.fgDiamonds.Enabled = True
            Me.fgDiamonds.Rows = fgDiamonds.Rows + 1
            Frame1.Enabled = True
            CuurentLogdata
            '        ChkAssplied_Click
            activateass
ChkItemMakingNew_Click
        Case 2
 If SystemOptions.WorkWithGroupCode = True Then

            If DCPreFix.text = "" Then
              MsgBox "Õœœ «·Ã“¡ «·À«Ì  „‰ «·„Ã„Ê⁄Â"
                'DCPreFix.SetFocus
            '    SendKeys "{F4}"

             '   Exit Sub
            End If
End If

            Dim currentcode As String

            If txtid.text = "" Then
                currentcode = get_coding(branch_id, "TblItems", 3, Me.DCPreFix.text)

                If currentcode = "miniError" Then
                    MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Exit Sub
                                
                ElseIf currentcode = "Manual" Then
                    MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                Else
                    txtid = currentcode
                End If

            Else
                currentcode = txtid
            End If

            XPTxtCode = DCPreFix.text & currentcode

            If val(XPTxtSall.text) = 0 Then
                XPTxtSall.text = val(Me.TxtPrice(0).text)
            End If
  Me.C1Tab1.CurrTab = 0
  If XPTxtNamee.text = "" Then
  XPTxtNamee.text = XPTxtName.text
  End If
  
            SaveData
             
  
      
  
            'SaveData_Unites
         ' Frame1.Enabled = False

        Case 3
            Undo

        Case 4
    
            If CheckItemsIntransactions(val(XPTxtID)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰ Õ–› Â–« «·’‰› ·ÊÃÊœ Õ—ﬂ«  ⁄·Ì…", vbCritical
                Else
                    MsgBox "Cant Delete", vbCritical
            
                End If

                Exit Sub
    
            End If
    
            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Item

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            Load FrmItemSearch
            FrmItemSearch.RetrunType = 0
            FrmItemSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If
 If C1Tab1.CurrTab = 7 Then
print_report
Else
            PrintReport
End If
        Case 8
            AddNewFgRow

        Case 9
            DeleteFgRow

        Case 10
            AddNewFgAttachRow

        Case 11
            DeleteFgAttachRow
   Case 27
   RemoveFgRow7
 Case 26
   RemoveFgRow6
       Case 25
    DeleteFgRowAther
Case 30
Load FrmInputBarcode
FrmInputBarcode.show
   
        Case 20
    
    
            AddNewRow
DcboUnits.Enabled = True
TxtUnitFactor.Enabled = True
        Case 21

            If CheckItemsIntransactions(val(XPTxtID)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Â–« «·’‰› ·ÊÃÊœ Õ—ﬂ«  ⁄·Ì…", vbCritical
                Else
                    MsgBox "Cant Modify", vbCritical
                        
                End If

                Exit Sub
    
            End If
    
            RemoveFgRow

        Case 22
            Unload Me

        Case 23
            SaveData_Unites
 
        Case 14
            AddNewRow1

        Case 15
            RemoveFgRow1

        Case 18
            AddNewRow2
            '      Case 19
            '       RemoveFgRow2
           Case 24
            AddNewFgRowother
           Case 28
            Me.fgCameo.Clear flexClearScrollable, flexClearEverything
            fgCameo.Rows = 2
            Me.fgCameo.Enabled = True
           Case 29
           Me.fgDiamonds.Clear flexClearScrollable, flexClearEverything
            fgDiamonds.Rows = 2
            Me.fgDiamonds.Enabled = True
        
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub AddNewFgRowother()

    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long

    If val(Me.Dcbiteem.BoundText) = 0 Then
        Msg = "  ÌÃ»  ÕœÌœ «”„ «·’‰›"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.Dcbiteem.SetFocus
        Exit Sub
    End If

   ' If Me.TxtModFlg.text = "E" Then
   '     If val(Me.DcboItems.BoundText) = val(Me.XPTxtID.text) Then
   '         Msg = "?????? ?? ???? ????? ??? ?? ????....!!!"
   '         MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
   '         Me.DcboItems.SetFocus
   '         Exit Sub
   '     End If
   ' End If

    If val(Me.TxtItemQty(2).text) = 0 Then
        Msg = " ÌÃ»  ÕœÌœ ﬂ„ÌÂ «·’‰›"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtItemQty(2).SetFocus
        Exit Sub
    End If

    If val(Me.TxtItemPrice(2).text) = 0 Then
        Msg = " ÌÃ»  ÕœÌœ  ﬂ·›… «·’‰›"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtItemPrice(2).SetFocus
        Exit Sub
    End If

    If val(Me.Dcbuniit.BoundText) = 0 Then
        Msg = " ÌÃ»  ÕœÌœ ÊÕœ… «·’‰›"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.Dcbuniit.SetFocus
        Exit Sub
    End If

    With Me.VSFlexGrid1
        LngFindRow = .FindRow(val(Me.Dcbiteem.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)

        If LngFindRow <> -1 Then
            Msg = "Â–« «·’‰› „ÊÃÊœ ›⁄·«"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            .SetFocus
            Exit Sub
        End If

    End With

    LngNewRow = ModFgLib.SetFgForNewRow(VSFlexGrid1, VSFlexGrid1.ColIndex("ItemID"))

    With Me.VSFlexGrid1
        .TextMatrix(LngNewRow, .ColIndex("ItemID")) = Me.Dcbiteem.BoundText
    
        .TextMatrix(LngNewRow, .ColIndex("ItemCode")) = Trim$(Me.TxtCodeAother.text)
        .TextMatrix(LngNewRow, .ColIndex("ItemName")) = Me.Dcbiteem.BoundText
    
        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = Me.Dcbuniit.BoundText
        .TextMatrix(LngNewRow, .ColIndex("UnitName")) = Me.Dcbuniit.text
        .TextMatrix(LngNewRow, .ColIndex("ItemQty")) = val(Me.TxtItemQty(2).text)
        .TextMatrix(LngNewRow, .ColIndex("ItemPrice")) = val(Me.TxtItemPrice(2).text)
        .TextMatrix(LngNewRow, .ColIndex("Remarks")) = Me.TxtRemark.text
        .AutoSize 0, .Cols - 1, False
    End With

    Me.lbl(38).Caption = ModFgLib.GetItemsInFg(VSFlexGrid1, VSFlexGrid1.ColIndex("ItemID"))

    Me.TxtCodeAother.text = ""
    Me.DcboItems.BoundText = ""
    Me.TxtItemQty(2).text = ""
    Me.TxtItemPrice(2).text = ""
    TxtRemark.text = ""
    Me.TxtCodeAother.SetFocus
End Sub



Private Sub DcboItems_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

      Load FrmItemSearch
            FrmItemSearch.RetrunType = 31
            FrmItemSearch.show vbModal
End If

If KeyCode = vbKeyF5 Then
    Dim Dcombos As New ClsDataCombos
   
    Dcombos.GetItemsNames Me.DcboItems
    
End If

End Sub

Private Sub TxtCodeAother_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtCodeAother.text = "" Then
            Me.Dcbiteem.BoundText = ""
        Else
            Me.Dcbiteem.BoundText = GetItemID(Trim$(Me.TxtCodeAother.text))
        End If
    End If
End Sub


Private Sub Dcbiteem_Change()
 Dim unitid As Long
    Dim Unitname As String
    Me.TxtCodeAother.text = GetItemCode(val(Me.Dcbiteem.BoundText))
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsUnits·byitemid Me.Dcbuniit, val(Me.Dcbiteem.BoundText)
  
    GetDefaultItemUnit val(Me.Dcbiteem.BoundText), unitid, Unitname
    Dcbuniit.text = Unitname
    Dcbuniit.BoundText = unitid
    Me.TxtItemPrice(2).text = ModItemCostPrice.GetCostItemPrice(val(Me.Dcbiteem.BoundText), , , , SystemOptions.SysMainStockCostMethod, , , Date, , unitid)
End Sub

Private Sub Dcbiteem_Click(Area As Integer)
 Dcbiteem_Change
End Sub


Private Sub CmdAttach_Click()
            On Error Resume Next
ShowAttachments DCPreFix.text & txtid.text, "0701201407"
 

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub
Function addtotable(NoOfRow As Integer, code As String, cost As Double, Optional PartNo As String = "", Optional name As String = "" _
, Optional NameE As String, Optional Color As String, Optional size As String, Optional Class As String, Optional itemcode As String)
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim i As Integer

    str = "select * from   TblPrintBarCode where 1=-1"
    
   rs.Open str, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  For i = 1 To NoOfRow
        rs.AddNew
   
        rs("PartNo").value = PartNo
        rs("code").value = code
        rs("cost").value = val(cost)
        rs("Name").value = name
'        rs("NameE").value = NameE
        rs("Color").value = Color
        rs("size").value = size
        rs("class").value = Class
        rs.update
    Next i
'
End Function

Public Sub PrintBarCode(Optional rowcode As Integer = 0, Optional nameBar As String)
  Dim str, code As String

    Dim RowNum As Integer
    Dim ItemCount As Integer
    str = "Delete  TblPrintBarCode"
    Cn.Execute str
DoEvents
Dim LngItemID As Long
Dim LngUnitID As Long
    'cBarcode.AddItem
    ' cBarcode.ClearItems
  

   ' LngItemID = val(TxtItemID.text)
   ' LngUnitID = val(TxtUnitID.text)
  
code = TxtbarCodeNO.text

       ' If Grid.Cell(flexcpChecked, RowNum, Grid.ColIndex("Print")) = flexChecked Then
       '     If Not IsNull(Grid.TextMatrix(RowNum, Grid.ColIndex("ActCount"))) Then
           
      addtotable rowcode, code, val(FgUnites.TextMatrix(1, FgUnites.ColIndex("UnitSalesPrice"))), TxtPartNo.text, XPTxtName.text, XPTxtNamee.text
      'val(Grid.TextMatrix(RowNum, Grid.ColIndex("ActCount"))), Grid.TextMatrix(RowNum, Grid.ColIndex("ParrtNoCode")), GetItemPrice(LngItemID, 1, LngUnitID), g
     '
     '       End If
     '   End If



    printCodeBarcode WindowTarget, nameBar
End Sub
Private Sub CmdPic_Click(Index As Integer)
On Error GoTo ErrTrap
    Select Case Index

        Case 0

            With cdg
                '*.jpg,*.jpeg,*.jpe,*.jfif
                .CancelError = False
                .DialogTitle = " ≈Œ Ì«— ’Ê—…"
                'Set The Filter to show pictures only
                .Filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|"  ' choose formats to include
             '& "GIF (*.gif)|*.gif|All Files|*.*"
                .ShowOpen

                If .FileName <> "" Then
                    Set Me.ImgPic.Picture = LoadPicture(.FileName)
                End If

            End With

        Case 1
            Set Me.ImgPic.Picture = Nothing
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " ÕÃ„ «·’Ê—… €Ì— „œ⁄Ê„", vbCritical
Else
MsgBox " image Size Not Siutable, vbCritical"
End If


End Sub

Private Sub Command1_Click()

    StrSQL = "Delete  From TblItemsUnits "
    Cn.Execute StrSQL, , adExecuteNoRecords
 
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    rs.Open "TblItemsUnits", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    Rs1.Open "TblItems", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    For i = 1 To Rs1.RecordCount
        rs.AddNew
        rs("ItemID").value = val(Rs1("ItemID").value)  'Val(Me.DcboItems1.BoundText)
        rs("UnitID").value = 1
        rs("UnitFactor").value = 1
        rs("UnitSalesPrice").value = 0
        rs("UnitPurPrice").value = 0
           
        rs("DefaultUnit").value = 1
          
        rs("SecOrder").value = 1
           
        rs("FactorByDefaultUnit").value = 1
            
        rs("FactorBySmallUnit").value = 1
            
        rs.update
        Rs1.MoveNext
    Next i

    MsgBox "Done"
   
End Sub

Private Sub Command2_Click()
    StrSQL = "Delete  From TblSalesPrices "
    Cn.Execute StrSQL, , adExecuteNoRecords
 
    Dim rs As New ADODB.Recordset
    Dim rsBranch As New ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Dim unitid As Long
    Dim Unitname As String
    Dim i As Integer
    Dim J  As Integer
    Set Rs1 = New ADODB.Recordset
 
    Rs1.Open "TblItems", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rsBranch.Open "TblBranchesData", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.Open "TblSalesPrices", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    For i = 1 To Rs1.RecordCount
        rsBranch.MoveFirst
        GetDefaultItemUnit val(XPTxtID.text), unitid, Unitname

        For J = 1 To rsBranch.RecordCount
             
            rs.AddNew
            rs("ItemID").value = val(Rs1("ItemID").value)  'Val(Me.DcboItems1.BoundText)
            rs("UnitID").value = unitid
            rs("Price1").value = IIf(IsNull(Rs1("SallingPrice").value), 0, (Rs1("SallingPrice").value))
            rs("Price2").value = IIf(IsNull(Rs1("CustomerPrice").value), 0, (Rs1("CustomerPrice").value))
            rs("Price3").value = IIf(IsNull(Rs1("DealerPrice").value), 0, (Rs1("DealerPrice").value))
            rs("BranchId").value = val(rsBranch("branch_id").value)
            rs("Price4").value = 0
            rs("Price5").value = 0
            rs("Price6").value = 0
            rs("Discount1").value = 0
            rs("Discount2").value = 0
            rs("Discount3").value = 0
            rs("Discount4").value = 0
            rs("Discount5").value = 0
            rs("Discount6").value = 0
                       
            rs.update
            
            rsBranch.MoveNext
        Next J

        Rs1.MoveNext
    Next i

    MsgBox "Done"

End Sub

Private Sub DcboItemID1_Change()
    Me.TxtAttachedItemCode.text = GetItemCode(val(Me.DcboItemID1.BoundText))
End Sub

Private Sub DcboItemID1_Click(Area As Integer)
    DcboItemID1_Change
End Sub

Private Sub DcboItems_Change()
    Dim unitid As Long
    Dim Unitname As String
    Me.TxtItemCode.text = GetItemCode(val(Me.DcboItems.BoundText))
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsUnits·byitemid Me.dcItemunit, val(Me.DcboItems.BoundText)
    GetDefaultItemUnit val(Me.DcboItems.BoundText), unitid, Unitname
    dcItemunit.text = Unitname
    dcItemunit.BoundText = unitid
    Me.TxtItemPrice(0).text = ModItemCostPrice.GetCostItemPrice(val(Me.DcboItems.BoundText), , , , SystemOptions.SysMainStockCostMethod, , , Date, , unitid)

End Sub

Private Sub DcboItems_Click(Area As Integer)
    DcboItems_Change
End Sub

Public Sub DcboItems1_Change()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'If Val(Me.DcboItems1.BoundText) = 0 Then
    '    Me.FgUnites.Rows = Me.FgUnites.FixedRows
    '    Exit Sub
    'End If

    StrSQL = "SELECT TblItemsUnits.JunckID, TblItemsUnits.ItemID, TblItemsUnits.UnitID," & "TblUnites.UnitName,TblUnites.UnitNamee, TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder,TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice,TblItemsUnits.UnitPurPrice"
    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits ON TblUnites.UnitID =" & "TblItemsUnits.UnitID "
    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & val(Me.XPTxtID.text)
StrSQL = StrSQL & "order by UnitFactor"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FgUnites
            .Rows = Me.FgUnites.FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1

                If rs("DefaultUnit").value = 1 Then
                    .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexUnchecked
                End If

                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(rs("UnitID").value), "", rs("UnitID").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs("UnitName").value), "", rs("UnitName").value)
                Else
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs("UnitNamee").value), "", rs("UnitNamee").value)
                End If

                .TextMatrix(i, .ColIndex("UnitFactor")) = IIf(IsNull(rs("UnitFactor").value), "", rs("UnitFactor").value)
            
                .TextMatrix(i, .ColIndex("UnitSalesPrice")) = IIf(IsNull(rs("UnitSalesPrice").value), "", rs("UnitSalesPrice").value)
                .TextMatrix(i, .ColIndex("UnitPurPrice")) = IIf(IsNull(rs("UnitPurPrice").value), "", rs("UnitPurPrice").value)
            
                .TextMatrix(i, .ColIndex("SecOrder")) = IIf(IsNull(rs("SecOrder").value), "", rs("SecOrder").value)
                WriteDes CLng(i)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    Else
        Me.FgUnites.Rows = Me.FgUnites.FixedRows
        Exit Sub
    End If

    rs.Close
    Set rs = Nothing
    ViewPrices
End Sub

Function PrepareFgSalePrice()
    Dim i As Integer
    Dim RsPrepareFgSalePrice As ADODB.Recordset

    'StrSQL = "SELECT  * from TblSalesPrices    "
 
    'Prepare Grid1$$$$$$$$$$$$4
    Dim column_location As Integer

    For i = 0 To 5
        lblPrice(i).Visible = False
        lblDiscount(i).Visible = False
               
        TxtPrice(i).Visible = False
        TxtDiscount(i).Visible = False
        
    Next i
     
    Dim NoOfColumns As Integer

    With Me.FgSalePrice
        StrSQL = "SELECT  * from TblSalePriceNames    "
        Set RsPrepareFgSalePrice = New ADODB.Recordset
        RsPrepareFgSalePrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsPrepareFgSalePrice.RecordCount > 0 Then
            NoOfColumns = RsPrepareFgSalePrice.RecordCount

            If NoOfColumns > 6 Then
                NoOfColumns = 6
            End If

            For i = 0 To NoOfColumns - 1
                '              On Error Resume Next
                .ColHidden(.ColIndex("Price" & RsPrepareFgSalePrice.Fields("id").value)) = False
                .ColHidden(.ColIndex("Discount" & RsPrepareFgSalePrice.Fields("id").value)) = False
     
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(0, .ColIndex("Price" & RsPrepareFgSalePrice.Fields("id").value)) = IIf(IsNull(RsPrepareFgSalePrice.Fields("PriceName").value), 0, RsPrepareFgSalePrice.Fields("PriceName").value)
                    .TextMatrix(0, .ColIndex("Discount" & RsPrepareFgSalePrice.Fields("id").value)) = IIf(IsNull(RsPrepareFgSalePrice.Fields("DiscountName").value), 0, RsPrepareFgSalePrice.Fields("DiscountName").value)
                    lblPrice(i).Caption = IIf(IsNull(RsPrepareFgSalePrice.Fields("PriceName").value), 0, RsPrepareFgSalePrice.Fields("PriceName").value)
                    lblDiscount(i).Caption = IIf(IsNull(RsPrepareFgSalePrice.Fields("DiscountName").value), 0, RsPrepareFgSalePrice.Fields("DiscountName").value)
                Else
                    .TextMatrix(0, .ColIndex("Price" & RsPrepareFgSalePrice.Fields("id").value)) = IIf(IsNull(RsPrepareFgSalePrice.Fields("PriceNameE").value), 0, RsPrepareFgSalePrice.Fields("PriceNameE").value)
                    .TextMatrix(0, .ColIndex("Discount" & RsPrepareFgSalePrice.Fields("id").value)) = IIf(IsNull(RsPrepareFgSalePrice.Fields("DiscountNameE").value), 0, RsPrepareFgSalePrice.Fields("DiscountNameE").value)
                    lblPrice(i).Caption = IIf(IsNull(RsPrepareFgSalePrice.Fields("PriceNameE").value), 0, RsPrepareFgSalePrice.Fields("PriceNameE").value)
                    lblDiscount(i).Caption = IIf(IsNull(RsPrepareFgSalePrice.Fields("DiscountNameE").value), 0, RsPrepareFgSalePrice.Fields("DiscountNameE").value)
   
                End If
        
                TxtPrice(i).Visible = True
                TxtDiscount(i).Visible = True
                lblPrice(i).Visible = True
                lblDiscount(i).Visible = True
                RsPrepareFgSalePrice.MoveNext
            Next i

        End If

    End With

    '$$$$$$$$$$$$$$$$$$$$$$$$$$

End Function

Function ViewPrices()

    Dim rs As ADODB.Recordset
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = " SELECT     dbo.TblSalesPrices.ItemID, dbo.TblSalesPrices.Price1, dbo.TblSalesPrices.Price2, dbo.TblSalesPrices.Price3, dbo.TblSalesPrices.Price5, dbo.TblSalesPrices.Price4, "
    StrSQL = StrSQL + " dbo.TblSalesPrices.Price6, dbo.TblSalesPrices.Discount1, dbo.TblSalesPrices.Discount2, dbo.TblSalesPrices.Discount3, dbo.TblSalesPrices.Discount4,"
    StrSQL = StrSQL + " dbo.TblSalesPrices.Discount5, dbo.TblSalesPrices.Discount6, dbo.TblUnites.UnitName, dbo.TblSalesPrices.UnitID, dbo.TblSalesPrices.BranchId,"
    StrSQL = StrSQL + " dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL + "  FROM         dbo.TblSalesPrices LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.TblBranchesData ON dbo.TblSalesPrices.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.TblUnites ON dbo.TblSalesPrices.UnitID = dbo.TblUnites.UnitID"

    StrSQL = StrSQL + " Where  ItemID=" & val(Me.XPTxtID.text)
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FgSalePrice
            .Rows = Me.FgSalePrice.FixedRows + rs.RecordCount
            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                End If
                                    
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(rs("UnitID").value), "", rs("UnitID").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs("UnitName").value), "", rs("UnitName").value)
                .TextMatrix(i, .ColIndex("Price1")) = IIf(IsNull(rs("Price1").value), "", rs("Price1").value)
                .TextMatrix(i, .ColIndex("Price2")) = IIf(IsNull(rs("Price2").value), "", rs("Price2").value)
                .TextMatrix(i, .ColIndex("Price3")) = IIf(IsNull(rs("Price3").value), "", rs("Price3").value)
                .TextMatrix(i, .ColIndex("Price4")) = IIf(IsNull(rs("Price4").value), "", rs("Price4").value)
                .TextMatrix(i, .ColIndex("Price5")) = IIf(IsNull(rs("Price5").value), "", rs("Price5").value)
                .TextMatrix(i, .ColIndex("Price6")) = IIf(IsNull(rs("Price6").value), "", rs("Price6").value)
                .TextMatrix(i, .ColIndex("Discount1")) = IIf(IsNull(rs("Discount1").value), "", rs("Discount1").value)
                .TextMatrix(i, .ColIndex("Discount2")) = IIf(IsNull(rs("Discount2").value), "", rs("Discount2").value)
                .TextMatrix(i, .ColIndex("Discount3")) = IIf(IsNull(rs("Discount3").value), "", rs("Discount3").value)
                .TextMatrix(i, .ColIndex("Discount4")) = IIf(IsNull(rs("Discount4").value), "", rs("Discount4").value)
                .TextMatrix(i, .ColIndex("Discount5")) = IIf(IsNull(rs("Discount5").value), "", rs("Discount5").value)
                .TextMatrix(i, .ColIndex("Discount6")) = IIf(IsNull(rs("Discount6").value), "", rs("Discount6").value)
             
                rs.MoveNext
            Next i

            ' .AutoSize 0, .Cols - 1, False
        End With

    Else
        Me.FgSalePrice.Rows = Me.FgSalePrice.FixedRows
        '    Exit Function
    End If

    rs.Close
    Set rs = Nothing
 
    'VENDOR pRICES
    StrSQL = " SELECT     dbo.TblVendorContractDetails.TblVendorContractD, dbo.TblVendorContractDetails.UnitID, dbo.TblVendorContractDetails.ItemID, dbo.TblVendorContractDetails.Discount, "
    StrSQL = StrSQL & "    dbo.TblVendorContractDetails.Price, dbo.TblUnites.UnitName, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.TblVendorContract.VendorId,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee"
    StrSQL = StrSQL & " FROM         dbo.TblVendorContractDetails INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblItems ON dbo.TblVendorContractDetails.ItemID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "                         dbo.TblVendorContract ON dbo.TblVendorContractDetails.TblVendorContractD = dbo.TblVendorContract.TblVendorContractD LEFT OUTER JOIN"
    StrSQL = StrSQL & "    dbo.TblCustemers ON dbo.TblVendorContract.VendorId = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "           dbo.TblUnites ON dbo.TblVendorContractDetails.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "     WHERE     (dbo.TblVendorContractDetails.ItemID = " & val(Me.XPTxtID.text) & ")"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.FgVendorPrice
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                '.TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(RsDev("ItemId").value), _
                 "", RsDev("ItemId").value)
            
                '         .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), _
                          "", RsDev("ItemCode").value)
            
                '         .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), _
                          "", RsDev("ItemName").value)
                '         .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsDev("UnitId").value), _
                          "", RsDev("UnitId").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
            
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, val(RsDev("Price").value))
            
                .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), 0, val(RsDev("Discount").value))
            
                RsDev.MoveNext
            Next i
 
        End With

    Else
        Me.FgVendorPrice.Rows = Me.FgVendorPrice.FixedRows

    End If
 
    ReLineGrid
    Exit Function
ErrTrap:

End Function

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
   With Me.VSFlexGrid2

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("CatlogName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With
    IntCounter = 0
    With Me.FgVendorPrice

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("CusName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With
    
   '''//////////
    IntCounter = 0


    With Me.fgDiamonds

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("type")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("NumIndex")) = IntCounter
  
            End If

        Next i
   
    End With
    
    
     IntCounter = 0
  

    With Me.fgCameo

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("type")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("NumIndex")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Function ViewPricesold()

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    StrSQL = "SELECT  * from TblItemsPrices    "
 
    StrSQL = StrSQL + " Where  CustomerOrVendor=0 and   ItemID=" & val(Me.XPTxtID.text)
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FgPrices
            .Rows = Me.FgPrices.FixedRows + rs.RecordCount
            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1

                If rs("DefaultUnit").value = True Then
                    .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexUnchecked
                End If
            
                .TextMatrix(i, .ColIndex("PriceName")) = IIf(IsNull(rs("PriceName").value), "", rs("PriceName").value)
                .TextMatrix(i, .ColIndex("Pricevalue")) = IIf(IsNull(rs("Pricevalue").value), "", rs("Pricevalue").value)
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(rs("Des").value), "", rs("Des").value)
             
                rs.MoveNext
            Next i

            ' .AutoSize 0, .Cols - 1, False
        End With

    Else
        Me.FgPrices.Rows = Me.FgPrices.FixedRows
        '    Exit Function
    End If

    rs.Close
    Set rs = Nothing

    StrSQL = "SELECT  * from TblItemsPrices   "
 
    StrSQL = StrSQL + " Where CustomerOrVendor=1 and  ItemID=" & val(Me.XPTxtID.text)
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FgPrices1
            .Rows = Me.FgPrices1.FixedRows + rs.RecordCount
            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1

                If rs("DefaultUnit").value = True Then
                    .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("DefaultUnit")) = flexUnchecked
                End If
            
                .TextMatrix(i, .ColIndex("PriceName")) = IIf(IsNull(rs("PriceName").value), "", rs("PriceName").value)
                .TextMatrix(i, .ColIndex("Pricevalue")) = IIf(IsNull(rs("Pricevalue").value), "", rs("Pricevalue").value)
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(rs("Des").value), "", rs("Des").value)
             
                rs.MoveNext
            Next i

            ' .AutoSize 0, .Cols - 1, False
        End With

    Else
        Me.FgPrices1.Rows = Me.FgPrices1.FixedRows
        Exit Function
    End If

    rs.Close
    Set rs = Nothing

End Function

Private Sub DcboItems1_Click(Area As Integer)
    DcboItems1_Change
End Sub

Private Sub WriteDes(LngRow As Long)
    Dim StrTemp1 As String
    Dim StrTemp2 As String

    With Me.FgUnites

        If LngRow = 1 Then
            .TextMatrix(LngRow, .ColIndex("FactorDes")) = "«·ÊÕœ… «·√Ê·Ï"
        Else
            StrTemp1 = .TextMatrix(LngRow - 1, .ColIndex("UnitName"))
            StrTemp2 = StrTemp1 & "=" & .TextMatrix(LngRow, .ColIndex("UnitFactor")) & .TextMatrix(LngRow, .ColIndex("UnitName"))
            .TextMatrix(LngRow, .ColIndex("FactorDes")) = StrTemp2
        End If

    End With

End Sub

Private Sub dcItemunit_Change()
    Me.TxtItemPrice(0).text = ModItemCostPrice.GetCostItemPrice(val(Me.DcboItems.BoundText), , , , SystemOptions.SysMainStockCostMethod, , , Date, , val(dcItemunit.BoundText))
End Sub

Private Sub dcItemunit_Click(Area As Integer)
    dcItemunit_Change
End Sub

Private Sub DcTemplate_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FixedAssetsSearch1.RetrunType = 1
        FixedAssetsSearch1.show vbModal
  
    End If
    
End Sub

Private Sub fgCameo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim k As Integer
Dim StrComboList As String
            
    
    With fgCameo

        Select Case .ColKey(Col)
     Case "type"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode


              
                     

   

                   End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

       
    End With

    ReLineGrid
End Sub

Private Sub fgCameo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With fgCameo

      
        Select Case .ColKey(Col)
            
            Case "weight"
          
            
               fgCameo.ComboList = ""
            
            
               
        End Select

    End With
End Sub

Private Sub fgCameo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmItemCameoSearch
            FrmItemCameoSearch.show

'
End If
End Sub

Private Sub fgCameo_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With fgCameo

        Select Case .ColKey(Col)
 Case "type"
     StrSQL = " select code,name,nameE from TblGemstones "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "name", "code")
                Else
                    StrComboList = Fg.BuildComboList(rs, "nameE", "code")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
       
 Case "unite"
     StrSQL = " select UnitID,UnitName,UnitNamee from TblUnites "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "UnitName", "UnitID")
                Else
                    StrComboList = Fg.BuildComboList(rs, "UnitNamee", "UnitID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                'LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
         
 
        End Select

    End With
End Sub

Private Sub fgDiamonds_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim k As Integer
Dim StrComboList As String
            
    
    With fgDiamonds

        Select Case .ColyKe(Col)
              Case "type"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode


              
                     

   

                   End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

       
    End With

    ReLineGrid
End Sub

Private Sub fgDiamonds_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With fgDiamonds

      
        Select Case .ColKey(Col)
            
            Case "weight"
          
            
               fgDiamonds.ComboList = ""
            
            
               
        End Select

    End With
End Sub

Private Sub fgDiamonds_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmItemDiamoSearch
            FrmItemDiamoSearch.show

'
End If
End Sub

Private Sub fgDiamonds_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With fgDiamonds

        Select Case .ColKey(Col)
 Case "type"
     StrSQL = " select code,name,nameE from TblDiamonds "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "name", "code")
                Else
                    StrComboList = Fg.BuildComboList(rs, "nameE", "code")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
       
 Case "unite"
     StrSQL = " select UnitID,UnitName,UnitNamee from TblUnites "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "UnitName", "UnitID")
                Else
                    StrComboList = Fg.BuildComboList(rs, "UnitNamee", "UnitID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                'LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
         
 Case "ÛQuality"
     StrSQL = " select code,name,nameE from TblQuPices "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "name", "code")
                Else
                    StrComboList = Fg.BuildComboList(rs, "nameE", "code")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                  
          Case "Gestonf"
     StrSQL = " select code,name,nameE from TblGestonesFrm "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "name", "code")
                Else
                    StrComboList = Fg.BuildComboList(rs, "nameE", "code")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                  
                  Case "color"
     StrSQL = " select ColorID,ColorName  from TblItemsColors "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "ColorName", "ColorID")
                Else
                    StrComboList = Fg.BuildComboList(rs, "ColorName", "ColorID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
        End Select

    End With

End Sub

Private Sub Form_Activate()

    If SystemOptions.UserInterface = EnglishInterface And first_run = True Then
        '  SetInterface Me
        '  ChangeLang
        first_run = False
    End If

    'XPTxtID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
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
'sa If mdifrmmain.GoldMenu.Visible = True Then
'sa C1Tab1.TabVisible(7) = True
'sa Else
'sa C1Tab1.TabVisible(7) = False
'sa End If

    
    ScreenNameArabic = " »Ì«‰«  «·√’‰«›  "
    ScreenNameEnglish = " Items Data "
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"
    
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "Update TblItems Set RequestLimit=0 "
        StrSQL = StrSQL + " Where RequestLimit Is Null"
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    'On Error GoTo ErrTrap

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
   
    End If

    With Me.CboItemCase
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÃœÌœ"
            .AddItem "„” ⁄„·"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .AddItem "New"
            .AddItem "Used"
        End If

    End With

    With Me.CboItemType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "”·⁄…"
            .AddItem "Œœ„…"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .AddItem "Goods"
            .AddItem "Services"
        End If

    End With

    'Me.Width = 9930
    'Me.Height = 8085
    'Resize_Form Me
    'FillGroupCmbo
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemSGroups Me.XPCboGroup, False
    Dcombos.GetPrefix Me.DCPreFix, 3, val(branch_id)
    Set cDboSearch(0) = New clsDCboSearch
    Set cDboSearch(0).Client = Me.XPCboGroup
Dcombos.GetItemsNames Me.Dcbiteem
    Dcombos.GetItemsNames Me.DcboItems
    Set cDboSearch(1) = New clsDCboSearch
    Set cDboSearch(1).Client = Me.DcboItems

    Dcombos.GetItemsNames Me.DcboItemID1
    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True


Dcombos.GetTemplates Me.DcTemplate

    Set cDboSearch(2) = New clsDCboSearch
    Set cDboSearch(2).Client = Me.DcboItemID1
    TreeItems.ImageList = mdifrmmain.ImgLstTree
    '-------------------------------------------
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        ModFgLib.LinkFgColWithDataCombo Fg, Fg.ColIndex("ItemName"), Me.DcboItems
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgAttachs
        Set .WallPaper = GrdBack.Picture
        ModFgLib.LinkFgColWithDataCombo FgAttachs, FgAttachs.ColIndex("ItemName"), Me.DcboItemID1
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgSalePrice
        Set .WallPaper = GrdBack.Picture
     
    End With

    PrepareFgSalePrice

    With Me.FgVendorPrice
        Set .WallPaper = GrdBack.Picture
     
    End With

    '-------------------------------------------
    Set rs = New ADODB.Recordset
    rs.Open "[TblItems]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    LoadMenus
    LoadTreeGroups TreeItems
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    AddTip
    C1Tab1.CurrTab = 0

    ''''unites'''''''''''''''''''''
    Set GrdBack = New ClsBackGroundPic

    With Me.FgUnites
        .Rows = .FixedRows
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
        .ExtendLastCol = True
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
    End With

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DcboItems1
    Set cSearch(0) = New clsDCboSearch
    Set cSearch(0).Client = Me.DcboItems1
Dcombos.GetItemsUnits Me.Dcbuniit
    Dcombos.GetItemsUnits Me.DcboUnits
    Dcombos.GetItemsUnits Me.DcUnit
    Dcombos.GetBranches Dcbranch

    Set cSearch(1) = New clsDCboSearch
    Set cSearch(1).Client = Me.DcboUnits

    'Resize_Form Me
If FrmItems.CALLEDFPRM = False Then Exit Sub

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub
 
Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " ﬂÊœ «·’‰›  " & DCPreFix & txtid.text & Chr(13) & "  «”„ «·’‰› " & XPTxtName & Chr(13) & " ‰Ê⁄ «·’‰›   " & CboItemType.text & Chr(13) & " «·„Ã„Ê⁄Â  " & XPCboGroup.text & Chr(13) & " «Œ— ”⁄— ‘—«¡  " & XPTxtPurchase.text & Chr(13) & "”⁄— «·»Ì⁄ «·Õ«·Ì „” Â·ﬂ  " & XPTxtSall.text & Chr(13) & "  ”⁄— «·»Ì⁄ «·Õ«·Ì  ⁄„Ì·  " & TxtCusPrice.text
        LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Code  " & DCPreFix & txtid.text & Chr(13) & "    Name " & XPTxtNamee & Chr(13) & " Type   " & CboItemType.text & Chr(13) & " Group  " & XPCboGroup.text & Chr(13) & " Last Purchase Price  " & XPTxtPurchase.text & Chr(13) & "Sales Price Customer  " & XPTxtSall.text & Chr(13) & "  Sales Price Dealer  " & TxtCusPrice.text
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    On Error GoTo ErrTrap
' FrmItems.CALLEDFPRM = False
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish

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
    Set ItemReport = Nothing

    For i = LBound(cDboSearch) To UBound(cDboSearch)
        Set cDboSearch(i) = Nothing
    Next i

    For i = LBound(cSearch) To UBound(cSearch)
        Set cSearch(i) = Nothing
    Next i

    Erase cSearch

    Exit Sub
ErrTrap:
End Sub

Private Sub FgUnites_DblClick()
    '            If CheckItemsIntransactions(val(XPTxtID)) = True Then
    '                        If SystemOptions.UserInterface = ArabicInterface Then
    '                        MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Â–« «·’‰› ·ÊÃÊœ Õ—ﬂ«  ⁄·Ì…", vbCritical
    '                        Else
    '                        MsgBox "Cant Modify", vbCritical
    '
    '                        End If
    '                        Exit Sub
    '
    '            End If
    
    With Me.FgUnites

        If .Row <= 0 Then Exit Sub
        If .Col = -1 Then Exit Sub
    
        Me.TxtRowNumber.text = .Row

        If .Cell(flexcpChecked, .Row, .ColIndex("DefaultUnit")) = flexChecked Then
            Me.ChkDef.value = vbChecked
        Else
            Me.ChkDef.value = vbUnchecked
        End If

        Me.DcboUnits.BoundText = .TextMatrix(.Row, .ColIndex("UnitID"))
        DcboUnits.Enabled = False
        TxtUnitFactor.Enabled = False
        Me.TxtUnitFactor.text = .TextMatrix(.Row, .ColIndex("UnitFactor"))
        
        Me.TxtUnitSalesPrice.text = .TextMatrix(.Row, .ColIndex("UnitSalesPrice"))
        Me.TxtUnitPurPrice.text = .TextMatrix(.Row, .ColIndex("UnitPurPrice"))

    End With

End Sub

 

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)
    Me.lbl(14).ToolTipText = "›Ï Õ«·… «—”«· ’‰› «·Ï «·«—‘Ì› ·« ÌŸÂ— Â–« «·’‰› ›Ï «·›Ê« Ì— «·ÃœÌœ… »‘—ÿ «‰ ÌﬂÊ‰ —’Ìœ… ›Ï «·„Œ“‰ ’›— "
End Sub

Private Sub SearchCashCustomer_Click()
FrmModelsSearch.calltype = 1
Load FrmModelsSearch
FrmModelsSearch.show

End Sub

Private Sub TxtDealerPrice_LostFocus()
    On Error Resume Next

    If val(TxtDealerPrice.text) > val(XPTxtSall.text) Or val(TxtDealerPrice.text) > val(TxtCusPrice.text) Then
        MsgBox "⁄›Ê« ”⁄— «·œÌ·— «⁄·Ï  ", vbOKOnly, App.title
        TxtDealerPrice.SetFocus
        Exit Sub
    End If

    If val(TxtDealerPrice.text) < val(XPTxtPurchase.text) Then
        MsgBox "⁄›Ê« ”⁄— »Ì⁄ «·„” Â·ﬂ «ﬁ· „‰ ”⁄— «·‘—«¡ ", vbOKOnly, App.title
        TxtDealerPrice.SetFocus
        Exit Sub
    End If

End Sub

Private Sub TxtDiscount_KeyPress(Index As Integer, _
                                 KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDiscount(Index).text, 0)
End Sub

Private Sub TxtFreeQty_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtFreeQty.text, 0)
End Sub

Private Sub TxtPrice_KeyPress(Index As Integer, _
                              KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPrice(Index).text, 0)
End Sub

Private Sub TxtUnitPurPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtUnitPurPrice.text, 0)
End Sub

Private Sub TxtUnitSalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtUnitSalesPrice.text, 0)
End Sub

Private Sub ImgPic_DblClick()
    Load FrmViewPic
    Set FrmViewPic.MainView.Picture = ImgPic.Picture
    FrmViewPic.show vbModal
End Sub

Private Sub LblCostPrice_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    'Me.LblCostPrice.ToolTipText = WriteNo(CStr(Val(Me.LblCostPrice.Caption)), 0)

    Me.LblCostPrice.ToolTipText = "”⁄— «· ﬂ·›… «·Õ«·Ï ÂÊ „ Ê”ÿ ”⁄— «·’‰› »«· ﬂ·›… ÊÌŸÂ— »⁄œ «Ê· ⁄„·Ì… ‘—«¡ «Ê —’Ìœ «›  «ÕÏ "
End Sub

Private Sub TreeItems_MouseUp(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    On Error GoTo ErrTrap
    Dim tp            As POINTAPI
    Dim lX            As Single
    Dim lY            As Single
    Dim tr            As RECT
    Dim XNodeSeelcted As MSComctlLib.Node

    If Me.TreeItems.SelectedItem Is Nothing Then
        Exit Sub
    End If

    'TxtMenuState_Change
    'If right(TreeItems.SelectedItem.Key, 1) = "I" Then
    '    XPPopUp.Menus(1).MenuItems(1).Enabled = True
    '    XPPopUp.Menus(1).MenuItems(3).Enabled = True
    '    XPPopUp.Menus(1).MenuItems(4).Enabled = True
    '    XPPopUp.Menus(1).MenuItems(5).Enabled = False
    '    XPPopUp.Menus(1).MenuItems(6).Enabled = True
    '    XPPopUp.Menus(1).MenuItems(7).Enabled = False
    'Else
    '    XPPopUp.Menus(1).MenuItems(1).Enabled = False
    '    XPPopUp.Menus(1).MenuItems(3).Enabled = False
    '    XPPopUp.Menus(1).MenuItems(4).Enabled = False
    '    XPPopUp.Menus(1).MenuItems(5).Enabled = False
    '    XPPopUp.Menus(1).MenuItems(6).Enabled = False
    '    XPPopUp.Menus(1).MenuItems(7).Enabled = False
    'End If
    If Button = vbRightButton Then
        GetCursorPos tp
        lX = (tp.X) * Screen.TwipsPerPixelX
        lY = tp.Y * Screen.TwipsPerPixelY
 '       XPPopUp.PopupMenu "mnuDropMenu1", lX, lY
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TreeItems_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim NodeKey As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        If right(Node.key, 1) = "G" Then
        
            XPCboGroup.BoundText = val(Node.key)
            XPCboGroup_Click (0)
        End If

        Exit Sub
    End If

    If right(Node.key, 1) = "G" Then
        Exit Sub
    End If

    NodeKey = left(Node.key, Len(Node.key) - 1)

    If NodeKey <> "" Then
        Retrive (NodeKey)
        DcboItems1_Change
        Retriveshow (NodeKey)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtAttachedItemCode_KeyDown(KeyCode As Integer, _
                                        Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If TxtAttachedItemCode.text = "" Then
            Me.DcboItemID1.BoundText = ""
        Else
            Me.DcboItemID1.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode.text))
        End If
    End If

End Sub

Private Sub TxtCusPrice_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    'If Val(TxtCusPrice.text) > Val(XPTxtSall.text) Then
    'MsgBox "⁄›Ê« ”⁄— »Ì⁄ «·⁄„Ì· «⁄·Ï „‰ ”⁄— »Ì⁄ «·„” Â·ﬂ ", vbOKOnly, App.Title
    'TxtCusPrice.SetFocus
    'Exit Sub
    'End If
    '
    '
    'If Val(TxtCusPrice.text) < Val(XPTxtPurchase.text) Then
    'MsgBox "⁄›Ê« ”⁄— »Ì⁄ «·⁄„Ì· «ﬁ· „‰ ”⁄— «·‘—«¡ ", vbOKOnly, App.Title
    'TxtCusPrice.SetFocus
    'Exit Sub
    'End If
    'End If

    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCusPrice.text, 0)
End Sub

Private Sub TxtDealerPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDealerPrice.text, 0)
End Sub

Private Sub TxtGuarValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtGuarValue.text, 1)
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If TxtItemCode.text = "" Then
            Me.DcboItems.BoundText = ""
        Else
            Me.DcboItems.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If

End Sub

Private Sub TxtItemPrice_KeyPress(Index As Integer, _
                                  KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtItemPrice(Index).text, 0)
End Sub

Private Sub TxtItemQty_KeyPress(Index As Integer, _
                                KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtItemQty(Index).text, 0)
End Sub

'Private Sub TxtMenuState_Change()
'If right(TreeItems.SelectedItem.Key, 1) = "I" Then
'    XPPopUp.Menus(1).MenuItems(1).Enabled = True
'    XPPopUp.Menus(1).MenuItems(3).Enabled = True
'    XPPopUp.Menus(1).MenuItems(4).Enabled = True
'    XPPopUp.Menus(1).MenuItems(5).Enabled = True
'    XPPopUp.Menus(1).MenuItems(6).Enabled = True
'    XPPopUp.Menus(1).MenuItems(7).Enabled = True
'Else
'    XPPopUp.Menus(1).MenuItems(1).Enabled = False
'    XPPopUp.Menus(1).MenuItems(3).Enabled = False
'    XPPopUp.Menus(1).MenuItems(4).Enabled = False
'    XPPopUp.Menus(1).MenuItems(5).Enabled = False
'    XPPopUp.Menus(1).MenuItems(6).Enabled = False
'    XPPopUp.Menus(1).MenuItems(7).Enabled = False
'End If
'Select Case TxtMenuState.Text
'    Case "N"
'        If right(TreeItems.SelectedItem.Key, 1) = "I" Then
'            XPPopUp.Menus(1).MenuItems(8).Enabled = False
'        Else
'            XPPopUp.Menus(1).MenuItems(8).Enabled = False
'        End If
'        Me.XPBtnMove(0).Enabled = True
'        Me.XPBtnMove(1).Enabled = True
'        Me.XPBtnMove(2).Enabled = True
'        Me.XPBtnMove(3).Enabled = True
'    Case "C"
'        If right(TreeItems.SelectedItem.Key, 1) = "G" Then
'            XPPopUp.Menus(1).MenuItems(8).Enabled = True
'        Else
'            XPPopUp.Menus(1).MenuItems(8).Enabled = False
'        End If
'        Me.XPBtnMove(0).Enabled = False
'        Me.XPBtnMove(1).Enabled = False
'        Me.XPBtnMove(2).Enabled = False
'        Me.XPBtnMove(3).Enabled = False
'End Select
'Exit Sub
'End Sub
Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·√’‰«›"
            Else
                Me.Caption = "Items Data"
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
        
            Me.XPTxtCode.locked = True
            Me.XPTxtName.locked = True
            Me.XPCboGroup.locked = True
            TxtRequired.locked = True
            XPChkSerial.Enabled = False
            Me.ChkAr.Enabled = False
            XPTxtPurchase.locked = True
            XPTxtSall.locked = True
            Me.TxtCusPrice.locked = True
            Me.TxtDealerPrice.locked = True
            Me.ChkGuar.Enabled = False
            Me.TxtGuarValue.locked = True
            Me.Ele(0).Enabled = False
        
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            Else
                '            TxtMenuState.Text = "N"
            End If

            TreeItems.Enabled = True
        
            Me.CmdPic(0).Enabled = False
            Me.CmdPic(1).Enabled = False
        
            Me.ChkAssplied.Enabled = False
            ChkItemMakingNew.Enabled = False
            Me.lbl(17).Enabled = False
            Me.lbl(18).Enabled = False
            Me.lbl(19).Enabled = False
            Me.lbl(20).Enabled = False
            Me.TxtItemCode.Enabled = False
            Me.DcboItems.Enabled = False
            Me.TxtItemPrice(0).Enabled = False
            Me.TxtItemQty(0).Enabled = False
            Me.Cmd(8).Enabled = False
            Me.Cmd(9).Enabled = False
            '------------------------------
            Me.ChkRelated.Enabled = False
            ' Me.Lbl(23).Enabled = False
            Me.lbl(24).Enabled = False
            Me.lbl(25).Enabled = False
            Me.lbl(26).Enabled = False
            Me.TxtAttachedItemCode.Enabled = False
            Me.DcboItemID1.Enabled = False
            Me.TxtItemPrice(1).Enabled = False
            Me.TxtItemQty(1).Enabled = False
            Me.Cmd(10).Enabled = False
            Me.Cmd(11).Enabled = False
        
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·√’‰«›( ÃœÌœ )"
            Else
                Me.Caption = "Items Data(New Record)."
            End If

            LblCostPrice.Caption = 0
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
            XPChkSerial.value = Unchecked

            Me.XPTxtCode.locked = False
            Me.XPTxtName.locked = False
            TxtRequired.locked = False
            Me.XPCboGroup.locked = False
            XPChkSerial.Enabled = True
            Me.ChkAr.Enabled = True
            '  TreeItems.Enabled = False
            XPTxtPurchase.locked = False
            XPTxtSall.locked = False
            Me.TxtCusPrice.locked = False
            Me.TxtDealerPrice.locked = False
            Me.CmdPic(0).Enabled = True
            Me.CmdPic(1).Enabled = True
            Me.ChkGuar.Enabled = True
            Me.TxtGuarValue.locked = False
            Me.Ele(0).Enabled = True
        
            ChkAssplied.Enabled = True
            ChkItemMakingNew.Enabled = True
            ChkAssplied_Click
            ChkItemMakingNew_Click
            ChkRelated.Enabled = True
            ChkRelated_Click
            DcboUnits.Enabled = True
            TxtUnitFactor.Enabled = True

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·√’‰«›(  ⁄œÌ· )"
            Else
                Me.Caption = "Items Data(Edit Record)."
            End If

            DcboUnits.Enabled = True
            TxtUnitFactor.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            TxtRequired.locked = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.XPTxtCode.locked = False
            Me.XPTxtName.locked = False
            Me.XPCboGroup.locked = False
            XPChkSerial.Enabled = True
            Me.ChkAr.Enabled = True
            TreeItems.Enabled = False
            XPTxtPurchase.locked = False
            XPTxtSall.locked = False
            Me.TxtCusPrice.locked = False
            Me.TxtDealerPrice.locked = False
            Me.CmdPic(0).Enabled = True
            Me.CmdPic(1).Enabled = True
            Me.ChkGuar.Enabled = True
            Me.TxtGuarValue.locked = False
            Me.Ele(0).Enabled = True
            Me.ChkAssplied.Enabled = True
            ChkItemMakingNew.Enabled = True
            ChkAssplied_Click
            ChkItemMakingNew_Click
            ChkRelated.Enabled = True
            ChkRelated_Click
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtRequired_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtRequired.text, 1)
End Sub



Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid2

        Select Case .ColKey(Col)
 Case "CatloPath"
' CommonDialog1.InitDir = App.path & "\ REPORTS"""
'CommonDialog1.ShowOpen

 .TextMatrix(Row, .ColIndex("CatloPath1")) = CommonDialog1.FileName
 
End Select

  If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If
   End With
   ReLineGrid
End Sub

Private Sub VSFlexGrid2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
  With Me.VSFlexGrid2

        Select Case .ColKey(Col)

                 Case "view"
                 ' LngRow = Row
                 FilePath = .TextMatrix(Row, .ColIndex("CatloPath1"))
ShellExecute 0&, vbNullString, FilePath, vbNullString, vbNullString, vbNormalFocus
 
             

                    
                End Select
                End With
End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With VSFlexGrid2

        Select Case .ColKey(Col)
 Case "CatloPath"
 CommonDialog1.Filter = "PDF File|*.PDF"
 CommonDialog1.InitDir = App.path & "\ REPORTS"""
CommonDialog1.ShowOpen

 .TextMatrix(Row, .ColIndex("CatloPath1")) = CommonDialog1.FileName
 Case "view"
 .ColComboList(.ColIndex("view")) = "..."
End Select
     
   End With
End Sub

Private Sub VSFlexGrid3_Click()
  With VSFlexGrid3

        Select Case .Col
 
         

            Case 3
FrmPO5.show
                FrmPO5.Retrive val(.TextMatrix(.Row, 2))

          
        End Select

    End With
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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
    Exit Sub
ErrTrap:
End Sub
Sub Retriveshow(Optional IDitem As Integer = 0)
Dim sql As String
Dim Rsditails As ADODB.Recordset
Set Rsditails = New ADODB.Recordset
  VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid3.Rows = 2
sql = " SELECT    dbo.Transactions.Transaction_ID,  dbo.TblItems.HaveSerial, dbo.Transactions.Transaction_Date, dbo.Transactions.PODays, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
sql = sql & "                      dbo.Transaction_Details.Item_ID, dbo.TblItems.Fullcode, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitID,"
sql = sql & "                      dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.ShowQty,"
sql = sql & "                      dbo.Transaction_Details.showPrice, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile,"
sql = sql & "                      dbo.TblCustemers.Fullcode AS CusFullcode, dbo.Transactions.CusID"
sql = sql & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems INNER JOIN"
sql = sql & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
sql = sql & "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & " Where (dbo.Transactions.Transaction_Type = 46) And (dbo.TblItems.ItemID =" & val(IDitem) & ")"
sql = sql & " ORDER BY dbo.Transactions.Transaction_Date"
       Rsditails.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (Rsditails.BOF Or Rsditails.EOF) Then

                    With Me.VSFlexGrid3
                        .Rows = .FixedRows + Rsditails.RecordCount
                       
                  For i = 1 To .Rows - 1
                   
                  .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(Rsditails("Transaction_ID").value), "", Rsditails("Transaction_ID").value)
                  
                     .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rsditails("NoteSerial1").value), "", Rsditails("NoteSerial1").value)
                     .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(Rsditails("Transaction_Date").value), "", Rsditails("Transaction_Date").value)
                    .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rsditails("Price").value), "", Rsditails("Price").value)
                    .TextMatrix(i, .ColIndex("PODays")) = IIf(IsNull(Rsditails("PODays").value), "", Rsditails("PODays").value)
                    .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rsditails("CusID").value), "", Rsditails("CusID").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rsditails("CusName").value), "", Rsditails("CusName").value)
                    Else
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rsditails("CusNamee").value), "", Rsditails("CusNamee").value)
                    End If
                    Rsditails.MoveNext
                  Next i
                  

End With
End If
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsParts As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "ItemID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.EOF Or rs.BOF Then
            Exit Sub
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("ItemID").value), "", val(rs("ItemID").value))
'    On Error Resume Next

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If val(XPTxtID.text) <> 0 Then

        'Text1.text = get_item_qty(Val(XPTxtID.text))
        'Text4.text = get_item_Order_qty(Val(XPTxtID.text))
        'Text5.text = get_item_Reserved_qty(Val(XPTxtID.text))
    Else
        'Text1.text = 0
        'Text4.text = 0
        'Text5.text = 0

    End If

                            
                                Dim FirstPeriodDateInthisYear  As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
Dim LastPurchaseDate As String
Dim LastPurchasePrice As Double
Dim LastPurchaseqty As Double

    Fromdate = FirstPeriodDateInthisYear
  GetlastPurchasedata 46, val(XPTxtID.text), FirstPeriodDateInthisYear, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
         lstorderdate.text = LastPurchaseDate
          lastorderPrice.text = LastPurchasePrice
          
                           
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    XPTxtCode.text = IIf(IsNull(rs("ItemCode").value), "", Trim(rs("ItemCode").value))
    TxtPartNo.text = IIf(IsNull(rs("PartNo").value), "", Trim(rs("PartNo").value))

TxtFreeQty.text = IIf(IsNull(rs("FreeQty").value), 0, (rs("FreeQty").value))

TxtbarCodeNO.text = IIf(IsNull(rs("barCodeNO").value), "", (rs("barCodeNO").value))
TxtCatlogNO.text = IIf(IsNull(rs("CatlogNO").value), "", (rs("CatlogNO").value))
TxtFactoryNO.text = IIf(IsNull(rs("FactoryNO").value), "", (rs("FactoryNO").value))
 
Me.TxtOverHead.text = IIf(IsNull(rs("OverHead").value), 0, rs("OverHead").value)
Me.TxtWight.text = IIf(IsNull(rs("Wight").value), 0, rs("Wight").value)

Me.txtContent.text = IIf(IsNull(rs("Content").value), "", rs("Content").value)
 Me.txtDippre.text = IIf(IsNull(rs("Dippre").value), "", rs("Dippre").value)
 
  Me.TxtSource.text = IIf(IsNull(rs("Source").value), "", rs("Source").value)
   Me.txtTypenew.text = IIf(IsNull(rs("Typenew").value), "", rs("Typenew").value)
   

    XPTxtName.text = IIf(IsNull(rs("ItemName").value), "", Trim(rs("ItemName").value))
    XPTxtNamee.text = IIf(IsNull(rs("ItemNamee").value), "", Trim(rs("ItemNamee").value))

    XPTxtPurchase.text = IIf(IsNull(rs("PurchasePrice").value), "", Trim(rs("PurchasePrice").value))
    XPTxtSall.text = IIf(IsNull(rs("SallingPrice").value), "", Trim(rs("SallingPrice").value))
    TxtRequired.text = IIf(IsNull(rs("RequestLimit").value), "", Trim(rs("RequestLimit").value))

Txtminvalueqty.text = IIf(IsNull(rs("minvalueqty").value), 0, (rs("minvalueqty").value))

TxtMaxValueqty.text = IIf(IsNull(rs("MaxValueqty").value), 0, (rs("MaxValueqty").value))


    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    Me.txtid.text = IIf(IsNull(rs("code").value), "", rs("code").value)

    If Not IsNull(rs("ItemPhoto").value) Then
        If LenB(rs("ItemPhoto")) Then
            LoadPictureFromDB ImgPic, rs, "ItemPhoto"
        Else
            Set ImgPic.Picture = Nothing
        End If

    Else
        Set ImgPic.Picture = Nothing
    End If

    If Not IsNull(rs("GroupID")) Then
        XPCboGroup.BoundText = rs("GroupID").value
    Else
        XPCboGroup.BoundText = ""
    End If

    Me.DBCboClientName.BoundText = IIf(IsNull(rs("DefaultSupplier").value), "", rs("DefaultSupplier").value)

Me.DcTemplate.BoundText = IIf(IsNull(rs("TemplateID").value), "", rs("TemplateID").value)


    
    If IsNull(rs("ItemCase").value) Then
        Me.CboItemCase.ListIndex = -1
    ElseIf rs("ItemCase").value = 1 Then
        Me.CboItemCase.ListIndex = 0
    ElseIf rs("ItemCase").value = 2 Then
        Me.CboItemCase.ListIndex = 1
    End If
TxtItemMaxDiscount.text = IIf(IsNull(rs("ItemMaxDiscount").value), "0", (rs("ItemMaxDiscount").value))

    TxtCusPrice.text = IIf(IsNull(rs("CustomerPrice").value), "0", Trim(rs("CustomerPrice").value))
    TxtDealerPrice.text = IIf(IsNull(rs("DealerPrice").value), "0", Trim(rs("DealerPrice").value))

    XPChkSerial.value = IIf(rs("HaveSerial").value = True, vbChecked, vbUnchecked)

    Me.ChkGuar.value = IIf(rs("HaveGuarantee").value = True, vbChecked, vbUnchecked)
    Me.TxtGuarValue.text = IIf(IsNull(rs("GuaranteeValue").value) = True, "0", rs("GuaranteeValue").value)

    If Not IsNull(rs("GuaranteeType").value) Then
        If rs("GuaranteeType").value = 0 Then
            OptGaurType(0).value = True
            OptGaurType(1).value = False
        Else
            OptGaurType(1).value = True
            OptGaurType(0).value = False
        End If

    Else
        OptGaurType(0).value = True
    End If

    If IsNull(rs("IsArchive").value) Or rs("IsArchive").value = 0 Or rs("IsArchive").value = False Then
        Me.ChkAr.value = vbUnchecked
    Else
        Me.ChkAr.value = vbChecked
    End If

    If Not (IsNull(rs("ItemType").value)) Then
        If rs("ItemType").value = 0 Then
            Me.CboItemType.ListIndex = 0
        Else
            Me.CboItemType.ListIndex = 1
        End If

    Else
        Me.CboItemType.ListIndex = -1
    End If

    '---------------------------------------
    Me.TxtItemComment.text = IIf(IsNull(rs("ItemComment").value), "", Trim(rs("ItemComment").value))
    Me.TxtBinLocation.text = IIf(IsNull(rs("BinLocation").value), "", Trim(rs("BinLocation").value))
    

'BinLocation

    '------------------------
    If rs("AssbliedItem").value = True Then
        Me.ChkAssplied.value = vbChecked
        ChkAssplied.Visible = True
     
    ElseIf rs("AssbliedItem").value = False Then
        Me.ChkAssplied.value = vbUnchecked
    End If

    If rs("ItemMakingNew").value = True Then
        Me.ChkItemMakingNew.value = vbChecked
        ChkItemMakingNew.Visible = True
    ElseIf rs("ItemMakingNew").value = False Then
        Me.ChkItemMakingNew.value = vbUnchecked
    End If

    ' If ChkAssplied.Visible = True Then
    '        If ChkAssplied.Value = vbChecked Then
    '            Rs("AssbliedItem").Value = True
    '        ElseIf ChkAssplied.Value = vbUnchecked Then
    '            Rs("AssbliedItem").Value = False
    '        End If
    '    End If
    With Me.Fg
        .Rows = .FixedRows
        .AutoSize 0, .Cols - 1, False
    End With
With Me.VSFlexGrid1
        .Rows = .FixedRows
        .AutoSize 0, .Cols - 1, False
    End With
    If ChkAssplied.value = vbChecked Then
        If Not (IsNull(rs("AssbliedItem").value)) Then
            If rs("AssbliedItem").value = True Then
                Me.ChkAssplied.value = vbChecked
             
                Set RsParts = New ADODB.Recordset
                '   StrSQL = "Select * From TblItemsParts Where ItemID=" & rs("ItemID").value
                '   StrSQL = StrSQL + " Order By TableID"
                StrSQL = "SELECT     TOP 100 PERCENT dbo.TblItemsParts.Unitid, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
                StrSQL = StrSQL + "  dbo.TblItemsParts.ItemID , dbo.TblItemsParts.TableID, dbo.TblUnites.unitname, dbo.TblUnites.UnitNamee"
                StrSQL = StrSQL + " FROM         dbo.TblItemsParts INNER JOIN"
                StrSQL = StrSQL + " dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID"
                StrSQL = StrSQL + " Where (dbo.TblItemsParts.ItemID = " & rs("ItemID").value & ")"
                StrSQL = StrSQL + " ORDER BY dbo.TblItemsParts.TableID"
             
                RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsParts.BOF Or RsParts.EOF) Then

                    With Me.Fg
                        .Rows = .FixedRows + RsParts.RecordCount

                        For i = .FixedRows To .Rows - 1
                            .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsParts("PartItemID").value), "", RsParts("PartItemID").value)
                            .TextMatrix(i, .ColIndex("ItemCode")) = GetItemCode(val(.TextMatrix(i, .ColIndex("ItemID"))))
                            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsParts("PartItemID").value), "", RsParts("PartItemID").value)
                            .TextMatrix(i, .ColIndex("ItemQty")) = IIf(IsNull(RsParts("PartItemQty").value), "", RsParts("PartItemQty").value)
                            .TextMatrix(i, .ColIndex("ItemPrice")) = IIf(IsNull(RsParts("PartItemPrice").value), "", RsParts("PartItemPrice").value)
                            .TextMatrix(i, .ColIndex("Unitid")) = IIf(IsNull(RsParts("Unitid").value), "", RsParts("Unitid").value)

                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsParts("unitname").value), "", RsParts("unitname").value)
                            Else
                                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsParts("unitnamee").value), "", RsParts("unitnamee").value)
                            End If
                        
                            RsParts.MoveNext
                        Next i

                        .AutoSize 0, .Cols - 1, False
                    
                    End With

                End If

            ElseIf rs("AssbliedItem").value = False Then
                Me.ChkAssplied.value = vbUnchecked
            End If

        Else
            Me.ChkAssplied.value = vbUnchecked
        End If

        ChkAssplied_Click
    End If

    '’‰› „‰ Ã ÃœÌœ
    If ChkItemMakingNew.value = vbChecked Then
        If Not (IsNull(rs("ItemMakingNew").value)) Then
            If rs("ItemMakingNew").value = True Then
                Me.ChkItemMakingNew.value = vbChecked
             
                Set RsParts = New ADODB.Recordset
                '    StrSQL = "Select * From TblItemsParts Where ItemID=" & rs("ItemID").value
                '    StrSQL = StrSQL + " Order By TableID"
             
                StrSQL = "SELECT     TOP 100 PERCENT dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItemsParts.ItemID, dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.PartItemQty, "
                StrSQL = StrSQL & " dbo.TblItemsParts.PartItemPrice , dbo.TblItemsParts.unitid"
                StrSQL = StrSQL & "  FROM         dbo.TblItemsParts INNER JOIN"
                StrSQL = StrSQL & " dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID"
                StrSQL = StrSQL & "  Where (dbo.TblItemsParts.ItemID = " & rs("ItemID").value & ")"
                StrSQL = StrSQL & "  ORDER BY dbo.TblItemsParts.TableID"

                RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsParts.BOF Or RsParts.EOF) Then

                    With Me.Fg
                        .Rows = .FixedRows + RsParts.RecordCount + 1

                        For i = .FixedRows To .Rows - 1
                            .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsParts("PartItemID").value), "", RsParts("PartItemID").value)
                            .TextMatrix(i, .ColIndex("ItemCode")) = GetItemCode(val(.TextMatrix(i, .ColIndex("ItemID"))))
                            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsParts("PartItemID").value), "", RsParts("PartItemID").value)
                            .TextMatrix(i, .ColIndex("ItemQty")) = IIf(IsNull(RsParts("PartItemQty").value), "", RsParts("PartItemQty").value)
                            .TextMatrix(i, .ColIndex("ItemPrice")) = IIf(IsNull(RsParts("PartItemPrice").value), "", RsParts("PartItemPrice").value)
                            .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsParts("UnitName").value), "", RsParts("UnitName").value)
                            .TextMatrix(i, .ColIndex("unitid")) = IIf(IsNull(RsParts("unitid").value), "", RsParts("unitid").value)
                                         
                            RsParts.MoveNext
                        Next i

                        .AutoSize 0, .Cols - 1, False
                    
                    End With

                End If

            ElseIf rs("itemMakingNew").value = False Then
                Me.ChkItemMakingNew.value = vbUnchecked
            End If

        Else
            Me.ChkItemMakingNew.value = vbUnchecked
        End If

        ChkItemMakingNew_Click
    End If

    '’‰› „’‰⁄
    '------------------------
    Me.chkItemMaking.value = vbUnchecked

    'With Me.Fg
    '    .Rows = .FixedRows
    '    .AutoSize 0, .Cols - 1, False
    'End With
    If chkItemMaking.Visible = True Then
        If Not (IsNull(rs("ItemMaking").value)) Then
            If rs("ItemMaking").value = True Then
                Me.chkItemMaking.value = vbChecked
             
                '             Set RsParts = New ADODB.Recordset
                '             StrSQL = "Select * From TblItemsParts Where ItemID=" & Rs("ItemID").Value
                '             StrSQL = StrSQL + " Order By TableID"
                '             RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                '             If Not (RsParts.BOF Or RsParts.EOF) Then
                '                With Me.Fg
                '                    .Rows = .FixedRows + RsParts.RecordCount
                '                    For I = .FixedRows To .Rows - 1
                '                        .TextMatrix(I, .ColIndex("ItemID")) = IIf(IsNull(RsParts("PartItemID").Value), "", RsParts("PartItemID").Value)
                '                        .TextMatrix(I, .ColIndex("ItemCode")) = GetItemCode(Val(.TextMatrix(I, .ColIndex("ItemID"))))
                '                        .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(RsParts("PartItemID").Value), "", RsParts("PartItemID").Value)
                '                        .TextMatrix(I, .ColIndex("ItemQty")) = IIf(IsNull(RsParts("PartItemQty").Value), "", RsParts("PartItemQty").Value)
                '                        .TextMatrix(I, .ColIndex("ItemPrice")) = IIf(IsNull(RsParts("PartItemPrice").Value), "", RsParts("PartItemPrice").Value)
                '                        RsParts.MoveNext
                '                    Next I
                '                    .AutoSize 0, .Cols - 1, False
                '
                '                End With
                '             End If
            ElseIf rs("ItemMaking").value = False Then
                Me.chkItemMaking.value = vbUnchecked
            End If

        Else
            Me.chkItemMaking.value = vbUnchecked
        End If
    
    End If

    '------------------------------------------------
    '------------------------------------------------
    Me.ChkRelated.value = vbUnchecked

    With Me.FgAttachs
        .Rows = .FixedRows
        .AutoSize 0, .Cols - 1, False
    End With
''''''''''''''''
 Set RsParts = New ADODB.Recordset
 
      StrSQL = " SELECT     dbo.TblItems.ItemID, dbo.TblItemDiamonds.type, dbo.TblItemDiamonds.unite, dbo.TblItemDiamonds.weight, dbo.TblItemDiamonds.indexe"
StrSQL = StrSQL & " FROM         dbo.TblItems INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TblItemDiamonds ON dbo.TblItems.ItemID = dbo.TblItemDiamonds.ItemID"
StrSQL = StrSQL & " Where (dbo.TblItems.ItemID = " & rs("ItemID").value & ") And (dbo.TblItemDiamonds.indexe = 1)"

            RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                With Me.fgCameo
                
            If Not (RsParts.BOF Or RsParts.EOF) Then


                    .Rows = .FixedRows + RsParts.RecordCount

                    For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("NumIndex")) = i
                        .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsParts("ItemID").value), "", RsParts("ItemID").value)
                        .TextMatrix(i, .ColIndex("type")) = IIf(IsNull(RsParts("type").value), "", RsParts("type").value)
                        .TextMatrix(i, .ColIndex("unite")) = IIf(IsNull(RsParts("unite").value), "", RsParts("unite").value)
                        .TextMatrix(i, .ColIndex("weight")) = IIf(IsNull(RsParts("weight").value), "", RsParts("weight").value)
                        '.TextMatrix(i, .ColIndex("ItemPrice")) = IIf(IsNull(RsParts("AttachItemPrice").value), "", RsParts("AttachItemPrice").value)
                        RsParts.MoveNext
                    Next i

            ' .AutoSize 0, .Cols - 1, False
                
Else
     .Clear flexClearScrollable, flexClearEverything
     .Rows = .FixedRows

             
            End If
            End With
            
            
            
          Set RsParts = New ADODB.Recordset
        StrSQL = " select * from TblItemCatalog where ItemID=" & rs("ItemID").value & ""


           RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                With Me.VSFlexGrid2
                
            If Not (RsParts.BOF Or RsParts.EOF) Then


                    .Rows = .FixedRows + RsParts.RecordCount

                    For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("Ser")) = i
                        .TextMatrix(i, .ColIndex("CatlogName")) = IIf(IsNull(RsParts("CatlogName").value), "", RsParts("CatlogName").value)
                        .TextMatrix(i, .ColIndex("CatloPath1")) = IIf(IsNull(RsParts("CatloPath").value), "", RsParts("CatloPath").value)
    
                        RsParts.MoveNext
                    Next i

            '        .AutoSize 0, .Cols - 1, False
                
Else

 
    .Clear flexClearScrollable, flexClearEverything
     .Rows = .FixedRows

            End If


End With
'''/
         
                Set RsParts = New ADODB.Recordset
                '   StrSQL = "Select * From TblItemsParts Where ItemID=" & rs("ItemID").value
                '   StrSQL = StrSQL + " Order By TableID"
                StrSQL = " SELECT     dbo.TblAotherItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee, dbo.TblAotherItems.IDItem, "
               StrSQL = StrSQL + "       dbo.TblAotherItems.Remark, dbo.TblAotherItems.Valu, dbo.TblAotherItems.Quntity, dbo.TblAotherItems.UnitID, dbo.TblUnites.UnitName,"
               StrSQL = StrSQL + "       dbo.TblUnites.UnitNamee"
               StrSQL = StrSQL + "  FROM         dbo.TblAotherItems LEFT OUTER JOIN"
                StrSQL = StrSQL + "      dbo.TblUnites ON dbo.TblAotherItems.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
               StrSQL = StrSQL + "       dbo.TblItems ON dbo.TblAotherItems.ItemID = dbo.TblItems.ItemID"
                StrSQL = StrSQL + "      Where (dbo.TblAotherItems.IDItem = " & rs("ItemID").value & ")"
             
                RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
Me.lbl(38).Caption = RsParts.RecordCount
                If Not (RsParts.BOF Or RsParts.EOF) Then

                    With Me.VSFlexGrid1
                        .Rows = .FixedRows + RsParts.RecordCount

                        For i = .FixedRows To .Rows - 1
                            .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsParts("ItemID").value), "", RsParts("ItemID").value)
                            .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsParts("Fullcode").value), "", RsParts("Fullcode").value)
                           
                            .TextMatrix(i, .ColIndex("ItemQty")) = IIf(IsNull(RsParts("Quntity").value), "", RsParts("Quntity").value)
                            .TextMatrix(i, .ColIndex("ItemPrice")) = IIf(IsNull(RsParts("Valu").value), "", RsParts("Valu").value)
                            .TextMatrix(i, .ColIndex("Unitid")) = IIf(IsNull(RsParts("UnitID").value), "", RsParts("UnitID").value)

                            If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsParts("ItemName").value), "", RsParts("ItemName").value)
                                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsParts("unitname").value), "", RsParts("unitname").value)
                            Else
                            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsParts("ItemNamee").value), "", RsParts("ItemNamee").value)
                                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsParts("unitnamee").value), "", RsParts("unitnamee").value)
                            End If
                        .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsParts("Remark").value), "", RsParts("Remark").value)
                            RsParts.MoveNext
                        Next i

                        .AutoSize 0, .Cols - 1, False
                    
                    End With

                End If
''//


    If Not (IsNull(rs("RelatedItem").value)) Then
        If rs("RelatedItem").value = True Then
            Me.ChkRelated.value = vbChecked
            Set RsParts = New ADODB.Recordset
            StrSQL = "Select * From TblItemsAttach Where ItemID=" & rs("ItemID").value
            StrSQL = StrSQL + " Order By TableID"
            RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsParts.BOF Or RsParts.EOF) Then

                With Me.FgAttachs
                    .Rows = .FixedRows + RsParts.RecordCount

                    For i = .FixedRows To .Rows - 1
                        .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsParts("AttachItemID").value), "", RsParts("AttachItemID").value)
                        .TextMatrix(i, .ColIndex("ItemCode")) = GetItemCode(val(.TextMatrix(i, .ColIndex("ItemID"))))
                        .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsParts("AttachItemID").value), "", RsParts("AttachItemID").value)
                        .TextMatrix(i, .ColIndex("ItemQty")) = IIf(IsNull(RsParts("AttachItemQty").value), "", RsParts("AttachItemQty").value)
                        .TextMatrix(i, .ColIndex("ItemPrice")) = IIf(IsNull(RsParts("AttachItemPrice").value), "", RsParts("AttachItemPrice").value)
                        RsParts.MoveNext
                    Next i

                    .AutoSize 0, .Cols - 1, False
                End With

            End If

        ElseIf rs("RelatedItem").value = False Then
            Me.ChkRelated.value = vbUnchecked
        End If

    Else
        Me.ChkRelated.value = vbUnchecked
    End If

    ChkRelated_Click
    '-----------------------------------------
    Me.lbl(21).Caption = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("ItemID"))
    Me.lbl(28).Caption = ModFgLib.GetItemsInFg(FgAttachs, FgAttachs.ColIndex("ItemID"))
    '-----------------------------------------
    'Get The  Item Cost Price
    'Me.LblCostPrice.Caption = ModItemCostPrice.GetCostItemPrice(Val(Me.XPTxtID.text),  2)
    Dim unitid As Long
    GetDefaultItemUnit val(Me.XPTxtID.text), unitid
    Me.LblCostPrice.Caption = ModItemCostPrice.GetCostItemPrice(val(Me.XPTxtID.text), , , , SystemOptions.SysMainStockCostMethod, , , Date, , unitid)
Retriveshow val(XPTxtID.text)
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    LblItemCode.Caption = DCPreFix & txtid

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblItemName = XPTxtName.text
    Else
        Me.LblItemName = XPTxtNamee.text
    End If
ChkItemMakingNew_Click
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            SetMeForNew
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ItemID=" & val(XPTxtID.text), , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub AddNewRow2()

    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long

    If TxtPriceName1 = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»    «”„ «·”⁄— ...!!!"
        Else
            Msg = "must Specify Price Namet...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
  
    If TxtSalesPrice1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ    ”⁄— «·»Ì⁄  ...!!!"
        Else
            Msg = "must Enter Sales Price ...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 
    'If FgPrices1.Rows = 1 Then
    '    LngRow = Val(Me.TxtRowNumber.text)
    'Else
    Me.FgPrices1.Rows = Me.FgPrices1.Rows + 1
    LngRow = Me.FgPrices1.Rows - 1
    'End If
  
    On Error Resume Next

    With Me.FgPrices1
    
        If Me.ChkDefSalePrice1.value = vbChecked Then
            .Cell(flexcpChecked, .FixedRows, .ColIndex("DefaultUnit"), .Rows - 1, .ColIndex("DefaultUnit")) = flexUnchecked
            .Cell(flexcpChecked, LngRow, .ColIndex("DefaultUnit")) = flexChecked
        End If
    
        .TextMatrix(LngRow, .ColIndex("PriceName")) = Me.TxtPriceName1.text
        .TextMatrix(LngRow, .ColIndex("Pricevalue")) = val(Me.TxtSalesPrice1.text)
        .TextMatrix(LngRow, .ColIndex("des")) = Me.TxtPriceDes1.text
        .TextMatrix(LngRow, .ColIndex("CustomerOrVendor")) = 1
   
        '    .AutoSize 0, .Cols - 1, False
    End With

    Me.ChkDefSalePrice1.value = vbUnchecked
 
    Me.TxtPriceName1.text = ""
    Me.TxtSalesPrice1.text = ""
    Me.TxtPriceDes1.text = ""
 
End Sub

Private Sub AddNewRow1(Optional auto As Boolean = False, _
                       Optional saleprice1 As Double, _
                       Optional saleprice2 As Double, _
                       Optional saleprice3 As Double)

    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long

    If auto = False Then
        If saleprice1 = 0 Then
            Exit Sub
        End If
    End If

    If auto = False Then
        If DcUnit.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— «·ÊÕœÂ  ...!!!"
            Else
                Msg = "must Specify Unit Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
 
    End If
 
    'If FgSalePrice.Rows = 1 Then
    '    LngRow = Val(Me.TxtRowNumber.text)
    'Else
    Me.FgSalePrice.Rows = Me.FgSalePrice.Rows + 1
    LngRow = Me.FgSalePrice.Rows - 1
    'End If
 
    If auto = True Then
        optBranch(0).value = True
    End If
 
    If optBranch(0).value = True Then '  ﬂ· «·›—Ê⁄
        Dim rs As ADODB.Recordset
        Dim sql As String
        Dim unitid As Long
        Dim Unitname As String
        sql = "Select  *   from TblBranchesData ORDER BY branch_id"
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If rs.RecordCount = 0 Then Exit Sub
        
        For i = 1 To rs.RecordCount

            With Me.FgSalePrice
            
                .TextMatrix(LngRow, .ColIndex("BranchId")) = val(rs("branch_id").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(LngRow, .ColIndex("BranchName")) = rs("branch_name").value
                Else
                    .TextMatrix(LngRow, .ColIndex("BranchName")) = rs("branch_namee").value
                End If
                                   
                If auto = False Then
                    .TextMatrix(LngRow, .ColIndex("UnitID")) = val(Me.DcUnit.BoundText)
                                     
                    .TextMatrix(LngRow, .ColIndex("UnitName")) = Me.DcUnit.text
                    .TextMatrix(LngRow, .ColIndex("Price1")) = val(Me.TxtPrice(0).text)
                    .TextMatrix(LngRow, .ColIndex("Price2")) = val(Me.TxtPrice(1).text)
                    .TextMatrix(LngRow, .ColIndex("Price3")) = val(Me.TxtPrice(2).text)
                    .TextMatrix(LngRow, .ColIndex("Price4")) = val(Me.TxtPrice(3).text)
                    .TextMatrix(LngRow, .ColIndex("Price5")) = val(Me.TxtPrice(4).text)
                    .TextMatrix(LngRow, .ColIndex("Price6")) = val(Me.TxtPrice(5).text)
                    .TextMatrix(LngRow, .ColIndex("Discount1")) = val(Me.TxtDiscount(0).text)
                    .TextMatrix(LngRow, .ColIndex("Discount2")) = val(Me.TxtDiscount(1).text)
                    .TextMatrix(LngRow, .ColIndex("Discount3")) = val(Me.TxtDiscount(2).text)
                    .TextMatrix(LngRow, .ColIndex("Discount4")) = val(Me.TxtDiscount(3).text)
                    .TextMatrix(LngRow, .ColIndex("Discount5")) = val(Me.TxtDiscount(4).text)
                    .TextMatrix(LngRow, .ColIndex("Discount6")) = val(Me.TxtDiscount(5).text)
                Else
                    GetDefaultItemUnit val(XPTxtID.text), unitid, Unitname
                    .TextMatrix(LngRow, .ColIndex("UnitID")) = unitid
                                     
                    .TextMatrix(LngRow, .ColIndex("UnitName")) = Unitname
                    .TextMatrix(LngRow, .ColIndex("Price1")) = val(saleprice1)
                                       
                    .TextMatrix(LngRow, .ColIndex("Price2")) = val(saleprice2)
                    .TextMatrix(LngRow, .ColIndex("Price3")) = val(saleprice3)
                    .TextMatrix(LngRow, .ColIndex("Price4")) = 0
                    .TextMatrix(LngRow, .ColIndex("Price5")) = 0
                    .TextMatrix(LngRow, .ColIndex("Price6")) = 0
                    .TextMatrix(LngRow, .ColIndex("Discount1")) = 0
                    .TextMatrix(LngRow, .ColIndex("Discount2")) = 0
                    .TextMatrix(LngRow, .ColIndex("Discount3")) = 0
                    .TextMatrix(LngRow, .ColIndex("Discount4")) = 0
                    .TextMatrix(LngRow, .ColIndex("Discount5")) = 0
                    .TextMatrix(LngRow, .ColIndex("Discount6")) = 0
                End If
                                 
                .Rows = .Rows + 1
                LngRow = LngRow + 1
                rs.MoveNext
                '    .AutoSize 0, .Cols - 1, False
            End With

        Next i

    Else

        If val(Dcbranch.BoundText) = 0 Then
            MsgBox "Õœœ ›—⁄ «Ê·« "
            Exit Sub
        End If

        With Me.FgSalePrice
            
            .TextMatrix(LngRow, .ColIndex("BranchId")) = val(Me.Dcbranch.BoundText)
            .TextMatrix(LngRow, .ColIndex("BranchName")) = Me.Dcbranch.text
                                    
            .TextMatrix(LngRow, .ColIndex("UnitID")) = val(Me.DcUnit.BoundText)
            .TextMatrix(LngRow, .ColIndex("UnitName")) = Me.DcUnit.text
                                    
            .TextMatrix(LngRow, .ColIndex("Price1")) = val(Me.TxtPrice(0).text)
            .TextMatrix(LngRow, .ColIndex("Price2")) = val(Me.TxtPrice(1).text)
            .TextMatrix(LngRow, .ColIndex("Price3")) = val(Me.TxtPrice(2).text)
            .TextMatrix(LngRow, .ColIndex("Price4")) = val(Me.TxtPrice(3).text)
            .TextMatrix(LngRow, .ColIndex("Price5")) = val(Me.TxtPrice(4).text)
            .TextMatrix(LngRow, .ColIndex("Price6")) = val(Me.TxtPrice(5).text)
            .TextMatrix(LngRow, .ColIndex("Discount1")) = val(Me.TxtDiscount(0).text)
            .TextMatrix(LngRow, .ColIndex("Discount2")) = val(Me.TxtDiscount(1).text)
            .TextMatrix(LngRow, .ColIndex("Discount3")) = val(Me.TxtDiscount(2).text)
            .TextMatrix(LngRow, .ColIndex("Discount4")) = val(Me.TxtDiscount(3).text)
            .TextMatrix(LngRow, .ColIndex("Discount5")) = val(Me.TxtDiscount(4).text)
            .TextMatrix(LngRow, .ColIndex("Discount6")) = val(Me.TxtDiscount(5).text)
                                 
            '    .AutoSize 0, .Cols - 1, False
        End With

    End If
 
    For i = 0 To 5
        TxtPrice(i).text = ""
        TxtDiscount(i).text = ""
    Next i
 
End Sub

Private Sub AddNewRow1old()

    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long

    If TxtPriceName = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»    «”„ «·”⁄— ...!!!"
        Else
            Msg = "must Specify Price Namet...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
  
    If TxtSalesPrice.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ    ”⁄— «·»Ì⁄  ...!!!"
        Else
            Msg = "must Enter Sales Price ...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 
    'If FgPrices.Rows = 1 Then
    '    LngRow = Val(Me.TxtRowNumber.text)
    'Else
    Me.FgPrices.Rows = Me.FgPrices.Rows + 1
    LngRow = Me.FgPrices.Rows - 1
    'End If
  
    On Error Resume Next

    With Me.FgPrices
    
        If Me.ChkDefSalePrice.value = vbChecked Then
            .Cell(flexcpChecked, .FixedRows, .ColIndex("DefaultUnit"), .Rows - 1, .ColIndex("DefaultUnit")) = flexUnchecked
            .Cell(flexcpChecked, LngRow, .ColIndex("DefaultUnit")) = flexChecked
        End If
    
        .TextMatrix(LngRow, .ColIndex("PriceName")) = Me.TxtPriceName.text
        .TextMatrix(LngRow, .ColIndex("Pricevalue")) = val(Me.TxtSalesPrice.text)
        .TextMatrix(LngRow, .ColIndex("des")) = Me.TxtPriceDes.text
        .TextMatrix(LngRow, .ColIndex("CustomerOrVendor")) = 0
   
        '    .AutoSize 0, .Cols - 1, False
    End With

    Me.ChkDefSalePrice.value = vbUnchecked
 
    Me.TxtPriceName.text = ""
    Me.TxtSalesPrice.text = ""
    Me.TxtPriceDes.text = ""
 
End Sub

Private Sub AddNewRow()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long

    If val(Me.DcboUnits.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «·ÊÕœ…...!!!"
        Else
            Msg = "must select Unit...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If val(Me.TxtRowNumber.text) = 0 Then
        LngFindRow = FgUnites.FindRow(val(Me.DcboUnits.BoundText), FgUnites.FixedRows, FgUnites.ColIndex("UnitID"), False, True)

        If LngFindRow <> -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«Ì„ﬂ‰  ﬂ—«— «·ÊÕœ…  ...!!!"
            Else
                Msg = " Can't Repeat unit  ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    If val(Me.TxtUnitFactor.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ⁄·«ﬁ… «·ÊÕœ… ...!!!"
        Else
            Msg = "must Enter Unit factor ...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If val(Me.TxtRowNumber.text) <> 0 Then
        LngRow = val(Me.TxtRowNumber.text)
    Else
        Me.FgUnites.Rows = Me.FgUnites.Rows + 1
        LngRow = Me.FgUnites.Rows - 1
    End If

    If LngRow = 1 Then
        If val(Me.TxtUnitFactor.text) > 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "›Ï Õ«·… «‰  ﬂÊ‰ Â–Â «Ê· ÊÕœ… ··’‰› ·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ „⁄«„· «· ÕÊÌ· «ﬂ»— „‰ Ê«Õœ"
            Else
                Msg = "because this is the first unit for this items So Unit Factor must be 1"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtUnitFactor.text = 1
        End If
    End If

    On Error Resume Next

    With Me.FgUnites

        If Me.ChkDef.value = vbChecked Then
            .Cell(flexcpChecked, .FixedRows, .ColIndex("DefaultUnit"), .Rows - 1, .ColIndex("DefaultUnit")) = flexUnchecked
            .Cell(flexcpChecked, LngRow, .ColIndex("DefaultUnit")) = flexChecked
        End If

        .TextMatrix(LngRow, .ColIndex("UnitID")) = Me.DcboUnits.BoundText
        .TextMatrix(LngRow, .ColIndex("UnitName")) = Me.DcboUnits.text
        .TextMatrix(LngRow, .ColIndex("UnitFactor")) = Format(val(Me.TxtUnitFactor.text), "0.000")
        .TextMatrix(LngRow, .ColIndex("UnitSalesPrice")) = val(Me.TxtUnitSalesPrice.text)
        .TextMatrix(LngRow, .ColIndex("UnitPurPrice")) = val(Me.TxtUnitPurPrice.text)
        .TextMatrix(LngRow, .ColIndex("SecOrder")) = val(.TextMatrix(LngRow - 1, .ColIndex("SecOrder"))) + 1
        WriteDes LngRow
        .AutoSize 0, .Cols - 1, False
    End With

    Me.ChkDef.value = vbUnchecked

    Me.DcboUnits.BoundText = ""
    Me.TxtUnitFactor.text = ""
    Me.TxtUnitSalesPrice.text = ""
    Me.TxtUnitPurPrice.text = ""

    Me.TxtRowNumber.text = ""
    Me.DcboUnits.SetFocus
End Sub

Private Sub Del_Item()
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        If RelatedItemTrans = True Then
            Exit Sub
        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·’‰› —ﬁ„ " & Chr(13)
            Msg = Msg + (XPTxtID.text) & Chr(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = " delete item ID :  " & Chr(13)
            Msg = Msg + (XPTxtID.text) & Chr(13)
            Msg = Msg + " Delete y/n?"
    
        End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                TreeItems.Nodes.Remove (rs("ItemID").value & "I")
                CuurentLogdata ("D")
                rs.delete
                StrSQL = "Delete From TblItemCatalog Where ItemID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = "Delete From TblItemsUnits Where ItemID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblAotherItems Where IDItem=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "Delete From TblItemsPrices Where ItemID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                If ChkAssplied.Visible = True Then
                    StrSQL = "Delete From TblItemsParts Where ItemID=" & val(Me.XPTxtID.text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                End If
            
                If ChkItemMakingNew.Visible = True Then
                    StrSQL = "Delete From TblItemsParts Where ItemID=" & val(Me.XPTxtID.text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                End If
            
                If ChkRelated.Visible = True Then
                    StrSQL = "Delete From TblItemsAttach Where ItemID=" & val(Me.XPTxtID.text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                End If

                Set Me.ImgPic.Picture = Nothing
                rs.MoveFirst

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

    Else
        clear_all Me

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "invalid operations no items to delete"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·’‰› "
        Else
            Msg = "cant' delete this items .... data integrity " & Chr(13) & "this items founded in transactions"
        End If

        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean
    On Error GoTo ErrTrap

    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ’‰› ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–« «·’‰›" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·’‰› «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ’‰›" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ’‰›" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·√’‰«›", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hwnd, "Items Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Add New Item...", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print this item data", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit this item data", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the new item data " & Wrap & "Or save the editing in the current record", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo enter new item" & Wrap & "Or Undo in the current editing", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete..." & Wrap & "Delete this item data", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Search for an item", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist Record" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next" & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last" & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help" & Wrap & "Show the Help File" & Wrap & "" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim XNode As MSComctlLib.Node
    Dim Msg As String
    Dim BeginTrans As Boolean
    Dim IntFgItems As Integer
    Dim RsParts As ADODB.Recordset
    Dim RsAttachs As ADODB.Recordset
   '  On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If XPTxtName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·’‰›"
            Else
                Msg = "please Enter Item Name "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.C1Tab1.CurrTab = 0
            XPTxtName.SetFocus
            Exit Sub
        End If

        If XPTxtCode.text = "" Then

            'XPTxtCode.Text = XPTxtID.Text
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· ﬂÊœ «·’‰›"
            Else
                Msg = "please Enter Item Code "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.C1Tab1.CurrTab = 0
            XPTxtCode.SetFocus
            Exit Sub
        End If

        If Me.CboItemType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ ﬁ„ » ÕœÌœ Â· «·’‰› ”·⁄… √„ Œœ„…...!!"
            Else
                Msg = "please Specify this item is Goods or service? "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.C1Tab1.CurrTab = 0
            CboItemType.SetFocus
            Exit Sub
        End If
    
        If Me.ChkGuar.value = 1 Then
            If TxtGuarValue.text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "„‰ ›÷·ﬂ ﬁ„ »ﬂ «»… „œ… «·÷„«‰...!!"
                Else
                    Msg = "please Enter Gurantee Interval"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.C1Tab1.CurrTab = 0
                CboItemType.SetFocus
                Exit Sub
            End If
        End If
 
        If Me.ChkGuar.value = 1 Then
            If OptGaurType(0).value = False And OptGaurType(1).value = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "„‰ ›÷·ﬂ ﬁ„ » ÕœÌœ „œ… «·÷„«‰...!!"
                Else
                    Msg = "please Enter Gurantee Interval"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.C1Tab1.CurrTab = 0
                CboItemType.SetFocus
                Exit Sub
            End If
        End If
    
        If Me.XPCboGroup.BoundText = "1" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«Ì„ﬂ‰  ”ÃÌ· «·√’‰«› „»«‘—… ⁄·Ï ‘Ã—… «·√’‰«›"
            Else
                Msg = "Can't Add Items Directly At Items Tree"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.C1Tab1.CurrTab = 0
            XPCboGroup.SetFocus
            Exit Sub
        End If

        Select Case TxtModFlg.text

            Case "N"
                XPTxtID.text = CStr(new_id("TblItems", "ItemID", "", True))
                'XPTxtCode.Text
                StrSQL = "select * From TblItems where ItemName='" & Trim(XPTxtName.text) & "'"
           
           
           '    If TxtPartNo.text <> "" Then
           '    StrSQL = StrSQL & " and  PartNo='" & Trim(TxtPartNo.text) & "'"
           '    End If
               
 If SystemOptions.DuplicateitemsNames = False Then
                                        
                                        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        
                                        If RsTemp.RecordCount > 0 Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                Msg = "ÌÊÃœ ’‰› „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                                                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                                                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·’‰›"
                                            Else
                                                Msg = "This item Name Already Exist" & Chr(13)
                                            End If
                        
                                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                            Me.C1Tab1.CurrTab = 0
                                            XPTxtName.SetFocus
                                            Exit Sub
                                        End If
                        
                                        RsTemp.Close
 End If
 
                StrSQL = "select * From TblItems where ItemCode='" & Trim(XPTxtCode.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÊÃœ ’‰› „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·ﬂÊœ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ﬂÊœ «·’‰›"
                    Else
                        Msg = "This item Code Already Exist" & Chr(13)
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Me.C1Tab1.CurrTab = 0
                    '                XPTxtCode.SetFocus
                    Exit Sub
                End If

                RsTemp.Close

'check Barcode
'************************************************************************
If TxtbarCodeNO.text <> "" Then

                StrSQL = "select * From TblItems where barCodeNO='" & Trim(TxtbarCodeNO.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÊÃœ ’‰› „”Ã· „”»ﬁ« »Â–« «·»«—ﬂÊœ" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»«—ﬂÊœ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·»«—ﬂÊœ"
                    Else
                        Msg = "This item barcode Already Exist" & Chr(13)
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Me.C1Tab1.CurrTab = 0
                    TxtbarCodeNO.SetFocus
                    Exit Sub
                End If

                RsTemp.Close
      

End If
'**************************************************************************

            Case "E"
            ''
            StrSQL = "Delete From TblItemCatalog Where ItemID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblAotherItems Where IDItem=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblItemDiamonds Where ItemID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
           
            '''
                StrSQL = "select * From TblItems where ItemName='" & Trim(XPTxtName.text) & "'"
                
           '         If TxtPartNo.text <> "" Then
           '    StrSQL = StrSQL & " and  PartNo='" & Trim(TxtPartNo.text) & "'"
           '    End If
         If SystemOptions.DuplicateitemsNames = False Then
      
                                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                
                                If RsTemp.RecordCount > 0 Then
                                    If RsTemp("ItemID").value <> val(XPTxtID.text) Then
                                        If SystemOptions.UserInterface = ArabicInterface Then
                                            Msg = "ÌÊÃœ ’‰› „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·’‰›"
                                        Else
                                            Msg = "This item Name Already Exist" & Chr(13)
                                        End If
                
                                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                        Me.C1Tab1.CurrTab = 0
                                        XPTxtName.SetFocus
                                        Exit Sub
                                    End If
                                End If
                                RsTemp.Close
End If
                
                StrSQL = "select * From TblItems where ItemCode='" & Trim(XPTxtCode.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("ItemID").value <> val(XPTxtID.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÊÃœ ’‰› „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ" & Chr(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·ﬂÊœ «·’ÕÌÕ " & Chr(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ﬂÊœ «·’‰›"
                        Else
                            Msg = "This item Code Already Exist" & Chr(13)
                        End If

                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Me.C1Tab1.CurrTab = 0
                        XPTxtCode.SetFocus
                        Exit Sub
                    End If
                End If

                RsTemp.Close
        
 '********************************************************************************************
 If TxtbarCodeNO.text <> "" Then
              StrSQL = "select * From TblItems where barCodeNO='" & Trim(TxtbarCodeNO.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("ItemID").value <> val(XPTxtID.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÊÃœ ’‰› „”Ã· „”»ﬁ« »Â–« «·»«—ﬂÊœ" & Chr(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»«—ﬂÊœ «·’ÕÌÕ " & Chr(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“   «·»«—ﬂÊœ"
                        Else
                            Msg = "This item Name Already Exist" & Chr(13)
                        End If

                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Me.C1Tab1.CurrTab = 0
                        TxtbarCodeNO.SetFocus
                        Exit Sub
                    End If
                End If

                RsTemp.Close
           End If
           
   '********************************************************************************************
        
        
        End Select

        If XPCboGroup.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Ã„Ê⁄…" & Chr(13)
                Msg = Msg + "«· Ì Ì‰ „Ì «·ÌÂ« Â–« «·’‰›"
            Else
                Msg = "Please Specify item Group" & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.C1Tab1.CurrTab = 0
            XPCboGroup.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Me.CboItemCase.ListIndex = -1 Then
            Me.CboItemCase.ListIndex = 0
        End If

        If TxtRequired.text <> "" Then
            If Not IsNumeric(TxtRequired.text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœ «·ÿ·» ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…" & Chr(13)
                Else
                    Msg = "Required Quantity Must be Numeric Only" & Chr(13)
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.C1Tab1.CurrTab = 0
                TxtRequired.SetFocus
                SelectText TxtRequired
                Exit Sub
            End If
        End If

        If XPTxtPurchase.text <> "" Then
            If Not IsNumeric(XPTxtPurchase.text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "”⁄— «·‘—«¡ ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…" & Chr(13)
                Else
                    Msg = "Purchase price must be Numeric" & Chr(13)
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.C1Tab1.CurrTab = 0
                XPTxtPurchase.SetFocus
                Exit Sub
            End If
        End If

        If XPTxtSall.text <> "" Then
            If Not IsNumeric(XPTxtSall.text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "”⁄— «·»Ì⁄ ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…" & Chr(13)
                Else
                    Msg = "sale price must be Numeric" & Chr(13)
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.C1Tab1.CurrTab = 0
                XPTxtSall.SetFocus
                Exit Sub
            End If
        End If

        If ChkAssplied.value = vbChecked Then
            IntFgItems = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("ItemID"))

            If IntFgItems < 2 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "›Ï Õ«·… ﬂÊ‰ «·’‰› „Ã„⁄"
                    Msg = Msg & Chr(13) & "›«‰Â ÌÃ» ⁄·Ìﬂ «‰  ﬁÊ„ »≈œŒ«· ’‰›Ì‰ ⁄·Ï «·√ﬁ· "
                Else
                    Msg = "in Composite Item "
                    Msg = Msg & Chr(13) & "You must insert at least two items "
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.C1Tab1.CurrTab = 1
                Fg.SetFocus
                Exit Sub
            End If
        End If
    
        If ChkItemMakingNew.value = vbChecked Then
            '        IntFgItems = ModFgLib.GetItemsInFg(FG, FG.ColIndex("ItemID"))
            '                If IntFgItems < 2 Then
            '                                            If SystemOptions.UserInterface = ArabicInterface Then
            '                                                Msg = "›Ï Õ«·… ﬂÊ‰ «·’‰› „‰ Ã"
            '                                                Msg = Msg & Chr(13) & "›«‰Â ÌÃ» ⁄·Ìﬂ «‰  ﬁÊ„ »≈œŒ«· ’‰›Ì‰ ⁄·Ï «·√ﬁ· "
            '                                            Else
            '                                                Msg = "in fINIem "
            '                                                Msg = Msg & Chr(13) & "You must insert at least two items "
            '                                            End If
            ''                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '                   Me.C1Tab1.CurrTab = 1
            '                   FG.SetFocus
            '                   Exit Sub
            '               End If
        End If
    
        If ChkRelated.value = vbChecked Then
            IntFgItems = ModFgLib.GetItemsInFg(FgAttachs, FgAttachs.ColIndex("ItemID"))

            If IntFgItems < 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "›Ï Õ«·… ﬂÊ‰ «·’‰› √’‰«› „·Õﬁ…"
                    Msg = Msg & Chr(13) & "›«‰Â ÌÃ» ⁄·Ìﬂ «‰  ﬁÊ„ »≈œŒ«· ’‰› Ê«Õœ ⁄·Ï «·√ﬁ· "
                Else
                    Msg = "because this item have attached items So, "
                    Msg = Msg & Chr(13) & "You must insert at least one items "
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.C1Tab1.CurrTab = 1
                Fg.SetFocus
                Exit Sub
            End If
        End If

'********************************************************
 
    Dim lngCount As Long
    Dim IntDefUnitRow As Integer
    'If Val(Me.DcboItems1.BoundText) = 0 Then
    '    Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰› ...!!!"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If
    lngCount = ItemsInGrid()
    If lngCount <= 0 Then
        Msg = "ÌÃ» ≈œŒ«· ÊÕœ… ⁄·Ï «·√ﬁ· ....!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
ElseIf Me.FgUnites.FixedRows + 1 = Me.FgUnites.Rows Then
        With Me.FgUnites
           .Cell(flexcpChecked, 1, .ColIndex("DefaultUnit")) = flexChecked
       End With
    Else
        If GetFgCheckCount() = 0 Then
        Msg = "ÌÃ»  ÕœÌœ ÊÕœ… ≈› —«÷Ì… ··’‰› ....!!!"
           MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If
 '********************************************************
        Cn.BeginTrans
        
        BeginTrans = True

        If TxtModFlg.text = "N" Then
            rs.AddNew
            rs("ItemID").value = IIf(XPTxtID.text = "", 0, val(XPTxtID.text))

        End If

        If FgSalePrice.Rows = 1 Then
            AddNewRow1 True, val(XPTxtSall.text), val(TxtCusPrice.text), val(TxtDealerPrice.text)
        End If
        
        rs("ItemCode").value = IIf(Trim(XPTxtCode.text) = "", "", Trim(XPTxtCode.text))
       
        rs("ItemNamee").value = IIf(XPTxtNamee.text = "", Trim(XPTxtName.text), Trim(XPTxtNamee.text))
            rs("BinLocation").value = IIf(TxtBinLocation.text = "", "", Trim(TxtBinLocation.text))
        
            rs("minvalueqty").value = IIf(Txtminvalueqty.text = "", Null, val(Txtminvalueqty.text))
        
            rs("MaxValueqty").value = IIf(TxtMaxValueqty.text = "", Null, val(TxtMaxValueqty.text))
        
        
        rs("PartNo").value = IIf(TxtPartNo.text = "", "", Trim(TxtPartNo.text))
         rs("FreeQty").value = IIf(TxtFreeQty.text = "", 0, val(TxtFreeQty.text))
         
         
         rs("CatlogNO").value = IIf(TxtCatlogNO.text = "", "", Trim(TxtCatlogNO.text))
         rs("FactoryNO").value = IIf(TxtFactoryNO.text = "", "", Trim(TxtFactoryNO.text))
         
             rs("TemplateID").value = IIf(DcTemplate.BoundText = "", 0, val(DcTemplate.BoundText))

        rs("HaveSerial").value = XPChkSerial.value
        rs("PurchasePrice").value = IIf(XPTxtPurchase.text = "", Null, Trim(XPTxtPurchase.text))
        rs("SallingPrice").value = IIf(XPTxtSall.text = "", Null, Trim(XPTxtSall.text))
        rs("LastUpdate").value = Date

        If XPCboGroup.BoundText = "" Then
            rs("GroupID").value = Null
        Else
            rs("GroupID").value = val(XPCboGroup.BoundText)
        End If
     rs("OverHead").value = val(TxtOverHead.text)
     rs("Wight").value = val(TxtWight.text)
      rs("Content").value = (txtContent.text)
     rs("Dippre").value = (txtDippre.text)
     
     rs("Source").value = (TxtSource.text)
     rs("Typenew").value = (txtTypenew.text)
     
     
        rs("DefaultSupplier").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    
        If Me.CboItemCase.ListIndex = 0 Then
            rs("ItemCase").value = 1
        Else
            rs("ItemCase").value = 2
        End If

        rs("RequestLimit").value = IIf(TxtRequired.text = "", Null, Trim(TxtRequired.text))
    
        If ImgPic.Picture = 0 Then
            rs("ItemPhoto").value = Null
        Else

            If SavePictureToDB(ImgPic, rs, "ItemPhoto") = False Then
                GoTo ErrTrap
            End If
        End If

       rs("ItemMaxDiscount").value = val(Me.TxtItemMaxDiscount.text)
        rs("CustomerPrice").value = val(Me.TxtCusPrice.text)
        rs("DealerPrice").value = val(Me.TxtDealerPrice.text)

        If Me.ChkGuar.value = vbChecked Then
            rs("HaveGuarantee").value = Me.ChkGuar.value
            rs("GuaranteeValue").value = val(Me.TxtGuarValue.text)
            rs("GuaranteeType").value = IIf(OptGaurType(0).value = True, 0, 1)
        Else
            rs("HaveGuarantee").value = False
            rs("GuaranteeValue").value = 0
            rs("GuaranteeType").value = 0
        End If

        rs("IsArchive").value = IIf(Me.ChkAr.value = vbChecked, 1, 0)

        If Me.CboItemType.ListIndex = 0 Then
            rs("ItemType").value = 0
        Else
            rs("ItemType").value = 1
        End If

        If ChkAssplied.Visible = True Then
            If ChkAssplied.value = vbChecked Then
                rs("AssbliedItem").value = True
            ElseIf ChkAssplied.value = vbUnchecked Then
                rs("AssbliedItem").value = False
            End If
        End If

        '   ’‰› Ì „ «‰ «Ã…
        If ChkItemMakingNew.Visible = True Then
            If ChkItemMakingNew.value = vbChecked Then
                rs("ItemMakingNew").value = True
            ElseIf ChkItemMakingNew.value = vbUnchecked Then
                rs("ItemMakingNew").value = False
            End If
        End If
    
        '   ’‰› „’‰€
        If chkItemMaking.Visible = True Then
            If chkItemMaking.value = vbChecked Then
                rs("ItemMaking").value = True
            ElseIf chkItemMaking.value = vbUnchecked Then
                rs("ItemMaking").value = False
            End If
        End If
    
        If ChkRelated.Visible = True Then
            If ChkRelated.value = vbChecked Then
                rs("RelatedItem").value = True
            Else
                rs("RelatedItem").value = False
            End If
        End If

        rs("ItemComment").value = IIf(Trim(Me.TxtItemComment.text) = "", Null, Trim(Me.TxtItemComment.text))
        rs("Branch_NO").value = val(branch_id)
        rs("code").value = txtid.text
        rs("Fullcode").value = IIf(DCPreFix.text = "", Null, DCPreFix.text) & IIf(Trim(txtid.text) = "", Null, txtid.text)
        rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)
If TxtbarCodeNO.text = "" Then
TxtbarCodeNO = rs("Fullcode").value
End If

rs("barCodeNO").value = IIf(TxtbarCodeNO.text = "", "", Trim(TxtbarCodeNO.text))
 'XPTxtName.text = XPTxtName.text & Me.TxtbarCodeNO.text
        rs("ItemName").value = IIf(XPTxtName.text = "", "", Trim(XPTxtName.text))
        rs.update

        If ChkAssplied.value = vbChecked Then
            If Me.TxtModFlg.text = "E" Then
                StrSQL = "Delete From TblItemsParts Where ItemID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If

            If ChkAssplied.value = vbChecked Then
                Set RsParts = New ADODB.Recordset
           '     RsParts.Open "TblItemsParts", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            StrSQL = "SELECT     *  from dbo.TblItemsParts Where (1 = -1)"
               RsParts.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
               
                For i = 1 To Me.Fg.Rows - 1
                    RsParts.AddNew
                    RsParts("ItemID").value = val(Me.XPTxtID.text)
                    RsParts("PartItemID").value = val(Fg.TextMatrix(i, Fg.ColIndex("ItemID")))
                    
                    RsParts("PartItemQty").value = val(Fg.TextMatrix(i, Fg.ColIndex("ItemQty")))
                    RsParts("PartItemPrice").value = val(Fg.TextMatrix(i, Fg.ColIndex("ItemPrice")))
                    RsParts("UnitID").value = val(Fg.TextMatrix(i, Fg.ColIndex("UnitID")))
                      
                    RsParts.update
                Next i

            End If
        End If
    ''//«·»œ«∆·
        
                
        '’‰› Ì „ «‰ «Ã…
        If ChkItemMakingNew.value = vbChecked Then
            If Me.TxtModFlg.text = "E" Then
                StrSQL = "Delete From TblItemsParts Where ItemID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If

            If ChkItemMakingNew.value = vbChecked Then
                Set RsParts = New ADODB.Recordset
              '  RsParts.Open "TblItemsParts", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
StrSQL = "SELECT     *  from dbo.TblItemsParts Where (1 = -1)"
   RsParts.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
                For i = 1 To Me.Fg.Rows - 1
                    RsParts.AddNew
                    RsParts("ItemID").value = val(Me.XPTxtID.text)
                    RsParts("PartItemID").value = val(Fg.TextMatrix(i, Fg.ColIndex("ItemID")))
                    RsParts("PartItemQty").value = val(Fg.TextMatrix(i, Fg.ColIndex("ItemQty")))
                    RsParts("PartItemPrice").value = val(Fg.TextMatrix(i, Fg.ColIndex("ItemPrice")))
                    RsParts("unitid").value = val(Fg.TextMatrix(i, Fg.ColIndex("unitid")))
                    
                    RsParts.update
                Next i

            End If
        End If
    
        If Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From TblItemsAttach Where ItemID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If

        If ChkRelated.value = vbChecked Then
            Set RsAttachs = New ADODB.Recordset
          '  RsAttachs.Open "TblItemsAttach", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
StrSQL = "SELECT     *  from dbo.TblItemsAttach Where (1 = -1)"
   RsAttachs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
            For i = 1 To Me.FgAttachs.Rows - 1
                RsAttachs.AddNew
                RsAttachs("ItemID").value = val(Me.XPTxtID.text)
                RsAttachs("AttachItemID").value = val(FgAttachs.TextMatrix(i, FgAttachs.ColIndex("ItemID")))
                RsAttachs("AttachItemQty").value = val(FgAttachs.TextMatrix(i, FgAttachs.ColIndex("ItemQty")))
                RsAttachs("AttachItemPrice").value = val(FgAttachs.TextMatrix(i, FgAttachs.ColIndex("ItemPrice")))
                RsAttachs.update
            Next i

        End If
    '''''''''///////////////////
     Set RsParts = New ADODB.Recordset
              '  RsParts.Open "TblItemDiamonds", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
StrSQL = "SELECT     *  from dbo.TblItemDiamonds Where (1 = -1)"
    RsParts.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
With Me.fgCameo
                For i = 1 To .Rows - 1
                     If .TextMatrix(i, .ColIndex("type")) <> "" Then
                    RsParts.AddNew
                    RsParts("ItemID").value = val(Me.XPTxtID.text)
                    RsParts("type").value = (.TextMatrix(i, .ColIndex("type")))
                    
                    RsParts("unite").value = (.TextMatrix(i, .ColIndex("unite")))
                    RsParts("weight").value = (.TextMatrix(i, .ColIndex("weight")))
                    RsParts("indexe").value = 1
                      
                    RsParts.update
                    End If
                Next i
                
              End With
   ''//
             Set RsParts = New ADODB.Recordset
     '           RsParts.Open "TblItemDiamonds", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
StrSQL = "SELECT     *  from dbo.TblItemCatalog Where (1 = -1)"
   RsParts.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
With Me.VSFlexGrid2
                For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("CatlogName")) <> "" Then
                    RsParts.AddNew
                    RsParts("ItemID").value = val(Me.XPTxtID.text)
                    RsParts("CatlogName").value = (.TextMatrix(i, .ColIndex("CatlogName")))
                    
                    RsParts("CatloPath").value = (.TextMatrix(i, .ColIndex("CatloPath1")))
                   
                      
                    RsParts.update
                    End If
                Next i
                
              End With
 ''///
   '''/
              
          Set RsParts = New ADODB.Recordset
     '           RsParts.Open "TblItemDiamonds", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
StrSQL = "SELECT     *  from dbo.TblItemDiamonds Where (1 = -1)"
   RsParts.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
With Me.fgDiamonds
                For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("type")) <> "" Then
                    RsParts.AddNew
                    RsParts("ItemID").value = val(Me.XPTxtID.text)
                    RsParts("type").value = (.TextMatrix(i, .ColIndex("type")))
                    
                    RsParts("unite").value = (.TextMatrix(i, .ColIndex("unite")))
                    RsParts("weight").value = (.TextMatrix(i, .ColIndex("weight")))
                     RsParts("quality").value = (.TextMatrix(i, .ColIndex("ÛQuality")))
                    
                    RsParts("color").value = (.TextMatrix(i, .ColIndex("color")))
                    RsParts("Gestonf").value = (.TextMatrix(i, .ColIndex("weight")))
                  RsParts("indexe").value = 0
                      
                    RsParts.update
                    End If
                Next i
                
              End With
 ''///
  Set RsParts = New ADODB.Recordset
           '     RsParts.Open "TblItemsParts", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            StrSQL = " SELECT     *  from dbo.TblAotherItems Where (1 = -1)"
               RsParts.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
               
                For i = 1 To Me.VSFlexGrid1.Rows - 1
                    RsParts.AddNew
                    RsParts("IDItem").value = val(Me.XPTxtID.text)
                    RsParts("ItemID").value = val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("ItemID")))
                    
                    RsParts("Quntity").value = val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("ItemQty")))
                    RsParts("Valu").value = val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("ItemPrice")))
                    RsParts("UnitID").value = val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("UnitID")))
                    RsParts("Remark").value = VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("Remarks"))
                    RsParts.update
                Next i
 ''//
        Cn.CommitTrans
        BeginTrans = False

        If TxtModFlg.text = "E" Then
            TreeItems.Nodes.Remove (rs("ItemID").value & "I")
        End If

        Set XNode = TreeItems.Nodes.Add(Trim(rs("GroupID").value) & "G", tvwChild, rs("ItemID").value & "I", rs("ItemName").value, "Item")
        TreeItems.Nodes(rs("ItemID").value & "I").Selected = True
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"
                SaveData_Unites
                SaveData_Prices

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·’‰›" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Data was Saved " & Chr(13)
                    Msg = Msg + "do you want enter another item y/n?"
           
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                      Frame1.Enabled = True
                    Exit Sub
                End If
  
            Case "E"
                SaveData_Unites
                SaveData_Prices

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "changes Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.text = "R"
    End If

       Dim Dcombos As ClsDataCombos
            Set Dcombos = New ClsDataCombos
            Dcombos.GetItemsNames Me.DcboItems
            
           Retrive (val(XPTxtID.text))
           DcboItems1_Change

          DataPassing
         
            
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        If rs.EditMode <> adEditNone Then
            rs.CancelUpdate
        End If

        BeginTrans = False
        Cn.RollbackTrans
    End If
    
    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…" & Chr(13)
        Else
            Msg = "Can't Save ,Error in enterd Data  " & Chr(13)
        End If

        Msg = Msg + "Err.Description" & Err.description & Chr(13)
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
        Msg = "Sorry........ Error During Saving " & Chr(13)
    End If

    Msg = Msg + "Err.Description" & Err.description & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Set ItemReport = New ClsItemsReport
        ItemReport.ItemData XPTxtID.text
    End If

    Exit Sub
ErrTrap:
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
ErrTrap:         End Sub

Private Sub LoadMenus()
    On Error GoTo ErrTrap

    With Me.XPPopUp
        'Clear the Menu and ToolBars
        .ClearAll
        .SetImageList mdifrmmain.img16

        If SystemOptions.UserInterface = ArabicInterface Then
            .RightToLeft = True

            With .Menus.Add("mnuDropMenu1", tsSecondaryMenu, True)
                .MenuItems.Add tsMenuCaption, " ÕÊÌ· «·’‰› ≈·Ï „Ã„Ê⁄…", False, True, 11, , , , , "Convert", , , , " ÕÊÌ· «·’‰› ≈·Ï „Ã„Ê⁄…"
                .MenuItems.Add tsMenuCaption, "≈÷«›… ’‰›", False, True, 2, , , , , "AddItem", , , "≈÷«›… ’‰›"
                .MenuItems.Add tsMenuCaption, " ⁄œÌ· «·’‰›", False, True, 3, , , , , "EditItem", , , , " ⁄œÌ· «·’‰›"
                .MenuItems.Add tsMenuCaption, "Õ–› «·’‰›", False, True, 4, , , , , "DelItem", , , , "Õ–› «·’‰›"
                .MenuItems.Add tsMenuCaption, "„”Õ «·«Œ Ì«—", False, False, 5, , , True, , "ClearItem", , , , "„”Õ «·«Œ Ì«—"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "ﬁ’", False, False, 7, , , True, , "CutItem", , , , "ﬁ’"
                .MenuItems.Add tsMenuCaption, "·’ﬁ", False, False, 6, , , , , "PasteItem", , , , "·’ﬁ"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Œ’«∆’", False, False, 9, , , True, , "ItemProperties", , , , "Œ’«∆’"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "ÿ»«⁄… ", False, False, 10, , , True, , "PrintItem", , , , "ÿ»«⁄… ‘Ã—… «·√’‰«›"
            End With

        Else
            .RightToLeft = False

            With .Menus.Add("mnuDropMenu1", tsSecondaryMenu, True)
                .MenuItems.Add tsMenuCaption, "Convert item into group", False, True, 11, , , , , "Convert", , , , "Convert this item into group"
                .MenuItems.Add tsMenuCaption, "Add Item...", False, True, 2, , , , , "AddItem", , , "Add new item"
                .MenuItems.Add tsMenuCaption, "Edit Item...", False, True, 3, , , , , "EditItem", , , , "Eidt this item"
                .MenuItems.Add tsMenuCaption, "Delete Item...", False, True, 4, , , , , "DelItem", , , , "Delete this item"
                .MenuItems.Add tsMenuCaption, "Clear Cheked", False, False, 5, , , True, , "ClearItem", , , , "Clear Checked"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Cut", False, False, 7, , , True, , "CutItem", , , , "Cut"
                .MenuItems.Add tsMenuCaption, "Paste", False, False, 6, , , , , "PasteItem", , , , "Paste"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Properties", False, False, 9, , , True, , "ItemProperties", , , , "Properties"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Print", False, False, 10, , , True, , "PrintItem", , , , "Print Items Tree"
            End With

        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboGroup_Change()

If Me.TxtModFlg <> "R" Then
            If XPTxtName.text = "" Then
                    If SystemOptions.DecideItemName = True Then
                    XPTxtName.text = XPCboGroup.text
                     End If
                     
            End If

End If

End Sub

Private Sub XPCboGroup_Click(Area As Integer)
    On Error Resume Next
    Dim OverHead As Double
    
     GetGroupData val(XPCboGroup.BoundText), , , , , "groups", , , OverHead
     TxtOverHead.text = OverHead
If SystemOptions.WorkWithGroupCode = False Then Me.DCPreFix.text = "": Exit Sub
    If val(XPCboGroup.BoundText) = 0 Then Exit Sub
    Me.DCPreFix.text = GetPrefix(val(XPCboGroup.BoundText), "Groups")

     If Len(Me.DCPreFix.text) > 1 And (Mid(Me.DCPreFix.text, 1, 1)) = SystemOptions.itemSeprator Then
 
       Me.DCPreFix.text = Mid(Me.DCPreFix.text, 2, Len(Me.DCPreFix.text))
    End If
 
End Sub

Private Sub XPCboGroup_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetItemSGroups Me.XPCboGroup, False
        
    End If

End Sub

Private Sub XPChkSerial_Click()

    If Me.TxtModFlg.text = "E" Then
        If XPChkSerial.Tag = "" Then
            If RelatedItemTrans = True Then
                XPChkSerial.Tag = "Shown"
                XPChkSerial.value = IIf(rs("HaveSerial").value = True, vbChecked, vbUnchecked)
                XPChkSerial.Tag = ""
            End If
        End If
    End If

End Sub

Private Sub XPPopUp_MenuItemClick(ByVal MenuIndex As Integer, _
                                  ByVal MenuID As String, _
                                  ByVal MenuItemIndex As Integer, _
                                  ByVal MenuItemID As String)
    On Error GoTo ErrTrap
    Dim XNode As MSComctlLib.Node
    Dim RsTemp As ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim StrSQL As String
    Dim GroupID As Integer

    Select Case MenuItemID

        Case "Convert"
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open "Groups", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            StrSQL = "select * From Groups where GroupName='" & Trim(XPTxtName.text) & "'"
            RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTest.RecordCount > 0 Then
                Msg = " ÊÃœ „Ã„Ê⁄… „”Ã·… „”»ﬁ« »Â–« «·«”„" & Chr(13)
                Msg = Msg + "Ì„ﬂ‰ﬂ  ⁄œÌ· »Ì«‰«  Â–« «·’‰› " & Chr(13)
                Msg = Msg + "Ê«Œ Ì«— «”„ «·„Ã„Ê⁄…"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            RsTemp.AddNew
            GroupID = CStr(new_id("Groups", "GroupID", "", True))
            RsTemp("GroupID").value = GroupID
            RsTemp("GroupName").value = Trim(XPTxtName.text)
            RsTemp("ParentID").value = val(XPCboGroup.BoundText)
            RsTemp.update
            Dim Dcombos As ClsDataCombos
            Dcombos.GetItemSGroups Me.XPCboGroup
            cDboSearch(0).Refresh
        
            Set XNode = TreeItems.Nodes.Add(TreeItems.SelectedItem.Parent.key, tvwChild, GroupID & "G", Trim(XPTxtName.text), "Closed_Node", "Open_Node")
            StrSQL = "update TblItems set GroupID=" & val(GroupID) & " where ItemID=" & val(rs("ItemID").value)
            Cn.Execute StrSQL
            TreeItems.Nodes.Remove (TreeItems.SelectedItem.key)
            Set XNode = TreeItems.Nodes.Add(GroupID & "G", tvwChild, rs("ItemID") & "I", rs("ItemName"), "Item")
            Retrive (rs("ItemID"))

        Case "AddItem"
            Cmd_Click (0)

            Select Case right(TreeItems.SelectedItem.key, 1)

                Case "G"
                    XPCboGroup.BoundText = left(TreeItems.SelectedItem.key, Len(TreeItems.SelectedItem.key) - 1)

                Case "I"
                    XPCboGroup.BoundText = left(TreeItems.SelectedItem.Parent.key, Len(TreeItems.SelectedItem.Parent.key) - 1)
            End Select

        Case "EditItem"
            Cmd_Click (1)

        Case "DelItem"
            Cmd_Click (4)

        Case "ClearItem"

        Case "CutItem"
            TreeItems.SelectedItem.backcolor = vbGreen
            TxtCutKey.text = (TreeItems.SelectedItem.key)

            '        TxtMenuState.Text = "C"
        Case "PasteItem"
            TreeItems.Nodes.Remove (TxtCutKey.text)
            Set XNode = TreeItems.Nodes.Add(Trim(TreeItems.SelectedItem.key), tvwChild, rs("ItemID") & "I", rs("ItemName"), "Item")
            StrSQL = "update TblItems set GroupID=" & val(left(TreeItems.SelectedItem.key, Len(TreeItems.SelectedItem.key) - 1)) & " where ItemID=" & val(rs("ItemID").value)
            Cn.Execute StrSQL
            Retrive (val(rs("ItemID").value))

            '        TxtMenuState.Text = "N"
        Case "ItemProperties"

        Case "PrintItem"
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

 MySQL = " SELECT     dbo.TblItems.ItemID, dbo.TblItemDiamonds.type, dbo.TblItemDiamonds.unite, dbo.TblItemDiamonds.weight, dbo.TblItemDiamonds.indexe, dbo.TblItemDiamonds.Gestonf, dbo.TblItemDiamonds.color, dbo.TblItemDiamonds.quality"
MySQL = MySQL & " FROM         dbo.TblItems INNER JOIN"
 MySQL = MySQL & "      dbo.TblItemDiamonds ON dbo.TblItems.ItemID = dbo.TblItemDiamonds.ItemID"
MySQL = MySQL & " Where (dbo.TblItems.ItemID = " & val(XPTxtID.text) & ")"


'MySQL = MySQL & " Where (dbo.TblCommisRece.id =" & val(XPTxtID.text) & ")"

 

 
   
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepItemDiamondG.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepItemDiamondG.rpt"
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
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
Private Sub ChangeLang()
    'ChkRelated.Caption = "Assembled"
    lbl(34).Caption = "Catlog NO"
    lbl(35).Caption = "Factory NO"
    lbl(40).Caption = "Bin Location"
    lbl(45).Caption = "Free items %"
    lbl(46).Caption = "B.Code"
      lbl(43).Caption = "Template"
      lbl(44).Caption = "Max Disc."
    Cmd(26).Caption = "Delete"
    Cmd(27).Caption = "Delete"
    Cmd(28).Caption = "Delete All"
    Cmd(29).Caption = "Delete All"
Text1.Caption = "Avialble"
    lblLabel1.Caption = "Item Code"
    lblLabel2.Caption = "Item Name"
    lbl(25).Caption = "Qty"
    lbl(26).Caption = "price"
    lbl(27).Caption = "Items Count"
    Cmd(10).Caption = "Add"
    Cmd(11).Caption = "Delete"
    lbl(16).Caption = "Remark"
    chkItemMaking.Caption = "Item making"

    Frame2.Caption = "Quantities"
'    Label1.Caption = "Avilable"
    Label2.Caption = "Minimum"
    Label3.Caption = "Maximum"
    Label4.Caption = "Ord.QTY"
    Label5.Caption = "Rsv.QTY"

    lbl(33).Visible = False
    lbl(37).Visible = True

    With FgAttachs
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = " Item Code "
        .TextMatrix(0, .ColIndex("itemNAME")) = " item Name  "
        .TextMatrix(0, .ColIndex("ItemQty")) = "Item Qty"
        .TextMatrix(0, .ColIndex("ItemPrice")) = "Item Price"
    End With


    With Fg
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = " Item Code "
        .TextMatrix(0, .ColIndex("itemNAME")) = " item Name  "
        .TextMatrix(0, .ColIndex("ItemQty")) = "Item Qty"
        .TextMatrix(0, .ColIndex("ItemPrice")) = "Item Price"
        .TextMatrix(0, .ColIndex("UnitName")) = "UnitName"

    End With



    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = " Item Code "
        .TextMatrix(0, .ColIndex("itemNAME")) = " item Name  "
        .TextMatrix(0, .ColIndex("ItemQty")) = "Item Qty"
        .TextMatrix(0, .ColIndex("ItemPrice")) = "Item Price"
        .TextMatrix(0, .ColIndex("UnitName")) = "UnitName"
.TextMatrix(0, .ColIndex("Remarks")) = "Remarks"

    End With

lbl(42).Caption = "Item Code"
lbl(41).Caption = "Item Name"
lbl(38).Caption = "Unit"
lbl(39).Caption = "Price"
C1Tab1.TabCaption(7) = "Data Diamonds"
Cmd(24).Caption = "Add"
Cmd(25).Caption = "Del"
    With FgUnites
        .TextMatrix(0, .ColIndex("DefaultUnit")) = "Default Unit  "
        .TextMatrix(0, .ColIndex("UnitID")) = " Unit ID  "
        .TextMatrix(0, .ColIndex("UnitName")) = " Unit Name  "
        .TextMatrix(0, .ColIndex("UnitFactor")) = "Unit Factor"
        .TextMatrix(0, .ColIndex("UnitSalesPrice")) = "Unit SalesPrice"
        .TextMatrix(0, .ColIndex("UnitPurPrice")) = "Unit PurPrice"
        .TextMatrix(0, .ColIndex("SecOrder")) = "Sec Order"
    End With

    itemnamex(2).Caption = "Item Name"
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    lbl(29).Caption = "Status"
    lbl(30).Caption = "Average Cost"
    lbl(32).Caption = "Default  Supplier"
  
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    With Me.CboItemCase
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÃœÌœ"
            .AddItem "„” ⁄„·"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .AddItem "New"
            .AddItem "Used"
        End If

    End With

    With Me.CboItemType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "”·⁄…"
            .AddItem "Œœ„…"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .AddItem "Goods"
            .AddItem "Services"
        End If

    End With

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Items Data"
    Me.EleHeader.Caption = Me.Caption

    'Lbl(0).Caption = "Item Code"
    lbl(1).Caption = "Current Record:"
    lbl(2).Caption = "NO. Recordes:"

    lbl(3).Caption = " Name AR"
    lbl(31).Caption = " Name Eng"

    lbl(4).Caption = "Item Group"
    lbl(5).Caption = "Purchase Price"
    lbl(6).Caption = "Item ID"
    lbl(7).Caption = "Sale Price"
    lbl(8).Caption = "On Demand QTY"
    lbl(9).Caption = "Serial"
    lbl(10).Caption = "Customer Price"
    lbl(11).Caption = "Dealer Price"
    lbl(12).Caption = "Default Guarantee"
    lbl(13).Caption = "Guarantee"
    lbl(14).Caption = "Block"
    ChkAr.Caption = "Is Blocked"
    lbl(15).Caption = "Item Type"
    lbl(16).Caption = "Comments On Item"
    ChkGuar.Caption = "Use Guarantee"

    XPChkSerial.Caption = "Use Serial"
    Ele(4).Caption = "Item Prices"
    Ele(6).Caption = "Item Picture"
    CmdPic(0).Caption = "Add Picture"
    CmdPic(1).Caption = "Delete Picture"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"

    Me.CmdHelp.Caption = "Help"
    Me.C1Tab1.TabCaption(0) = "Item Data"
    Me.C1Tab1.TabCaption(1) = "Other Data "
    
  Me.C1Tab1.TabCaption(2) = " Units Data "
  Me.C1Tab1.TabCaption(3) = "Sales Prices "
     Me.C1Tab1.TabCaption(4) = "Purchase Prices"
    Me.C1Tab1.TabCaption(5) = " Items Details"
 
    
    
    Me.C1Tab1.TabCaption(6) = " Alternatives "
    lbl(0).Caption = "Part No"

    Me.OptGaurType(0).Caption = "Month"
    Me.OptGaurType(1).Caption = "Day"
    ImgPic.ToolTipText = "Double Click to View Maximize"
    '----------------------------------
    Me.ChkAssplied.Caption = "Assblied Item"
    Me.ChkItemMakingNew.Caption = "Product Item"
    Me.lbl(17).Caption = "Price"
    Me.lbl(18).Caption = "Qty"
    Me.lbl(19).Caption = "Item Name"
    Me.lbl(20).Caption = "Item Code"
    Me.lbl(22).Caption = "Items Count"
    Me.Cmd(8).Caption = "Add"
    Me.Cmd(9).Caption = "Del"

    Me.ChkRelated.Caption = "Has Attached Items"
    Me.lbl(26).Caption = "Price"
    Me.lbl(25).Caption = "Qty"
    Me.lbl(24).Caption = "Item Name"
    Me.lbl(23).Caption = "Item Code"
    Me.lbl(27).Caption = "Items Count"

    Me.Cmd(10).Caption = "Add"
    Me.Cmd(11).Caption = "Del"
    lbl(8).Caption = "Risk Qty"
    lblÊÕœ…≈› —«÷Ì…(3).Caption = "Default Unit"
    ChkDef.Caption = "Default Unit"
    lbl«”„«·ÊÕœ…(0).Caption = "Unit name"
    lbl«·⁄·«ﬁ…„⁄(1).Caption = "Relation with other"
    lbl”⁄—«·»Ì⁄(4).Caption = "sale Price"
    lbl”⁄—«·‘—«¡(5).Caption = "Purchase"
    Cmd(20).Caption = "Add"
    Cmd(21).Caption = "Delete"
    Cmd(23).Caption = "save"
    Cmd(22).Caption = "cancel"

    Frame3.Caption = "Sales Prices"

    With FgSalePrice
        .TextMatrix(0, .ColIndex("BranchName")) = "Branch Name  "
 
        .TextMatrix(0, .ColIndex("UnitName")) = " Unit Name  "
    End With

    optBranch(0).Caption = "All  Branches"
    optBranch(1).Caption = " Branch"
    lbl«”„«·ÊÕœ…(3).Caption = "Unit"
    Cmd(14).Caption = "Add"
    Cmd(15).Caption = "Del"

    Frame4.Caption = "Pruchase Price From Vendors"
 
    With FgVendorPrice
        .TextMatrix(0, .ColIndex("Ser")) = "Ser  "
        .TextMatrix(0, .ColIndex("CusName")) = "Vendor Name "
        .TextMatrix(0, .ColIndex("UnitName")) = " Unit Name  "
        .TextMatrix(0, .ColIndex("Price")) = "Price  "
        .TextMatrix(0, .ColIndex("discount")) = "Discount  "
 
    End With

    With FgSum
        .TextMatrix(0, .ColIndex("NumIndex")) = " Index"
        .TextMatrix(0, .ColIndex("Quantity")) = " Quantity"
        '.TextMatrix(0, .ColIndex("UnitName")) = " UnitName "
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"

    End With
    '''//////////
        With fgDiamonds
        .TextMatrix(0, .ColIndex("NumIndex")) = " Index"
        .TextMatrix(0, .ColIndex("type")) = " Type Diamonds"
        .TextMatrix(0, .ColIndex("unite")) = " UnitName "
        .TextMatrix(0, .ColIndex("weight")) = "Weight"
              .TextMatrix(0, .ColIndex("color")) = "Color"
        .TextMatrix(0, .ColIndex("ÛQuality")) = " Quality pieces "
        .TextMatrix(0, .ColIndex("Gestonf")) = "Forms emstones"

    End With
     With fgCameo
        .TextMatrix(0, .ColIndex("NumIndex")) = " Index"
        .TextMatrix(0, .ColIndex("type")) = " Type Cameo"
        .TextMatrix(0, .ColIndex("unite")) = " UnitName "
        .TextMatrix(0, .ColIndex("weight")) = "Weight"

    End With
 '''//////////
 
    With FG1
        .TextMatrix(0, .ColIndex("NumIndex")) = " Index"
        .TextMatrix(0, .ColIndex("Quantity")) = " Quantity"
        .TextMatrix(0, .ColIndex("UnitName")) = " UnitName "
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("x")) = "Expiry Date"

        .TextMatrix(0, .ColIndex("itemsize")) = "size"
        .TextMatrix(0, .ColIndex("ColorName")) = "Color"
        .TextMatrix(0, .ColIndex("ClassName")) = "Class"

    End With
 
End Sub

Private Sub XPTxtCode_KeyPress(KeyAscii As Integer)

    'KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtCode.text, 1)
    If KeyAscii = vbKeySpace Then
        '    KeyAscii = 0
    End If

End Sub

Private Sub XPTxtID_Change()

    Set Rsqty = New ADODB.Recordset
    Rsqty.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rsqty.RecordCount < 1 Then
        FG1.Clear flexClearScrollable, flexClearEverything
        FgSum.Clear flexClearScrollable, flexClearEverything
            GridItemsDetails2.Clear flexClearScrollable, flexClearEverything

      '  Exit Sub
    Else
        RetriveQTY
    End If
            RetriveQTY1 val(XPTxtID.text)
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    'On Error GoTo ErrTrap
        
    StrSQL = "SELECT     ItemSerial, SUM(dbo.Transaction_Details.showqty * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.TblStore.StoreName, dbo.TblUnites.UnitName, "
    StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
 
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FirstPeriodDateInthisYear, True) & ""
    StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Date, True) & ""
    StrSQL = StrSQL + " and Item_ID =" & val(XPTxtID.text)
 
    StrSQL = StrSQL & "  GROUP BY dbo.TblStore.StoreName, dbo.TblUnites.UnitName, dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName,"
    StrSQL = StrSQL & "  dbo.TblItemsColors.ColorName,ItemSerial"
    StrSQL = StrSQL & "  HAVING      (SUM(dbo.Transaction_Details.showqty * dbo.TransactionTypes.StockEffect) <> 0)"
    Build_Sql = StrSQL
    Exit Function
ErrTrap:
End Function

Private Sub XPTxtName_Change()

    If IsNull(DcboItems1.text) = False Then DcboItems1.text = Trim(XPTxtName.text)
End Sub

Private Sub XPTxtName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtNameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub XPTxtPurchase_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtPurchase.text, 0)
End Sub

Private Sub XPTxtSall_KeyPress(KeyAscii As Integer)

    'If KeyAscii = 13 Then
    'If Val(XPTxtSall.text) < Val(XPTxtPurchase.text) Then
    'MsgBox "⁄›Ê« ”⁄— »Ì⁄ «·„” Â·ﬂ «ﬁ· „‰ ”⁄— «·‘—«¡ ", vbOKOnly, App.Title
    'XPTxtSall.SetFocus
    'Exit Sub
    'End If
    'End If
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtSall.text, 0)
End Sub

Private Function RelatedItemTrans() As Boolean
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim IntRes As Integer
    Dim Reports As ClsRepoerts

    Set rs = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "select  count(Transaction_ID)as TransCount,TransactionTypeName "
        StrSQL = StrSQL + " From ("
        StrSQL = StrSQL + " SELECT distinct Transactions.Transaction_ID," & "Transactions.Transaction_Type, TransactionTypes.TransactionTypeName," & "Transactions.Transaction_Serial, Transaction_Details.Item_ID "
        StrSQL = StrSQL + " FROM (TransactionTypes INNER JOIN Transactions ON " & "TransactionTypes.Transaction_Type = Transactions.Transaction_Type) " & "INNER JOIN Transaction_Details ON Transactions.Transaction_ID =" & "Transaction_Details.Transaction_ID) "
        StrSQL = StrSQL + " Where Item_ID =" & Me.XPTxtID.text & ""
        StrSQL = StrSQL + " Group by Transaction_Type,TransactionTypeName"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "select  count(Transaction_ID)as TransCount,TransactionTypeName "
        StrSQL = StrSQL + " From ("
        StrSQL = StrSQL + " SELECT distinct Transactions.Transaction_ID," & "Transactions.Transaction_Type, TransactionTypes.TransactionTypeName," & "Transactions.Transaction_Serial, Transaction_Details.Item_ID "
        StrSQL = StrSQL + " FROM (TransactionTypes INNER JOIN Transactions ON " & "TransactionTypes.Transaction_Type = Transactions.Transaction_Type) " & "INNER JOIN Transaction_Details ON Transactions.Transaction_ID =" & "Transaction_Details.Transaction_ID)As xTable "
        StrSQL = StrSQL + " Where Item_ID =" & Me.XPTxtID.text & ""
        StrSQL = StrSQL + " Group by Transaction_Type,TransactionTypeName"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        RelatedItemTrans = True
        Msg = "⁄›Ê« ·«Ì„ﬂ‰  €Ì— ‰Ÿ«„ «·”Ì—Ì«· «·Œ«’ »«·’‰› "
        Msg = Msg & Chr(13) & "√Ê Õ–› «·’‰› Ê–·ﬂ ·ÊÃÊœ Õ—ﬂ«  ”Ã·  ·Â–« «·’‰›..."
        Msg = Msg & Chr(13) & ""
        Msg = Msg & Chr(13) & "»Ì«‰ «·Õ—ﬂ«  «· Ï ”Ã·  ··’‰›:-"
        Msg = Msg & Chr(13) & ""
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1
            Msg = Msg & Chr(13) & rs("TransactionTypeName").value & vbTab & rs("TransCount").value
            rs.MoveNext
        Next i

        Msg = Msg & Chr(13) & ""
        Msg = Msg & Chr(13) & "Â·  —Ìœ «‰  ‘«Âœ »Ì«‰«  Â–Â «·Õ—ﬂ«  »«· ›’Ì·..øø"
        IntRes = MsgBox(Msg, vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

        If IntRes = vbYes Then
            StrSQL = "select * From ItemsTrans where Item_ID=" & Me.XPTxtID.text & ""
            StrSQL = StrSQL + " order by Transaction_ID"
            Set Reports = New ClsRepoerts
            Reports.TransReport StrSQL
            Set Reports = Nothing
        End If

    Else
        RelatedItemTrans = False
    End If

End Function

Private Sub AddNewFgRow()

    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long

    If val(Me.DcboItems.BoundText) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «Ú”„ «·’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.DcboItems.SetFocus
        Exit Sub
    End If

    If Me.TxtModFlg.text = "E" Then
        If val(Me.DcboItems.BoundText) = val(Me.XPTxtID.text) Then
            Msg = "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·’‰› Ã“¡ „‰ ‰›”Â....!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.DcboItems.SetFocus
            Exit Sub
        End If
    End If

    If val(Me.TxtItemQty(0).text) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ ﬂ„Ì… «·’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtItemQty(0).SetFocus
        Exit Sub
    End If

    If val(Me.TxtItemPrice(0).text) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ  ﬂ·›… «·’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtItemPrice(0).SetFocus
        Exit Sub
    End If

    If val(Me.dcItemunit.BoundText) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ ÊÕœ…  «·’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.dcItemunit.SetFocus
        Exit Sub
    End If

    With Me.Fg
        LngFindRow = .FindRow(val(Me.DcboItems.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)

        If LngFindRow <> -1 Then
            Msg = "Â–« «·’‰› „ÊÃÊœ ›⁄·« ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            .SetFocus
            Exit Sub
        End If

    End With

    LngNewRow = ModFgLib.SetFgForNewRow(Fg, Fg.ColIndex("ItemID"))

    With Me.Fg
        .TextMatrix(LngNewRow, .ColIndex("ItemID")) = Me.DcboItems.BoundText
    
        .TextMatrix(LngNewRow, .ColIndex("ItemCode")) = Trim$(Me.TxtItemCode.text)
        .TextMatrix(LngNewRow, .ColIndex("ItemName")) = Me.DcboItems.BoundText
    
        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = Me.dcItemunit.BoundText
        .TextMatrix(LngNewRow, .ColIndex("UnitName")) = Me.dcItemunit.text
        .TextMatrix(LngNewRow, .ColIndex("ItemQty")) = val(Me.TxtItemQty(0).text)
        .TextMatrix(LngNewRow, .ColIndex("ItemPrice")) = val(Me.TxtItemPrice(0).text)
        .AutoSize 0, .Cols - 1, False
    End With

    Me.lbl(21).Caption = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("ItemID"))

    Me.TxtItemCode.text = ""
    Me.DcboItems.BoundText = ""
    Me.TxtItemQty(0).text = ""
    Me.TxtItemPrice(0).text = ""
    Me.TxtItemCode.SetFocus
End Sub

Private Sub SetMeForNew()
    clear_all Me
    Me.Fg.Rows = Me.Fg.FixedRows
    Me.FgSalePrice.Rows = Me.FgSalePrice.FixedRows
    Me.FgVendorPrice.Rows = Me.FgVendorPrice.FixedRows

    Me.CboItemCase.ListIndex = 0
    Me.CboItemType.ListIndex = 0
End Sub
Private Sub DeleteFgRowAther()

    With Me.VSFlexGrid1

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        .AutoSize 0, .Cols - 1, False
        Me.lbl(21).Caption = ModFgLib.GetItemsInFg(VSFlexGrid1, VSFlexGrid1.ColIndex("ItemID"))
    End With

End Sub

Private Sub DeleteFgRow()

    With Me.Fg

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        .AutoSize 0, .Cols - 1, False
        Me.lbl(21).Caption = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("ItemID"))
    End With

End Sub

Private Sub AddNewFgAttachRow()
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long

    If val(Me.DcboItemID1.BoundText) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «Ú”„ «·’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.DcboItemID1.SetFocus
        Exit Sub
    End If

    If Me.TxtModFlg.text = "E" Then
        If val(Me.DcboItemID1.BoundText) = val(Me.XPTxtID.text) Then
            Msg = "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·’‰› „·Õﬁ ·‰›”Â....!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.DcboItemID1.SetFocus
            Exit Sub
        End If
    End If

    If val(Me.TxtItemQty(1).text) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ ﬂ„Ì… «·’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtItemQty(1).SetFocus
        Exit Sub
    End If

    If val(Me.TxtItemPrice(1).text) = 0 Then
        '    Msg = "ÌÃ»  ÕœÌœ ”⁄— «·’‰› ...!!!"
        '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    Me.TxtItemPrice(1).SetFocus
        '    Exit Sub
    End If

    With Me.FgAttachs
        LngFindRow = .FindRow(val(Me.DcboItemID1.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)

        If LngFindRow <> -1 Then
            Msg = "Â–« «·’‰› „ÊÃÊœ ›⁄·« ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            .SetFocus
            Exit Sub
        End If

    End With

    LngNewRow = ModFgLib.SetFgForNewRow(FgAttachs, FgAttachs.ColIndex("ItemID"))

    With Me.FgAttachs
        .TextMatrix(LngNewRow, .ColIndex("ItemID")) = Me.DcboItemID1.BoundText
        .TextMatrix(LngNewRow, .ColIndex("ItemCode")) = Trim$(Me.TxtAttachedItemCode.text)
        .TextMatrix(LngNewRow, .ColIndex("ItemName")) = Me.DcboItemID1.BoundText
        .TextMatrix(LngNewRow, .ColIndex("ItemQty")) = val(Me.TxtItemQty(1).text)
        .TextMatrix(LngNewRow, .ColIndex("ItemPrice")) = val(Me.TxtItemPrice(1).text)
        .AutoSize 0, .Cols - 1, False
    End With

    Me.lbl(28).Caption = ModFgLib.GetItemsInFg(FgAttachs, FgAttachs.ColIndex("ItemID"))

    Me.TxtAttachedItemCode.text = ""
    Me.DcboItemID1.BoundText = ""
    Me.TxtItemQty(1).text = ""
    Me.TxtItemPrice(1).text = ""
    Me.TxtAttachedItemCode.SetFocus

End Sub

Private Sub DeleteFgAttachRow()

    With Me.FgAttachs

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        .AutoSize 0, .Cols - 1, False
        Me.lbl(28).Caption = ModFgLib.GetItemsInFg(FgAttachs, FgAttachs.ColIndex("ItemID"))
    End With

End Sub

Private Sub TxtUnitFactor_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtUnitFactor.text, 0)
End Sub



