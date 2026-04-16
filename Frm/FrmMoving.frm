VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMoving 
   Caption         =   " ÕÊÌ· „‰ „Œ“‰ ≈·Ï „Œ“‰ "
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   HelpContextID   =   380
   Icon            =   "FrmMoving.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   13935
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
      Height          =   8520
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   13935
      _cx             =   24580
      _cy             =   15028
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   4935
         Left            =   0
         TabIndex        =   72
         Top             =   2520
         Width           =   13935
         _cx             =   24580
         _cy             =   8705
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
         Caption         =   "»Ì«‰«  «·«’‰«›|Õ«·… «·«⁄ „«œ"
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
         Begin C1SizerLibCtl.C1Elastic EleFg 
            Height          =   4560
            Left            =   45
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   45
            Width           =   13845
            _cx             =   24421
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
            GridCols        =   5
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmMoving.frx":038A
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   2925
               Left            =   30
               TabIndex        =   74
               Top             =   1080
               Width           =   13755
               _cx             =   24262
               _cy             =   5159
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
               Cols            =   20
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmMoving.frx":0421
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
               TabIndex        =   75
               Top             =   4020
               Width           =   13275
               _ExtentX        =   23416
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1035
               Index           =   2
               Left            =   30
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   30
               Width           =   13785
               _cx             =   24315
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
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   5505
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   82
                  Top             =   705
                  Width           =   1245
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   2490
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   705
                  Width           =   1050
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   4065
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   705
                  Width           =   1455
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   1260
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   705
                  Width           =   1215
               End
               Begin VB.TextBox TxtShortName 
                  Height          =   270
                  Left            =   1260
                  TabIndex        =   78
                  Top             =   210
                  Width           =   6795
               End
               Begin VB.TextBox TxtItemCodeB1 
                  Alignment       =   1  'Right Justify
                  Height          =   270
                  Left            =   10380
                  TabIndex        =   77
                  Top             =   240
                  Width           =   1695
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   6810
                  TabIndex        =   83
                  Top             =   705
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
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   11340
                  TabIndex        =   84
                  Top             =   705
                  Width           =   2010
                  _ExtentX        =   3545
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   330
                  Left            =   690
                  TabIndex        =   85
                  Top             =   705
                  Width           =   390
                  _ExtentX        =   688
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
                  ButtonImage     =   "FrmMoving.frx":074D
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
               Begin ImpulseButton.ISButton CmdSearch 
                  Height          =   270
                  Left            =   3585
                  TabIndex        =   86
                  Top             =   720
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "..."
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
                  ButtonImage     =   "FrmMoving.frx":0AE7
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   255
                  Index           =   31
                  Left            =   11745
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   570
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   255
                  Index           =   30
                  Left            =   8685
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   540
                  Width           =   1455
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   255
                  Index           =   29
                  Left            =   5655
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   540
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   255
                  Index           =   28
                  Left            =   4290
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   540
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   255
                  Index           =   27
                  Left            =   2700
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   540
                  Width           =   690
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬂ·›…"
                  Height          =   255
                  Index           =   26
                  Left            =   1635
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   540
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»ÕÀ «·”—Ì⁄"
                  Height          =   285
                  Index           =   97
                  Left            =   8490
                  TabIndex        =   88
                  Top             =   210
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»«—ﬂÊœ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Index           =   95
                  Left            =   11955
                  TabIndex        =   87
                  Top             =   180
                  Width           =   1395
               End
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   330
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   3360
               Width           =   435
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   4560
            Left            =   14580
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   45
            Width           =   13845
            _cx             =   24421
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   4095
               Left            =   120
               TabIndex        =   97
               Tag             =   "1"
               Top             =   120
               Width           =   13740
               _cx             =   24236
               _cy             =   7223
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
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmMoving.frx":0E81
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
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   6510
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   4200
               Width           =   3315
            End
            Begin VB.Label Label1100 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   3150
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   4200
               Visible         =   0   'False
               Width           =   3315
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   7500
         Width           =   13890
         _cx             =   24500
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
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3750
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   120
            Width           =   1125
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8235
            TabIndex        =   6
            Top             =   45
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Accredit 
            Height          =   510
            Left            =   5670
            TabIndex        =   43
            Top             =   0
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   900
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··«⁄ „«œ"
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
            Caption         =   "«Ã„«·Ì  "
            Height          =   315
            Index           =   63
            Left            =   4995
            TabIndex        =   39
            Top             =   120
            Width           =   465
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
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   12780
            TabIndex        =   38
            Top             =   0
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   0
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   975
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Left            =   2115
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   105
            Width           =   435
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   135
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   315
            Index           =   1
            Left            =   11160
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   75
            Width           =   1110
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1770
         Index           =   0
         Left            =   15
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   765
         Width           =   13905
         _cx             =   24527
         _cy             =   3122
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
         Begin VB.TextBox txtOrderID 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   0
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox TxtInspectionReport 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   1425
            Width           =   2895
         End
         Begin VB.TextBox TXT_order_no 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   30
            Width           =   1245
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmMoving.frx":0FC4
            Left            =   1410
            List            =   "FrmMoving.frx":0FC6
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   0
            Width           =   1200
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   690
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   55
            Top             =   345
            Width           =   2535
         End
         Begin VB.TextBox TxtPhone 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   7860
            TabIndex        =   52
            Top             =   1380
            Width           =   1020
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   5685
            TabIndex        =   48
            Top             =   1095
            Width           =   705
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   9570
            TabIndex        =   47
            Top             =   1380
            Width           =   2670
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   11280
            TabIndex        =   46
            Top             =   1080
            Width           =   945
         End
         Begin VB.TextBox TxtStoreID1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   690
            Width           =   945
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   345
            Width           =   945
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  ﬁÌœ «·”‰œ"
            Height          =   585
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   1080
            Width           =   2085
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   240
               Width           =   1335
            End
            Begin ImpulseButton.ISButton Cmd 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
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
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -210
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   690
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   0
            Width           =   945
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   2490
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   -405
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   3765
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   -465
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   2850
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   -255
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   2325
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   -255
            Visible         =   0   'False
            Width           =   675
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   7875
            TabIndex        =   2
            Top             =   345
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   300
            Left            =   7905
            TabIndex        =   1
            Top             =   -15
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
            _Version        =   393216
            Format          =   245432321
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DCboSecondStore 
            Height          =   315
            Left            =   7875
            TabIndex        =   3
            Top             =   690
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   3420
            TabIndex        =   25
            Top             =   0
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch1 
            Height          =   315
            Left            =   3435
            TabIndex        =   27
            Top             =   345
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch2 
            Height          =   315
            Left            =   3435
            TabIndex        =   29
            Top             =   690
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   7875
            TabIndex        =   49
            Top             =   1080
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   3450
            TabIndex        =   50
            Top             =   1080
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            Height          =   255
            Index           =   65
            Left            =   6375
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   1425
            Width           =   1170
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»‰«¡ ⁄·Ì"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2730
            TabIndex        =   56
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2685
            TabIndex        =   54
            Top             =   465
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ·Ì›Ê‰"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9000
            TabIndex        =   53
            Top             =   1425
            Width           =   450
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   12510
            TabIndex        =   51
            Top             =   1425
            Width           =   1305
         End
         Begin VB.Label SalesPerson 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰œÊ»"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   6375
            TabIndex        =   42
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label lblCustomer 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„Ì·"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   12630
            TabIndex        =   41
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì »⁄ ›—⁄"
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   6525
            TabIndex        =   30
            Top             =   645
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì »⁄ ›—⁄"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   6360
            TabIndex        =   28
            Top             =   345
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄ „‰›– «·⁄„·ÌÂ"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   6360
            TabIndex        =   26
            Top             =   0
            Width           =   1230
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰ «·„ÕÊ· ≈·ÌÂ"
            Height          =   285
            Index           =   4
            Left            =   12390
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   750
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„Œ“‰ «·„ÕÊ· „‰…"
            Height          =   270
            Index           =   3
            Left            =   12390
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   345
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„” ‰œ"
            Height          =   270
            Index           =   6
            Left            =   10245
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   -15
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„” ‰œ"
            Height          =   240
            Index           =   5
            Left            =   12750
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   60
            Width           =   1065
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   705
         Left            =   15
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   15
         Width           =   13905
         _cx             =   24527
         _cy             =   1244
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
         Caption         =   " ÕÊÌ· „‰ „Œ“‰ ≈·Ï „Œ“‰ "
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
         Begin VB.CheckBox chkIgnorDetails 
            Alignment       =   1  'Right Justify
            Caption         =   " Ã«Â· «· ›«’Ì·"
            Height          =   270
            Left            =   9495
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   240
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.TextBox txtPassword 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   3960
            PasswordChar    =   "*"
            TabIndex        =   103
            Top             =   225
            Width           =   870
         End
         Begin VB.CommandButton cmdReSave 
            Caption         =   "÷»ÿ «·Õ—ﬂ« "
            Height          =   330
            Left            =   8100
            TabIndex        =   102
            Top             =   225
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   255
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   1425
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   405
            Index           =   0
            Left            =   1875
            TabIndex        =   21
            Top             =   135
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "FrmMoving.frx":0FC8
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
            Index           =   3
            Left            =   1005
            TabIndex        =   22
            Top             =   135
            Width           =   795
            _ExtentX        =   1402
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
            ButtonImage     =   "FrmMoving.frx":1362
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
            Left            =   2760
            TabIndex        =   23
            Top             =   135
            Width           =   765
            _ExtentX        =   1349
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
            ButtonImage     =   "FrmMoving.frx":16FC
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
            Left            =   150
            TabIndex        =   24
            Top             =   135
            Width           =   735
            _ExtentX        =   1296
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
            ButtonImage     =   "FrmMoving.frx":1A96
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin MSComCtl2.DTPicker txtFromDateReSave 
            Height          =   360
            Left            =   6510
            TabIndex        =   104
            Top             =   225
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
            _Version        =   393216
            Format          =   245432321
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtToDateReSave 
            Height          =   360
            Left            =   4950
            TabIndex        =   105
            Top             =   225
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
            _Version        =   393216
            Format          =   245432321
            CurrentDate     =   38784
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
            Height          =   495
            Index           =   7
            Left            =   3855
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   135
            Width           =   5565
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   1
         Left            =   15
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   7950
         Width           =   13905
         _cx             =   24527
         _cy             =   979
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
            Height          =   390
            Index           =   0
            Left            =   12765
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   135
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   688
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
            Height          =   375
            Index           =   1
            Left            =   11715
            TabIndex        =   61
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
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
            Height          =   375
            Index           =   2
            Left            =   10440
            TabIndex        =   62
            Top             =   90
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   661
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
            Height          =   375
            Index           =   3
            Left            =   9315
            TabIndex        =   63
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
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
            Height          =   375
            Index           =   4
            Left            =   7890
            TabIndex        =   64
            Top             =   90
            Width           =   1050
            _ExtentX        =   1852
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
            Height          =   375
            Index           =   5
            Left            =   6510
            TabIndex        =   65
            Top             =   90
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   661
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
            Height          =   375
            Index           =   6
            Left            =   1710
            TabIndex        =   66
            Top             =   90
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   661
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
            Height          =   375
            Index           =   7
            Left            =   5445
            TabIndex        =   67
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   661
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
            Height          =   375
            Left            =   2745
            TabIndex        =   68
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   661
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   8
            Left            =   4080
            TabIndex        =   71
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â PL"
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
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   375
            Left            =   120
            TabIndex        =   100
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··«⁄ „«œ"
            BackColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   -2147483635
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
Attribute VB_Name = "FrmMoving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(3) As clsDCboSearch
Public publicSearch As Boolean

Public BolPrint As Boolean
Dim general_noteid As Long
Dim SngTemp As Variant
Dim mIsFinishSave As Boolean
Dim IsSaveWithOutMsg As Boolean
Dim mIsStart As Boolean
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ ”‰œ «· ÕÊÌ·   " & TxtNoteSerial1.text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & " „‰ „Œ“‰ " & DCboStoreName.text & CHR(13) & " «·Ì „Œ“‰  " & DCboSecondStore.text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.text & CHR(13) & "»‰«¡ ⁄·Ï " & CBoBasedON & "»—ﬁ„   " & TXT_order_no & CHR(13) & "  «·„‰œÊ» " & DcboEmp.text & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " doc No " & TxtNoteSerial1.text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " from  " & DCboStoreName.text & CHR(13) & " to " & DCboSecondStore.text & CHR(13) & "customer" & DBCboClientName.text & CHR(13) & "Based On" & CBoBasedON & "No :   " & TXT_order_no & CHR(13) & "Payment Type" & CHR(13) & " GE NO" & TxtNoteSerial
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 190, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 190, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function
Public Sub RetriveSerials(ItemID As String, _
                          ItemName As String, _
                          seriallist As String, _
                          currentrow As Long, Optional Price As Double)
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
  
If val(Price) > 0 Then
            FG.TextMatrix(Num, FG.ColIndex("price")) = Price
        End If
        


        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.rows = FG.rows + 1
 
        Num = Num + 1
    Next
 
    TxtFillData.text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Private Sub C1Elastic6_DblClick()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim AskOption As Boolean
    Dim intDef    As Integer
    Dim Msg       As String
    Dim StrSQL    As String
    Dim RsTest    As ADODB.Recordset
    '  On Error GoTo ErrTrap

    Select Case Index

        Case 0
 
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.rows - 1
            CBoBasedON.ListIndex = 0
            Dim dstore       As Integer
            Dim dBox         As Integer
            Dim usertype     As Integer
            Dim EmpID        As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
            Cmd(2).Enabled = True
            Dim CUSTID As Integer
 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID, , CUSTID
            DBCboClientName.BoundText = CUSTID
            If usertype <> 0 Then 'admin
                DCBranch.Enabled = False
 
                DCboStoreName.Enabled = True
                '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                DCBranch.Enabled = True
                DCboStoreName.Enabled = True
 
                Me.DCBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
                '                TxtStoreID.Enabled = True
            End If

            If SystemOptions.usertype <> UserAdminAll Then
                If checkmanyBranches = False Then
                    Me.DCBranch.Enabled = True
                Else
                    Me.DCBranch.Enabled = True
                End If
                    
                If checkmanyStores = False Then
                    Me.DCboStoreName.Enabled = True
                    Me.DCboSecondStore.Enabled = False
                Else
                    Me.DCboStoreName.Enabled = True
                    Me.DCboSecondStore.Enabled = True
                End If
                                  
            End If
                        
            Me.DCBranch.BoundText = Current_branch

            Me.DCBranch.BoundText = Current_branch

        Case 1
            
            If ChekClodePeriod(XPDtbBill.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
            If ScreenAproved(val(XPTxtBillID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
                  
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
        
            If Check() = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "ﬂÊœ «·”‰œ „ÊÃÊœ „”»ﬁ« Ì—ÃÏ  €ÌÌ—Â"
                Else
                    MsgBox "Please Change Code Becouse It's Already Exists"
                End If
                Exit Sub
            End If
              
            If ChekClodePeriod(XPDtbBill.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
                  
            If Trim(DCBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.DCBranch.BoundText
 
            '      If Me.TxtModFlg.text = "N" Then
             
            '         End If
            Cmd(2).Enabled = False
            SaveData

        Case 3
            Call Undo
            '   Cmd(2).Enabled = True

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

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmMovingSearch.show vbModal

        Case 7
            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmPrintOptions.show vbModal
            End If
         
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport

        Case 6
            Unload Me
        
        Case 10
            ShowGL_cc TxtNoteSerial.text, , 200, val(Me.TXTNoteID.text)
        Case 8
           
            Dim SaleReport As ClsSaleReport
   
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowPrice XPTxtBillID.text, 160, DcboEmp.text
            Exit Sub

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub XPBtnAdd_Click()
    On Error GoTo ErrTrap

    If FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code")) <> "" Then
        FG.rows = FG.rows + 1
        NewGrid.GridDefaultValue FG.rows - 1
        FG.Row = FG.rows - 1
        FG.Col = FG.ColIndex("Code")
        FG.ShowCell FG.rows - 1, FG.ColIndex("Code")
        FG.SetFocus
    End If

    Exit Sub
ErrTrap:
End Sub

'
Private Sub cmdReSave_Click()
    On Error GoTo eh

    Dim s As String
    Dim rsDummy As ADODB.Recordset
    Dim Ids() As Long
    Dim cnt As Long, i As Long

    IsSaveWithOutMsg = True

    '·Ê œÌ » €Ì— rs ⁄«·„Ì/›·« —Ö „„ﬂ‰ ‰” €‰Ï ⁄‰Â« √À‰«¡ resave
    'XPBtnMove_Click 2
    'DoEvents

    Set rsDummy = New ADODB.Recordset

    s = " SELECT Transaction_ID " & _
        " FROM Transactions " & _
        " WHERE (Transaction_Type=10 OR Transaction_Type=992) " & _
        "   AND Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & _
        "   AND Transaction_Date <= " & SQLDate(txtToDateReSave.value, True) & _
        " ORDER BY Transaction_Date, BranchId, Transaction_ID"

    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly

    If rsDummy.EOF Then
        IsSaveWithOutMsg = False
        MsgBox "·«  ÊÃœ Õ—ﬂ«  ›Ì «·› —… «·„Õœœ…"
        Exit Sub
    End If

    '1) «Ã„⁄ «·‹ IDs
    rsDummy.MoveFirst
    Do While Not rsDummy.EOF
        cnt = cnt + 1
        ReDim Preserve Ids(1 To cnt)
        Ids(cnt) = CLng(val(rsDummy!Transaction_ID & ""))
        rsDummy.MoveNext
    Loop
    rsDummy.Close

    '2) ‰›¯– Ê«Õœ Ê«Õœ
    For i = 1 To cnt
        Resave_Transfer Ids(i)
        DoEvents
    Next i

    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"
    Exit Sub

eh:
    IsSaveWithOutMsg = False
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
End Sub

Private Sub Resave_Transfer(ByVal TransID As Long)
    On Error GoTo eh

    '«› Õ «·Õ—ﬂ… ’Õ (Â‰« ·«“„ Retrive  › Õ rs Õ Ï √À‰«¡ resave)
    Me.TxtModFlg.text = "R"
    Me.Retrive TransID, True   'True = ForceOpen

    '«œŒ· Edit
    Me.TxtModFlg.text = "E"

    '„Â„ Ãœ«: ·Ê «‰  „‘ ⁄«Ì“  €Ì— «· «—ÌŒ ›⁄·Ì«° »·«‘  ‰«œÌ Change
    NewGrid.DtpBillDate_Change
'NewGrid.Calculate 1
    '«Õ›Ÿ »’„  + „‰ resave
    SaveData True, True

    Exit Sub
eh:
    Debug.Print "Resave_Transfer failed TransID=" & TransID & " Err=" & Err.Number & " " & Err.Description
    Err.Clear
End Sub



Private Sub cmdReSave_ClickOld()
    Dim s       As String
    Dim rsDummy As ADODB.Recordset
    IsSaveWithOutMsg = True
    XPBtnMove_Click (2)
    DoEvents
    
    XPBtnMove_Click (1)
    DoEvents
    Set rsDummy = New ADODB.Recordset
    
    s = " SELECT * FROM Transactions WHERE (Transaction_Type=10 or Transaction_Type=992)"
    s = s & "   and ( Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & " and "
    s = s & "   Transaction_Date <=   " & SQLDate(txtToDateReSave.value, True) & " )"

    'If chkWithoutCost.value = vbChecked Then
        
'    s = s & "  and Transaction_ID in "
'    s = s & "  (SELECT        TT.Transaction_ID"
'    s = s & "   FROM            dbo.Transactions TT INNER JOIN"
'    s = s & "                 dbo.Transaction_Details ON TT.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    s = s & " Where (TT.Transaction_Type = 10 Or  Transaction_Type=992))"
    'Transaction_Details.Price > 3000) "
    'End If
        
    s = s & " ORDER BY  Transaction_Date, BranchId, Transaction_ID"
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rsDummy.EOF
        On Error GoTo NextRow
        
        mIsFinishSave = False
        mIsStart = True
        Me.TxtModFlg.text = "R"
        Me.Retrive val(rsDummy!Transaction_ID & "")
       
        DoEvents
11:
        DoEvents
        If mIsFinishSave And mIsStart Then
            IsSaveWithOutMsg = True
            Me.TxtModFlg.text = "E"
            
            NewGrid.DtpBillDate_Change
            DoEvents
            'Cmd_Click (1)
            TxtModFlg.text = "E"
            DoEvents
            DoEvents
            DoEvents
    
            SaveData True, True
            mIsStart = False
        Else
            GoTo 11
        End If
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        
        DoEvents
                 
NextRow:
        rsDummy.MoveNext
        
    Loop
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"
End Sub

Private Sub DBCboClientName_Change()
   On Error Resume Next
    Dim fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 1
    TxtSearchCode.text = fullcode

        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            Dim DefaultSalesPersonId As Integer
            GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId

            If Not DefaultSalesPersonId = 0 Then

                Me.DcboEmp.BoundText = DefaultSalesPersonId
            End If
         End If
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 15
        FrmCustemerSearch.show vbModal
    End If
End Sub

Function GetReceQty(Optional Transaction_ID As Double, Optional Item_ID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(dbo.Transaction_Details.ShowQty) AS SumShowQty"
sql = sql & " FROM         dbo.Transaction_Details INNER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
sql = sql & " WHERE     (dbo.Transactions.Transaction_Type = 10) AND (dbo.Transactions.BillBasedOn = 2) AND (dbo.Transactions.order_no = '" & TXT_order_no.text & "') AND"
sql = sql & "                      (dbo.Transactions.Transaction_ID <> " & Transaction_ID & ") and (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetReceQty = IIf(IsNull(rs2("SumShowQty").value), 0, rs2("SumShowQty").value)
Else
GetReceQty = 0
End If
End Function
Public Sub RetriveOrder(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim rsDummy As New ADODB.Recordset
    Dim mCostPrice As Double
    Dim s As String
    Dim mQtyE As Double
    Dim Num As Long
   On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

   
        StrSQL = "Select * from transactions where  Transaction_Type=" & Transaction_Type & " and noteserial1='" & order_no & "' "
 
        
            If Transaction_Type = 38 Then
              
              
              If SystemOptions.IsInternalMultiOrder Then

                    
                    If CheckAprroveScreen("FrmPO6") = True Then
                        StrSQL = StrSQL + " and approved=1"
                    End If
                    StrSQL = StrSQL & " and Transactions.Transaction_ID NOT IN ("
    
                    StrSQL = StrSQL & " SELECT"
      
                    StrSQL = StrSQL & " OrderID"
       
                    StrSQL = StrSQL & " FROM   Transaction_Details AS td"
                    StrSQL = StrSQL & " INNER JOIN Transactions t"
                    StrSQL = StrSQL & " ON  t.Transaction_ID = td.Transaction_ID"
                    StrSQL = StrSQL & " AND t.Transaction_Type = 10"
                    StrSQL = StrSQL & " AND t.BillBasedOn = 1"
                    StrSQL = StrSQL & " INNER JOIN Transaction_Details tt"
                    StrSQL = StrSQL & " ON  tt.Transaction_ID = t.OrderID"
                    StrSQL = StrSQL & " AND tt.Item_ID = td.Item_ID"
                    StrSQL = StrSQL & " AND tt.UnitId = td.UnitId"
                    StrSQL = StrSQL & " Group By"
                    StrSQL = StrSQL & " td.Item_ID,"
                    StrSQL = StrSQL & " td.UnitId,tt.ShowQty,"
                    StrSQL = StrSQL & " OrderID"
                    StrSQL = StrSQL & " Having SUM(td.ShowQty) >= (tt.ShowQty)"

                    StrSQL = StrSQL & " )"
            End If
            End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
       ' Me.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
      Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid1").value), "", rs("storeid1").value)
    If Transaction_Type = 22 Then
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
    Else
        Me.DCboSecondStore.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
    End If
        Me.DCBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)

        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    StrSQL = "SELECT TblItems.HaveSerial, * "
    If SystemOptions.IsInternalMultiOrder Then
        StrSQL = StrSQL & "  ,ShowQty-  IsNull(("
        StrSQL = StrSQL & "             SELECT SUM(ShowQty)"
        StrSQL = StrSQL & "             FROM   Transaction_Details AS td"
        StrSQL = StrSQL & "                    INNER JOIN Transactions AS t"
        StrSQL = StrSQL & "                         ON  t.Transaction_ID = td.Transaction_ID"
        StrSQL = StrSQL & "                         AND t.Transaction_Type = 10"
        StrSQL = StrSQL & "                         AND t.OrderID =" & val(rs("Transaction_ID").value)
        StrSQL = StrSQL & "                                    Where td.Item_ID = Transaction_Details.Item_ID"
        StrSQL = StrSQL & "                                           AND td.UnitId = Transaction_Details.UnitId"
        StrSQL = StrSQL & "                               ),0) AS ShoQty"
    End If
      
    StrSQL = StrSQL & "                             FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1
     
        For Num = 1 To RsDetails.RecordCount
        
             If Num = 15 Then
            mCostPrice = 0
        End If
        Dim movingqty As Double
        Dim actulaqty As Double
        Dim RsTest As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngColorID As Long
    Dim StrItemSize As String
    Dim LngClassId As Long
        txtOrderID = rs!Transaction_ID & ""
         LngItemID = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
         
              LngColorID = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            StrItemSize = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
          LngClassId = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
           If SystemOptions.IsInternalMultiOrder Then
            mQtyE = val(RsDetails!ShoQty & "")
        Else
            mQtyE = val(RsDetails!ShowQty & "")
        End If
        
              Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.DCboSecondStore.BoundText)), XPDtbBill.value, , , , , True, LngColorID, StrItemSize, LngClassId)
              If RsTest.EOF Or RsTest.BOF Then
                   actulaqty = 0
              Else
           
                  actulaqty = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
              
              End If
              
              If SystemOptions.poWithatotalQty = False Then ' Õ«·Â «·ﬂ„Ì… ﬂ«„·…
                         movingqty = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value)) - actulaqty
              
              Else
                         movingqty = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
              End If
              If Transaction_Type = 42 Then
                    movingqty = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
              End If
              
             If val(IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value)) = 0) Then GoTo skiploop
             
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            If Transaction_Type = 38 Then
                FG.TextMatrix(Num, FG.ColIndex("Count")) = mQtyE
            Else
                FG.TextMatrix(Num, FG.ColIndex("Count")) = movingqty ' IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value)) - GetReceQty(val(XPTxtBillID.Text), val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
            End If
            
            
            FG.TextMatrix(Num, FG.ColIndex("OriginalQty")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
         '   If Transaction_Type = 0 Then
                'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
         '   End If
      
       
    
       
       
       
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            
            FG.TextMatrix(Num, FG.ColIndex("ProfitType")) = IIf(IsNull(RsDetails("ProfitType")), "", (RsDetails("ProfitType").value))
            FG.TextMatrix(Num, FG.ColIndex("ProfitValue")) = IIf(IsNull(RsDetails("ProfitValue")), "", (RsDetails("ProfitValue").value))
            FG.TextMatrix(Num, FG.ColIndex("NetProfit")) = IIf(IsNull(RsDetails("NetProfit")), "", (RsDetails("NetProfit").value))
                            
            
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            If Transaction_Type = 42 Then
                s = "SELECT T2.* "
                s = s & " from  Transactions AS t "
                s = s & " Inner Join Transaction_Details T2 On T2.Transaction_ID = t.Transaction_ID"
                s = s & " WHERE t.Transaction_Type = 26 and t.OrderID =  " & val(txtOrderID)
                s = s & " and  T2.Item_ID = " & val(RsDetails("Item_ID").value & "")
                s = s & " and T2.UnitId= " & val(RsDetails("UnitId").value & "")
                Set rsDummy = New ADODB.Recordset
                
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rsDummy.EOF Then
                    mCostPrice = 0
                Else
                    mCostPrice = val(rsDummy!ShowPrice & "")
                End If
                           
            End If

            If mCostPrice <> 0 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = mCostPrice
            Else
                FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(Num, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.cell(flexcpData, Num, FG.ColIndex("UnitID"))))
            End If
            
   
skiploop:
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub


Private Sub Dcbranch_Change()
  If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
TxtNoteSerial1.text = ""
 TxtNoteSerial.text = ""
End If

End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    fill_bill_items_table

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 190

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////











End Sub

Private Sub ISButton1_Click()
    Dim BeginTrans As Boolean
If val(XPTxtBillID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
 
    SendTopost Me.Name, "Transactions", "Transaction_ID", 0, val(DCBranch.BoundText), val(XPTxtBillID.text), TxtNoteSerial1.text
  rs.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    ISButton1.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
    ISButton1.Caption = "Sent To Approval "
End If
    Retrive (val(Me.XPTxtBillID.text))
End Sub
Function fillapprovData()
 Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtBillID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"
RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    ISButton1.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
    ISButton1.Caption = "Sent To approval "
End If
ISButton1.Enabled = False
Else
ISButton1.Enabled = True
 If SystemOptions.UserInterface = ArabicInterface Then
    ISButton1.Caption = " «·«—”«· ··«⁄ „«œ"
Else
    ISButton1.Caption = "Sent To approval "
End If
End If
 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
        For Num = 1 To RsDetails.RecordCount
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
       Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
       Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label24.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label24.Caption = "Approved"
                                 End If
                            Label24.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label24.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label24.Caption = "Currently required Approve"
                            End If
                 Label24.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.rows = 1
    End If
RsDetails.Close

End Function

Private Sub Txt_order_no_Change()
    Dim Transaction_Type As Integer
    If CBoBasedON.ListIndex = 1 Then
              Transaction_Type = 38
    ElseIf CBoBasedON.ListIndex = 2 Then
              Transaction_Type = 22
    ElseIf CBoBasedON.ListIndex = 3 Then
            RetriveOrderDef Me.TXT_order_no, 0
        Exit Sub

    ElseIf CBoBasedON.ListIndex = 4 Then
            Transaction_Type = 42
        

    Else
     
         Exit Sub
    End If

   ' Transaction_ID = get_transactionData("order_no", Txt_order_no.text, "Transaction_ID", Transaction_Type)
'


    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TXT_order_no, Transaction_Type
    End If




End Sub


Public Sub RetriveOrderDef(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    Dim StoreId2 As Double
    Dim issuedQty As Double
    Dim rsDummy As New ADODB.Recordset
    
        Dim actulaqty As Double
        Dim RsTest As ADODB.Recordset
        Dim LngItemID As Long
        Dim movingqty As Double
    Dim mCostPrice As Double
    Dim s As String
   On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    txtOrderID = order_no
      
             StrSQL = "SELECT  TblDefComItem.StoreID2 ,TblDefComItem.StoreID3, TblDefComItem.order_no,TblDefComItem.OrderID, TblDefComItem.CusID,TblDefComItem.storeid,"
    StrSQL = StrSQL + " TblDefComItem.Branchid, dbo.TblDefComItemData.*,"
    StrSQL = StrSQL + " TblItems.*,TblUnites.UnitName,TblUnites.UnitNamee "
    StrSQL = StrSQL + " From TblDefComItemData"
    StrSQL = StrSQL + " Left Outer join TblDefComItem On TblDefComItem.ID = TblDefComItemData.IDDefCIT"
    StrSQL = StrSQL + " Left Outer Join TblItems On TblItems.ItemID = TblDefComItemData.ItemID"
    StrSQL = StrSQL + " Left Outer Join TblUnites On TblUnites.UnitID = TblDefComItemData.UnitId"
    StrSQL = StrSQL + " Where TblDefComItem.Id = " & val(txtOrderID)
    
    StrSQL = StrSQL & " order by TblDefComItemData.id "

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    

    If RsDetails.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(RsDetails("CusID").value), "", RsDetails("CusID").value)
       ' Me.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        
If Transaction_Type = 21 And SystemOptions.MultyStore = True Then

Else
Me.DCboStoreName.BoundText = IIf(IsNull(RsDetails("storeid").value), "", RsDetails("storeid").value)
Me.DCboSecondStore.BoundText = IIf(IsNull(RsDetails("StoreID2").value), "", RsDetails("StoreID2").value)
End If

        Me.DCBranch.BoundText = IIf(IsNull(RsDetails("Branchid").value), "", RsDetails("Branchid").value)
       ' TxtOldOpOrderID.Text = IIf(IsNull(rs("OldOpOrderID").value), "", rs("OldOpOrderID").value)
       ' TxtCashCustomerName.Text = IIf(IsNull(rs("CashCustomerName").value), "", rs("CashCustomerName").value)
        
        
        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If RsDetails.EOF Or RsDetails.BOF Then
        Exit Sub
    End If
  '  txtOrderID = rs!Transaction_ID & ""


    Screen.MousePointer = vbArrowHourglass
  


    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            StoreId2 = val(DCboStoreName.BoundText)
            LngItemID = IIf(IsNull(RsDetails("ItemID")), 0, (RsDetails("ItemID").value))
            
                  Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.DCboSecondStore.BoundText)), XPDtbBill.value, , , , , True)
              If RsTest.EOF Or RsTest.BOF Then
                   actulaqty = 0
              Else
           
                  actulaqty = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
              
              End If
              
              If SystemOptions.poWithatotalQty = False Then ' Õ«·Â «·ﬂ„Ì… ﬂ«„·…
                         movingqty = IIf(IsNull(RsDetails("Qty")), 0, (RsDetails("Qty").value)) - actulaqty
              
              Else
                         movingqty = IIf(IsNull(RsDetails("Qty")), 0, (RsDetails("Qty").value))
              End If
              
             If movingqty = 0 Then GoTo skiploop
             

            FG.TextMatrix(Num, FG.ColIndex("Count")) = movingqty ' IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value)) - GetReceQty(val(XPTxtBillID.Text), val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
            
            FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("hight")), "", (RsDetails("hight").value))
            FG.TextMatrix(Num, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))

            
            FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("widtj")), "", (RsDetails("widtj").value))
            

            FG.TextMatrix(Num, FG.ColIndex("OriginalQty")) = IIf(IsNull(RsDetails("qty")), "", (RsDetails("qty").value))

            
         '   FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Qty")), "", (RsDetails("Qty").value))
            
 
        issuedQty = GetIssuedQty(TXT_order_no, val(Me.XPTxtBillID), StoreId2, val(FG.TextMatrix(Num, FG.ColIndex("Code"))))


            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
         '   If Transaction_Type = 0 Then
                'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
         '   End If
      
       
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = 0 'IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("TotalDisc")), "", (RsDetails("TotalDisc").value))
            'FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            'FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            'FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
             
            If val(RsDetails!OrderID & "") <> 0 And SystemOptions.CostByProduction Then
                s = "SELECT T2.* "
                s = s & " from  Transactions AS t "
                s = s & " Inner Join Transaction_Details T2 On T2.Transaction_ID = t.Transaction_ID"
                s = s & " WHERE t.Transaction_Type = 26 and t.OrderID =  " & val(RsDetails!OrderID & "")
                s = s & " and  T2.Item_ID = " & val(RsDetails("ItemID").value & "")
                s = s & " and T2.UnitId= " & val(RsDetails("UnitId").value & "")
                Set rsDummy = New ADODB.Recordset
                
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rsDummy.EOF Then
                    mCostPrice = 0
                Else
                    mCostPrice = val(rsDummy!ShowPrice & "")
                End If
                           
            End If

            If mCostPrice <> 0 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = mCostPrice
            Else
                FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(Num, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.cell(flexcpData, Num, FG.ColIndex("UnitID"))), val(Me.DCboStoreName.BoundText))
            End If
          '  FG.TextMatrix(Num, FG.ColIndex("SalesPrice")) = GetItemPrice(val(FG.TextMatrix(Num, FG.ColIndex("Code"))), 0, val(FG.Cell(flexcpData, Num, FG.ColIndex("UnitID"))))
          '  FG.TextMatrix(Num, FG.ColIndex("TotalSalesPrice")) = val(FG.TextMatrix(Num, FG.ColIndex("SalesPrice"))) * val(FG.TextMatrix(Num, FG.ColIndex("Count")))
skiploop:
            RsDetails.MoveNext
            Debug.Print Num
      
            

         

        Next Num

    End If

    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
'    XPTxtCurrent.Caption = rs.AbsolutePosition
'    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub


Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, Shift As Integer)
 If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
    If KeyCode = vbKeyF3 Then
    
    
        If CBoBasedON.ListIndex = 1 Then
        FrmBuySearch.DealingForm = GridTransType.internalorder
            FrmBuySearch.Index = 2
           FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…"
        ElseIf CBoBasedON.ListIndex = 2 Then
            FrmBuySearch.DealingForm = GridTransType.PurchaseTransaction
            FrmBuySearch.Index = 2
           FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ «·„‘ —Ì« "
        End If
            FrmBuySearch.show vbModal
    End If
End If
End Sub

Sub SerchItems(Optional str As String)
 
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
sql = " select  ItemID,barCodeNO   from  dbo.TblItems where 1=1"
If SystemOptions.UserInterface = ArabicInterface Then
SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
Else
SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
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
                                 sql = " select  ItemID,barCodeNO   from  dbo.TblItems where 1=1"
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                 SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
                                     SQL1 = SQL1 + " Order BY ItemName "
                                 Else
                                 SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
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
sql = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
Else
sql = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
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

Private Sub txtPassword_Change()

    If Trim(txtPassword) = "Alex2025" Then
        cmdReSave.Visible = True
        txtFromDateReSave.Visible = True
        txtToDateReSave.Visible = True
        chkIgnorDetails.Visible = True
        chkIgnorDetails.value = 1
    Else
        cmdReSave.Visible = False
        txtFromDateReSave.Visible = False
        txtToDateReSave.Visible = False
        chkIgnorDetails.Visible = False
    
    End If
End Sub

Private Sub TxtShortName_KeyDown(KeyCode As Integer, Shift As Integer)
'   LoadSpecificItems
SerchItems (TxtShortName.text)
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

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub
Private Sub DBCboClientName_Click(Area As Integer)
DBCboClientName_Change
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 10
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboSecondStore_Change()
    On Error Resume Next
 TxtStoreID1.text = getStoreCoding(val(DCboSecondStore.BoundText))
 
    If val(DCboSecondStore.BoundText) <> 0 Then
        Dcbranch2.BoundText = GetInventoryBranch(DCboSecondStore.BoundText)
    End If

End Sub

Private Sub DCboSecondStore_Click(Area As Integer)
    DCboSecondStore_Change
End Sub

Private Sub DCboSecondStore_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboSecondStore, , , 1
    End If

End Sub

Private Sub DCboStoreName_Change()
    On Error Resume Next
 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If val(DCboStoreName.BoundText) <> 0 Then
        dcBranch1.BoundText = GetInventoryBranch(DCboStoreName.BoundText)
    End If


    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(DCBranch.BoundText), 12) = True Then
     TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""

     End If
     
    End If


End Sub

Private Sub DCboStoreName_Click(Area As Integer)
    DCboStoreName_Change
End Sub

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName, , , 1
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches DCBranch
    End If

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1
    End If

End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreID
    End If
End Sub

Private Sub TxtStoreID1_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID1)
        Me.DCboSecondStore.BoundText = StoreID
    End If
End Sub

Public Sub XPBtnMove_Click(Index As Integer)
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
        Me.TxtModFlg.text = ""
        Dim StrSQL As String
             Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
         
    StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=10 or Transaction_Type=992)"
  StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"


       If IsSaveWithOutMsg Then
            StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=10 or Transaction_Type=992)"
            StrSQL = StrSQL & "   and ( Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & " and "
            StrSQL = StrSQL & "   Transaction_Date <=   " & SQLDate(txtToDateReSave.value, True) & " )"
    
'            StrSQL = StrSQL & "  and Transaction_ID in "
'            StrSQL = StrSQL & "  (SELECT        TT.Transaction_ID"
'            StrSQL = StrSQL & "   FROM            dbo.Transactions TT INNER JOIN"
'            StrSQL = StrSQL & "                 dbo.Transaction_Details ON TT.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'            StrSQL = StrSQL & " Where (TT.Transaction_Type = 10 Or  Transaction_Type=992) )"
        'Transaction_Details.Price > 3000) "
        'End If
        End If
    
        
        
            
            
            
    
       If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  UserID = " & user_id
  
             End If
  
        Me.DCBranch.Enabled = True
      
      
    End If
    If IsSaveWithOutMsg Then
            
        StrSQL = StrSQL & " Order by Transaction_Date Desc"
    Else
     StrSQL = StrSQL & " Order by Transaction_ID"
    End If
  Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If
Me.TxtModFlg.text = "R"
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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            '       Cmd_Click (0)
        Else
            '       SendKeys "{TAB}"
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
            XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            '   XPBtnRemove_Click
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

Private Sub ChangeLang()
    'CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(95).Caption = "Barcode"
    Me.Caption = "Transfer items between Stores"
    lbl(97).Caption = "Smart Search"
    C1Elastic6.Caption = Me.Caption
    Label3.Caption = "Branch"
    lblCustomer.Caption = "Customer "
    Label4.Caption = "Cash Customer Name"
    Label5.Caption = "Tel"
    lbl(65).Caption = "Check No"
    SalesPerson.Caption = "Sales Person"

    lbl(63).Caption = "Totals"
    Accredit.Caption = "Send To Approve"

Label7.Caption = "Based On"
Label6.Caption = "Remarks"
    lbl(5).Caption = "Voucher ID"
    lbl(6).Caption = "Date"
    lbl(3).Caption = "From Store"
    lbl(4).Caption = "To Store"
    Label1.Caption = "From Branch "
    Label2.Caption = "To Branch "
    Frame3.Caption = "GE Data"
    Cmd(10).Caption = "Print GE"
    lbl(1).Caption = " By:"
    lbl(0).Caption = "Curr rec."
    lbl(2).Caption = "Rec.count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = " Case"
    lbl(28).Caption = " Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Cmd(8).Caption = "Print PL"
    Me.CmdHelp.Caption = "Help"
    With Me.FG
         .TextMatrix(0, .ColIndex("OriginalQty")) = "Original Qty"
    End With
    With Grid2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
    End With
        
       With CBoBasedON
    .Clear
        .AddItem "Na"
     .AddItem "Internal request"
    .AddItem "Purchase Order"
    .AddItem "Assemply"
    .AddItem "Quot"
    
    End With
        Label1100.Caption = "Approval Requested by "
        Label24.Caption = "Approval Requested by "
        ISButton1.Caption = "Send for Approval"
        C1Tab1.TabCaption(0) = "Items"
        'C1Tab1.TabCaption(0) = "Items"
        C1Tab1.TabCaption(1) = "Approval Status"
End Sub
Function Check() As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "SELECT * FROM Transactions WHERE Transaction_Type=10 and NoteSerial1 ='" & TxtNoteSerial1.text & "' and Transaction_ID <> " & val(XPTxtBillID.text) & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Check = True
Else
Check = False
End If
End Function
Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
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
Dim ScreenNameArabic As String
Dim ScreenNameEnglish As String
  ScreenNameArabic = "    ”‰œ  ÕÊÌ· "
    ScreenNameEnglish = " Moving Items    "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 190

If ChekSanNumber(Current_branch, 12) = True Then
TxtNoteSerial1.Enabled = False
Else
TxtNoteSerial1.Enabled = True
End If
chkIgnorDetails.value = vbUnchecked
    Dim My_SQL As String
   ' My_SQL = "  select branch_id,branch_name from TblBranchesData   "
   ' fill_combo dcBranch, My_SQL
  
  
'    My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    
'    fill_combo dcBranch1, My_SQL
  
'    My_SQL = "  select branch_id,branch_name from TblBranchesData   "
'    fill_combo dcBranch2, My_SQL


    If SystemOptions.usertype <> UserAdminAll Then
        Me.DCBranch.Enabled = True
    End If



 
    Set NewGrid.Grid = FG
    NewGrid.GridTrans = MoveItems
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.GrdTBar = Me.TBar
        Set NewGrid.DtpBillDate = Me.XPDtbBill
        
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
        Set NewGrid.StoreName = DCboStoreName
          Set NewGrid.TxtItemCodeB1 = TxtItemCodeB1


    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.CmdAddSerialLIst = Me.CmdSearch
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.txtTotal = XPTxtSum
    Resize_Form Me, TransactionSize
    With CBoBasedON
    .Clear
        .AddItem "»·«"
     .AddItem "ÿ·» œ«Œ·Ì "
    .AddItem "›« Ê—… „‘ —Ì« "
    .AddItem "”‰œ  Ã„Ì⁄"
    .AddItem "⁄—÷ ”⁄—"
    
    End With
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FG.WallPaper = BGround.Picture
    AddTip
    XPDtbBill.value = Date
'    StrSQL = "SELECT * From TblStore"
'    fill_combo Me.DCboStoreName, StrSQL
'    fill_combo Me.DCboSecondStore, StrSQL
    
    StrSQL = "SELECT * From TblUsers"
    fill_combo DCboUserName, StrSQL
      Dim Dcombos As New ClsDataCombos
      
     Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
Dcombos.GetStores DCboStoreName, , , 1
Dcombos.GetStores DCboSecondStore, , , 1
'Dcombos.GetStores

Dcombos.GetBranches DCBranch
Dcombos.GetBranches dcBranch1
Dcombos.GetBranches Dcbranch2
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
    
    
            Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
         
         
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=-10"
       StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
 
       
       If SystemOptions.usertype <> UserAdminAll Then
 
                    If SystemOptions.FixedCustomer = 1 Then
                      StrSQL = StrSQL & " and  UserID = " & user_id
                     ' StrSQL = StrSQL & " or  storeId = " & dstore
                      
                       End If
                    Me.DCBranch.Enabled = True
      
    End If
    
     StrSQL = StrSQL & " Order by Transaction_ID"
     
    'StrSql = "SELECT * FROM QryMovingItems"
    'StrSql = "SELECT Transactions.*, Transactions_1.StoreID FROM TblStore " & _
    '"INNER JOIN (TblStore AS TblStore_1 INNER JOIN (Transactions INNER JOIN Transactions " & _
    '"AS Transactions_1 ON Transactions.Transaction_ID=Transactions_1.ReturnID) ON " & _
    '"TblStore_1.StoreID=Transactions_1.StoreID) ON TblStore.StoreID=Transactions.StoreID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    NewGrid.FillGrid
 If SystemOptions.HideCost = True Then
XPTxtSum.Visible = False
 TxtPrice.Visible = False
       FG.ColHidden(FG.ColIndex("Price")) = True
       FG.ColHidden(FG.ColIndex("Valu")) = True

 End If
   ' XPBtnMove_Click 2
   Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer

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

RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 1


    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
    Dim RsTest As ADODB.Recordset
    Dim StrSQL As String

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = " ÕÊÌ· «·»÷«⁄… „‰ «·„Œ“‰"
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
       
            Me.XPDtbBill.Enabled = False
            Me.DCboStoreName.locked = True
            DCboSecondStore.locked = True
            FG.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
            '    Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If

            Ele(2).Enabled = False

        Case "N"
            '        Me.Caption = " ÕÊÌ· «·»÷«⁄… „‰ «·„Œ“‰( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
        
            FG.Enabled = True
            FG.rows = 2
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DCboStoreName.locked = False
            DCboSecondStore.locked = False
            FG.Editable = flexEDKbdMouse
            Ele(2).Enabled = True
            CboItemCase.ListIndex = 0

        Case "E"
            '        Me.Caption = " ÕÊÌ· «·»÷«⁄… „‰ «·„Œ“‰(  ⁄œÌ· )"
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
        
            FG.Enabled = True
            Me.XPDtbBill.Enabled = True
            Me.DCboStoreName.locked = False
            DCboSecondStore.locked = False
            FG.Editable = flexEDKbdMouse
            Ele(2).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, Optional ByVal ForceOpen As Boolean = False)


    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim Num As Long
     On Error GoTo ErrTrap
If Lngid = 0 Then
Else


If Lngid <> 0 Then
    If ForceOpen Or (Not IsSaveWithOutMsg) Then
        StrSQL = "SELECT * FROM Transactions " & _
                 " WHERE (Transaction_Type=10 OR Transaction_Type=992) " & _
                 "   AND Transaction_ID=" & Lngid & _
                 "   AND BranchId in(" & Current_branchSql & ")"

        '«ﬁ›· rs «·ﬁœÌ„ (⁄·‘«‰ „«Ì»ﬁ«‘ „«”ﬂ ‘«‘…  «‰Ì…)
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then rs.Close
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    End If
End If

If Lngid <> 0 Then
    If rs.EOF Or rs.BOF Then Exit Sub
    If CLng(val(rs("Transaction_ID").value & "")) <> CLng(Lngid) Then
        'Õ„«Ì…: ·Ê ·√Ì ”»» rs „‘ ⁄·Ï ‰›” «·”Ã·
        Exit Sub
    End If
End If


            If Not (rs.EOF Or rs.BOF) Then
               ' If Lngid = 0 Then
                rs.MoveLast
          '      End If
            End If
Me.TxtModFlg.text = "R"
End If
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
'Me.TxtModFlg.text = "R"

    Screen.MousePointer = vbArrowHourglass
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
    txtOrderID.text = IIf(IsNull(rs.Fields("OrderID").value), 0, rs.Fields("OrderID").value)
    lbl(7).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    DCBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    
    TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
TxtInspectionReport.text = IIf(IsNull(rs("InspectionReport").value), "", (rs("InspectionReport").value))

CBoBasedON.ListIndex = IIf(IsNull(rs("BillBasedOn").value), 0, rs("BillBasedOn").value)
 
TXT_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
 

Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.text = ""
    End If
    
    TxtPhone.text = IIf(IsNull(rs("Phone").value), "", (rs("Phone").value))
    TxtFillData.text = "T"
    Set RsTemp = New ADODB.Recordset
    StrSQL = "select * From Transactions where ReturnID = " & val(rs("Transaction_ID").value)
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Me.DCboSecondStore.BoundText = IIf(IsNull(RsTemp("StoreID").value), "", RsTemp("StoreID").value)
    End If

    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)

            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
            FG.TextMatrix(Num, FG.ColIndex("OriginalQty")) = IIf(IsNull(RsDetails("OriginalQty")), "", (RsDetails("OriginalQty").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
        
        
            FG.TextMatrix(Num, FG.ColIndex("ProfitType")) = IIf(IsNull(RsDetails("ProfitType")), "", (RsDetails("ProfitType").value))
            FG.TextMatrix(Num, FG.ColIndex("ProfitValue")) = IIf(IsNull(RsDetails("ProfitValue")), "", (RsDetails("ProfitValue").value))
            FG.TextMatrix(Num, FG.ColIndex("NetProfit")) = IIf(IsNull(RsDetails("NetProfit")), "", (RsDetails("NetProfit").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassId")) = IIf(IsNull(RsDetails("ClassId")), 1, (RsDetails("ClassId").value))
             FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
             FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
             FG.cell(flexcpData, Num, FG.ColIndex("IsExpirDate")) = IIf(IsNull(RsDetails("IsExpirDate")), "", (RsDetails("IsExpirDate").value))
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Else
                FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
            End If
            FG.TextMatrix(Num, FG.ColIndex("ItemBalance")) = IIf(IsNull(RsDetails("ItemBalance")), "", (RsDetails("ItemBalance").value))
            FG.TextMatrix(Num, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
                 FG.TextMatrix(Num, FG.ColIndex("ItemsDetailsNewidea")) = IIf(IsNull(RsDetails("ItemsDetailsNewidea")), "", (RsDetails("ItemsDetailsNewidea").value))
            FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
            FG.TextMatrix(Num, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))

            
            FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))


            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If
 fillapprovData
     
    mIsFinishSave = True
    
    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
   
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            If Not rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Retrive
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    Dim Msg As String
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial1.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
             CuurentLogdata ("D")
                Deletepost Me.Name, "Transactions", "Transaction_ID", 0, val(DCBranch.BoundText), val(XPTxtBillID.text), TxtNoteSerial1.text
                StrSQL = "delete ItemsDetails   where Transaction_ID= " & IIf(IsNull(rs("ReturnID").value), 0, rs("ReturnID").value)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "delete ItemsDetails   where Transaction_ID= " & val(Me.XPTxtBillID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                
                StrSQL = "Delete  FROM Transactions Where ReturnID=" & Me.XPTxtBillID.text & ""
                StrSQL = StrSQL + " AND Transactions.Transaction_Type=11"
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete  FROM Transactions Where Transaction_ID=" & Me.XPTxtBillID.text & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.text)
    
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                rs.Requery

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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "· ÕÊÌ· «·»÷«⁄… „‰ „Œ“‰ ≈·Ï „Œ“‰ ¬Œ—" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ «·»Ì«‰«  «·Õ«·Ì…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì…  ÕÊÌ·" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    'With TTP
    '   .Create Me.hwnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnAdd, _
    '    "≈÷«›… «·√’‰«› ..." & Wrap & _
    '    " ·«÷«›… ’‰› ÃœÌœ" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnRemove, _
    '    "Õ–› ’‰› ..." & Wrap & _
    '    "·Õ–› √Õœ «·√’‰«›" & Wrap & _
    '    " ÕœœÂ Ê«÷€ÿ Â‰«", True
    'End With
    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ÕÊÌ· «·»÷«⁄… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Function SaveItemsData(Optional Transaction_ID As Double = 0, Optional effect As Integer = 0)
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
    Cn.Execute "delete ItemsDetails   where Transaction_ID= " & Transaction_ID
    
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
                        RsgGrantee("Transaction_ID").value = Transaction_ID
                        RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                       RsgGrantee("EffectN").value = effect
                    RsgGrantee.update
                                    Next intX
                Else
    '            RsgGrantee.AddNew
    '          RsgGrantee("ParrtNoCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))
    '        RsgGrantee("count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
    '        RsgGrantee("unitid").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
    '      RsgGrantee("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
    '        RsgGrantee("sizeid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
    '        RsgGrantee("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
    '        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
    '       RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
    '      RsgGrantee("ItemDetailedCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))
    '      RsgGrantee("EffectN").value = -1
    '       RsgGrantee.update
    '
                   
                   End If
                   

 
                
  
                    
            End If

       

    Next RowNum

End Function

Private Sub SaveData(Optional ByVal IsSaveWithOutMsg As Boolean = False, _
                     Optional ByVal fromResave As Boolean = False)
                     
    Dim Msg               As String
    Dim RowNum            As Integer
    Dim RSTransDetails    As ADODB.Recordset
    Dim RsNotes           As ADODB.Recordset
    Dim RsTemp            As New ADODB.Recordset
    Dim RsTest            As New ADODB.Recordset
    Dim RsRepeat          As ADODB.Recordset
    Dim RsDetalis         As ADODB.Recordset
    Dim RsTrans           As ADODB.Recordset
    Dim StrSQL            As String
    Dim StrSqlDel         As String
    Dim note_id           As Integer
    Dim BeginTrans        As Boolean
    Dim LngItemID         As Long
    Dim Posted            As Integer
    Dim Transaction_Type  As Integer
    Dim Transaction_Type2 As Integer
    '****************************
    '· Ã«Â· Õ›Ÿ «· ›«’Ì· „⁄ «⁄«œÂ Ÿ»ÿ «·Õ—ﬂ« 
    Dim mSaveDetails      As Boolean
    mSaveDetails = True
    'mSaveDetails = Not ((fromResave And chkIgnorDetails.value = 1) Or Not fromResave)
    '***********************
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
    
    
   Dim CurrType As Long
   If Not rs.EOF Then
        CurrType = val(rs("Transaction_Type").value & "")
    End If

If CurrType = 992 Then
    Transaction_Type = 992
    Transaction_Type2 = 993
    Posted = 1
Else
    Transaction_Type = 10
    Transaction_Type2 = 11
    Posted = 0
End If
 
 
    
'    If IsSaveWithOutMsg Then
'        Transaction_Type = 10
'        Transaction_Type2 = 11
'        '  Posted = 0
'    Else
'        If CheckAprroveScreen(Me.Name) = True Then
'            Transaction_Type = 992
'            Transaction_Type2 = 993
'            Posted = 1
'        Else
'            Transaction_Type = 10
'            Transaction_Type2 = 11
'            Posted = 0
'        End If
'    End If
    
    If Me.TxtModFlg.text <> "R" Then
        If DCboStoreName.text = "" Then
                    
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰ «·–Ì ”Ì „ ‰ﬁ· «·»÷«⁄… „‰Â"
            Else
                Msg = "Specify From Store"
            End If
    
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Cmd(2).Enabled = True
            Exit Sub
        End If

        If DCboSecondStore.BoundText = "" Then
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰ «·–Ì ”Ì „ ‰ﬁ· «·»÷«⁄… ≈·ÌÂ"
            Else
                Msg = "Specify From To"
            End If
            If Not IsSaveWithOutMsg Then
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCboSecondStore.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Cmd(2).Enabled = True
           End If
            Exit Sub
        Else
        
        End If

        If DCboStoreName.BoundText = DCboSecondStore.BoundText Then
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«Ì„ﬂ‰ ≈Ã—«¡ ⁄„·Ì… «· ÕÊÌ· ≈·Ï ‰›” «·„Œ“‰"
            Else
                Msg = "  From  store and To store are the same"
            End If
    
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Screen.MousePointer = vbDefault
            Cmd(2).Enabled = True
            Exit Sub
        End If

        If XPDtbBill.value = "" Then
            Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ  ”ÃÌ· Â–Â «·⁄„·Ì…"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPDtbBill.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Cmd(2).Enabled = True
            Exit Sub
        End If
    
        '---------------------------------
        If TxtNoteSerial.text = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Cmd(2).Enabled = True:   Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Cmd(2).Enabled = True: Exit Sub
                Else
                    TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
        Dim NoteSerial1str As String
        
        If TxtNoteSerial1.text = "" Then
            NoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 12, 190, , 10, , val(DCboStoreName.BoundText))
            
            If NoteSerial1str = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÕÊÌ·  „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Cmd(2).Enabled = True: Exit Sub
            Else
                       
                If NoteSerial1str = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Cmd(2).Enabled = True: Exit Sub
                Else
                    TxtNoteSerial1.text = NoteSerial1str
                End If
            End If
        
        End If
     If Not IsSaveWithOutMsg Then
        If NewGrid.CheckDataEntered = False Then
            Cmd(2).Enabled = True
            Exit Sub
        End If
        
        End If

        Dim RsNotesGeneral As ADODB.Recordset
        Set RsNotesGeneral = New ADODB.Recordset
        Dim NoteID     As Long
        Dim NoteDate   As Date
        Dim NoteSerial As String
        Dim Notevalue  As Double
        Dim des        As String

        '     RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        '   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
        '   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
        '**************************************************************
        '        RsNotesGeneral.AddNew
        '        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        '        general_noteid = RsNotesGeneral("NoteID").value
        '        TxtNoteID.text = general_noteid
        '        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        '        RsNotesGeneral("NoteType").value = 190 ' «–‰ «÷«›…
        '        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) 'ﬁÌœ
        '        RsNotesGeneral("Note_Value").value = SngTemp
        '        RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        '        RsNotesGeneral("Remark").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        '        RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        '        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        '        RsNotesGeneral("numbering_type1").value = sand_numbering_type(12) '  «–‰ ’—›
        '        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        '        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        '        RsNotesGeneral("note_value_by_characters").value = WriteNo(Format(SngTemp, "0.00"), 0, True, ".")
        '        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        '        RsNotesGeneral.update
        '**************************************************************
 
        '............................................
           
        ' If MDIFrmMain.ActiveForm.name = "FrmReturnSalling" Then
        '''„Ì‰«) „—«Ã⁄… «·—’Ìœ ›Ï «·„Œ“‰
    
        Cn.BeginTrans
        BeginTrans = True
        CuurentLogdata
        If Me.TxtModFlg.text = "N" Then
            
            Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=10 or Transaction_Type=992"))
            '  XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            rs.AddNew
            rs("Transaction_ID").value = CStr(new_id("Transactions", "Transaction_ID", "", True))
            XPTxtBillID.text = rs("Transaction_ID").value
            ' rs.update
          
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
            If TxtNoteSerial1.text = "" Then
                TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 12, 190, , 10, , val(DCboStoreName.BoundText))
            End If
   
        ElseIf Me.TxtModFlg.text = "E" Then
           ' If mSaveDetails Then
                StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
                Cn.Execute StrSqlDel, , adExecuteNoRecords
           ' End If
            StrSqlDel = "delete From Transactions where ReturnID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
         
            StrSqlDel = "delete From Notes where noteid=" & val(TXTNoteID.text)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
          
            ' general_noteid = val(TxtNoteID.text)
          
            If TxtNoteSerial.text = "" Then
                TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 12, 190, , 10, , val(DCboStoreName.BoundText))
            End If
          
        End If
  
        Screen.MousePointer = vbArrowHourglass
        '     rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    
        '        rs("NoteId").value = val(TxtNoteID.text)
        rs("Transaction_Serial").value = Me.TxtTransSerial.text
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Type").value = Transaction_Type
        rs("OrderID").value = val(txtOrderID.text)
        rs("UserID").value = user_id
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("BranchId").value = IIf(Me.DCBranch.BoundText = "", 0, val(DCBranch.BoundText))
        Dim FromstoreAr As String
        Dim FromstoreEn As String
         
        Dim TostoreAr   As String
        Dim TostoreEn   As String
             
        getStorenames val(DCboStoreName.BoundText), FromstoreAr, FromstoreEn
        getStorenames val(DCboSecondStore.BoundText), TostoreAr, TostoreEn
      
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)

        If Trim$(Me.TxtCashCustomerName.text) <> "" Then
            rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
        Else
            rs("CashCustomerName").value = Null
        End If
        rs("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))
        rs("InspectionReport").value = IIf(Trim$(TxtInspectionReport.text) = "", Null, Trim$(TxtInspectionReport.text))
  
        rs("BillBasedOn").value = val(CBoBasedON.ListIndex)
        rs("order_no").value = TXT_order_no.text
   
        rs("Phone").value = IIf(Trim$(TxtPhone.text) = "", Null, Trim$(TxtPhone.text))
   
        rs.update

        Dim mProfitValueTotal As Double
        mSaveDetails = True
        If mSaveDetails Then
            For RowNum = 1 To FG.rows - 1
            
                If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then

                    'Check Repeat Serial
                    If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                        StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                        StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.text
                        Set RsTemp = New ADODB.Recordset
                        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                        If Not (RsTemp.EOF Or RsTemp.BOF) Then
                            Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                            Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                            Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            RsTemp.Close
                            FG.Row = RowNum
                            FG.Col = FG.ColIndex("name")
                            FG.ShowCell RowNum, FG.ColIndex("name")
                            FG.SetFocus
                            Screen.MousePointer = vbDefault
                            BeginTrans = False
                            Cn.RollbackTrans
                            Cmd(2).Enabled = True
                            Exit Sub
                        End If

                        RsTemp.Close
                    End If
                    Set RSTransDetails = New ADODB.Recordset
                    '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                    StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
                    RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                    RSTransDetails.AddNew
                    RSTransDetails("FromstoreAr").value = FromstoreAr
                    RSTransDetails("TostoreAr").value = TostoreAr
                    RSTransDetails("FromstoreEn").value = FromstoreEn
                    RSTransDetails("TostoreEn").value = TostoreEn
                    RSTransDetails("IsExpirDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("IsExpirDate")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("IsExpirDate"))))
                    RSTransDetails("OriginalQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("OriginalQty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("OriginalQty"))))
                    RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
                    RSTransDetails("OrderArrivalDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
                    RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
                    RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
                    RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                    RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                    If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                        StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                        If Not (RsTemp.EOF Or RsTemp.BOF) Then
                            If RsTemp("HaveSerial").value = True Then
                                RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                            End If
                        End If

                        RsTemp.Close
                    End If
                    
                    RSTransDetails("ItemBalance").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemBalance")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemBalance"))))
                    RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                    RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            
                    RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                    RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
               
                    RSTransDetails("ProfitType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProfitType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ProfitType"))))
                    RSTransDetails("ProfitValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProfitValue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ProfitValue"))))
                    RSTransDetails("NetProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NetProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("NetProfit"))))
                
                    mProfitValueTotal = mProfitValueTotal + val(RSTransDetails("NetProfit").value & "")
                    RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
                    RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
            
                    RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                    RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                    RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))

                    RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))

                    RSTransDetails("BranchId").value = Me.dcBranch1.BoundText
                    RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
                    RSTransDetails("ItemsDetailsNewidea").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")))

                    Dim RsUnitData   As ADODB.Recordset
                    Dim LngCurItemID As Long
                    Dim LngUnitID    As Long
                    Dim DblQty       As Double
        
                    LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                    DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                    StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                    StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                    Set RsUnitData = New ADODB.Recordset
                    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                        RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                
                        RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                        RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
                    End If
                    
                    RSTransDetails("Height").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Height")))
                    RSTransDetails("length").value = val(FG.TextMatrix(RowNum, FG.ColIndex("length")))
                    RSTransDetails("Width").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Width")))
            
                    'RSTransDetails("price").value = Round(Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) / RSTransDetails("Quantity").value, 2)
                    RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
                    RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
                    RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
                    Dim OldQty  As Double
                    Dim OldCost As Double
                    Dim NewQty  As Double
                    Dim NewCost As Double
               
                    getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.text), OldQty, OldCost, NewQty, NewCost, , LngUnitID
                    RSTransDetails("OldQty").value = NewQty
                    RSTransDetails("OldCost").value = NewCost
       
                    RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
                    RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       
                    RSTransDetails.update
                    
                End If

            Next RowNum
       
            SaveItemsData val(XPTxtBillID.text), -1
        End If
        '≈÷«›… «·»÷«∆⁄ ≈·Ï «·„Œ“‰ «·ÃœÌœ
        rs.AddNew
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        rs("NoteId").value = val(TXTNoteID.text)
        rs("Transaction_ID").value = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Type").value = Transaction_Type2
        rs("UserID").value = user_id
        rs("StoreID").value = IIf(DCboSecondStore.BoundText = "", Null, val(DCboSecondStore.BoundText))
        rs("ReturnID").value = val(XPTxtBillID.text)
        rs("BranchId").value = IIf(Me.DCBranch.BoundText = "", 0, val(DCBranch.BoundText))
     
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
             
        rs.update
        If mSaveDetails Then
            For RowNum = 1 To FG.rows - 1
                If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                    RSTransDetails.AddNew
                    RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
                    RSTransDetails("OrderArrivalDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
                    RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
                    RSTransDetails("Transaction_ID").value = rs("Transaction_ID").value
                    RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
             
                    RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))

                    If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                        StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                        If Not (RsTemp.EOF Or RsTemp.BOF) Then
                            If RsTemp("HaveSerial").value = True Then
                                RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                            End If
                        End If

                        RsTemp.Close
                    End If

                    RSTransDetails("BranchId").value = Me.Dcbranch2.BoundText
            
                    RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
 
                    RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                    RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            
                    RSTransDetails("ProfitType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProfitType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ProfitType"))))
                    RSTransDetails("ProfitValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProfitValue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ProfitValue"))))
                    RSTransDetails("NetProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NetProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("NetProfit"))))

                    RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
                    RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
            
                    RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                    RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                    RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
                    If val(RSTransDetails("NetProfit").value & "") <> 0 Then
                        RSTransDetails("ShowPrice").value = (val(RSTransDetails("NetProfit").value & "") / val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                    Else
                        RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                
                    End If
                    RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

                    '---------------------------------------------------
                    Dim RsUnitData1 As ADODB.Recordset
        
                    LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                    DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                    StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                    StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                    Set RsUnitData1 = New ADODB.Recordset
                    RsUnitData1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (RsUnitData1.BOF And RsUnitData1.EOF) Then
                        RSTransDetails("QtyBySmalltUnit").value = RsUnitData1("UnitFactor").value
                        RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                        RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
                    End If

                    'RSTransDetails("price").Value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").Value, 2)
                    RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
                    RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
                    RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
                    RSTransDetails("FromstoreAr").value = FromstoreAr
                    RSTransDetails("TostoreAr").value = TostoreAr
                    RSTransDetails("FromstoreEn").value = FromstoreEn
                    RSTransDetails("TostoreEn").value = TostoreEn
                    getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboSecondStore.BoundText), val(Me.XPTxtBillID.text), OldQty, OldCost, NewQty, NewCost, , LngUnitID
                    RSTransDetails("OldQty").value = NewQty
                    RSTransDetails("OldCost").value = NewCost
       
                    RSTransDetails("NewQty").value = RSTransDetails("Quantity").value + RSTransDetails("OldQty").value
                    If (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value) <> 0 Then
                        RSTransDetails("NewCost").value = ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
                    Else
                        RSTransDetails("NewCost").value = 0
                    End If
       
                    RSTransDetails.update
                End If
            Next RowNum
            SaveItemsData rs("Transaction_ID").value, 1
        End If
        Cn.CommitTrans
        BeginTrans = False
    
        CreateNotes NoteID, (XPDtbBill.value), val(DCBranch.BoundText), 190, 0, TxtNoteSerial.text, TxtNoteSerial1, "Transactions", "Transaction_ID", val(XPTxtBillID.text), TxtNoteSerial1.text, ToHijriDate(XPDtbBill.value)
        TXTNoteID.text = NoteID
        '      If TxtNoteSerial.text = "" Then
        '      TxtNoteSerial.text = NoteSerial
        '      End If
           
        general_noteid = NoteID
    
        Dim LngDevID           As Long
        Dim LngDevNO           As Integer
        Dim StrTempAccountCode As String
        Dim StrTempDes         As String

        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        '----------------
        Dim Account_Code_dynamic As String
        'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) 'ﬁÌœ

        If SngTemp > 0 Then
            '1 work with branch
            '2 work with inventory
            '3 work with groups

            If detect_inventory_work_type = 1 Then
                ' 1«·„Œ“Ê‰ ›Ì «·›—⁄
                Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    Cmd(2).Enabled = True
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        Cmd(2).Enabled = True
                        GoTo ErrTrap
         
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
    
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "√–‰  ÕÊÌ· »÷«∆⁄ »Ì‰ «·„Œ«“‰  —ﬁ„ " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = "  Moving Items Vchr  No. " & Me.TxtNoteSerial1.text
                End If
        
                LngDevNO = 0

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch1.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
     
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.Dcbranch2.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
    
            ElseIf detect_inventory_work_type = 2 Then
                Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

                If Account_Code_dynamic = "" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄      " & DCboStoreName.text, vbCritical
                    GoTo ErrTrap
                End If
    
                StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

                ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "√–‰  ÕÊÌ· »Ì‰ «·„Œ«“‰   —ﬁ„ " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = " Moving Items Vchr  No. " & Me.TxtNoteSerial1.text
                End If
    
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch1.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If

                '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
                Account_Code_dynamic = get_store_Account(DCboSecondStore.BoundText, "Account_Code")

                If Account_Code_dynamic = "" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    " & DCboSecondStore.text, vbCritical
                    GoTo ErrTrap
                End If
                If mProfitValueTotal <> 0 Then
                    Dim Account_Code_dynamic157 As String
                    Account_Code_dynamic157 = get_account_code_branch(157, val(Dcbranch2.BoundText))
                    If Account_Code_dynamic157 = "" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·—»ÕÌ… ·Â–« «·›—⁄    " & DCboSecondStore.text, vbCritical
                        GoTo ErrTrap
                    End If
                End If
                StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

                ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "√–‰  ÕÊÌ· »Ì‰ «·„Œ«“‰   —ﬁ„ " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = " Moving Items Vchr  No. " & Me.TxtNoteSerial1.text
                End If
    
                LngDevNO = LngDevNO + 1
                If mProfitValueTotal <> 0 Then
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, mProfitValueTotal, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.Dcbranch2.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                        GoTo ErrTrap
                    End If
                Else
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.Dcbranch2.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                        GoTo ErrTrap
                    End If
                End If
                If mProfitValueTotal <> 0 Then

                    If Account_Code_dynamic157 <> "" Then
                        LngDevNO = LngDevNO + 1
                        '
                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic157, mProfitValueTotal - SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch1.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                            GoTo ErrTrap
                        End If
                        
                    End If
                End If
                Dim BranchId1   As Integer
                Dim BranchID2   As Integer
                Dim DeptSide1   As String
                Dim CreditSide1 As String
                Dim noteid1     As Double
                BranchId1 = val(dcBranch1.BoundText)
                BranchID2 = val(Dcbranch2.BoundText)
                LngDevNO = LngDevNO + 1
                If BranchId1 <> BranchID2 Then

                    DeptSide1 = getBranchCurrentAccount(BranchId1)
                    CreditSide1 = getBranchCurrentAccount(BranchID2)
                    LngDevNO = LngDevNO + 1

                    Dim mTotal As Double
                    If mProfitValueTotal <> 0 Then
                        mTotal = mProfitValueTotal
                    Else
                        mTotal = SngTemp
                    End If

                    If CreditSide1 <> "" Then
                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, CreditSide1, mTotal, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , BranchId1, , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                            GoTo ErrTrap
                        End If
                    End If
                    LngDevNO = LngDevNO + 1
                    If DeptSide1 <> "" Then
                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, DeptSide1, mTotal, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                            GoTo ErrTrap
                        End If
                    End If
                    noteid1 = val(general_noteid)
                    updateNotesValueAndNobytext noteid1, CDbl(mTotal)
                
                End If
            ElseIf detect_inventory_work_type = 3 Then
                Dim groupAccount As String
             
                Dim line_value   As Single
                Dim i            As Integer

                With FG

                    For i = 1 To FG.rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                            groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …" & DCboStoreName.text
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined" & DCboStoreName.text
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "√–‰  ÕÊÌ· Ì÷«∆⁄ »Ì‰ «·„Œ«“‰  —ﬁ„ " & Me.TxtNoteSerial1.text
                            Else
                                StrTempDes = "moving items   No. " & Me.TxtNoteSerial1.text
                            End If

                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch1.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With
 
                With FG

                    For i = 1 To FG.rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                            groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboSecondStore.BoundText, 0)

                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …" & DCboSecondStore.text
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined" & DCboSecondStore.text
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "√–‰  ÕÊÌ·   —ﬁ„ " & Me.TxtNoteSerial1.text
                            Else
                                StrTempDes = " Moving Items No. " & Me.TxtNoteSerial1.text
                            End If

                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.Dcbranch2.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If

            '----------------
            'LngDevID = LngDevID + 1
            'LngDevNO = 0
        End If
    
        '----------------------------------------------------------------
        '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
        If publicSearch = False Then
          
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_ID=" & val(XPTxtBillID.text) & ""
            If SystemOptions.usertype <> UserAdminAll Then
 
                If SystemOptions.FixedCustomer = 1 Then
                    StrSQL = StrSQL & " and  UserID = " & user_id
                End If
      
            End If
    
        Else
        
            StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=10 or Transaction_Type=992)"
            StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"

            If SystemOptions.usertype <> UserAdminAll Then
 
                If SystemOptions.FixedCustomer = 1 Then
                    StrSQL = StrSQL & " and  UserID = " & user_id
                    ' StrSQL = StrSQL & " or  storeId = " & dstore
                      
                End If
                Me.DCBranch.Enabled = True
      
            End If
    
            StrSQL = StrSQL & " Order by Transaction_ID"
   
        End If
 
        ' Set rs = New ADODB.Recordset
        '     rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        '     Me.Retrive val(Me.XPTxtBillID.Text)
        '----------------------------------------------------------------
        If IsSaveWithOutMsg Then Exit Sub
        
        Select Case Me.TxtModFlg.text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Success" & CHR(13)
                    Msg = Msg + "Enter Another Record ?    yes/ No"
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
                    MsgBox "changes Was Saved ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
                lbl(7).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
            
        End Select
        TxtModFlg.text = "R"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Me.Retrive val(Me.XPTxtBillID.text)
        '---------------------------------------
        TxtModFlg.text = "R"
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Cmd(2).Enabled = True
End Sub

Private Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap

    If FG.rows > 1 Then
        If FG.rows = 2 Then
            FG.Clear flexClearScrollable, flexClearEverything
        Else

            If FG.rows > 1 Then
                If FG.Row <> FG.FixedRows - 1 Then
                    FG.RemoveItem (FG.Row)
                End If
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport()

    Dim BuyReport As ClsBuyReport
    On Error GoTo ErrTrap
    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowMovingVOucherData XPTxtBillID.text
        End If

    Else

        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowMovingVOucherData XPTxtBillID.text, True
        End If
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

Public Sub Convert()
    Cmd_Click (0)
End Sub

Private Sub XPFillData_Click()
    On Error GoTo ErrTrap

    If FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code")) <> "" Then
        FG.rows = FG.rows + 1
        NewGrid.GridDefaultValue FG.rows - 1
    End If

  '  FrmFillItems.DealingForm = MoveItems
  '  FrmFillItems.show vbModal
  '  NewGrid.Calculate 1, , , True
    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub XPDtbBill_DblClick()
  If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
TxtNoteSerial1.text = ""
 TxtNoteSerial.text = ""
End If

End Sub

