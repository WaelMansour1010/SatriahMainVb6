VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmVacationSettings 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”Ì«”… «Õ ”«» «·«Ã«“« "
   ClientHeight    =   8160
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   13335
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmVacationSettings.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   13335
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8160
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13335
      _cx             =   23521
      _cy             =   14393
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
         Height          =   7230
         Left            =   30
         TabIndex        =   1
         Top             =   -90
         Width           =   18615
         _cx             =   32835
         _cy             =   12753
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
         Caption         =   "."
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6810
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   18525
            _cx             =   32676
            _cy             =   12012
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
               Height          =   660
               Index           =   5
               Left            =   0
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   0
               Width           =   18525
               _cx             =   32676
               _cy             =   1164
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
               Picture         =   "FrmVacationSettings.frx":038A
               Caption         =   ""
               Align           =   1
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
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   270
                  Index           =   0
                  Left            =   1470
                  TabIndex        =   27
                  Top             =   60
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmVacationSettings.frx":1064
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
                  Height          =   270
                  Index           =   2
                  Left            =   765
                  TabIndex        =   28
                  Top             =   60
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmVacationSettings.frx":13FE
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
                  Height          =   270
                  Index           =   1
                  Left            =   1815
                  TabIndex        =   29
                  Top             =   60
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmVacationSettings.frx":1798
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
                  Height          =   270
                  Index           =   3
                  Left            =   1110
                  TabIndex        =   30
                  Top             =   60
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmVacationSettings.frx":1B32
                  ColorHighlight  =   4194304
                  ColorHoverText  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
                  ColorToggledHoverText=   16777215
                  ColorTextShadow =   16777215
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "”Ì«”… «Õ ”«» «·«Ã«“« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Index           =   8
                  Left            =   10875
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   90
                  Width           =   2715
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6780
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   17790
               _cx             =   31380
               _cy             =   11959
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
               Begin XtremeSuiteControls.CheckBox ChCommContract 
                  Height          =   255
                  Left            =   600
                  TabIndex        =   56
                  Top             =   720
                  Width           =   1815
                  _Version        =   786432
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«·«· “«„ »„œ… «·⁄ﬁœ"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox TxtNoMonth 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   735
                  Width           =   885
               End
               Begin XtremeSuiteControls.RadioButton RdType 
                  Height          =   390
                  Index           =   0
                  Left            =   8235
                  TabIndex        =   33
                  Top             =   615
                  Width           =   1350
                  _Version        =   786432
                  _ExtentX        =   2381
                  _ExtentY        =   688
                  _StockProps     =   79
                  Caption         =   "ÿ»ﬁ« ··⁄ﬁœ"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6570
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Text            =   "0"
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   585
               End
               Begin VB.TextBox xptxtid 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9750
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   735
                  Width           =   2355
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   -4575
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   9495
                  Width           =   2550
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6810
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   2550
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3810
                  Left            =   0
                  TabIndex        =   7
                  Top             =   2310
                  Width           =   13185
                  _cx             =   23257
                  _cy             =   6720
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
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmVacationSettings.frx":1ECC
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
               Begin XtremeSuiteControls.RadioButton RdType 
                  Height          =   390
                  Index           =   1
                  Left            =   3555
                  TabIndex        =   34
                  Top             =   615
                  Width           =   4560
                  _Version        =   786432
                  _ExtentX        =   8043
                  _ExtentY        =   688
                  _StockProps     =   79
                  Caption         =   "„‰  «—ÌŒ «Œ— ⁄Êœ…  »‘—ÿ  Ã«Ê“ „œ… «·Œœ„… "
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic1 
                  Height          =   1065
                  Left            =   0
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   1230
                  Width           =   13305
                  _cx             =   23469
                  _cy             =   1879
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
                  Begin ImpulseButton.ISButton ISButton2 
                     Height          =   450
                     Left            =   13890
                     TabIndex        =   38
                     TabStop         =   0   'False
                     ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   794
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
                     ButtonImage     =   "FrmVacationSettings.frx":1F6A
                     ColorButton     =   14737632
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton ISButton3 
                     Height          =   450
                     Left            =   14895
                     TabIndex        =   39
                     TabStop         =   0   'False
                     ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
                     Top             =   315
                     Visible         =   0   'False
                     Width           =   990
                     _ExtentX        =   1746
                     _ExtentY        =   794
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
                     ButtonImage     =   "FrmVacationSettings.frx":2304
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton ISButton4 
                     Height          =   390
                     Left            =   16305
                     TabIndex        =   40
                     TabStop         =   0   'False
                     Top             =   210
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   688
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
                     ButtonImage     =   "FrmVacationSettings.frx":269E
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin MSComCtl2.DTPicker FrmDate 
                     Height          =   315
                     Left            =   9675
                     TabIndex        =   42
                     Top             =   600
                     Width           =   2040
                     _ExtentX        =   3598
                     _ExtentY        =   556
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   95617025
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker ToDate 
                     Height          =   315
                     Left            =   6030
                     TabIndex        =   44
                     Top             =   600
                     Width           =   2055
                     _ExtentX        =   3625
                     _ExtentY        =   556
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   95617025
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker AlowDate 
                     Height          =   315
                     Left            =   2415
                     TabIndex        =   46
                     Top             =   600
                     Width           =   2025
                     _ExtentX        =   3572
                     _ExtentY        =   556
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   95617025
                     CurrentDate     =   38784
                  End
                  Begin ImpulseButton.ISButton ISButton5 
                     Height          =   555
                     Left            =   120
                     TabIndex        =   48
                     ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                     Top             =   480
                     Width           =   2175
                     _ExtentX        =   3836
                     _ExtentY        =   979
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
                     ButtonImage     =   "FrmVacationSettings.frx":2A38
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "› —… «·”„«Õ ﬁ»·"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   6
                     Left            =   4545
                     TabIndex        =   47
                     Top             =   600
                     Width           =   1410
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Ï  «—ÌŒ"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   4
                     Left            =   8190
                     TabIndex        =   45
                     Top             =   600
                     Width           =   1410
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‰  «—ÌŒ"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   3
                     Left            =   11835
                     TabIndex        =   43
                     Top             =   600
                     Width           =   1395
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "› —«  «·«Ã«“…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   390
                     Index           =   2
                     Left            =   5775
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   120
                     Width           =   3555
                  End
               End
               Begin ImpulseButton.ISButton ISButton6 
                  Height          =   315
                  Left            =   9390
                  TabIndex        =   49
                  ToolTipText     =   "Õ–› «·ﬂ·"
                  Top             =   6270
                  Width           =   1650
                  _ExtentX        =   2910
                  _ExtentY        =   556
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
                  ButtonImage     =   "FrmVacationSettings.frx":929A
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton ISButton7 
                  Height          =   315
                  Left            =   11820
                  TabIndex        =   50
                  ToolTipText     =   "Õ–› «·ﬂ·"
                  Top             =   6270
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› "
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
                  ButtonImage     =   "FrmVacationSettings.frx":FAFC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Index           =   0
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   735
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   5
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   1695
                  Width           =   2070
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Index           =   7
                  Left            =   11820
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   735
                  Width           =   1275
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   390
                  Left            =   16155
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   975
                  Width           =   960
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   9855
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   90
               Width           =   1290
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   975
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   7185
         Width           =   13335
         _cx             =   23521
         _cy             =   1720
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
            Height          =   330
            Left            =   13920
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   105
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
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
            ButtonImage     =   "FrmVacationSettings.frx":1635E
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   360
            Left            =   14925
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   225
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVacationSettings.frx":166F8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   300
            Left            =   16335
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   529
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
            ButtonImage     =   "FrmVacationSettings.frx":16A92
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   0
            Left            =   9630
            TabIndex        =   17
            Top             =   255
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   900
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
            Height          =   510
            Index           =   1
            Left            =   8580
            TabIndex        =   18
            Top             =   285
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   900
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
            Height          =   510
            Index           =   2
            Left            =   7635
            TabIndex        =   19
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   900
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
            Height          =   510
            Index           =   3
            Left            =   6405
            TabIndex        =   20
            Top             =   285
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   900
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
            Height          =   510
            Index           =   4
            Left            =   5205
            TabIndex        =   21
            Top             =   285
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   900
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
            Height          =   510
            Index           =   6
            Left            =   4110
            TabIndex        =   22
            Top             =   285
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   900
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
            Height          =   510
            Index           =   5
            Left            =   4980
            TabIndex        =   23
            Top             =   285
            Visible         =   0   'False
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   900
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
            Height          =   510
            Index           =   7
            Left            =   4065
            TabIndex        =   31
            Top             =   285
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   900
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
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   300
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   555
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   330
            Left            =   2205
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   360
            Width           =   510
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   300
            Index           =   10
            Left            =   930
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   255
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   300
            Index           =   9
            Left            =   2715
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   240
            Width           =   840
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
            Height          =   225
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   105
            Width           =   2070
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
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   255
            Width           =   1785
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
      ButtonImage     =   "FrmVacationSettings.frx":16E2C
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmVacationSettings"
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

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
 If RdType(1).value = True Then
      '  If val(Me.TxtNoMonth.Text) = 0 Then
      '      If SystemOptions.UserInterface = EnglishInterface Then
      '          Msg = "Please Enter Duration of service No. !!"
      '      Else
      '          Msg = "ÌÃ»  «œŒ«·  „œ… «·Œœ„…..!!"
      '      End If
      '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      '      TxtNoMonth.SetFocus
      '      SendKeys "{F4}"
      '      Exit Sub
      '  End If
  End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TblVacationSettingsDet where VacSetID=" & val(Me.XPTxtID.Text)
    End If
    rs("ID").value = XPTxtID.Text
    rs("NoMonth").value = val(TxtNoMonth.Text)
    rs("FrmDate").value = FrmDate.value
    rs("ToDate").value = ToDate.value
    rs("AlowDate").value = AlowDate.value
    rs("UserID").value = user_id
    If Me.RdType(1).value = True Then
    rs("Typ").value = 1
    Else
    rs("Typ").value = 0
    End If
    If ChCommContract.value = vbChecked Then
    rs("CommContract").value = 1
    Else
    rs("CommContract").value = 0
    End If
    rs.update
    Set RsDev = New ADODB.Recordset
    RsDev.Open "TblVacationSettingsDet", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("FrmDate")) <> "" Then
                RsDev.AddNew
                RsDev("FrmDate").value = IIf(.TextMatrix(i, .ColIndex("FrmDate")) = "", Null, .TextMatrix(i, .ColIndex("FrmDate")))
                RsDev("ToDate").value = IIf(.TextMatrix(i, .ColIndex("ToDate")) = "", Null, .TextMatrix(i, .ColIndex("ToDate")))
                RsDev("AlowDate").value = IIf(.TextMatrix(i, .ColIndex("AlowDate")) = "", Null, .TextMatrix(i, .ColIndex("AlowDate")))
                RsDev("VacSetID").value = Me.XPTxtID.Text
                RsDev.update
            End If
        Next i
    End With
    Cn.CommitTrans
    BeginTrans = False
    Select Case Me.TxtModFlg.Text
        Case "N"
        '################### Khaled Was Here ##########################
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Data saved" & CHR(13)
                Msg = Msg + "Do you want to enter another recored"
            Else
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            End If
        '##############################################################
            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Edits saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.Text = "R"
    Retrive
    'End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Data can not be saved" & CHR(13)
            Msg = Msg + "An invalid values was entered" & CHR(13)
            Msg = Msg + "make sure of the data and try again"
        Else
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Sorry,something went wrong while saving the data" & CHR(13)
    Else
         Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    End If
    
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   End If
End Sub
Sub RemoveGrid()
If Me.TxtModFlg.Text <> "R" Then
With Grid
If .Row <= 0 Then Exit Sub
.RemoveItem .Row
End With
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
     On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.Text = "N"
            clear_all Me
            Me.XPTxtID.Text = CStr(new_id("TblVacationSettings", "ID", "", True))
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            ChCommContract.value = vbUnchecked
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "E"
        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

             Del_Trans
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
        Case 6
            Unload Me

        Case 7
                 If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            print_report
         End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "This recored will be deleted" & CHR(13)
            Msg = Msg + "Do you want to continue ?"
        Else
            Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        End If
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblVacationSettings Where id=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                 Cn.Execute "delete TblVacationSettingsDet where VacSetID=" & val(Me.XPTxtID)

                If rs.RecordCount < 1 Then
                Grid.Clear flexClearScrollable, flexClearEverything
                Grid.Rows = 2
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
        If SystemOptions.UserInterface = EnglishInterface Then
             Msg = "This process is not allowed because there no recoreds"
        Else
             Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Sorry,something went wrong while saving the data" & CHR(13)
    Else
         Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture

    Dim My_SQL2 As String

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    Set BKGrndPic = New ClsBackGroundPic
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
    StrSQL = "select * From TblVacationSettings  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
'################################# Khaled Was Here ############################
    FrmVacationSettings.Caption = "Vacation Policy"
    ChCommContract.RightToLeft = False
    ChCommContract.Caption = "Contract Period"
    lbl(8).Caption = "Vacation Policy"
    lbl(7).Caption = "Ser No."
    RdType(0).Caption = "Based on contract"
    RdType(1).Caption = "From the last date of return in condition of passing the service duration"
    lbl(0).Caption = "Month"
    lbl(3).Caption = "From date"
    lbl(4).Caption = "To Date"
    lbl(6).Caption = "Allowance duration before"
    ISButton5.Caption = "Add"
    lbl(2).Caption = "Vacation Periods"
    With Grid
        .TextMatrix(0, .ColIndex("Ser")) = "No."
        .TextMatrix(0, .ColIndex("FrmDate")) = "From Date."
        .TextMatrix(0, .ColIndex("ToDate")) = "To Date."
        .TextMatrix(0, .ColIndex("AlowDate")) = "Allowance duration"
    End With
    
    ISButton7.Caption = "Delete"
    ISButton6.Caption = "Delete All"
    
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(6).Caption = "Exit"
    Cmd(7).Caption = "Print"
    
    lbl(9).Caption = "Current Record"
    lbl(10).Caption = "No. of Record"
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

Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = " SELECT     dbo.TbVisaDeti.ID, dbo.TbVisaDeti.VisaID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TbVisaDeti.HododNo, dbo.TbVisaDeti.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
MySQL = MySQL & "                      dbo.TbVisaDeti.NotionalID, dbo.Nationality.name, dbo.Nationality.namee, dbo.TbVisaDeti.CityID, dbo.TblCountriesGovernments.GovernmentName,"
MySQL = MySQL & "                      dbo.TbVisa.OrderNo, dbo.TbVisa.VisaNo, dbo.TbVisa.Priod, dbo.TbVisa.DMYPriod, dbo.TbVisa.StarDate, dbo.TbVisa.StarDateH, dbo.TbVisa.EndDate,"
MySQL = MySQL & "                      dbo.TbVisa.EndDateH, dbo.TbVisa.ID AS IDM, dbo.TbVisaDeti.Place, dbo.TbVisaDeti.Type, dbo.TbVisaDeti.[count], dbo.TbVisaDeti.Price"
MySQL = MySQL & " FROM         dbo.Nationality RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TbVisa LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TbVisaDeti ON dbo.TbVisa.ID = dbo.TbVisaDeti.VisaID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments ON dbo.TbVisaDeti.CityID = dbo.TblCountriesGovernments.GovernmentID ON"
MySQL = MySQL & "                      dbo.Nationality.id = dbo.TbVisaDeti.NotionalID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TbVisaDeti.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & " Where (dbo.TbVisa.id = " & val(XPTxtID.Text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVisaData.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVisaDataE.rpt"
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
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "There's no recoreds to show"
        Else
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        End If

 
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
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
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
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
  If Lngid <> 0 Then
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    Me.XPTxtID.Text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
    Me.TxtNoMonth.Text = IIf(IsNull(rs("NoMonth").value), "", rs("NoMonth").value)
    FrmDate.value = IIf(IsNull(rs("FrmDate").value), Date, rs("FrmDate").value)
    ToDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
    AlowDate.value = IIf(IsNull(rs("AlowDate").value), Date, rs("AlowDate").value)
    If Not IsNull(rs("Typ").value) Then
    If (rs("Typ").value) = 1 Then
    RdType(1).value = True
    Else
    RdType(0).value = True
    End If
    Else
    RdType(0).value = True
    End If
    
    If Not IsNull(rs("CommContract").value) Then
    If (rs("CommContract").value) = 1 Then
    ChCommContract.value = vbChecked
    Else
    ChCommContract.value = vbUnchecked
    End If
    Else
    ChCommContract.value = vbUnchecked
    End If
''///////////
   StrSQL = "select * from TblVacationSettingsDet where VacSetID=" & Me.XPTxtID.Text
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
        With Me.Grid
            .Rows = .FixedRows + RsDev.RecordCount
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("FrmDate")) = IIf(IsNull(RsDev("FrmDate").value), "", RsDev("FrmDate").value)
                .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(RsDev("ToDate").value), "", RsDev("ToDate").value)
                .TextMatrix(i, .ColIndex("AlowDate")) = IIf(IsNull(RsDev("AlowDate").value), "", RsDev("AlowDate").value)
                RsDev.MoveNext
            Next i
        End With
    End If
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount

    Exit Sub
ErrTrap:
End Sub
Private Sub GRID2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub

Private Sub ISButton5_Click()
If Me.TxtModFlg.Text <> "R" Then
FillGrid
End If
End Sub

Private Sub ISButton6_Click()
If Me.TxtModFlg.Text <> "R" Then
  Grid.Clear flexClearScrollable, flexClearEverything
  Grid.Rows = 1
End If
End Sub

Private Sub ISButton7_Click()
RemoveGrid
End Sub

Private Sub RdType_Click(Index As Integer)
If RdType(0).value = True Then
TxtNoMonth.Enabled = False
TxtNoMonth.Text = ""
Else
TxtNoMonth.Enabled = True
End If
End Sub
Sub FillGrid()
Dim k As Integer
With Grid
.Rows = .Rows + 1
k = .Rows - 1
.TextMatrix(k, .ColIndex("FrmDate")) = FrmDate.value
.TextMatrix(k, .ColIndex("ToDate")) = ToDate.value
.TextMatrix(k, .ColIndex("AlowDate")) = AlowDate.value
.TextMatrix(k, .ColIndex("Ser")) = k
End With
End Sub
Private Sub TxtModFlg_Change()
    If Me.TxtModFlg.Text = "N" Then
        ELe(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        ELe(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        ELe(1).Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub
Private Sub TxtNoMonth_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNoMonth.Text, 0)
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
