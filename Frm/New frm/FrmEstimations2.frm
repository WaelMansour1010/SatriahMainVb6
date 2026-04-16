VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmAllocationToContract 
   BackColor       =   &H00E2E9E9&
   Caption         =   "   ‘«‘… «·œ›⁄«  «·„” Õﬁ…   "
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
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
   Icon            =   "FrmEstimations2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9525
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18960
      _cx             =   33443
      _cy             =   16801
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
         Height          =   7905
         Left            =   30
         TabIndex        =   1
         Top             =   720
         Width           =   18900
         _cx             =   33338
         _cy             =   13944
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
         Caption         =   "»Ì«‰«  «·«” Õﬁ«ﬁ« |‘—Õ «·„Ê«“‰…"
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
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7485
            Index           =   1
            Left            =   45
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Width           =   18810
            _cx             =   33179
            _cy             =   13203
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Begin VB.CommandButton cmdCreateEntry 
               Caption         =   "«‰‘«¡ «·ﬁÌœ"
               Height          =   345
               Left            =   4770
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   6750
               Width           =   1665
            End
            Begin VB.TextBox txtSMS 
               Alignment       =   1  'Right Justify
               Height          =   540
               Left            =   4950
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   110
               Top             =   1020
               Width           =   6315
            End
            Begin VB.PictureBox Picture1 
               Height          =   90
               Left            =   960
               ScaleHeight     =   30
               ScaleWidth      =   30
               TabIndex        =   109
               Top             =   6600
               Width           =   90
            End
            Begin VB.CommandButton cmdPrintInst 
               Appearance      =   0  'Flat
               Caption         =   "ÿ»«⁄… «·›Ê« Ì—"
               Height          =   375
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   6720
               Width           =   1935
            End
            Begin VB.CommandButton Command2 
               Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
               Height          =   345
               Left            =   8190
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   6720
               Width           =   1665
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   1290
               Left            =   20670
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   405
               Width           =   2910
               Begin VB.OptionButton DistType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· Ê“Ì⁄ ⁄·Ï „—«ﬂ“  ﬂ·›…"
                  Height          =   210
                  Index           =   1
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   480
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
                  TabIndex        =   79
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.OptionButton DistType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· Ê“Ì⁄ ⁄·Ï  «·›—Ê⁄"
                  Height          =   210
                  Index           =   2
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   720
                  Width           =   2055
               End
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   495
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   -405
               Visible         =   0   'False
               Width           =   2640
            End
            Begin VB.TextBox txtid 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   0
               Left            =   -4845
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   11790
               Width           =   2655
            End
            Begin VB.TextBox TxtTransID 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   16185
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   0
               Width           =   1215
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   930
               Left            =   19110
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   1035
               Width           =   2655
               Begin VB.OptionButton PercentagType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”» «·ÌÂ"
                  Height          =   210
                  Index           =   0
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.OptionButton PercentagType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”» ÌœÊÌÂ"
                  Height          =   210
                  Index           =   1
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   480
                  Width           =   1335
               End
            End
            Begin VB.TextBox TxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   690
               Left            =   4950
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   67
               Top             =   360
               Width           =   6315
            End
            Begin VB.Frame Frame1 
               Caption         =   "«· Ê“Ì⁄ ⁄·Ï «Õ”«»« "
               Height          =   1065
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   9075
               Width           =   17295
               Begin VB.TextBox TxtPercentage 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.TextBox TxtRemarks1 
                  Alignment       =   1  'Right Justify
                  Height          =   615
                  Left            =   2160
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   59
                  Top             =   120
                  Width           =   3615
               End
               Begin MSDataListLib.DataCombo DCAccountDist 
                  Height          =   315
                  Left            =   8760
                  TabIndex        =   61
                  Top             =   240
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   582
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  TabIndex        =   62
                  Top             =   240
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
                  FontName        =   "Arial"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmEstimations2.frx":038A
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   240
                  TabIndex        =   63
                  Top             =   240
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  FontName        =   "Arial"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmEstimations2.frx":0724
                  DrawFocusRectangle=   0   'False
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
                  TabIndex        =   66
                  Top             =   240
                  Width           =   1080
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
                  TabIndex        =   65
                  Top             =   240
                  Width           =   615
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
                  TabIndex        =   64
                  Top             =   240
                  Width           =   840
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Õœœ ”‰Ê«  «·„ﬁ«—‰…"
               Height          =   1560
               Left            =   19545
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   405
               Width           =   5295
               Begin VSFlex8Ctl.VSFlexGrid GridIntervals1 
                  Height          =   915
                  Left            =   120
                  TabIndex        =   57
                  Top             =   240
                  Width           =   4545
                  _cx             =   8017
                  _cy             =   1614
                  Appearance      =   2
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEstimations2.frx":0CBE
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
            Begin VB.Frame Frame6 
               Caption         =   "Õœœ «·„Ê«“‰«  «·”«»ﬁ…"
               Height          =   1560
               Left            =   20025
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   405
               Width           =   8475
               Begin VSFlex8Ctl.VSFlexGrid GridOldEstimation 
                  Height          =   915
                  Left            =   120
                  TabIndex        =   55
                  Top             =   240
                  Width           =   8265
                  _cx             =   14579
                  _cy             =   1614
                  Appearance      =   2
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  FormatString    =   $"FrmEstimations2.frx":0DA3
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
            Begin VB.TextBox Percentage 
               Alignment       =   1  'Right Justify
               Height          =   435
               Left            =   19725
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Text            =   "0"
               Top             =   2070
               Width           =   1095
            End
            Begin VB.ComboBox OperatorsID 
               Height          =   330
               ItemData        =   "FrmEstimations2.frx":0E41
               Left            =   19215
               List            =   "FrmEstimations2.frx":0E51
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Text            =   " "
               Top             =   2070
               Width           =   1365
            End
            Begin VB.OptionButton OptAlarms 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ–Ì— ›ﬁÿ"
               Height          =   315
               Index           =   0
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   -4410
               Width           =   1425
            End
            Begin VB.OptionButton OptAlarms 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«› «·Õ”«»"
               Height          =   315
               Index           =   1
               Left            =   5430
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   -4410
               Width           =   1860
            End
            Begin VB.Frame Frame7 
               Caption         =   "«œŒ«· «·”‰Ê«  «·„«÷Ì…"
               Height          =   780
               Left            =   -5550
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1680
               Width           =   3930
               Begin VB.OptionButton OptActual 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÌœÊÌ"
                  Height          =   195
                  Index           =   0
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   240
                  Width           =   1305
               End
               Begin VB.OptionButton OptActual 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·Ì"
                  Height          =   195
                  Index           =   1
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   240
                  Width           =   1305
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Õœœ «· «—ÌŒ"
               Height          =   1095
               Left            =   12360
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   405
               Width           =   6120
               Begin VB.CheckBox Check17 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  Height          =   195
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   840
                  Width           =   1695
               End
               Begin MSComCtl2.DTPicker Fromdate 
                  Height          =   330
                  Left            =   3135
                  TabIndex        =   40
                  Top             =   240
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   210173953
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker todate 
                  Height          =   330
                  Left            =   840
                  TabIndex        =   41
                  Top             =   240
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   210173953
                  CurrentDate     =   41640
               End
               Begin Dynamic_Byte.NourHijriCal Fromdate√H 
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   42
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   510
                  Index           =   9
                  Left            =   120
                  TabIndex        =   43
                  Top             =   480
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   900
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
                  FontName        =   "Arial"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmEstimations2.frx":0E6D
                  DrawFocusRectangle=   0   'False
               End
               Begin Dynamic_Byte.NourHijriCal todateH 
                  Height          =   255
                  Left            =   840
                  TabIndex        =   44
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈«·Ï"
                  Height          =   435
                  Index           =   14
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   240
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·› —… „‰"
                  Height          =   315
                  Index           =   0
                  Left            =   4980
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   240
                  Width           =   945
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "«Ã„«·Ì« "
               Height          =   1065
               Left            =   -1035
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   7635
               Width           =   12255
               Begin VB.TextBox TxtTotalContract 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   10320
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtInsuranceValue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   6240
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtWater 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtElectricity 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   2160
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   360
                  Width           =   945
               End
               Begin VB.TextBox TxtCommiValue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   8280
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtPhone 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   360
                  Width           =   945
               End
               Begin VB.TextBox TxtTotalTo 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   10320
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   720
                  Width           =   1065
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·«ÌÃ«—"
                  Height          =   195
                  Index           =   6
                  Left            =   11505
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   360
                  Width           =   990
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " √„Ì‰"
                  Height          =   195
                  Index           =   19
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   360
                  Width           =   510
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„Ì«Â"
                  Height          =   195
                  Index           =   20
                  Left            =   5385
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   480
                  Width           =   750
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬂÂ—»«¡"
                  Height          =   195
                  Index           =   21
                  Left            =   2985
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   480
                  Width           =   750
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄Ì/—”Ê„"
                  Height          =   405
                  Index           =   25
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   360
                  Width           =   810
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œœ„« "
                  Height          =   195
                  Index           =   27
                  Left            =   1035
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   480
                  Width           =   990
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„” Õﬁ ··€Ì—"
                  Height          =   195
                  Index           =   1
                  Left            =   11520
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   720
                  Width           =   990
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "»Ì«‰«  „Õ«”»Ì…"
               Height          =   825
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   6570
               Width           =   4830
               Begin VB.CommandButton Command9 
                  Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                  Height          =   375
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   195
                  Index           =   35
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   360
                  Width           =   990
               End
            End
            Begin MSDataListLib.DataCombo DCAccountMaster 
               Height          =   330
               Left            =   22875
               TabIndex        =   68
               Top             =   630
               Width           =   6405
               _ExtentX        =   11298
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   4905
               Left            =   20700
               TabIndex        =   75
               Top             =   2550
               Width           =   18495
               _cx             =   32623
               _cy             =   8652
               Appearance      =   2
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Cols            =   28
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEstimations2.frx":1207
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
               Height          =   255
               Left            =   12510
               TabIndex        =   76
               Top             =   0
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   450
               _Version        =   393216
               Format          =   210239489
               CurrentDate     =   41640
            End
            Begin MSDataListLib.DataCombo DcBranch 
               Height          =   330
               Left            =   5010
               TabIndex        =   81
               Top             =   0
               Width           =   6270
               _ExtentX        =   11060
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
               Height          =   4725
               Left            =   390
               TabIndex        =   82
               Top             =   1620
               Width           =   18390
               _cx             =   32438
               _cy             =   8334
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   12
               Cols            =   53
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEstimations2.frx":163E
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
            Begin Dynamic_Byte.NourHijriCal recordDateH 
               Height          =   285
               Left            =   13875
               TabIndex        =   83
               Top             =   0
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
            End
            Begin ImpulseButton.ISButton CmdPrint 
               Height          =   345
               Left            =   6720
               TabIndex        =   106
               Top             =   6720
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄Â «·‘«‘…"
               BackColor       =   14871017
               FontName        =   "Arial"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmEstimations2.frx":1E6F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSComDlg.CommonDialog cd 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   510
               Index           =   7
               Left            =   16710
               TabIndex        =   113
               Top             =   6630
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   900
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               FontName        =   "Arial"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmEstimations2.frx":2209
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   510
               Index           =   8
               Left            =   15660
               TabIndex        =   114
               Top             =   6630
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   900
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               FontName        =   "Arial"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmEstimations2.frx":25A3
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "SMS"
               Height          =   360
               Index           =   1
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   1080
               Width           =   945
            End
            Begin VB.Shape Shape1 
               BorderWidth     =   2
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   1425
               Left            =   195
               Top             =   135
               Width           =   4545
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   510
               Left            =   16995
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   1185
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ—ﬂ…"
               Height          =   315
               Index           =   7
               Left            =   16905
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   0
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Õ—ﬂ…"
               Height          =   315
               Index           =   8
               Left            =   15150
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   0
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «· Ê“Ì⁄"
               Height          =   330
               Index           =   3
               Left            =   19965
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1560
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·› —… „‰ "
               Height          =   495
               Index           =   4
               Left            =   18990
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   630
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   360
               Index           =   2
               Left            =   11220
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   510
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Â–… «·‘«‘…  ﬁÊ„ »«À»«  «” Õﬁ«ﬁ «·œ›⁄«  «·„” Õﬁ…"
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
               Height          =   1200
               Index           =   38
               Left            =   225
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   270
               Width           =   4275
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
               Height          =   285
               Index           =   37
               Left            =   -1395
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   1950
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   210
               Index           =   13
               Left            =   11220
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   0
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «· ﬁœÌ— „ Ê”ÿ „«”»ﬁ"
               Height          =   465
               Index           =   15
               Left            =   19200
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   2070
               Width           =   2355
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‰”»…"
               Height          =   390
               Index           =   0
               Left            =   21030
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   2070
               Width           =   885
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   390
               Left            =   11280
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   -2610
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰œ „Œ«·›… «· ﬁœÌ—Ï"
               ForeColor       =   &H000000FF&
               Height          =   480
               Index           =   16
               Left            =   8385
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   -4410
               Width           =   2370
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   885
         Left            =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   8640
         Width           =   18900
         _cx             =   33338
         _cy             =   1561
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   315
            Left            =   11220
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   75
            Visible         =   0   'False
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
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
            ButtonImage     =   "FrmEstimations2.frx":293D
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   300
            Left            =   12195
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   420
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
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
            ButtonImage     =   "FrmEstimations2.frx":2CD7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   465
            Index           =   0
            Left            =   10470
            TabIndex        =   8
            Top             =   465
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Height          =   465
            Index           =   1
            Left            =   9630
            TabIndex        =   9
            Top             =   465
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Height          =   465
            Index           =   2
            Left            =   8835
            TabIndex        =   10
            Top             =   465
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Height          =   465
            Index           =   3
            Left            =   7890
            TabIndex        =   11
            Top             =   465
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Height          =   465
            Index           =   4
            Left            =   6915
            TabIndex        =   12
            Top             =   465
            Visible         =   0   'False
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            Enabled         =   0   'False
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Height          =   465
            Index           =   6
            Left            =   4980
            TabIndex        =   13
            Top             =   465
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Height          =   465
            Index           =   5
            Left            =   6045
            TabIndex        =   14
            Top             =   465
            Visible         =   0   'False
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Height          =   345
            Left            =   15630
            TabIndex        =   15
            Tag             =   "Delete Row"
            Top             =   0
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   609
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            MICON           =   "FrmEstimations2.frx":3071
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   345
            Left            =   2880
            TabIndex        =   107
            Top             =   480
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â  "
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEstimations2.frx":308D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker txtFromDate 
            Height          =   255
            Left            =   16560
            TabIndex        =   115
            Top             =   510
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   450
            _Version        =   393216
            Format          =   209911809
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker txtToDate 
            Height          =   255
            Left            =   15090
            TabIndex        =   116
            Top             =   510
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   450
            _Version        =   393216
            Format          =   209911809
            CurrentDate     =   41640
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   225
            Left            =   1470
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì"
            Height          =   225
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   225
            Width           =   930
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   450
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   210
            Width           =   975
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   225
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   765
         Index           =   5
         Left            =   0
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   0
         Width           =   19035
         _cx             =   33576
         _cy             =   1349
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
         Picture         =   "FrmEstimations2.frx":3427
         Caption         =   "   ‘«‘… «·œ›⁄«  «·„” Õﬁ…   "
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
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   " Õ„Ì· Õ Ì «·⁄ﬁÊœ «·„’›«Â"
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   240
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.TextBox TxtRowNumber 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Text            =   "Text4"
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6780
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Text            =   "Text1"
            Top             =   180
            Visible         =   0   'False
            Width           =   375
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1695
            TabIndex        =   100
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEstimations2.frx":4101
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
            TabIndex        =   101
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEstimations2.frx":449B
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
            TabIndex        =   102
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEstimations2.frx":4835
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
            TabIndex        =   103
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEstimations2.frx":4BCF
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   6120
            Picture         =   "FrmEstimations2.frx":4F69
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "⁄—÷"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmEstimations2.frx":8BD1
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmAllocationToContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
    Dim rsV As ADODB.Recordset
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim printSQL  As String

Dim StrSQL  As String
Dim Account_Code_dynamic80 As String
Dim Account_Code_dynamic81 As String
Dim Account_Code_dynamic82 As String
Dim Account_Code_dynamic83 As String
Dim Account_Code_dynamic84 As String
Dim Account_Code_dynamic85 As String
 Dim Account_Code_dynamic123 As String
 Dim Account_Code_dynamic125 As String
 
Dim Account_Code_dynamic153 As String
 Dim Account_Code_dynamic154 As String
 Dim Account_Code_dynamic155 As String
 Dim Account_Code_dynamic156 As String
 
 Dim vaTAccount As String
 Dim vataccount2 As String

Dim rs As ADODB.Recordset


Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Private Sub Del_Trans()
    Dim Msg As String
    On Error GoTo ErrTrap
   Dim StrSQL As String
   Dim rsDummy As ADODB.Recordset
   Set rsDummy = New ADODB.Recordset
   StrSQL = "select installid from tblContractInsAllocationsDetails where transid=" & TxtTransID & " and ISNULL(zatcaStatus, 0) = 1"
   rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
   If Not rsDummy.EOF Then
        MsgBox "·«Ì„ﬂ‰ Õ–› «·”Ã· ‰Ÿ—« ·ÊÃÊœ ›Ê« Ì—  „ —›⁄Â« ··ÂÌ∆…"
        Exit Sub
   End If
   rsDummy.Close
    If TxtTransID.Text <> "" Then
     Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtTransID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                 
         StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
                   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
     Cn.Execute " update  TblContractInstallments set  allocations=0 where id in( " & " select installid from tblContractInsAllocationsDetails where transid=" & TxtTransID & ")"
                rs.delete
             
        
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                
                    TxtModFlg_Change
                    LabCurrRec.Caption = 0
                    LabCountRec.Caption = 0
                Else
                    clear_all Me
                    Retrive
                End If

                '--------
            
                '-------
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub




Private Sub BtnPrint_Click()
 printAnyreport printSQL, Me.Name, "«·œ›⁄«  «·„” Õﬁ… "
End Sub

'
'Private Sub ALLButton1_Click()
'    FrmShowCol1.show
'End Sub









 

'Private Sub CboYear_Click()
'    CmdOk_Click
'End Sub

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.GridInstallments
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = True
            Next i

        End With

    Else

        With Me.GridInstallments

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = False
            Next i

        End With

    End If

ReLineGrid
End Sub







Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
  Dim RsDev2 As ADODB.Recordset
  
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
 

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
           Me.TxtTransID.Text = CStr(new_id("tblContractInsAllocations", "transID", "", True))
    ElseIf Me.TxtModFlg.Text = "E" Then
          Cn.Execute "delete tblContractInsAllocationsDetails where transID=" & val(Me.TxtTransID.Text)
          Cn.Execute "delete tblContractInsAllocationsDetails1 where transID=" & val(Me.TxtTransID.Text)
      StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords


    End If
    
  rs("transID").value = TxtTransID.Text
    rs("recordDate").value = XPDtbTrans.value
    rs("RecorddateH").value = recordDateH.value
     rs("Fromdate").value = Fromdate.value
           rs("todate").value = todate.value
       rs("Fromdateh").value = ToHijriDate(Fromdate.value)
           rs("todateh").value = ToHijriDate(todate.value)
        rs("BranchId").value = IIf(Me.Dcbranch.BoundText = "", Null, val(Me.Dcbranch.BoundText))
  
      
    rs("Remarks").value = IIf(Me.tXtRemarks.Text = "", "", Me.tXtRemarks.Text)
   

    rs.update
    
 
    Set RsDetails1 = New ADODB.Recordset
 
           StrSQL = "SELECT  *  from dbo.tblContractInsAllocationsDetails Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Dim i As Integer
      Dim InvoiceTypeCodeID As Integer
      Dim PaymentMeansCode As String
       Dim paymentnote
       Dim Export
    Dim SerialMap As Object
Set SerialMap = CreateObject("Scripting.Dictionary")


    With Me.GridInstallments
'Selected
        For i = 1 To .rows - 1
   
        If val(.TextMatrix(i, .ColIndex("value"))) <> 0 And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
            RsDetails1.AddNew
            RsDetails1("transID").value = Me.TxtTransID.Text
            RsDetails1("VATValue").value = val(.TextMatrix(i, .ColIndex("VATValue")))
            RsDetails1("hijri").value = (.TextMatrix(i, .ColIndex("hijri")))
            RsDetails1("NoteSerial").value = (.TextMatrix(i, .ColIndex("NoteSerial1")))
            RsDetails1("ContNo").value = val(.TextMatrix(i, .ColIndex("ContNo")))
            
            RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
            RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
            RsDetails1("Installdate").value = .TextMatrix(i, .ColIndex("Due_Date"))
            RsDetails1("InstalldateH").value = .TextMatrix(i, .ColIndex("Due_DateH"))
             
'            If Trim(.TextMatrix(i, .ColIndex("NoteSerial1H"))) = "" Then
'                RsDetails1("NoteSerial1H").value = Voucher_coding(val(dcBranch.BoundText), .TextMatrix(i, .ColIndex("Due_Date")), 85, 5002)
'                .TextMatrix(i, .ColIndex("NoteSerial1H")) = RsDetails1("NoteSerial1H").value & ""
'            End If
            '.TextMatrix(i, .ColIndex("NoteSerial1")) & .TextMatrix(i, .ColIndex("Installid"))
            Dim BranchID As Integer
Dim dueDate As Date
Dim SerialKey As String

BranchID = val(Dcbranch.BoundText)
dueDate = .TextMatrix(i, .ColIndex("Due_Date"))
SerialKey = BranchID & Format(dueDate, "yyMM")  ' „À«·: 1242412

    If Trim(.TextMatrix(i, .ColIndex("NoteSerial1H"))) = "" Then
        If Not SerialMap.Exists(SerialKey) Then
            ' √Ê· „—… ‰ﬁ«»· «· «—ÌŒ Ê«·›—⁄ œÂ° ‰Ê·œ √Ê· —ﬁ„ ”Ì—Ì«· »«” Œœ«„ «·›‰ﬂ‘‰
            SerialMap(SerialKey) = Voucher_coding(BranchID, dueDate, 85, 5002, , , , , , , "")
        End If

    ' ‰” Œœ„ «·—ﬁ„ Ê‰Œ“¯‰Â
    RsDetails1("NoteSerial1H").value = SerialMap(SerialKey)
    .TextMatrix(i, .ColIndex("NoteSerial1H")) = CStr(SerialMap(SerialKey))

    ' ‰⁄œ «·”Ì—Ì«· œ«Œ· ‰›” «·„Ã„Ê⁄…
    SerialMap(SerialKey) = SerialMap(SerialKey) + 1
End If
RsDetails1("NoteSerial1H").value = .TextMatrix(i, .ColIndex("NoteSerial1H"))
           
           
               
            RsDetails1("CIBAN").value = ""
            'vat data
            RsDetails1("RecTime").value = Time
            
            RsDetails1("tablename").value = "tblContractInsAllocationsDetails"
            
  
            InvoiceTypeCodeID = 388
            
            RsDetails1("InvoiceTypeCodeID").value = InvoiceTypeCodeID
            
            
            
            If val(val(.TextMatrix(i, .ColIndex("value")))) >= 1000 Then
            
            
                If Export = 1 Then
                    RsDetails1("InvoiceTypeCodename").value = "0100100"
                Else
                    RsDetails1("InvoiceTypeCodename").value = "0100000"
                End If
            
            Else
                RsDetails1("InvoiceTypeCodename").value = "0200000"
            End If
            
            RsDetails1("DocumentCurrencyCode").value = "SAR"
            RsDetails1("TaxCurrencyCode").value = "SAR"
            RsDetails1("ActualDeliveryDate").value = Date
            RsDetails1("LatestDeliveryDate").value = Date

         
            '10 In cash
            '30 Credit
            '42 Payment to bank account
            '48 Bank card
            '1 Instrument not defined(Free text)
           
'        If CboPayMentType.ListIndex = 0 Then ' ‰ﬁœ«
'                  PaymentMeansCode = "10"
'                      paymentnote = "Payment by Cash"
'        ElseIf CboPayMentType.ListIndex = 1 Then ' ¬Ã·
'                 PaymentMeansCode = "30"
'                 paymentnote = "Payment by Credit"
'         ElseIf CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then  '  ÕÊÌ· »‰ﬂÌ
'                    If SystemOptions.AllowSalesMultyPayed = True Then
'                     PaymentMeansCode = "48" 'ﬂ«— 
'                      paymentnote = "Payment by Bank Card"
'                    Else
'                    PaymentMeansCode = "42" '»‰ﬂ Õ”«»
'                    paymentnote = "Payment by Bank Account"
'                    End If
'
'         End If
                PaymentMeansCode = "30"
                paymentnote = "Payment by Credit"
                RsDetails1("PaymentMeansCode").value = PaymentMeansCode
                
                RsDetails1("paymentnote").value = paymentnote
                
                
                RsDetails1("nextinstalldate").value = IIf(.TextMatrix(i, .ColIndex("nextinstalldate")) = "", Null, .TextMatrix(i, .ColIndex("nextinstalldate")))
                RsDetails1("nextinstalldateH").value = .TextMatrix(i, .ColIndex("nextinstalldateH"))
                
                
                RsDetails1("installValue").value = val(.TextMatrix(i, .ColIndex("value")))
                RsDetails1("RentValue").value = val(.TextMatrix(i, .ColIndex("RentValue")))
                RsDetails1("Commissions").value = val(.TextMatrix(i, .ColIndex("Commissions")))
                RsDetails1("Insurance").value = val(.TextMatrix(i, .ColIndex("Insurance")))
                RsDetails1("Water").value = val(.TextMatrix(i, .ColIndex("Water")))
                RsDetails1("Electric").value = val(.TextMatrix(i, .ColIndex("Electric")))
                RsDetails1("TelandNet").value = val(.TextMatrix(i, .ColIndex("TelandNet")))
                RsDetails1("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
                RsDetails1("Countsofall").value = val(.TextMatrix(i, .ColIndex("Countsofall")))
                RsDetails1("Iqar").value = val(.TextMatrix(i, .ColIndex("Iqar")))
                RsDetails1("commisiontype").value = val(.TextMatrix(i, .ColIndex("commisiontype")))
                
                RsDetails1("AmolaValus").value = val(.TextMatrix(i, .ColIndex("AmolaValus")))
                RsDetails1("ownerid").value = val(.TextMatrix(i, .ColIndex("ownerid")))
                

        
                RsDetails1("allocations").value = 1
                Cn.Execute " update  TblContractInstallments set  allocations=1 ,NoteSerial1H = '" & Trim(.TextMatrix(i, .ColIndex("NoteSerial1H"))) & "' where id=" & val(.TextMatrix(i, .ColIndex("Installid")))

                        '     If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
'                            RsDetails1("Select").value = 1
'                            RsDetails1("allocations").value = 1
'                            Cn.Execute " update  TblContractInstallments set  allocations=1 where id=" & val(.TextMatrix(i, .ColIndex("Installid")))
                        '    Else
'                            RsDetails1("Select").value = 0
                        '    RsDetails1("allocations").value = 0
                        '    Cn.Execute " update  TblContractInstallments set  allocations=0 where id=" & val(.TextMatrix(i, .ColIndex("Installid")))
                        '    End If
           
            RsDetails1.update
        Else
       
                Cn.Execute " update  TblContractInstallments set  allocations=0 where id=" & val(.TextMatrix(i, .ColIndex("Installid")))
        End If
           Next i
        RsDetails1.Close
    End With
    
 
''***********************************************************************************************
'   Set RsDetails1 = New ADODB.Recordset
'
'           StrSQL = "SELECT  *  from dbo.tblContractInsAllocationsDetails1 Where (1 = -1)"
'   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'
'      Dim Countsofall As Double
'      Dim j As Integer
'          Dim SngAllValue As Single
'
'    Dim IntNoOFQast As Integer
'    Dim IntRes As Integer
'    Dim SngOnePor As Single
'    Dim FirstDate As Date
'    Dim PreDate As Date
'    Dim NewDate As Date
'    Dim DateInterval As String
'        Dim NewDateH As String
'           Dim PreDateH As String
'           Dim hijriorJerojian As Integer
'            Dim LastDate As Date
'            Dim LastDateH As String
'                    Dim FirstDate1 As Date
'            Dim FirstDateH1 As String
'
'    Dim DateNumber As Integer
'
'Dim watervalue As Double
'Dim Electricity As Double
'Dim noOfRemaindays As Integer
'Dim noOfRemaindays1 As Integer
'Dim endpartdays As Integer
'Dim MonthLastDay1 As Date
'Dim onedayvale As Double
'     Dim onedayRentValue As Double
'   Dim onedayCommissions As Double
'Dim onedayInsurance As Double
'Dim onedayWater As Double
'Dim onedayElectric As Double
'Dim onedayTelandNet As Double
'
'    With Me.GridInstallments
''Selected
'                       For i = 1 To .rows - 1
'
'                       If val(.TextMatrix(i, .ColIndex("value"))) <> 0 And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("commisiontype"))) <> 1 Then
'               .TextMatrix(i, .ColIndex("Due_DateH")) = ToHijriDate(.TextMatrix(i, .ColIndex("Due_Date")))
'                    Countsofall = val(.TextMatrix(i, .ColIndex("Countsofall")))
'                    VBA.Calendar = vbCalGreg
'             '       LastDate = DateAdd("M", Countsofall, (.TextMatrix(i, .ColIndex("Due_Date"))))
'            '      VBA.Calendar = vbCalHijri
'            '       LastDateH = DateAdd("M", Countsofall, (.TextMatrix(i, .ColIndex("Due_DateH"))))
'
'                   LastDate = DateAdd("M", Countsofall, (.TextMatrix(i, .ColIndex("Due_Date"))))
'                     LastDate = DateAdd("d", -1, LastDate)
'                  VBA.Calendar = vbCalHijri
'                   LastDateH = DateAdd("M", Countsofall, (.TextMatrix(i, .ColIndex("Due_DateH"))))
'                      LastDateH = DateAdd("d", -1, LastDateH)
'
'
'
'              '«· √ﬂœ «‰ «· «—ÌŒ ·Ì” «Ê· «·‘Â—
'
'                      hijriorJerojian = 1
'                       If hijriorJerojian = 1 Then 'jorjian
'                         VBA.Calendar = vbCalGreg
'                  FirstDate1 = dhFirstDayInMonth(.TextMatrix(i, .ColIndex("Due_Date")))
'                       noOfRemaindays1 = DateDiff("D", .TextMatrix(i, .ColIndex("Due_Date")), FirstDate1)
'                   Else
'                   VBA.Calendar = vbCalHijri
'
'                  FirstDateH1 = dhFirstDayInMonth(.TextMatrix(i, .ColIndex("Due_DateH")))
'                   noOfRemaindays1 = DateDiff("D", .TextMatrix(i, .ColIndex("Due_DateH")), FirstDateH1)
'                   End If
'
'                   If noOfRemaindays1 = 0 Then GoTo ll
'
'                       hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
'                  hijriorJerojian = 1
'                       If hijriorJerojian = 1 Then 'jorjian
'
'                         VBA.Calendar = vbCalGreg
'                      noOfRemaindays = DateDiff("D", .TextMatrix(i, .ColIndex("Due_Date")), MonthLastDay(.TextMatrix(i, .ColIndex("Due_Date"))))
'
'                      Else
'
'                     VBA.Calendar = vbCalHijri
'                      noOfRemaindays = DateDiff("D", .TextMatrix(i, .ColIndex("Due_DateH")), MonthLastDay(.TextMatrix(i, .ColIndex("Due_DateH"))))
'
'                      End If
'                      If noOfRemaindays > 0 Then
'                      Countsofall = Countsofall - 1
'                      End If
'                        endpartdays = 30 - noOfRemaindays
'
'                      onedayvale = val(.TextMatrix(i, .ColIndex("value"))) / IIf(val(.TextMatrix(i, .ColIndex("Countsofall"))) = 0, 1, val(.TextMatrix(i, .ColIndex("Countsofall")))) / 30
'                       onedayRentValue = val(.TextMatrix(i, .ColIndex("RentValue"))) / IIf(val(.TextMatrix(i, .ColIndex("Countsofall"))) = 0, 1, val(.TextMatrix(i, .ColIndex("Countsofall")))) / 30
'                      onedayCommissions = val(.TextMatrix(i, .ColIndex("Commissions"))) / IIf(val(.TextMatrix(i, .ColIndex("Countsofall"))) = 0, 1, val(.TextMatrix(i, .ColIndex("Countsofall")))) / 30
'                       onedayInsurance = val(.TextMatrix(i, .ColIndex("Insurance"))) / IIf(val(.TextMatrix(i, .ColIndex("Countsofall"))) = 0, 1, val(.TextMatrix(i, .ColIndex("Countsofall")))) / 30
'                       onedayWater = val(.TextMatrix(i, .ColIndex("Water"))) / IIf(val(.TextMatrix(i, .ColIndex("Countsofall"))) = 0, 1, val(.TextMatrix(i, .ColIndex("Countsofall")))) / 30
'                       onedayElectric = val(.TextMatrix(i, .ColIndex("Electric"))) / IIf(val(.TextMatrix(i, .ColIndex("Countsofall"))) = 0, 1, val(.TextMatrix(i, .ColIndex("Countsofall")))) / 30
'                       onedayTelandNet = val(.TextMatrix(i, .ColIndex("TelandNet"))) / IIf(val(.TextMatrix(i, .ColIndex("Countsofall"))) = 0, 1, val(.TextMatrix(i, .ColIndex("Countsofall")))) / 30
'
'
'                  '*****************part one of month
'                        If noOfRemaindays > 0 Then
'                       VBA.Calendar = vbCalGreg
'                            NewDate = (.TextMatrix(i, .ColIndex("Due_Date")))
'                   NewDateH = Format((.TextMatrix(i, .ColIndex("Due_DateH"))), "DD/MM/YYYY")
'
'                     RsDetails1.AddNew
'                     RsDetails1("transID").value = Me.TxtTransID.text
'                   RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
'                   hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
'         hijriorJerojian = 1
'                       RsDetails1("NoteSerial").value = (.TextMatrix(i, .ColIndex("NoteSerial1")))
'                         RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
'
'                         RsDetails1("Installdate").value = (NewDate)
'                         RsDetails1("InstalldateH").value = NewDateH
'                         RsDetails1("hijri").value = val(.TextMatrix(i, .ColIndex("hijri")))
'                        RsDetails1("installValue").value = Round(onedayvale * noOfRemaindays, 2)
'                   Dim commission As Double
'                      If val(.TextMatrix(i, .ColIndex("AmolaValus"))) > 0 Then
'
'                      commission = onedayRentValue * val(.TextMatrix(i, .ColIndex("AmolaValus"))) / 100
'                        RsDetails1("RentValue").value = Round(commission * noOfRemaindays, 2)
'                        RsDetails1("Commission").value = 1
'                    Else
'
'                    RsDetails1("RentValue").value = Round(onedayRentValue * noOfRemaindays, 2)
'                     RsDetails1("Commission").value = 0
'                      End If
'
'
'                         RsDetails1("RentValue").value = Round(onedayRentValue * noOfRemaindays, 2)
'                         RsDetails1("Commissions").value = Round(onedayCommissions * noOfRemaindays, 2)
'                         RsDetails1("Insurance").value = Round(onedayInsurance * noOfRemaindays, 2)
'                         RsDetails1("Water").value = Round(onedayWater * noOfRemaindays, 2)
'                         RsDetails1("Electric").value = Round(onedayElectric * noOfRemaindays, 2)
'                         RsDetails1("TelandNet").value = Round(onedayTelandNet * noOfRemaindays, 2)
'                        RsDetails1("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
'
'
'                          RsDetails1.update
'
'                         End If
'                 '***********************end of first part*******************************
'
'                   VBA.Calendar = vbCalGreg
'                           NewDate = MonthLastDay(.TextMatrix(i, .ColIndex("Due_Date")))
'                           VBA.Calendar = vbCalHijri
'                   NewDateH = MonthLastDay(.TextMatrix(i, .ColIndex("Due_DateH")))
'
'                             VBA.Calendar = vbCalGreg
'
'                           NewDate = DateAdd("D", 1, NewDate)
'                           VBA.Calendar = vbCalHijri
'                   NewDateH = DateAdd("D", 1, NewDateH)
'll:
'      If noOfRemaindays = 0 Then
'                       VBA.Calendar = vbCalGreg
'                            NewDate = (.TextMatrix(i, .ColIndex("Due_Date")))
'                   NewDateH = Format((.TextMatrix(i, .ColIndex("Due_DateH"))), "DD/MM/YYYY")
'         End If
'
'                                 For j = 1 To Countsofall
'
'                                  RsDetails1.AddNew
'                                 RsDetails1("transID").value = Me.TxtTransID.text
'                              RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
'                              hijriorJerojian = 1
'                  '   hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
'
'
'                               If j = 1 Then
'
'                        Else
'                              VBA.Calendar = vbCalGreg
'                            PreDate = NewDate
'
'                            If hijriorJerojian = 1 Then 'jorijan
'                            VBA.Calendar = vbCalGreg
'                            NewDate = DateAdd("m", 1, NewDate)
'                            NewDateH = ToHijriDate(NewDate)
'                            End If
'
'                                 PreDateH = NewDateH
'
'            If hijriorJerojian = 0 Then 'hijri
'            VBA.Calendar = vbCalHijri
'                            NewDateH = (DateAdd("m", 1, NewDateH))
'                              VBA.Calendar = vbCalGreg
'            NewDate = ToGregorianDate(NewDateH)
'            End If
'
'            End If
'                                     RsDetails1("NoteSerial").value = (.TextMatrix(i, .ColIndex("NoteSerial1")))
'                                     RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
'                                           VBA.Calendar = vbCalGreg
'                                     RsDetails1("Installdate").value = (NewDate)
'                                     RsDetails1("InstalldateH").value = NewDateH
'                                     RsDetails1("hijri").value = val(.TextMatrix(i, .ColIndex("hijri")))
'                                    RsDetails1("installValue").value = Round(val(.TextMatrix(i, .ColIndex("value"))) / val(.TextMatrix(i, .ColIndex("Countsofall"))), 2)
'                                     RsDetails1("RentValue").value = Round(val(.TextMatrix(i, .ColIndex("RentValue"))) / val(.TextMatrix(i, .ColIndex("Countsofall"))), 2)
'                                     RsDetails1("Commissions").value = Round(val(.TextMatrix(i, .ColIndex("Commissions"))) / val(.TextMatrix(i, .ColIndex("Countsofall"))), 2)
'                                     RsDetails1("Insurance").value = Round(val(.TextMatrix(i, .ColIndex("Insurance"))) / val(.TextMatrix(i, .ColIndex("Countsofall"))), 2)
'                                     RsDetails1("Water").value = Round(val(.TextMatrix(i, .ColIndex("Water"))) / val(.TextMatrix(i, .ColIndex("Countsofall"))), 2)
'                                     RsDetails1("Electric").value = Round(val(.TextMatrix(i, .ColIndex("Electric"))) / val(.TextMatrix(i, .ColIndex("Countsofall"))), 2)
'                                     RsDetails1("TelandNet").value = Round(val(.TextMatrix(i, .ColIndex("TelandNet"))) / val(.TextMatrix(i, .ColIndex("Countsofall"))), 2)
'                                    RsDetails1("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
'
'                                      RsDetails1.update
'                               Next j
'
'
'
'       '*****************  Last part of month
'                     If noOfRemaindays1 = 0 Then GoTo xx
'                        If noOfRemaindays > 0 Then
'
'          If hijriorJerojian = 1 Then ' jorjia then
'          VBA.Calendar = vbCalGreg
'                    NewDate = DateAdd("m", 1, NewDate)
'                    NewDateH = ToHijriDate(NewDate)
'           Else
'            VBA.Calendar = vbCalHijri
'               NewDateH = DateAdd("m", 1, NewDateH)
'                VBA.Calendar = vbCalGreg
'                    NewDate = ToGregorianDate(NewDateH)
'
'           End If
'
'
'
'        '   Calendar = vbCalGreg
'                    '        NewDateH = ToHijriDate(NewDate)
'
'                           If hijriorJerojian = 1 Then 'jorjian
'                           VBA.Calendar = vbCalGreg
'                      noOfRemaindays = DateDiff("D", NewDate, LastDate)
'                      Else
'                      VBA.Calendar = vbCalHijri
'                      noOfRemaindays = DateDiff("D", NewDateH, LastDateH)
'                      End If
'                      noOfRemaindays = noOfRemaindays + 1
'                     RsDetails1.AddNew
'                     RsDetails1("transID").value = Me.TxtTransID.text
'                   RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
'                   hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
'
'                        RsDetails1("NoteSerial").value = (.TextMatrix(i, .ColIndex("NoteSerial1")))
'                         RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
'                             VBA.Calendar = vbCalGreg
'                         RsDetails1("Installdate").value = NewDate
'                         RsDetails1("InstalldateH").value = NewDateH
'                         RsDetails1("hijri").value = val(.TextMatrix(i, .ColIndex("hijri")))
'                        RsDetails1("installValue").value = Round(onedayvale * endpartdays, 2)
'                        RsDetails1("RentValue").value = Round(onedayRentValue * endpartdays, 2)
'                         RsDetails1("Commissions").value = Round(onedayCommissions * endpartdays, 2)
'                         RsDetails1("Insurance").value = Round(onedayInsurance * endpartdays, 2)
'                         RsDetails1("Water").value = Round(onedayWater * endpartdays, 2)
'                         RsDetails1("Electric").value = Round(onedayElectric * endpartdays, 2)
'                         RsDetails1("TelandNet").value = Round(onedayTelandNet * endpartdays, 2)
'                        RsDetails1("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
'                         RsDetails1.update
'
'                         End If
'                 '*********************************************************************
'xx:
'     Else
'
'                        '    Cn.Execute " update  TblContractInstallments set  allocations=0 where id=" & val(.TextMatrix(i, .ColIndex("Installid")))
'
'        End If
'           Next i
'        RsDetails1.Close
'    End With
 '***********************************************************************************************
' ≈⁄«œ… › Õ Recordset ·ÃœÊ· tblContractInsAllocationsDetails1 ·≈÷«›… «·”ÿÊ— «·„ﬁ”„… ⁄·Ï «·‘ÂÊ—/«·√Ì«„
Set RsDetails1 = New ADODB.Recordset
 
StrSQL = "SELECT * FROM dbo.tblContractInsAllocationsDetails1 WHERE (1 = -1)"
RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

'Dim i As Integer
Dim TotalValue As Double
Dim TotalRent As Double
Dim TotalCommissions As Double
Dim TotalInsurance As Double
Dim TotalWater As Double
Dim TotalElectric As Double
Dim TotalTelandNet As Double

Dim StartDate As Date
Dim EndDate As Date
Dim TotalDays As Long
Dim DayValue As Double
Dim DayRent As Double
Dim DayCommissions As Double
Dim DayInsurance As Double
Dim DayWater As Double
Dim DayElectric As Double
Dim DayTelandNet As Double

Dim CurrentStart As Date
Dim MonthEnd As Date
Dim PartEnd As Date
Dim DaysPart As Long
Dim Countsofall As Long

With Me.GridInstallments

    For i = 1 To .rows - 1

        ' ‰›” ‘—Êÿ «·«Œ Ì«— «·ﬁœÌ„…
        If val(.TextMatrix(i, .ColIndex("value"))) <> 0 _
           And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked _
           And val(.TextMatrix(i, .ColIndex("commisiontype"))) <> 1 Then

            ' ﬁ—«¡… «·ﬁÌ„ «·√’·Ì… „‰ «·Ã—Ìœ
            TotalValue = val(.TextMatrix(i, .ColIndex("value")))
            TotalRent = val(.TextMatrix(i, .ColIndex("RentValue")))
            TotalCommissions = val(.TextMatrix(i, .ColIndex("Commissions")))
            TotalInsurance = val(.TextMatrix(i, .ColIndex("Insurance")))
            TotalWater = val(.TextMatrix(i, .ColIndex("Water")))
            TotalElectric = val(.TextMatrix(i, .ColIndex("Electric")))
            TotalTelandNet = val(.TextMatrix(i, .ColIndex("TelandNet")))

            ' »œ«Ì… «·ﬁ”ÿ
            StartDate = .TextMatrix(i, .ColIndex("Due_Date"))
            ' ⁄œœ «·‘ÂÊ— „‰ «·Ã—Ìœ
            Countsofall = val(.TextMatrix(i, .ColIndex("Countsofall")))

            If Countsofall <= 0 Or TotalValue = 0 Then
                GoTo NextInstallment
            End If

            ' ‰Â«Ì… «·ﬁ”ÿ = »⁄œ Countsofall ‘ÂÊ— - ÌÊ„
            VBA.Calendar = vbCalGreg
            EndDate = DateAdd("m", Countsofall, StartDate)
            EndDate = DateAdd("d", -1, EndDate)

            ' ≈Ã„«·Ì ⁄œœ «·√Ì«„ ›Ì «·› —…
            TotalDays = DateDiff("d", StartDate, EndDate) + 1
            If TotalDays <= 0 Then
                GoTo NextInstallment
            End If

            ' ﬁÌ„… «·ÌÊ„ «·Ê«Õœ ·ﬂ· »‰œ
            DayValue = TotalValue / TotalDays
            DayRent = TotalRent / TotalDays
            DayCommissions = TotalCommissions / TotalDays
            DayInsurance = TotalInsurance / TotalDays
            DayWater = TotalWater / TotalDays
            DayElectric = TotalElectric / TotalDays
            DayTelandNet = TotalTelandNet / TotalDays

            ' ‰⁄œÌ ⁄·Ï «·› —… ‘Â— »‘Â—
            CurrentStart = StartDate

      Do While CurrentStart <= EndDate

    ' «·ÌÊ„ «·√Ê· ›Ì «·‘Â—
        Dim FirstOfMonth As Date
            FirstOfMonth = DateSerial(year(CurrentStart), Month(CurrentStart), 1)
        
            ' «·ÌÊ„ «·√ŒÌ— ›Ì «·‘Â—
            Dim LastOfMonth As Date
            LastOfMonth = DateSerial(year(CurrentStart), Month(CurrentStart) + 1, 0)
        
            ' «·Ã“¡ «·›⁄·Ì „‰ «·‘Â— œ«Œ· › —… «·ﬁ”ÿ
            Dim PeriodStart As Date: PeriodStart = CurrentStart
            Dim PeriodEnd As Date: PeriodEnd = IIf(LastOfMonth < EndDate, LastOfMonth, EndDate)
        
            DaysPart = DateDiff("d", PeriodStart, PeriodEnd) + 1
        
            RsDetails1.AddNew
              RsDetails1("transID").value = Me.TxtTransID.Text
              RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
              RsDetails1("NoteSerial").value = .TextMatrix(i, .ColIndex("NoteSerial1"))
              RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
        
              RsDetails1("Installdate").value = PeriodStart
              RsDetails1("InstalldateH").value = ToHijriDate(PeriodStart)
                                            RsDetails1("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
              RsDetails1("installValue").value = Round(DayValue * DaysPart, 2)
              RsDetails1("RentValue").value = Round(DayRent * DaysPart, 2)
              RsDetails1("Commissions").value = Round(DayCommissions * DaysPart, 2)
              RsDetails1("Insurance").value = Round(DayInsurance * DaysPart, 2)
              RsDetails1("Water").value = Round(DayWater * DaysPart, 2)
              RsDetails1("Electric").value = Round(DayElectric * DaysPart, 2)
              RsDetails1("TelandNet").value = Round(DayTelandNet * DaysPart, 2)
        
            RsDetails1.update
        
            ' «·«‰ ﬁ«· ·√Ê· ÌÊ„ ›Ì «·‘Â— «· «·Ì
            CurrentStart = DateAdd("m", 1, FirstOfMonth)

    Loop

NextInstallment:

        Else
            ' ·Ê „‘ „ ⁄·„ √Ê «·ﬁÌ„… = 0 √Ê ‰Ê⁄ «·⁄„Ê·… = 1° »‰⁄œ¯Ì
        End If

    Next i

End With

RsDetails1.Close
'***********************************************************************************************

'**********************************************************************************************
        createVoucher
    Cn.CommitTrans
    BeginTrans = False
  'VBA.Calendar = vbCalGreg
    
 
    Select Case Me.TxtModFlg.Text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Function addInterval()
 
 
End Function

 Private Sub Cmd_Click(Index As Integer)
    'On Error GoTo ErrTrap
    Select Case Index
    Dim X As Integer

Case 8
SaveAlloca
Case 7
FillGrid2
Case 9
If Me.TxtModFlg.Text = "E" Then
X = MsgBox("”Ì „ «·€«¡ «· Œ’Ì’ «·Õ«·Ì", vbCritical + vbOKCancel)
            If X = vbOK Then
                 Cn.Execute " update  TblContractInstallments set  allocations=0 where id in( " & " select installid from tblContractInsAllocationsDetails where transid=" & TxtTransID & ")"
        Else
        
        Exit Sub
            End If
End If



FillGrid

        Case 0
 
            TxtModFlg.Text = "N"
            clear_all Me
        OperatorsID.ListIndex = 0
       OptAlarms(0).value = True
       OptActual(1).value = True
            Me.XPDtbTrans.value = Date
            recordDateH.value = ToHijriDate(Date)
            
            Me.Fromdate.value = Date
            Me.todate.value = Date
            Check17.value = vbChecked
            Me.Fromdate√H.value = ToHijriDate(Date)
todateH.value = ToHijriDate(Date)

   Me.Dcbranch.BoundText = Current_branch
       
            'XPDtbTrans.SetFocus
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 1
            GridInstallments.Enabled = True
      
 Check1.value = vbChecked
 

        Case 1
                    If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                 Dim rsDummy As ADODB.Recordset
                StrSQL = "select installid from tblContractInsAllocationsDetails where transid=" & TxtTransID & " and ISNULL(zatcaStatus, 0) = 1"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                     MsgBox "·«Ì„ﬂ‰  ⁄œÌ· «·”Ã· ‰Ÿ—« ·ÊÃÊœ ›Ê« Ì—  „ —›⁄Â« ··ÂÌ∆…"
                     Exit Sub
                End If
                rsDummy.Close
            TxtModFlg.Text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True

        Case 2
        
             PercentgValueAddedAccount_Transec XPDtbTrans, 8, 1, vaTAccount
             PercentgValueAddedAccount_Transec XPDtbTrans, 21, 1, vataccount2
            'AccountVat.BoundText = vataccount
       If vaTAccount = "" Then
       MsgBox "Ì—ÃÏ  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
       Exit Sub
       End If
         If vataccount2 = "" Then
       MsgBox "Ì—ÃÏ  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„»Ì⁄« "
       Exit Sub
       End If
       
       
                           If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
       If val(Me.Dcbranch.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
            Else
                MsgBox "Select Branch Firstly    ", vbCritical
            End If

            Dcbranch.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
CheckAcconts
If TxtNoteSerial.Text = "" Then     'ÃœÌœ ›ﬁÿ
                        If Notes_coding(val(my_branch), Me.XPDtbTrans.value) = "error" Then
                            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                        Else
                                       
                                        If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                                            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                                        Else
                                             
                                        End If
                        End If
 End If
 
            SaveData
           
        Case 3
            Undo

        Case 4
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
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

            '  If DoPremis(Do_Search, Me.name, True) = False Then
            '      Exit Sub
            '  End If
            '  Load FrmNotesSearch
            '  FrmNotesSearch.SearchType = 3
            'FrmNotesSearch.Show vbModal
        Case 6
            Unload Me

        Case 7
            addInterval

            '   ViewDataList
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

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

    If val(Me.TxtRowNumber.Text) <> 0 Then
        LngRow = val(Me.TxtRowNumber.Text)
    Else
        Me.Grid.rows = Me.Grid.rows + 1
        LngRow = Me.Grid.rows - 1
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
    
        .TextMatrix(LngRow, .ColIndex("AName")) = Me.DCAccountDist.Text
    
        .TextMatrix(LngRow, .ColIndex("Percentage")) = val(Me.TxtPercentage.Text)
    
        .TextMatrix(LngRow, .ColIndex("Remarks")) = (Me.TxtRemarks1.Text)
     
        .AutoSize 0, .Cols - 1, False
    End With

    Me.DCAccountDist.BoundText = ""
    Me.TxtPercentage.Text = ""
    Me.TxtRemarks1.Text = ""
  
    ReLineGrid
 
End Sub

Private Sub Undo()
   ' On Error GoTo ErrTrap

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





Private Sub cmdCreateENtry_Click()
Dim StrSQL  As String
StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
 createVoucher
 MsgBox " „ «⁄«œ… «‰‘«¡ «·ﬁÌœ"
End Sub

Private Sub CmdPrint_Click()
    
    
    
    
    On Error Resume Next
    Dim i As Integer
 



    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    GridInstallments.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.GridInstallments.PrintGrid " ﬁ—Ì—   «·„” Õﬁ«  ", True, 1, 1, 1500

 

End Sub

Private Sub cmdPrintInst_Click()
PeintInstalMent 0, , , , True
End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    If Grid.rows > 1 Then
        If Grid.rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.rows > 1 Then
                If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.Row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub
 
Private Sub Command2_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = "C:\" & "\File.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
   
      On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.FileName = "Payroll"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.FileName & ".xls"
Me.GridInstallments.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    

End Sub

Private Sub Command9_Click()
       ShowGL_cc Me.TxtNoteSerial.Text, , 200, val(Me.TxtNoteID.Text)
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
    Grid.rows = 1
          
    Select Case Index
        
        Case 0
            Frame1.Caption = "«· Ê“Ì⁄ ⁄·Ï «·Õ”«»«  "
            lbl(5).Caption = "«·Õ”«» "

            With Me.Grid
                .TextMatrix(0, .ColIndex("ASerial")) = "ﬂÊœ «·Õ”«»"
                .TextMatrix(0, .ColIndex("AName")) = "«”„ «·Õ”«»"
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
                .TextMatrix(0, .ColIndex("ASerial")) = "ﬂÊœ «·„—ﬂ“"
                .TextMatrix(0, .ColIndex("AName")) = "«”„ «·„—ﬂ“"
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
                .TextMatrix(0, .ColIndex("ASerial")) = "ﬂÊœ «·›— ⁄ "
                .TextMatrix(0, .ColIndex("AName")) = "«”„ «·›— ⁄ "
            End With
            
            Set Dcombos = New ClsDataCombos

            If SystemOptions.UserInterface = ArabicInterface Then
                Dcombos.GetBranches DCAccountDist
            Else
                Dcombos.GetBranches DCAccountDist

            End If

    End Select

End Sub

Function CheckAcconts() As Boolean
CheckAcconts = False

            Account_Code_dynamic80 = get_account_code_branch(80, my_branch)
            Account_Code_dynamic81 = get_account_code_branch(81, my_branch)
            Account_Code_dynamic82 = get_account_code_branch(82, my_branch)
            Account_Code_dynamic83 = get_account_code_branch(83, my_branch)
            Account_Code_dynamic84 = get_account_code_branch(84, my_branch)
            Account_Code_dynamic85 = get_account_code_branch(85, my_branch)
            Account_Code_dynamic123 = get_account_code_branch(123, my_branch)
             Account_Code_dynamic125 = get_account_code_branch(125, my_branch)
            Account_Code_dynamic153 = get_account_code_branch(153, my_branch)
            
            Account_Code_dynamic154 = get_account_code_branch(154, my_branch)
            
            Account_Code_dynamic155 = get_account_code_branch(155, my_branch)
            
            Account_Code_dynamic156 = get_account_code_branch(156, my_branch)
            
               If Account_Code_dynamic125 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»      ⁄„Ê·«  „” Õﬁ… ·«„·«ﬂ «·€Ì— ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
            If Account_Code_dynamic125 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»      «·«ÌÃ«—«  «·„” Õﬁ… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
              
             If Account_Code_dynamic153 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «Ì—«œ«  „Ì«Â ”⁄Ì", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
            If Account_Code_dynamic153 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»    Õ”«»   «Ì—«œ«  „Ì«Â ”⁄Ì", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
               
              
              If Account_Code_dynamic154 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «Ì—«œ«  „Ì«Â „” Õﬁ…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
            If Account_Code_dynamic154 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»    Õ”«»   «Ì—«œ«  „Ì«Â „” Õﬁ…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
              
              
              
               If Account_Code_dynamic155 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «Ì—«œ«    „” Õﬁ… ﬂÂ—»«¡", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
            If Account_Code_dynamic155 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»    Õ”«»   «Ì—«œ«  ﬂÂ—»«¡ „” Õﬁ…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
             
             
             
              If Account_Code_dynamic156 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «Ì—«œ«  Œœ„«  „” Õﬁ…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
            If Account_Code_dynamic156 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»    Õ”«»   «Ì—«œ«  Œœ„«  „” Õﬁ…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
             
             
             
             
If 1 = 1 Then ' ÃœÌœ
                If (val(TxtCommiValue)) > 0 Then
                            Account_Code_dynamic81 = get_account_code_branch(81, my_branch)
                            If Account_Code_dynamic81 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»         «·”⁄Ì Ê «·—”Ê„ «·«œ«—Ì… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
              
              
               If (val(TxtInsuranceValue)) > 0 Then
                            Account_Code_dynamic82 = get_account_code_branch(82, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «· √„Ì‰ «·„” —œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
                 
              
              
                    If (val(txtWater)) > 0 Then
                            Account_Code_dynamic83 = get_account_code_branch(83, my_branch)
                            If Account_Code_dynamic83 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «·„Ì«Â «·„ﬁœ„… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
                 
              
              
               If (val(TxtElectricity)) > 0 Then
                            Account_Code_dynamic84 = get_account_code_branch(84, my_branch)
                            If Account_Code_dynamic84 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «·ﬂÂ—»«¡ «·„ﬁœ„… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
                 
              
                      If (val(TxtPhone)) > 0 Then
                            Account_Code_dynamic85 = get_account_code_branch(85, my_branch)
                            If Account_Code_dynamic85 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  « ·Œœ„«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
              
                
End If


           Account_Code_dynamic123 = get_account_code_branch(123, my_branch)
                            If Account_Code_dynamic123 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»        «ÌÃ«—«  „” Õﬁ… ··€Ì— ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
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


Function createVoucher()
Dim noteId As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = "«À»«  «” Õﬁ«ﬁ ⁄‰ «·› —… „‰  " '& Fromdate√H.value & "  Õ Ï  " & TodateH.value & Chr(13)
des = des & " „‰ " & Fromdate.value & "  Õ Ï  " & todate.value & CHR(13)
des = des & " «·„Ê«›ﬁ „‰ " & Fromdate√H.value & "  Õ Ï  " & todateH.value & CHR(13)
des = des & " ··›—⁄ " & Dcbranch.Text & "     " & tXtRemarks

Dim tablename As String
Dim Filedname As String
Dim contNo As Long
Dim sql As String
tablename = "tblContractInsAllocations"
Filedname = "transID"
contNo = TxtTransID
Notevalue = 0

'If SystemOptions.WorkWithFirstInstallOnly = False Then
'Notevalue = val(TxtTotalContract) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone)
Notevalue = val(TxtTotalContract) + val(TxtCommiValue) + val(txtWater) + val(TxtElectricity) + val(TxtPhone)
'Else

'With GridInstallments

'If .Rows > 1 Then
'Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("RentValue")))
'Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Commissions")))
'Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Insurance")))
'Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Water")))
'Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Electric")))
'Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("TelandNet")))
'
 
'End If

'End With


'End If

 
 
If Me.TxtModFlg = "N" Then
CreateNotes noteId, (XPDtbTrans.value), val(Dcbranch.BoundText), 61, Notevalue, NoteSerial, TxtTransID, tablename, Filedname, contNo, des, recordDateH.value
 
 TxtNoteID.Text = noteId
TxtNoteSerial.Text = NoteSerial
Else
    Cn.Execute "Delete Notes where NoteId = " & val(TxtNoteID.Text)
    CreateNotes noteId, (XPDtbTrans.value), val(Dcbranch.BoundText), 61, Notevalue, NoteSerial, TxtTransID, tablename, Filedname, contNo, des, recordDateH.value
sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
TxtNoteID = noteId
'sql = sql & ",NoteSerial1='" & Me.TxtTransID & "'"
sql = sql & ",NoteSerial1='" & Me.TxtTransID & "',remark='" & des & "'"
  sql = sql & " where NoteID=" & val(TxtNoteID.Text)
   Cn.Execute sql
End If


CREATE_VOUCHER_GE val(TxtNoteID.Text), val(Dcbranch.BoundText), user_id, XPDtbTrans.value
rs.Resync adAffectCurrent
 

End Function



Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
 Dim TelandNet As Double
 Dim commisiontype As Double
Dim Insurance As Double
Dim Water As Double
Dim Electric As Double
Dim AmolaValus As Double
Dim OtherInformation As New ClsGLOther
Dim ownerid As Double
Dim commission As Double
 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim actiondesdes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 CheckAcconts
 Dim StrSQL As String
 
   PercentgValueAddedAccount_Transec XPDtbTrans, 8, 1, vaTAccount
             PercentgValueAddedAccount_Transec XPDtbTrans, 21, 1, vataccount2
             
             
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        
 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    
    
   
    my_branch = BranchID
 
   With Me.GridInstallments
'Selected
                          For i = 1 To .rows - 1
                     
                                   If val(.TextMatrix(i, .ColIndex("value"))) <> 0 And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                                   If SystemOptions.DUEDOCUMENTbyinstallDate = True Then
                               NoteDate = .TextMatrix(i, .ColIndex("Due_Date"))
                               End If
                               
                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))))
                     
                      '          StrTempDes = " «À»«  «” Õﬁ«ﬁ ⁄‰ «·› —…  «· Ì  »œ√ » «—ÌŒ " & (.TextMatrix(1, .ColIndex("Due_Date"))) & CHR(13)
                      '           StrTempDes = StrTempDes & "  «·„Ê«›ﬁ " & (.TextMatrix(1, .ColIndex("Due_DateH"))) & "  ·„œ…  " & (.TextMatrix(1, .ColIndex("Countsofall"))) & "  ‘ÂÊ—" & CHR(13)
                                  
                      'salim here
                      If SystemOptions.DUEDOCUMENTbyinstallDate = True Then
                   '    StrTempDes = " «À»«  «” Õﬁ«ﬁ  «·œ›⁄Â „‰    «—ÌŒ : " & .TextMatrix(i, .ColIndex("Due_DateH")) & CHR(13)
                   '    StrTempDes = StrTempDes & "  «·„Ê«›ﬁ  " & .TextMatrix(i, .ColIndex("Due_Date")) & CHR(13)
                      Else
                   '            StrTempDes = " «À»«  «” Õﬁ«ﬁ ⁄‰ «·› —…  «· Ì  »œ√ » «—ÌŒ " & FromDate.value & "  «·Ì " & ToDate.value & CHR(13)
                   '              StrTempDes = StrTempDes & "  «·„Ê«›ﬁ „‰" & Fromdate√H.value & "  «·Ì  " & todateH.value & CHR(13)
                        End If
                                          
                                     StrTempDes = " ⁄ﬁœ —ﬁ„ " & .TextMatrix(i, .ColIndex("NoteSerial1")) & CHR(13)
                                 StrTempDes = StrTempDes & " «À»«  «” Õﬁ«ﬁ ⁄‰ «·› —…  «· Ì  »œ√ » «—ÌŒ " & .TextMatrix(i, .ColIndex("Due_Date")) & "  «·Ì " & .TextMatrix(i, .ColIndex("nextinstalldate")) & CHR(13)
                                 StrTempDes = StrTempDes & "  «·„Ê«›ﬁ ÂÃ—Ì  „‰ " & .TextMatrix(i, .ColIndex("Due_DateH")) & "  «·Ì " & .TextMatrix(i, .ColIndex("nextinstalldateH")) & CHR(13)
                                 
                                  
                                 StrTempDes = StrTempDes & " «·„«·ﬂ " & .TextMatrix(i, .ColIndex("OwnerName")) & CHR(13)
                                 
                                  StrTempDes = StrTempDes & " «·„” √Ã— " & .TextMatrix(i, .ColIndex("CusName")) & "     " & .TextMatrix(i, .ColIndex("Remarks")) & CHR(13)
      StrTempDes = StrTempDes & "   «·⁄ﬁ«— " & .TextMatrix(i, .ColIndex("aqarname")) & CHR(13)
      StrTempDes = StrTempDes & "   «·ÊÕœ… " & .TextMatrix(i, .ColIndex("UnitTypeName")) & "  —ﬁ„ " & .TextMatrix(i, .ColIndex("unitno")) & CHR(13)
       
                              Notevalue = val(.TextMatrix(i, .ColIndex("RentValue")))
                          AmolaValus = val(.TextMatrix(i, .ColIndex("AmolaValus")))
                          ownerid = val(.TextMatrix(i, .ColIndex("ownerid")))
                          commission = val(.TextMatrix(i, .ColIndex("Commissions")))
                          commisiontype = val(.TextMatrix(i, .ColIndex("commisiontype")))
                          TelandNet = val(.TextMatrix(i, .ColIndex("TelandNet")))
                           Water = val(.TextMatrix(i, .ColIndex("Water")))
                           Electric = val(.TextMatrix(i, .ColIndex("Electric")))
                           
                                  If Notevalue > 0 Then
                                    LngDevNO = LngDevNO + 1
                                         actiondesdes = "ﬁÌ„… «· ⁄«ﬁœ "
                                           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                              
                              
                                                    LngDevNO = LngDevNO + 1
                                           actiondesdes = "ﬁÌ„… «·«ÌÃ«—  " '
                                           If commisiontype > 0 Then
                                                            If SystemOptions.Create2account4Supp = False Then
                                                                     StrTempAccountCode = Account_Code_dynamic123
                                                            Else
                                                            '
                                                            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_code1")
                                                            End If
                                                            If AmolaValus > 0 Then
                                                                    commission = Notevalue * AmolaValus / 100
                                                                    commission = Round(commission, 2)
                                                                    Notevalue = Notevalue - commission
                                                           End If
                                           
                                           Else
                                           
                                           
                                                          StrTempAccountCode = Account_Code_dynamic80
                                           
                                           '«·«ÌÃ«—«  «·„” Õﬁ…
                                           End If
                                           
                                           
                                           
                                                      
                    '     OtherInformation.Unitss = IIf(.TextMatrix(i, .ColIndex("Unitss")) = "", "", .TextMatrix(i, .ColIndex("Unitss")))
                    '  OtherInformation.StrUnit = IIf(.TextMatrix(i, .ColIndex("StrUnit")) = "", "", .TextMatrix(i, .ColIndex("StrUnit")))
                      OtherInformation.uintid = val(.TextMatrix(i, .ColIndex("uintid")))
                    '  OtherInformation.mType = IIf(.TextMatrix(i, .ColIndex("type")) = "", 0, val(.TextMatrix(i, .ColIndex("type"))))
                      OtherInformation.iqarid = IIf(.TextMatrix(i, .ColIndex("Iqar")) = "", 0, val(.TextMatrix(i, .ColIndex("Iqar"))))

   
   
   
                                           
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID, , , , , , , , , , val(.TextMatrix(i, .ColIndex("Iqar"))), , val(.TextMatrix(i, .ColIndex("uintid"))), , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                                     
                                                     
                                  End If
                                    

                                         
                                       
                            If commission > 0 Then
                          '   LngDevNO = LngDevNO + 1
                                        'StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_code1")
                                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))))
                                          '     If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, commission, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                          '              GoTo ErrTrap
                                         'End If
                              
                   
                            
                                          StrTempAccountCode = Account_Code_dynamic125
                                    LngDevNO = LngDevNO + 1
                                         actiondesdes = "⁄„Ê·«  «œ«—… «„·«ﬂ «·€Ì— "
                                           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, commission, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                              
                                  End If
                                  
                                  '    RentValue Commissions        Insurance Water   Electric  TelandNet
                                          Insurance = val(.TextMatrix(i, .ColIndex("Insurance")))
                                          If Insurance > 0 Then
                                          actiondesdes = "«· √„Ì‰ "
                                               LngDevNO = LngDevNO + 1
                                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))))
                                               If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Insurance, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                                          
                
                                                   If SystemOptions.CreateInsuranceAccountForCustomers = True Then
                                                           StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))), "InsuranceAccount")
                
                                                   Else
                                                   
                                                        StrTempAccountCode = Account_Code_dynamic82
                                                    End If
                                    LngDevNO = LngDevNO + 1
                                         actiondesdes = "«· √„Ì‰ "
                                           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Insurance, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                              
                                        End If
                                        
                                         Water = val(.TextMatrix(i, .ColIndex("Water")))
                                          If Water > 0 Then
                                          actiondesdes = "«·„Ì«… "
                                          LngDevNO = LngDevNO + 1
                                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))))
                                               If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Water, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                                         
                                           StrTempAccountCode = Account_Code_dynamic154
                                           
                                           LngDevNO = LngDevNO + 1
                                           actiondesdes = "«·„Ì«… "
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Water, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                              
                                        End If
                                        
                                                                      Notevalue = val(.TextMatrix(i, .ColIndex("Electric")))
                                          If Electric > 0 Then
                                          actiondesdes = "«·ﬂÂ—»«¡ "
                                                              LngDevNO = LngDevNO + 1
                                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))))
                                               If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Electric, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                                         
                                           StrTempAccountCode = Account_Code_dynamic155
                                           
                                           
                                           
                                           LngDevNO = LngDevNO + 1
                                           actiondesdes = "«·ﬂÂ—»«¡ "
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Electric, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                              
                                        End If
                                        
                                  
                                      TelandNet = val(.TextMatrix(i, .ColIndex("TelandNet")))
                                      
                                          If TelandNet > 0 Then
                                          actiondesdes = "«·Œœ„«  "
                                                   LngDevNO = LngDevNO + 1
                                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))))
                                               If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TelandNet, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                                         
                                           StrTempAccountCode = Account_Code_dynamic156
                                           
                                         
                                           
                                           LngDevNO = LngDevNO + 1
                                           actiondesdes = "«·Œœ„«  "
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TelandNet, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                              
                                        
                                        End If
                                  Dim VATValue As Double
                         VATValue = val(.TextMatrix(i, .ColIndex("VATValue")))
                        
                        
                                       If VATValue > 0 Then
                                          actiondesdes = "«·ﬁÌ„Â «·„÷«›… "
                                                   LngDevNO = LngDevNO + 1
                                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))))
                                               If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, VATValue, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                                         
                                           StrTempAccountCode = vaTAccount
                                           
                                         
                                           
                                           LngDevNO = LngDevNO + 1
                                           actiondesdes = "«·ﬁÌ„Â «·„÷«›… "
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, VATValue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                              
                                        
                                        End If
                   ''///////////////
                               Dim VATVCom As Double
                        
                        VATVCom = commission * 5 / 100
                        
                                       If VATValue > 0 And VATVCom > 0 Then
                                          actiondesdes = "«·ﬁÌ„Â «·„÷«›… ··⁄„Ê·… "
                                                   LngDevNO = LngDevNO + 1
                                                    'StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID"))))
                                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_code1")
                                               If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, VATVCom, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                                         
                                           StrTempAccountCode = vataccount2
                                           
                                         
                                           
                                           LngDevNO = LngDevNO + 1
                                           actiondesdes = "«·ﬁÌ„Â «·„÷«›… ··⁄„Ê·…"
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, VATVCom, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                              
                                        
                                        End If
                                  
                                  
                                  End If
                         
                         
                         Next i
       
       End With
 
 
  
  
             
            
'                  StrTempDes = "«À»«  «” Õﬁ«ﬁ ⁄‰ «·› —… „‰  " & Fromdate√H.value & "  Õ Ï  " & TodateH.value & Chr(13)
'                StrTempDes = StrTempDes & " «·„Ê«›ﬁ " & Fromdate.value & "  Õ Ï  " & todate.value & Chr(13)
'                StrTempDes = StrTempDes & " ··›—⁄ " & dcBranch.text & "     " & TxtRemarks.text
                
                
            LngDevNO = LngDevNO + 1
 

 
 
 If val(TxtTotalContract.Text) - val(TxtTotalTo.Text) > 0 Then
   '    StrTempAccountCode = Account_Code_dynamic80
                      
            
   ' Notevalue = val(TxtTotalContract.text) - val(TxtTotalTo.text)
   '                      LngDevNO = LngDevNO + 1
   '                     If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «· ⁄«ﬁœ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchId) = False Then
   '                         GoTo ErrTrap
   '
   '
   '                     End If
  End If
  
  
 
 

 
 If val(TxtTotalTo.Text) > 0 Then
'       StrTempAccountCode = Account_Code_dynamic123
 
'      Notevalue = val(TxtTotalTo.text)
   
'                         LngDevNO = LngDevNO + 1
'                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «· ⁄«ﬁœ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchId) = False Then
'                            GoTo ErrTrap
                        
                        
'                        End If
  End If
  
  '”⁄Ì
 If (val(TxtCommiValue.Text)) > 0 Then
       StrTempAccountCode = Account_Code_dynamic153
       
             ' If SystemOptions.WorkWithFirstInstallOnly = False Then
             Notevalue = (val(TxtCommiValue.Text))
     'Else
     'Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("Commissions")))
     'End If
    
   
 '  LngDevNO = LngDevNO + 1
 '           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "      ⁄„Ê·«  Ê—”Ê„ «œ«—Ì… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
 '               GoTo ErrTrap
 '           End If
  End If
  
  
   If val(TxtInsuranceValue.Text) > 0 Then
       StrTempAccountCode = Account_Code_dynamic82
       
     '         If SystemOptions.WorkWithFirstInstallOnly = False Then
      Notevalue = val(TxtInsuranceValue.Text)
     'Else
     'Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("Insurance")))
     'End If
            
            

    
 '  LngDevNO = LngDevNO + 1
 '           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "     √„Ì‰ „” —œ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
 '               GoTo ErrTrap
 '           End If
  End If
  
  
     If val(txtWater.Text) > 0 Then
       StrTempAccountCode = Account_Code_dynamic154
'
   
     '         If SystemOptions.WorkWithFirstInstallOnly = False Then
 '   Notevalue = val(TxtWater.text)
     'Else
     'Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("Water")))
     'End If
     
   
   LngDevNO = LngDevNO + 1
 '           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "    „Ì«Â ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
 '               GoTo ErrTrap
 '           End If
  End If
  

       If val(TxtElectricity.Text) > 0 Then
       StrTempAccountCode = Account_Code_dynamic155
     '  Notevalue = val(TxtElectricity.text)
   
     '           If SystemOptions.WorkWithFirstInstallOnly = False Then
    Notevalue = val(TxtElectricity.Text)
     'Else
     'Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("Electric")))
     'End If
     
 '  LngDevNO = LngDevNO + 1
 '           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "      ﬂÂ—»«¡ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
 '               GoTo ErrTrap
 '           End If
  End If
  
  
       If (val(TxtPhone.Text)) > 0 Then
       StrTempAccountCode = Account_Code_dynamic156
'       Notevalue = (val(TxtPhone.text) + val(TxtEnternet.text))
   
     '           If SystemOptions.WorkWithFirstInstallOnly = False Then
    Notevalue = val(TxtPhone.Text)
     'Else
     'Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("TelandNet")))
     'End If
     
 '  LngDevNO = LngDevNO + 1
 '           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "   Œœ„«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
 '               GoTo ErrTrap
 '           End If
  End If
    
    updateNotesValueAndNobytext (general_noteid)
ErrTrap:
End Function


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
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With

    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic

 
Dcombos.GetBranches Dcbranch

    With Me.Grid
        .rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset

StrSQL = "select * From tblContractInsAllocations  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
 
    lbl(3).Caption = "Select "
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = " Account Distubution"
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

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("ASerial")) = "Code"
        .TextMatrix(0, .ColIndex("AName")) = "Name"
'        .TextMatrix(0, .ColIndex("Percentage")) = "Percentage"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"

    End With

    Me.C1Tab1.TabCaption(0) = "Account Distributions "
    Me.C1Tab1.TabCaption(1) = "Distributions Period"
'    Frame4.Caption = "Distributions Period"
'    lbl(10).Caption = "From"
'    lbl(11).Caption = "To"
'    lbl(12).Caption = "Remarks"

'    Cmd(7).Caption = "Add"
'    Cmd(8).Caption = "Del"
 
    
End Sub

Public Sub FillNewGrid()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from TblyearsData "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Me.GridIntervals1

        .rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
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


    sql = "Select * from tblContractInsAllocations "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With GridOldEstimation

        .rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .rows = Rs3.RecordCount + 1
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
 
    Rs3.Close
End Sub
Function getnextDate(Optional newinstallNo As Double, Optional ByRef installdate, Optional ByRef installdateh, Optional contNo As Double)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    
    MySQL = " SELECT    installdate,installdateH "
     
    MySQL = MySQL & "      FROM         dbo.TblContract LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
    MySQL = MySQL & "        Where (dbo.TblContract.ContNo = " & contNo & ") And (dbo.TblContractInstallments.InstallNo =" & newinstallNo & ")"
   Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
    Else
    installdate = IIf(IsNull(RsData("installdate").value), Null, RsData("installdate").value)
    installdateh = IIf(IsNull(RsData("installdateH").value), Null, RsData("installdateH").value)
    
    
    End If
    
End Function
 
Public Sub FillGrid()

  '  On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
  Dim newinstallNo  As Double
Dim nextinstalldate As Date
Dim nextinstalldateH As String
Dim notpayed As Double
notpayed = 0
 
'My_SQL = " SELECT  dbo.TblContract.EndContract   , dbo.TblContract.Iqar,  dbo.TblContractInstallments.*, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee ,dbo.TblContract.NoteSerial1 ,dbo.TblContract.CusID"
'My_SQL = My_SQL & " FROM         dbo.TblContractInstallments INNER JOIN"
'My_SQL = My_SQL & "  dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
'My_SQL = My_SQL & " dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID"
 
 My_SQL = "SELECT       dbo.TblContract.EndDate, dbo.TblContract.TodateH, dbo.TblContract.NoteSerial1 AS ContractNoteSerial1, dbo.TblContract.EndContract,TblContractInstallments.NoteSerial1H ,"
My_SQL = My_SQL & "                       dbo.TblContract.Iqar, dbo.TblContractInstallments.id, dbo.TblContractInstallments.ContNo, dbo.TblContractInstallments.InstallNo,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstalldateH, dbo.TblContractInstallments.installValue, dbo.TblContractInstallments.Status,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.NoteSerial, dbo.TblContractInstallments.NoteSerial1, dbo.TblContractInstallments.NoteID, dbo.TblContractInstallments.RentValue,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.Commissions, dbo.TblContractInstallments.Insurance, dbo.TblContractInstallments.Water, dbo.TblContractInstallments.Electric,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.TelandNet, dbo.TblContractInstallments.payed, dbo.TblContractInstallments.Remains, dbo.TblContractInstallments.RentValuePayed,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.CommissionsPayed, dbo.TblContractInstallments.InsurancePayed, dbo.TblContractInstallments.WaterPayed,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.ElectricPayed, dbo.TblContractInstallments.TelandNetPayed, dbo.TblContractInstallments.lastPayedDate,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.lastPayedDateH, dbo.TblContractInstallments.allocations, dbo.TblContractInstallments.Countsofall, dbo.TblContractInstallments.Doneofall,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.hijri, dbo.TblContractInstallments.OldValueDate, dbo.TblContractInstallments.OldValueDateH, dbo.TblContractInstallments.OldValue,"
My_SQL = My_SQL & "          dbo.TblContractInstallments.des, dbo.TblContractInstallments.NpayedValue, dbo.TblContractInstallments.Rent1, dbo.TblContractInstallments.RentArbon,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.NetRent, dbo.TblContractInstallments.Commissions1, dbo.TblContractInstallments.CommissionsArbon,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.NetCommissions, dbo.TblContractInstallments.Insurance1, dbo.TblContractInstallments.InsuranceArbon,"
My_SQL = My_SQL & "          dbo.TblContractInstallments.NetInsurance, dbo.TblContractInstallments.Water1, dbo.TblContractInstallments.WaterArbon, dbo.TblContractInstallments.NetWater,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.Electric1, dbo.TblContractInstallments.ElectricArbon, dbo.TblContractInstallments.NetElectric, dbo.TblContractInstallments.ServiceArbon,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.WivrID, dbo.TblContractInstallments.OldValuePayed, dbo.TblContractInstallments.TempInstal, dbo.TblContractInstallments.VATPayed,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.VATValue, dbo.TblContractInstallments.DevID, dbo.TblContractInstallments.Prefix, dbo.TblContractInstallments.VATArboon,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.ContractReVouchID, dbo.TblContractInstallments.ContractReVouchID2, dbo.TblContractInstallments.VATValueOld,"
My_SQL = My_SQL & "                                dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblContract.CusID, dbo.TblAqar.aqarname, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.unitdesc,"
My_SQL = My_SQL & "                                dbo.TblAkarUnit.name AS UnitTypeName, dbo.TblAkarUnit.namee, dbo.TblContract.ownerid, TblCustemers_1.CusName AS OwnerName,"
My_SQL = My_SQL & "                                TblCustemers_1.CusNamee AS OwnerNameE,TblContract.UnitNo as uintid"
My_SQL = My_SQL & "          FROM         dbo.TblContractInstallments INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id INNER JOIN"
 My_SQL = My_SQL & "                               dbo.TblCustemers AS TblCustemers_1 ON dbo.TblContract.ownerid = TblCustemers_1.CusID"
 
My_SQL = My_SQL & " WHERE    "
My_SQL = My_SQL & " (dbo.TblContractInstallments.allocations = 0 or dbo.TblContractInstallments.allocations IS NULL)  AND (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL)"
 





If Check1.value = vbUnchecked Then
My_SQL = My_SQL & " and     (dbo.TblContract.EndContract = 0 or dbo.TblContract.EndContract IS NULL)  "

'SELECT ContractNo,FilterDate,* FROM TblFiterWaiver
'SALIM HERE «»œ „‰ «· ⁄œÌ·

End If
My_SQL = My_SQL & "  and     dbo.TblContract.NoteSerial1 Not In (Select  VV.ContractNo from TblFiterWaiver VV where FilterDate <=" & SQLDate(todate, True) & ")"

   
        My_SQL = My_SQL + " and (Installdate >=" & SQLDate(Me.Fromdate, True) & ""
     

 
        My_SQL = My_SQL + " and Installdate <=" & SQLDate(todate, True) & " )"
  If Me.Dcbranch.BoundText <> "" Then
    My_SQL = My_SQL + "   AND (dbo.TblContract.Branch_NO = " & val(Me.Dcbranch.BoundText) & ")"
End If

My_SQL = My_SQL + "   order by Installdate "
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .rows = rs.RecordCount + 1
           rs.MoveFirst
Dim Iqar As Double
Dim commisiontype As Integer
 Dim AmolaValus As Double
 Dim ownerid As Double
 
            For i = 1 To .rows - 1
           Iqar = (IIf(IsNull(rs.Fields("Iqar").value), 0, rs.Fields("Iqar").value))
           
           
           commisiontype = AqarCommisionType(Iqar, AmolaValus, ownerid)
           .TextMatrix(i, .ColIndex("AmolaValus")) = AmolaValus
           .TextMatrix(i, .ColIndex("OwnerName")) = (IIf(IsNull(rs.Fields("OwnerName").value), "", rs.Fields("OwnerName").value))
            .TextMatrix(i, .ColIndex("ownerid")) = ownerid
             .TextMatrix(i, .ColIndex("Iqar")) = Iqar
             .TextMatrix(i, .ColIndex("commisiontype")) = commisiontype
              .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(rs.Fields("id").value), 0, rs.Fields("id").value))
               .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("InstallNo").value), 0, rs.Fields("InstallNo").value))
               .TextMatrix(i, .ColIndex("ContNo")) = (IIf(IsNull(rs.Fields("ContNo").value), "", rs.Fields("ContNo").value))
                .TextMatrix(i, .ColIndex("uintid")) = (IIf(IsNull(rs.Fields("uintid").value), "", rs.Fields("uintid").value))
.TextMatrix(i, .ColIndex("NoteSerial1H")) = (IIf(IsNull(rs.Fields("NoteSerial1H").value), "", rs.Fields("NoteSerial1H").value))
               .TextMatrix(i, .ColIndex("aqarname")) = (IIf(IsNull(rs.Fields("aqarname").value), "", rs.Fields("aqarname").value))
               .TextMatrix(i, .ColIndex("UnitTypeName")) = (IIf(IsNull(rs.Fields("UnitTypeName").value), "", rs.Fields("UnitTypeName").value))
               .TextMatrix(i, .ColIndex("unitno")) = (IIf(IsNull(rs.Fields("unitno").value), "", rs.Fields("unitno").value))
               
               
               newinstallNo = val(.TextMatrix(i, .ColIndex("InstallNo"))) + 1
getnextDate newinstallNo, nextinstalldate, nextinstalldateH, val(.TextMatrix(i, .ColIndex("ContNo")))

        
         .TextMatrix(i, .ColIndex("nextinstalldate")) = IIf(year(nextinstalldate) <> 1899, nextinstalldate, rs.Fields("EndDate").value)
  .TextMatrix(i, .ColIndex("nextinstalldateH")) = IIf(IsDate(nextinstalldateH), nextinstalldateH, rs.Fields("todateH").value)
 
  
.TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("ContractNoteSerial1").value), "", rs.Fields("ContractNoteSerial1").value))
                
                Debug.Print .TextMatrix(i, .ColIndex("NoteSerial1"))
  .TextMatrix(i, .ColIndex("VATValue")) = (IIf(IsNull(rs.Fields("VATValue").value), "", rs.Fields("VATValue").value))
                'ContNo
                
                
 .TextMatrix(i, .ColIndex("Due_DateH")) = (IIf(IsNull(rs.Fields("Installdateh").value), ToHijriDate(Date), rs.Fields("Installdateh").value))
  .TextMatrix(i, .ColIndex("Due_Date")) = IIf(IsNull(rs.Fields("Installdate").value), Date, rs.Fields("Installdate").value)
        

        
        
        
    .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("installValue").value), 0, rs.Fields("installValue").value))
     .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value))
   
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
   Else
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
   End If
 .TextMatrix(i, .ColIndex("hijri")) = (IIf(IsNull(rs.Fields("hijri").value), 0, rs.Fields("hijri").value))   '
   .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
 
    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("TelandNet").value), 0, rs.Fields("TelandNet").value))
 
    
       .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(rs.Fields("allocations").value), 0, rs.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(rs.Fields("Countsofall").value), 0, rs.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))
.TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Electric"))) + val(.TextMatrix(i, .ColIndex("Water"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("Insurance")))

        rs.MoveNext
            Next i
 
            rs.Close
        End If
  ' .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With
ReLineGrid
End Sub
 
 Public Sub FillGrid2()

  '  On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
  Dim newinstallNo  As Double
Dim nextinstalldate As Date
Dim nextinstalldateH As String
Dim notpayed As Double
notpayed = 0
 
'My_SQL = " SELECT  dbo.TblContract.EndContract   , dbo.TblContract.Iqar,  dbo.TblContractInstallments.*, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee ,dbo.TblContract.NoteSerial1 ,dbo.TblContract.CusID"
'My_SQL = My_SQL & " FROM         dbo.TblContractInstallments INNER JOIN"
'My_SQL = My_SQL & "  dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
'My_SQL = My_SQL & " dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID"
 
 My_SQL = "SELECT       dbo.TblContract.EndDate, dbo.TblContract.TodateH, dbo.TblContract.NoteSerial1 AS ContractNoteSerial1, dbo.TblContract.EndContract,TblContractInstallments.NoteSerial1H ,"
My_SQL = My_SQL & "                       dbo.TblContract.Iqar, dbo.TblContractInstallments.id, dbo.TblContractInstallments.ContNo, dbo.TblContractInstallments.InstallNo,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstalldateH, dbo.TblContractInstallments.installValue, dbo.TblContractInstallments.Status,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.NoteSerial, dbo.TblContractInstallments.NoteSerial1, dbo.TblContractInstallments.NoteID, dbo.TblContractInstallments.RentValue,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.Commissions, dbo.TblContractInstallments.Insurance, dbo.TblContractInstallments.Water, dbo.TblContractInstallments.Electric,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.TelandNet, dbo.TblContractInstallments.payed, dbo.TblContractInstallments.Remains, dbo.TblContractInstallments.RentValuePayed,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.CommissionsPayed, dbo.TblContractInstallments.InsurancePayed, dbo.TblContractInstallments.WaterPayed,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.ElectricPayed, dbo.TblContractInstallments.TelandNetPayed, dbo.TblContractInstallments.lastPayedDate,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.lastPayedDateH, dbo.TblContractInstallments.allocations, dbo.TblContractInstallments.Countsofall, dbo.TblContractInstallments.Doneofall,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.hijri, dbo.TblContractInstallments.OldValueDate, dbo.TblContractInstallments.OldValueDateH, dbo.TblContractInstallments.OldValue,"
My_SQL = My_SQL & "          dbo.TblContractInstallments.des, dbo.TblContractInstallments.NpayedValue, dbo.TblContractInstallments.Rent1, dbo.TblContractInstallments.RentArbon,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.NetRent, dbo.TblContractInstallments.Commissions1, dbo.TblContractInstallments.CommissionsArbon,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.NetCommissions, dbo.TblContractInstallments.Insurance1, dbo.TblContractInstallments.InsuranceArbon,"
My_SQL = My_SQL & "          dbo.TblContractInstallments.NetInsurance, dbo.TblContractInstallments.Water1, dbo.TblContractInstallments.WaterArbon, dbo.TblContractInstallments.NetWater,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.Electric1, dbo.TblContractInstallments.ElectricArbon, dbo.TblContractInstallments.NetElectric, dbo.TblContractInstallments.ServiceArbon,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.WivrID, dbo.TblContractInstallments.OldValuePayed, dbo.TblContractInstallments.TempInstal, dbo.TblContractInstallments.VATPayed,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.VATValue, dbo.TblContractInstallments.DevID, dbo.TblContractInstallments.Prefix, dbo.TblContractInstallments.VATArboon,"
My_SQL = My_SQL & "                                dbo.TblContractInstallments.ContractReVouchID, dbo.TblContractInstallments.ContractReVouchID2, dbo.TblContractInstallments.VATValueOld,"
My_SQL = My_SQL & "                                dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblContract.CusID, dbo.TblAqar.aqarname, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.unitdesc,"
My_SQL = My_SQL & "                                dbo.TblAkarUnit.name AS UnitTypeName, dbo.TblAkarUnit.namee, dbo.TblContract.ownerid, TblCustemers_1.CusName AS OwnerName,"
My_SQL = My_SQL & "                                TblCustemers_1.CusNamee AS OwnerNameE"
My_SQL = My_SQL & "          FROM         dbo.TblContractInstallments INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id INNER JOIN"
My_SQL = My_SQL & "                                dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id INNER JOIN"
 My_SQL = My_SQL & "                               dbo.TblCustemers AS TblCustemers_1 ON dbo.TblContract.ownerid = TblCustemers_1.CusID"
 
My_SQL = My_SQL & " WHERE   (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL)"
'My_SQL = My_SQL & " (dbo.TblContractInstallments.allocations = 0 or dbo.TblContractInstallments.allocations IS NULL)  AND (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL)"
 





If Check1.value = vbUnchecked Then
My_SQL = My_SQL & " and     (dbo.TblContract.EndContract = 0 or dbo.TblContract.EndContract IS NULL)  "

'SELECT ContractNo,FilterDate,* FROM TblFiterWaiver
'SALIM HERE «»œ „‰ «· ⁄œÌ·

End If
My_SQL = My_SQL & "  and     dbo.TblContract.NoteSerial1 Not In (Select  VV.ContractNo from TblFiterWaiver VV where FilterDate <=" & SQLDate(todate, True) & ")"

   
        My_SQL = My_SQL + " and (Installdate >=" & SQLDate(Me.Fromdate, True) & ""
     

 
        My_SQL = My_SQL + " and Installdate <=" & SQLDate(todate, True) & " )"
  If Me.Dcbranch.BoundText <> "" Then
    My_SQL = My_SQL + "   AND (dbo.TblContract.Branch_NO = " & val(Me.Dcbranch.BoundText) & ")"
End If

My_SQL = My_SQL + "   order by Installdate "
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .rows = rs.RecordCount + 1
           rs.MoveFirst
Dim Iqar As Double
Dim commisiontype As Integer
 Dim AmolaValus As Double
 Dim ownerid As Double
 
            For i = 1 To .rows - 1
           Iqar = (IIf(IsNull(rs.Fields("Iqar").value), 0, rs.Fields("Iqar").value))
           
           
           commisiontype = AqarCommisionType(Iqar, AmolaValus, ownerid)
           .TextMatrix(i, .ColIndex("AmolaValus")) = AmolaValus
           .TextMatrix(i, .ColIndex("OwnerName")) = (IIf(IsNull(rs.Fields("OwnerName").value), "", rs.Fields("OwnerName").value))
            .TextMatrix(i, .ColIndex("ownerid")) = ownerid
             .TextMatrix(i, .ColIndex("Iqar")) = Iqar
             .TextMatrix(i, .ColIndex("commisiontype")) = commisiontype
              .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(rs.Fields("id").value), 0, rs.Fields("id").value))
               .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("InstallNo").value), 0, rs.Fields("InstallNo").value))
               .TextMatrix(i, .ColIndex("ContNo")) = (IIf(IsNull(rs.Fields("ContNo").value), "", rs.Fields("ContNo").value))
                
.TextMatrix(i, .ColIndex("NoteSerial1H")) = (IIf(IsNull(rs.Fields("NoteSerial1H").value), "", rs.Fields("NoteSerial1H").value))
               .TextMatrix(i, .ColIndex("aqarname")) = (IIf(IsNull(rs.Fields("aqarname").value), "", rs.Fields("aqarname").value))
               .TextMatrix(i, .ColIndex("UnitTypeName")) = (IIf(IsNull(rs.Fields("UnitTypeName").value), "", rs.Fields("UnitTypeName").value))
               .TextMatrix(i, .ColIndex("unitno")) = (IIf(IsNull(rs.Fields("unitno").value), "", rs.Fields("unitno").value))
               
               
               newinstallNo = val(.TextMatrix(i, .ColIndex("InstallNo"))) + 1
getnextDate newinstallNo, nextinstalldate, nextinstalldateH, val(.TextMatrix(i, .ColIndex("ContNo")))

        
         .TextMatrix(i, .ColIndex("nextinstalldate")) = IIf(year(nextinstalldate) <> 1899, nextinstalldate, rs.Fields("EndDate").value)
  .TextMatrix(i, .ColIndex("nextinstalldateH")) = IIf(IsDate(nextinstalldateH), nextinstalldateH, rs.Fields("todateH").value)
 
  
.TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("ContractNoteSerial1").value), "", rs.Fields("ContractNoteSerial1").value))
                
                Debug.Print .TextMatrix(i, .ColIndex("NoteSerial1"))
  .TextMatrix(i, .ColIndex("VATValue")) = (IIf(IsNull(rs.Fields("VATValue").value), "", rs.Fields("VATValue").value))
                'ContNo
                
                
 .TextMatrix(i, .ColIndex("Due_DateH")) = (IIf(IsNull(rs.Fields("Installdateh").value), ToHijriDate(Date), rs.Fields("Installdateh").value))
  .TextMatrix(i, .ColIndex("Due_Date")) = IIf(IsNull(rs.Fields("Installdate").value), Date, rs.Fields("Installdate").value)
        

        
        
        
    .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("installValue").value), 0, rs.Fields("installValue").value))
     .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value))
   
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
   Else
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
   End If
 .TextMatrix(i, .ColIndex("hijri")) = (IIf(IsNull(rs.Fields("hijri").value), 0, rs.Fields("hijri").value))   '
   .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
 
    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("TelandNet").value), 0, rs.Fields("TelandNet").value))
 
    
       .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(rs.Fields("allocations").value), 0, rs.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(rs.Fields("Countsofall").value), 0, rs.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))
.TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Electric"))) + val(.TextMatrix(i, .ColIndex("Water"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("Insurance")))

        rs.MoveNext
            Next i
 
            rs.Close
        End If
  ' .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With
ReLineGrid
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

Private Sub FromDate_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         Me.Fromdate√H.value = ToHijriDate(Fromdate.value)
       
End If
End Sub

Private Sub Fromdate√H_LostFocus()
     If Me.TxtModFlg.Text <> "R" Then
             
            Fromdate.value = ToGregorianDate(Fromdate√H.value)
               
        End If

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
   
        If Row = .rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
 
    Dim Percenrage As Double
 
 
    IntCounter = 0
  Me.TxtTotalContract.Text = 0
  TxtTotalTo.Text = 0
  Me.TxtCommiValue.Text = 0
    Me.TxtInsuranceValue.Text = 0
      Me.txtWater.Text = 0
      Me.TxtElectricity.Text = 0
        Me.TxtPhone.Text = 0
     
    With Me.GridInstallments

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("Value")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                
                
                
                     If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                     If val(.TextMatrix(i, .ColIndex("commisiontype"))) = 1 Then
  Me.TxtTotalTo.Text = val(Me.TxtTotalTo.Text) + .TextMatrix(i, .ColIndex("RentValue"))
  Else
  Me.TxtTotalContract.Text = val(Me.TxtTotalContract.Text) + .TextMatrix(i, .ColIndex("RentValue"))
  End If
  
   
  Me.TxtCommiValue.Text = val(Me.TxtCommiValue.Text) + .TextMatrix(i, .ColIndex("Commissions"))
  Me.TxtInsuranceValue.Text = val(Me.TxtInsuranceValue.Text) + .TextMatrix(i, .ColIndex("Insurance"))
  Me.txtWater.Text = val(Me.txtWater.Text) + .TextMatrix(i, .ColIndex("Water"))
  Me.TxtElectricity.Text = val(Me.TxtElectricity.Text) + .TextMatrix(i, .ColIndex("Electric"))
  Me.TxtPhone.Text = val(Me.TxtPhone.Text) + .TextMatrix(i, .ColIndex("TelandNet"))
  
  End If
  
     
         
            End If

        Next i
   
    End With

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim RsDev1 As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.rows = 1
          
 
    
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtTransID.Text = IIf(IsNull(rs("transID").value), "", rs("transID").value)
 
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
  recordDateH.value = IIf(IsNull(rs("recordDateH").value), ToHijriDate(Date), rs("recordDateH").value)
  Dcbranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    Fromdate.value = IIf(IsNull(rs("Fromdate").value), Date, rs("Fromdate").value)
 Me.Fromdate√H.value = IIf(IsNull(rs("FromDateh").value), ToHijriDate(Date), rs("FromDateh").value)
    
        todate.value = IIf(IsNull(rs("todate").value), Date, rs("todate").value)
  todateH.value = IIf(IsNull(rs("todateH").value), ToHijriDate(Date), rs("todateH").value)
    
    Me.TxtNoteID.Text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
   Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    tXtRemarks.Text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
 
 
 

 '   StrSQL = "  SELECT     dbo.tblContractInsAllocationsDetails.*, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee"
'StrSQL = StrSQL & "  FROM         dbo.tblContractInsAllocationsDetails INNER JOIN"
'StrSQL = StrSQL & "   dbo.TblCustemers ON dbo.tblContractInsAllocationsDetails.CusID = dbo.TblCustemers.CusID"

StrSQL = "SELECT     dbo.tblContractInsAllocationsDetails.id, tblContractInsAllocationsDetails.NoteSerial1H,  dbo.tblContractInsAllocationsDetails.transID,tblContractInsAllocationsDetails.ContNo, dbo.tblContractInsAllocationsDetails.CusID, "
StrSQL = StrSQL & "                        dbo.tblContractInsAllocationsDetails.InstallNo, dbo.tblContractInsAllocationsDetails.Installdate, dbo.tblContractInsAllocationsDetails.InstalldateH,"
StrSQL = StrSQL & "                                    dbo.tblContractInsAllocationsDetails.installValue, dbo.tblContractInsAllocationsDetails.RentValue, dbo.tblContractInsAllocationsDetails.Commissions,"
StrSQL = StrSQL & "                                    dbo.tblContractInsAllocationsDetails.Insurance, dbo.tblContractInsAllocationsDetails.Water, dbo.tblContractInsAllocationsDetails.Electric,"
StrSQL = StrSQL & "                                    dbo.tblContractInsAllocationsDetails.TelandNet, dbo.tblContractInsAllocationsDetails.allocations, dbo.tblContractInsAllocationsDetails.Countsofall,"
StrSQL = StrSQL & "                                    dbo.tblContractInsAllocationsDetails.Doneofall, dbo.tblContractInsAllocationsDetails.Installid, dbo.tblContractInsAllocationsDetails.NoteSerial,"
StrSQL = StrSQL & "                                    dbo.tblContractInsAllocationsDetails.hijri, dbo.tblContractInsAllocationsDetails.Iqar, dbo.tblContractInsAllocationsDetails.commisiontype,"
StrSQL = StrSQL & "                                    dbo.tblContractInsAllocationsDetails.AmolaValus, dbo.tblContractInsAllocationsDetails.ownerid, dbo.tblContractInsAllocationsDetails.VATValue,"
StrSQL = StrSQL & "                                    dbo.tblContractInsAllocationsDetails.nextinstalldateH, dbo.tblContractInsAllocationsDetails.nextinstalldate, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "                                    dbo.TblCustemers.CusNamee, dbo.TblContractInstallments.ContNo, TblCustemers_1.CusName AS oWNERnAME, TblCustemers_1.CusNamee AS oWNERnAMEE,"
StrSQL = StrSQL & "                                    dbo.TblAqar.aqarname, dbo.TblContract.UnitNo as uintid, dbo.TblAqarDetai.unitno, dbo.TblAkarUnit.name AS UnitTypeName, dbo.TblAkarUnit.namee AS UnitTypeNameE"
StrSQL = StrSQL & "              FROM         dbo.tblContractInsAllocationsDetails INNER JOIN"
StrSQL = StrSQL & "                                    dbo.TblCustemers ON dbo.tblContractInsAllocationsDetails.CusID = dbo.TblCustemers.CusID INNER JOIN"
StrSQL = StrSQL & "                                    dbo.TblContractInstallments ON dbo.tblContractInsAllocationsDetails.Installid = dbo.TblContractInstallments.id INNER JOIN"
StrSQL = StrSQL & "                                    dbo.TblCustemers AS TblCustemers_1 ON dbo.tblContractInsAllocationsDetails.ownerid = TblCustemers_1.CusID INNER JOIN"
StrSQL = StrSQL & "                                    dbo.TblAqar ON dbo.tblContractInsAllocationsDetails.Iqar = dbo.TblAqar.Aqarid INNER JOIN"
StrSQL = StrSQL & "                                    dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
StrSQL = StrSQL & "                                    dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id INNER JOIN"
StrSQL = StrSQL & "                                    dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id"
 StrSQL = StrSQL & "   Where (dbo.tblContractInsAllocationsDetails.TransID = " & val(Me.TxtTransID.Text) & ")"

'StrSQL = StrSQL & "  WHERE     (dbo.tblContractInsAllocations.transID = " & val(Me.TxtTransID.text) & ") "
    'StrSQL = StrSQL & "  where transID=" & val(Me.TxtTransID.text)
    printSQL = StrSQL
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.GridInstallments
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
                           .TextMatrix(i, .ColIndex("aqarname")) = (IIf(IsNull(RsDev.Fields("aqarname").value), "", RsDev.Fields("aqarname").value))
               .TextMatrix(i, .ColIndex("UnitTypeName")) = (IIf(IsNull(RsDev.Fields("UnitTypeName").value), "", RsDev.Fields("UnitTypeName").value))
               .TextMatrix(i, .ColIndex("unitno")) = (IIf(IsNull(RsDev.Fields("unitno").value), "", RsDev.Fields("unitno").value))
               .TextMatrix(i, .ColIndex("OwnerName")) = (IIf(IsNull(RsDev.Fields("OwnerName").value), "", RsDev.Fields("OwnerName").value))
               .TextMatrix(i, .ColIndex("ContNo")) = (IIf(IsNull(RsDev.Fields("ContNo").value), "", RsDev.Fields("ContNo").value))
               .TextMatrix(i, .ColIndex("uintid")) = (IIf(IsNull(RsDev.Fields("uintid").value), "", RsDev.Fields("uintid").value))
               
               .TextMatrix(i, .ColIndex("NoteSerial1H")) = (IIf(IsNull(RsDev.Fields("NoteSerial1H").value), "", RsDev.Fields("NoteSerial1H").value))
               
               
           .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(RsDev.Fields("Installid").value), 0, RsDev.Fields("Installid").value))
               .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(RsDev.Fields("InstallNo").value), 0, RsDev.Fields("InstallNo").value))
 .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(RsDev.Fields("NoteSerial").value), "", RsDev.Fields("NoteSerial").value))
                
 .TextMatrix(i, .ColIndex("Due_DateH")) = (IIf(IsNull(RsDev.Fields("Installdateh").value), ToHijriDate(Date), RsDev.Fields("Installdateh").value))
  .TextMatrix(i, .ColIndex("Due_Date")) = IIf(IsNull(RsDev.Fields("Installdate").value), Date, RsDev.Fields("Installdate").value)
  
  
   .TextMatrix(i, .ColIndex("nextinstalldate")) = (IIf(IsNull(RsDev.Fields("nextinstalldate").value), "", RsDev.Fields("nextinstalldate").value))
  .TextMatrix(i, .ColIndex("nextinstalldateH")) = IIf(IsNull(RsDev.Fields("nextinstalldateH").value), "", RsDev.Fields("nextinstalldateH").value)
  
  
   
              
  
            .TextMatrix(i, .ColIndex("VATValue")) = (IIf(IsNull(RsDev.Fields("VATValue").value), 0, RsDev.Fields("VATValue").value))
            
    .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(RsDev.Fields("installValue").value), 0, RsDev.Fields("installValue").value))
     .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(RsDev.Fields("CusID").value), "", RsDev.Fields("CusID").value))
   
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(RsDev.Fields("CusName").value), "", RsDev.Fields("CusName").value))
   Else
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(RsDev.Fields("CusNamee").value), "", RsDev.Fields("CusNamee").value))
   End If
  'hijri
  .TextMatrix(i, .ColIndex("hijri")) = (IIf(IsNull(RsDev.Fields("hijri").value), 0, RsDev.Fields("hijri").value))   '
   .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
 
    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(RsDev.Fields("RentValue").value), 0, RsDev.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(RsDev.Fields("Commissions").value), 0, RsDev.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(RsDev.Fields("Insurance").value), 0, RsDev.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(RsDev.Fields("Water").value), 0, RsDev.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(RsDev.Fields("Electric").value), 0, RsDev.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(RsDev.Fields("TelandNet").value), 0, RsDev.Fields("TelandNet").value))
 

 
 .TextMatrix(i, .ColIndex("Iqar")) = (IIf(IsNull(RsDev.Fields("Iqar").value), 0, RsDev.Fields("Iqar").value))
 .TextMatrix(i, .ColIndex("commisiontype")) = (IIf(IsNull(RsDev.Fields("commisiontype").value), 0, RsDev.Fields("commisiontype").value))
  .TextMatrix(i, .ColIndex("AmolaValus")) = (IIf(IsNull(RsDev.Fields("AmolaValus").value), 0, RsDev.Fields("AmolaValus").value))
   .TextMatrix(i, .ColIndex("ownerid")) = (IIf(IsNull(RsDev.Fields("ownerid").value), 0, RsDev.Fields("ownerid").value))

    
       .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(RsDev.Fields("allocations").value), 0, RsDev.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(RsDev.Fields("Countsofall").value), 0, RsDev.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(RsDev.Fields("Doneofall").value), 0, RsDev.Fields("Doneofall").value))
             .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Electric"))) + val(.TextMatrix(i, .ColIndex("Water"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("Insurance")))
             
                RsDev.MoveNext
            Next i
 
        End With

    End If
 RsDev.Close
 
 
    LabCurrRec.Caption = rs.AbsolutePosition
    LabCountRec.Caption = rs.RecordCount
 
  
 
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With GridInstallments
 
    'If .ColKey(Col) <> "Due_DateH" And .ColKey(Col) <> "Status" Then
 
         If .ColKey(Col) = "Select" Or .ColKey(Col) = "VATValue" Then
         
         ElseIf .ColKey(Col) = "Print" Or .ColKey(Col) = "SMS" Then
            Cancel = False
   Else
        
        Cancel = True
        
        End If
 
        
    End With
End Sub

 

Private Sub GridInstallments_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
Dim newinstallNo  As Double
Dim nextinstalldate As Date
Dim nextinstalldateH As String
Dim contNo As Long
Dim cusId As Long
Dim cusName As String
Dim Due_Date As String
Dim mValue As Double
newinstallNo = 0
Dim TxtPhone As String
Dim s As String
Dim CurrentMessage As String
Dim mCusName As String
Dim rsDummy As New ADODB.Recordset
With GridInstallments
    newinstallNo = val(.TextMatrix(Row + 1, .ColIndex("InstallNo")))
    cusId = val(.TextMatrix(Row, .ColIndex("CusId")))
    mValue = val(.TextMatrix(Row, .ColIndex("Value")))
    mCusName = Trim(.TextMatrix(Row, .ColIndex("CusName")))
    Due_Date = Trim(.TextMatrix(Row, .ColIndex("Due_Date")))
    
    contNo = val(.TextMatrix(Row + 1, .ColIndex("ContNo")))
Select Case .ColKey(Col)

Case "SMS"
    s = "Select TblCustemers.Cus_mobile from TblCustemers  where cusid =  " & cusId
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummy.EOF Then
        TxtPhone = CheckPhoneNumber(Trim(rsDummy!Cus_mobile & ""), Row)
        
        CurrentMessage = txtSMS & " «·œ›⁄… » «—ÌŒ" & Due_Date & " »ﬁÌ„… " & mValue
        SENDSMS TxtPhone, CurrentMessage
        MsgBox " „ «·«—”«· »‰Ã«Õ"
    End If
Case "Print"
contNo = val(.TextMatrix(Row, .ColIndex("ContNo")))
getnextDate newinstallNo, nextinstalldate, nextinstalldateH
PeintInstalMent val(.TextMatrix(Row, .ColIndex("InstallNo"))), nextinstalldate, nextinstalldateH, contNo

Case "PrintJE"
ShowGL_cc .TextMatrix(Row, .ColIndex("NoteSerial")), , 200
Case "RecalcVAt"

End Select
End With

End Sub

Private Sub GridInstallments_DblClick()
'RSContract.show
'RSContract.FindRec val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("ContNo"))), True
'    RSContract.RereivID = val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("ContNo")))
    
    
    GridInstallments_CellButtonClick GridInstallments.Row, GridInstallments.ColIndex("Print")
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption


End Sub

 

Private Sub PercentagType_Click(Index As Integer)

    Select Case Index
        
        Case 0
            TxtPercentage.locked = True
            TxtPercentage.Text = ""

        Case 1
            TxtPercentage.locked = False
            TxtPercentage.Text = ""

    End Select

End Sub

Private Sub RecordDateH_LostFocus()
     If Me.TxtModFlg.Text <> "R" Then
             
            XPDtbTrans.value = ToGregorianDate(recordDateH.value)
               
        End If
End Sub

Private Sub ToDate_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         todateH.value = ToHijriDate(todate.value)
       
End If
End Sub

Private Sub ToDateH_LostFocus()
     If Me.TxtModFlg.Text <> "R" Then
             
            todate.value = ToGregorianDate(todateH.value)
               
        End If
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
    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = True

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

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

 
Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         recordDateH.value = ToHijriDate(XPDtbTrans.value)
       
End If
End Sub


Function PeintInstalMent(Optional InstallNo As Double, Optional nextinstalldate As Date, Optional nextinstalldateH As String, Optional ByVal contNo As Long = 0, Optional ByVal IsAll As Boolean = False)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    
    
    Dim newinstallNo  As Double
Dim nextinstalldate2 As Date
Dim nextinstalldateH2 As String
Dim ContNo2 As Double
Dim newinstallNo2 As Long
newinstallNo = 0
Dim RentValue As Double
   Dim VATValue As Double
   Dim mmValue As Double
    Dim mValueSt As String



    Dim Msg As String
    Dim i As Long
    Dim mMonth As Long
     Dim mYear As Long
    Dim rsDummy As New ADODB.Recordset
    Dim MySQL2  As String
    If IsAll Then
        MySQL = ""
        For i = 1 To GridInstallments.rows - 1
            ContNo2 = val(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("ContNo")))
            newinstallNo2 = val(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("InstallNo")))
            
            If i + 1 < GridInstallments.rows - 1 Then
            
            
            newinstallNo = val(GridInstallments.TextMatrix(i + 1, GridInstallments.ColIndex("InstallNo")))
            
            getnextDate newinstallNo, nextinstalldate2, nextinstalldateH2, ContNo2
            End If
            If i <> 1 Then
                MySQL = MySQL & "   Union all "
            End If
            
                nextinstalldate2 = CDate((GridInstallments.TextMatrix(i, GridInstallments.ColIndex("Due_Date"))))
                nextinstalldateH2 = CDate(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("nextinstalldate")))
                mMonth = DateDiff("m", nextinstalldate2, nextinstalldateH2)
                mYear = DateDiff("yyyy", nextinstalldate2, nextinstalldateH2)
                
                     MySQL = MySQL & "                     SELECT  '" & XPDtbTrans.value & "' as XPDtbTrans ," & mMonth & " as mMonth " & ", " & mYear & " as mYear " & ", '" & nextinstalldate2 & "' as installdateNext ,'" & nextinstalldateH2 & "' as installdateHNext,   dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblContract.ownerid, dbo.TblContract.UnitType, "
                    MySQL2 = "                     SELECT  '" & XPDtbTrans.value & "' as XPDtbTrans ," & mMonth & " as mMonth " & ", " & mYear & " as mYear " & ", '" & nextinstalldate2 & "' as installdateNext ,'" & nextinstalldateH2 & "' as installdateHNext,   dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblContract.ownerid, dbo.TblContract.UnitType, "
                RentValue = val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("RentValue")))
     VATValue = val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("VATValue")))
     mmValue = VATValue + RentValue
 mValueSt = WriteNo(Format(val(mmValue), "0.00"), 0, True, ".", , 0)
            '    MySQL = " SELECT    installdate,installdateH "
            '
            '    MySQL = MySQL & "      FROM    TblContractInstallments DDD     "
            ' , dbo.TblCustemers.VATNO as CusVatNo,TblCustemers.Cus_mobile,TblCustemers.RecordNo
            '    MySQL = MySQL & "        Where (DDD.ContNo = " & ContNo & ") And (DD.InstallNo =" & newinstallNo & ")"
            '
                MySQL = MySQL & "                  '" & mValueSt & "' mValueSt,TblContractInstallments.ID as mIID,TblContractInstallments.QrCodeImage, "
                MySQL = MySQL & "                   dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqarDetai.unitno, dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
                MySQL = MySQL & "                   dbo.TblCustemers.Fullcode, dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.TodateH,"
                MySQL = MySQL & "                   dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
                 MySQL = MySQL & "                      case IsNull(TblContractInstallments.NoteSerial1H,'') when '' then"
                MySQL = MySQL & "                  dbo.TblContract.NoteSerial1 + CAST(TblContractInstallments.InstallNo AS NVARCHAR(10))"
                MySQL = MySQL & "                  Else"
                MySQL = MySQL & "                  TblContractInstallments.NoteSerial1H"
                MySQL = MySQL & "                  end As NoteSerial1H2,"
                
                'dbo.TblContract.NoteSerial1 + cast (TblContractInstallments.InstallNo AS NVARCHAR(10)) AS NoteSerial1H ,"  ' 5555
                MySQL = MySQL & "                   dbo.TblContract.NetValue, dbo.TblContract.FATYou, dbo.TblContract.FATValue, dbo.TblContract.TotalValue, dbo.TblContractInstallments.*,"
                MySQL = MySQL & "                   dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.Address, dbo.TblCustemers.E_mail,"
                MySQL = MySQL & "                   dbo.TblCustemers.FaxNumber, dbo.TblCustemers.Remark2, dbo.TblCustemers.CustGID, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTitle,"
                MySQL = MySQL & "                   dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel, dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2,"
                MySQL = MySQL & "                   dbo.TblCustemers.BoxMil , dbo.TblCustemers.VATNO as CusVatNo,TblCustemers.Cus_mobile,TblCustemers.RecordNo,"
                MySQL = MySQL & "                   TblCustemers22.CusName OwnerName"
                MySQL = MySQL & "      FROM         dbo.TblContract LEFT OUTER JOIN"
                MySQL = MySQL & "                   dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN"
                MySQL = MySQL & "                   dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
                MySQL = MySQL & "                   dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
                MySQL = MySQL & "                   dbo.TblCustemers TblCustemers22  ON dbo.TblContract.ownerid= TblCustemers22.CusID LEFT OUTER JOIN"
                
                
                MySQL = MySQL & "                   dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
                MySQL = MySQL & "                   dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
                MySQL = MySQL & "                   dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
                
                MySQL = MySQL & "        Where (dbo.TblContract.ContNo = " & val(ContNo2) & ") And (dbo.TblContractInstallments.InstallNo =" & newinstallNo2 & ")"
        
                
                
                    
                    MySQL2 = "                     SELECT  '" & XPDtbTrans.value & "' as XPDtbTrans ," & mMonth & " as mMonth " & ", " & mYear & " as mYear " & ", '" & nextinstalldate2 & "' as installdateNext ,'" & nextinstalldateH2 & "' as installdateHNext,   dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblContract.ownerid, dbo.TblContract.UnitType, "
                RentValue = val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("RentValue")))
     VATValue = val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("VATValue")))
     mmValue = VATValue + RentValue
 mValueSt = WriteNo(Format(val(mmValue), "0.00"), 0, True, ".", , 0)
            '    MySQL = " SELECT    installdate,installdateH "
            '
            '    MySQL = MySQL & "      FROM    TblContractInstallments DDD     "
            '
            '    MySQL = MySQL & "        Where (DDD.ContNo = " & ContNo & ") And (DD.InstallNo =" & newinstallNo & ")"
            '
            
            
                MySQL2 = MySQL2 & "                  '" & mValueSt & "' mValueSt,TblContractInstallments.ID as mIID,TblContractInstallments.QrCodeImage, "
                MySQL2 = MySQL2 & "                                     dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqarDetai.unitno, dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
                MySQL2 = MySQL2 & "                                     dbo.TblCustemers.Fullcode, dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.TodateH,"
                MySQL2 = MySQL2 & "                                     dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "    ' 5555
                MySQL2 = MySQL2 & "                      case IsNull(TblContractInstallments.NoteSerial1H,'') when '' then"
                MySQL2 = MySQL2 & "                  dbo.TblContract.NoteSerial1 + CAST(TblContractInstallments.InstallNo AS NVARCHAR(10))"
                MySQL2 = MySQL2 & "                  Else"
                MySQL2 = MySQL2 & "                  TblContractInstallments.NoteSerial1H"
                MySQL2 = MySQL2 & "                  end As NoteSerial1H2,"
                
                MySQL2 = MySQL2 & "                                     dbo.TblContract.NetValue, dbo.TblContract.FATYou, dbo.TblContract.FATValue, dbo.TblContract.TotalValue, dbo.TblContractInstallments.*,"
                MySQL2 = MySQL2 & "                                     dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.Address, dbo.TblCustemers.E_mail,"
                MySQL2 = MySQL2 & "                                     dbo.TblCustemers.FaxNumber, dbo.TblCustemers.Remark2, dbo.TblCustemers.CustGID, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTitle,"
                MySQL2 = MySQL2 & "                                     dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel, dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2,"
                MySQL2 = MySQL2 & "                                     dbo.TblCustemers.BoxMil , dbo.TblCustemers.VATNO as CusVatNo,TblCustemers.Cus_mobile,TblCustemers.RecordNo,"
                MySQL2 = MySQL2 & "                                     TblCustemers22.CusName OwnerName"
                MySQL2 = MySQL2 & "                        FROM         dbo.TblContract LEFT OUTER JOIN"
                MySQL2 = MySQL2 & "                                     dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN"
                MySQL2 = MySQL2 & "                                     dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
                MySQL2 = MySQL2 & "                                     dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
                MySQL2 = MySQL2 & "                                     dbo.TblCustemers TblCustemers22  ON dbo.TblContract.ownerid= TblCustemers22.CusID LEFT OUTER JOIN"
                
                
                MySQL2 = MySQL2 & "                                     dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
                MySQL2 = MySQL2 & "                                     dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
                MySQL2 = MySQL2 & "                                     dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
                
                MySQL2 = MySQL2 & "                          Where (dbo.TblContract.ContNo = " & val(ContNo2) & ") And (dbo.TblContractInstallments.InstallNo =" & newinstallNo2 & ")"
                 
                
                rsDummy.Open MySQL2, Cn, adOpenForwardOnly, adLockReadOnly
                'Picture1.cl = -1
            SaveQRCode "TblContractInstallments", "ID", val(rsDummy!mIID & ""), Trim(rsDummy!NoteSerial1H & ""), (nextinstalldate2 & ""), CStr(val(rsDummy!RentValue & "") + val(rsDummy!VATValue & "")), Picture1, 0, CStr(rsDummy!VATValue & ""), CStr(val(rsDummy!RentValue & "") + val(rsDummy!VATValue & ""))
            rsDummy.Close
         Next
    
    Else
'    .TextMatrix(i, .ColIndex("NoteSerial1")) & .TextMatrix(i, .ColIndex("Installid"))
  newinstallNo = val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("Installid")))
        RentValue = val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("RentValue")))
     VATValue = val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("VATValue")))
     mmValue = VATValue + RentValue
 mValueSt = WriteNo(Format(val(mmValue), "0.00"), 0, True, ".", , 0)
    nextinstalldate = CDate(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("Due_Date")))
    nextinstalldateH = CDate(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("nextinstalldate")))
    mMonth = DateDiff("m", nextinstalldate, nextinstalldateH)
    mYear = DateDiff("yyyy", nextinstalldate, nextinstalldateH)
    
    MySQL = " SELECT  '" & XPDtbTrans.value & "' as XPDtbTrans ," & mMonth & " as mMonth " & ", " & mYear & " as mYear " & ",     '" & nextinstalldate & "' as installdateNext ,'" & nextinstalldateH & "' as installdateHNext,   dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblContract.ownerid, dbo.TblContract.UnitType, "
    
    
'    MySQL = " SELECT    installdate,installdateH "
'
'    MySQL = MySQL & "      FROM    TblContractInstallments DDD     "
'
'    MySQL = MySQL & "        Where (DDD.ContNo = " & ContNo & ") And (DD.InstallNo =" & newinstallNo & ")"
'

     
            
            

    
    MySQL = MySQL & "                  '" & mValueSt & "' mValueSt,TblContractInstallments.ID as mIID,TblContractInstallments.QrCodeImage, "
    MySQL = MySQL & "                   dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqarDetai.unitno, dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    MySQL = MySQL & "                   dbo.TblCustemers.Fullcode, dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.TodateH,"
    MySQL = MySQL & "                   dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
                     MySQL = MySQL & "                      case IsNull(TblContractInstallments.NoteSerial1H,'') when '' then"
                MySQL = MySQL & "                  dbo.TblContract.NoteSerial1 + CAST(TblContractInstallments.InstallNo AS NVARCHAR(10))"
                MySQL = MySQL & "                  Else"
                MySQL = MySQL & "                  TblContractInstallments.NoteSerial1H"
                MySQL = MySQL & "                  end As NoteSerial1H2,"
    MySQL = MySQL & "                   dbo.TblContract.NetValue, dbo.TblContract.FATYou, dbo.TblContract.FATValue, dbo.TblContract.TotalValue, dbo.TblContractInstallments.*,"
    MySQL = MySQL & "                   dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.Address, dbo.TblCustemers.E_mail,"
    MySQL = MySQL & "                   dbo.TblCustemers.FaxNumber, dbo.TblCustemers.Remark2, dbo.TblCustemers.CustGID, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTitle,"
    MySQL = MySQL & "                   dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel, dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2,"
    MySQL = MySQL & "                   dbo.TblCustemers.BoxMil , dbo.TblCustemers.VATNO CusVatNo,TblCustemers.Cus_mobile,TblCustemers.RecordNo,"
    MySQL = MySQL & "                   TblCustemers22.CusName OwnerName"
    MySQL = MySQL & "      FROM         dbo.TblContract LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblCustemers TblCustemers22  ON dbo.TblContract.ownerid= TblCustemers22.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
    MySQL = MySQL & "        Where (dbo.TblContract.ContNo = " & val(contNo) & ") And (dbo.TblContractInstallments.InstallNo =" & InstallNo & ")"
    
    
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    SaveQRCode "TblContractInstallments", "ID", val(RsData!mIID & ""), Trim(RsData!NoteSerial1H & ""), (nextinstalldate2 & ""), CStr(val(RsData!RentValue & "") + val(RsData!VATValue & "")), Picture1, 0, CStr(RsData!VATValue & ""), CStr(val(RsData!RentValue & "") + val(RsData!VATValue & ""))
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBiilRent.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBiilRent.rpt"
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
            Msg = "There's no data to show"
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
        StrReportTitle = "" '& StrAccountName
    Else
        StrReportTitle = ""
    End If
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function





Function savenewelectroncic()
   'vat data
'    Dim InvoiceTypeCodeID As Integer
'    rs("CIBAN").value = ""
'    'vat data
'      rs("RecTime").value = Time
'
'
'
'
' InvoiceTypeCodeID = 388
'
'  rs("InvoiceTypeCodeID").value = InvoiceTypeCodeID
'
'
'
'
'
'    If Export = 1 Then
'    rs("InvoiceTypeCodename").value = "0100100"
'    Else
'      rs("InvoiceTypeCodename").value = "0100000"
'   End If
'
'
'
'
'   Else
'    rs("InvoiceTypeCodename").value = "0200000"
'   End If
'
'   rs("DocumentCurrencyCode").value = Dccurrency.text
'   rs("TaxCurrencyCode").value = Dccurrency.text
'  rs("ActualDeliveryDate").value = txtDateRec.value
' rs("LatestDeliveryDate").value = txtDateRec.value
'Dim PaymentMeansCode As String
'
'            '10 In cash
'            '30 Credit
'            '42 Payment to bank account
'            '48 Bank card
'            '1 Instrument not defined(Free text)
'            Dim paymentnote
''        If CboPayMentType.ListIndex = 0 Then ' ‰ﬁœ«
''                  PaymentMeansCode = "10"
''                      paymentnote = "Payment by Cash"
''        ElseIf CboPayMentType.ListIndex = 1 Then ' ¬Ã·
''                 PaymentMeansCode = "30"
''                 paymentnote = "Payment by Credit"
''         ElseIf CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then  '  ÕÊÌ· »‰ﬂÌ
''                    If SystemOptions.AllowSalesMultyPayed = True Then
''                     PaymentMeansCode = "48" 'ﬂ«— 
''                      paymentnote = "Payment by Bank Card"
''                    Else
''                    PaymentMeansCode = "42" '»‰ﬂ Õ”«»
''                    paymentnote = "Payment by Bank Account"
''                    End If
''
''         End If
'         PaymentMeansCode = "30"
'                 paymentnote = "Payment by Credit"
'         rs("PaymentMeansCode").value = PaymentMeansCode
'
'rs("paymentnote").value = paymentnote
'rs.update
End Function



Private Sub SaveAlloca()
Dim RsDetails1 As ADODB.Recordset
Set RsDetails1 = New ADODB.Recordset
 Cn.Execute "delete tblContractInsAllocationsDetails1 where transID=" & val(Me.TxtTransID.Text)
StrSQL = "SELECT * FROM dbo.tblContractInsAllocationsDetails1 WHERE (1 = -1)"
RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

'Dim i As Integer
Dim TotalValue As Double
Dim TotalRent As Double
Dim TotalCommissions As Double
Dim TotalInsurance As Double
Dim TotalWater As Double
Dim TotalElectric As Double
Dim TotalTelandNet As Double

Dim StartDate As Date
Dim EndDate As Date
Dim TotalDays As Long
Dim DayValue As Double
Dim DayRent As Double
Dim DayCommissions As Double
Dim DayInsurance As Double
Dim DayWater As Double
Dim DayElectric As Double
Dim DayTelandNet As Double

Dim CurrentStart As Date
Dim MonthEnd As Date
Dim PartEnd As Date
Dim DaysPart As Long
Dim Countsofall As Long
Dim i As Long
Dim s As String
Dim mTransID As Long
s = "Select transID from tblContractInsAllocations where  Month(fromdate) =" & Month(Fromdate.value) & " And year(installdate) = " & year(Fromdate.value)
Dim rsDummy As New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
   mTransID = val(rsDummy!TransID & "")
End If
With Me.GridInstallments

    For i = 1 To .rows - 1

        ' ‰›” ‘—Êÿ «·«Œ Ì«— «·ﬁœÌ„…
        If val(.TextMatrix(i, .ColIndex("value"))) <> 0 _
           And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked _
           And val(.TextMatrix(i, .ColIndex("commisiontype"))) <> 1 Then

            ' ﬁ—«¡… «·ﬁÌ„ «·√’·Ì… „‰ «·Ã—Ìœ
            TotalValue = val(.TextMatrix(i, .ColIndex("value")))
            TotalRent = val(.TextMatrix(i, .ColIndex("RentValue")))
            TotalCommissions = val(.TextMatrix(i, .ColIndex("Commissions")))
            TotalInsurance = val(.TextMatrix(i, .ColIndex("Insurance")))
            TotalWater = val(.TextMatrix(i, .ColIndex("Water")))
            TotalElectric = val(.TextMatrix(i, .ColIndex("Electric")))
            TotalTelandNet = val(.TextMatrix(i, .ColIndex("TelandNet")))

            ' »œ«Ì… «·ﬁ”ÿ
            StartDate = .TextMatrix(i, .ColIndex("Due_Date"))
            ' ⁄œœ «·‘ÂÊ— „‰ «·Ã—Ìœ
            Countsofall = val(.TextMatrix(i, .ColIndex("Countsofall")))

            If Countsofall <= 0 Or TotalValue = 0 Then
                GoTo NextInstallment
            End If

            ' ‰Â«Ì… «·ﬁ”ÿ = »⁄œ Countsofall ‘ÂÊ— - ÌÊ„
            VBA.Calendar = vbCalGreg
            EndDate = DateAdd("m", Countsofall, StartDate)
            EndDate = DateAdd("d", -1, EndDate)

            ' ≈Ã„«·Ì ⁄œœ «·√Ì«„ ›Ì «·› —…
            TotalDays = DateDiff("d", StartDate, EndDate) + 1
            If TotalDays <= 0 Then
                GoTo NextInstallment
            End If

            ' ﬁÌ„… «·ÌÊ„ «·Ê«Õœ ·ﬂ· »‰œ
            DayValue = TotalValue / TotalDays
            DayRent = TotalRent / TotalDays
            DayCommissions = TotalCommissions / TotalDays
            DayInsurance = TotalInsurance / TotalDays
            DayWater = TotalWater / TotalDays
            DayElectric = TotalElectric / TotalDays
            DayTelandNet = TotalTelandNet / TotalDays

            ' ‰⁄œÌ ⁄·Ï «·› —… ‘Â— »‘Â—
            CurrentStart = StartDate

      Do While CurrentStart <= EndDate

    ' «·ÌÊ„ «·√Ê· ›Ì «·‘Â—
        Dim FirstOfMonth As Date
            FirstOfMonth = DateSerial(year(CurrentStart), Month(CurrentStart), 1)
        
            ' «·ÌÊ„ «·√ŒÌ— ›Ì «·‘Â—
            Dim LastOfMonth As Date
            LastOfMonth = DateSerial(year(CurrentStart), Month(CurrentStart) + 1, 0)
        
            ' «·Ã“¡ «·›⁄·Ì „‰ «·‘Â— œ«Œ· › —… «·ﬁ”ÿ
            Dim PeriodStart As Date: PeriodStart = CurrentStart
            Dim PeriodEnd As Date: PeriodEnd = IIf(LastOfMonth < EndDate, LastOfMonth, EndDate)
        
            DaysPart = DateDiff("d", PeriodStart, PeriodEnd) + 1
        
            RsDetails1.AddNew
              RsDetails1("transID").value = mTransID 'Me.TxtTransID.text
              RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
              RsDetails1("NoteSerial").value = .TextMatrix(i, .ColIndex("NoteSerial1"))
              RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
        
              RsDetails1("Installdate").value = PeriodStart
              RsDetails1("InstalldateH").value = ToHijriDate(PeriodStart)
                                            RsDetails1("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
              RsDetails1("installValue").value = Round(DayValue * DaysPart, 2)
              RsDetails1("RentValue").value = Round(DayRent * DaysPart, 2)
              RsDetails1("Commissions").value = Round(DayCommissions * DaysPart, 2)
              RsDetails1("Insurance").value = Round(DayInsurance * DaysPart, 2)
              RsDetails1("Water").value = Round(DayWater * DaysPart, 2)
              RsDetails1("Electric").value = Round(DayElectric * DaysPart, 2)
              RsDetails1("TelandNet").value = Round(DayTelandNet * DaysPart, 2)
        
            RsDetails1.update
        
            ' «·«‰ ﬁ«· ·√Ê· ÌÊ„ ›Ì «·‘Â— «· «·Ì
            CurrentStart = DateAdd("m", 1, FirstOfMonth)

    Loop

NextInstallment:

        Else
            ' ·Ê „‘ „ ⁄·„ √Ê «·ﬁÌ„… = 0 √Ê ‰Ê⁄ «·⁄„Ê·… = 1° »‰⁄œ¯Ì
        End If

    Next i

End With

RsDetails1.Close
'***********************************************************************************************

'************************
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap

    Dim movementIds As Collection
    Dim i As Long
    Dim movementCount As Long
    Dim loadedCount As Long
    Dim patchedMovements As Long
    Dim patchedLinesTotal As Long
    Dim skippedMovements As Long
    Dim failedLoads As Long
    Dim patchResult As Long
    Dim reportText As String

    Set movementIds = GetMovementIDsInRange(reportText)
    If movementIds Is Nothing Then Exit Sub

    If movementIds.count = 0 Then
        MsgBox "·«  ÊÃœ Õ—ﬂ«  —∆Ì”Ì… œ«Œ· «·› —… «·„Õœœ….", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    movementCount = movementIds.count

    For i = 1 To movementIds.count
        If LoadMovementByID(CLng(movementIds(i)), reportText) Then
            loadedCount = loadedCount + 1
            DoEvents

            patchResult = PatchCurrentLoadedMovementSafely(reportText)

            If patchResult > 0 Then
                patchedMovements = patchedMovements + 1
                patchedLinesTotal = patchedLinesTotal + patchResult
            Else
                skippedMovements = skippedMovements + 1
            End If
        Else
            failedLoads = failedLoads + 1
        End If

        DoEvents
    Next i

    MsgBox "⁄œœ «·Õ—ﬂ«  œ«Œ· «·› —…: " & movementCount & vbCrLf & _
           "⁄œœ «·Õ—ﬂ«  «· Ì  „  Õ„Ì·Â«: " & loadedCount & vbCrLf & _
           "⁄œœ «·Õ—ﬂ«  «· Ì  „  ⁄œÌ·Â«: " & patchedMovements & vbCrLf & _
           "⁄œœ ”ÿÊ— «·ﬁÌœ «·„⁄œ·…: " & patchedLinesTotal & vbCrLf & _
           "⁄œœ «·Õ—ﬂ«  «·„ Œÿ«…: " & skippedMovements & vbCrLf & _
           "⁄œœ «·Õ—ﬂ«  «· Ì ›‘·  Õ„Ì·Â«: " & failedLoads, _
           vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

    If Trim$(reportText) <> "" Then
        Debug.Print "===== Batch Safe Patch Report ====="
        Debug.Print reportText
    End If

    Exit Sub

ErrTrap:
    MsgBox "ÕœÀ Œÿ√ √À‰«¡  ‰›Ì– «· ÕœÌÀ «·œ›⁄Ì." & vbCrLf & Err.Description, _
           vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Function GetMovementIDsInRange(ByRef reportText As String) As Collection
    On Error GoTo ErrTrap

    Dim rsRange As ADODB.Recordset
    Dim sql As String
    Dim fromDateValue As Date
    Dim toDateValue As Date
    Dim result As Collection

    fromDateValue = CDate(txtFromDate.value)
    toDateValue = CDate(txtToDate.value)

    If DateValue(fromDateValue) > DateValue(toDateValue) Then
        MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰  «—ÌŒ „‰ √ﬂ»— „‰  «—ÌŒ ≈·Ï.", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    sql = "SELECT transID " & _
          "FROM dbo.tblContractInsAllocations " & _
          "WHERE RecordDate >= " & SQLDate(fromDateValue, True) & " " & _
          "AND RecordDate <= " & SQLDate(toDateValue, True) & " " & _
          "ORDER BY RecordDate, transID"

    Set rsRange = New ADODB.Recordset
    rsRange.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Set result = New Collection

    Do While Not rsRange.EOF
        result.Add CLng(val("" & rsRange.Fields("transID").value))
        rsRange.MoveNext
    Loop

    rsRange.Close
    Set rsRange = Nothing

    Set GetMovementIDsInRange = result
    Exit Function

ErrTrap:
    MsgBox "ÕœÀ Œÿ√ √À‰«¡ Ã·» «·Õ—ﬂ«  œ«Œ· «·› —…." & vbCrLf & Err.Description, _
           vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

Private Function LoadMovementByID(ByVal movementId As Long, ByRef reportText As String) As Boolean
    On Error GoTo ErrTrap

    If rs Is Nothing Then
        reportText = reportText & "Movement " & movementId & ": rs €Ì— „ÂÌ√." & vbCrLf
        Exit Function
    End If

    If rs.RecordCount <= 0 Then
        reportText = reportText & "Movement " & movementId & ": ·«  ÊÃœ ”Ã·«  ›Ì „’œ— »Ì«‰«  «·‘«‘…." & vbCrLf
        Exit Function
    End If

    rs.MoveFirst
    rs.Find "transID=" & movementId, , adSearchForward, 1

    If rs.EOF Then
        reportText = reportText & "Movement " & movementId & ": ·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·Õ—ﬂ… ›Ì rs." & vbCrLf
        Exit Function
    End If

    Retrive
    DoEvents

    If val(Me.TxtTransID.Text) <> movementId Then
        reportText = reportText & "Movement " & movementId & ": «·‘«‘… ·„  Õ„· ‰›” —ﬁ„ «·Õ—ﬂ…." & vbCrLf
        Exit Function
    End If

    If Trim$(Me.TxtNoteID.Text & "") = "" Then
        reportText = reportText & "Movement " & movementId & ":  „  Õ„Ì· «·Õ—ﬂ… ·ﬂ‰ TXTNoteID ›«—€." & vbCrLf
        Exit Function
    End If

    If Me.GridInstallments.rows <= Me.GridInstallments.FixedRows Then
        reportText = reportText & "Movement " & movementId & ":  „  Õ„Ì· «·Õ—ﬂ… ·ﬂ‰ GridInstallments ›«—€." & vbCrLf
        Exit Function
    End If

    LoadMovementByID = True
    Exit Function

ErrTrap:
    reportText = reportText & "Movement " & movementId & ": Œÿ√ √À‰«¡  Õ„Ì· «·”Ã·. " & Err.Description & vbCrLf
End Function

Private Function PatchCurrentLoadedMovementSafely(ByRef reportText As String) As Long
    On Error GoTo ErrTrap

    Dim BeginTrans As Boolean
    Dim rsV As ADODB.Recordset
    Dim sql As String
    Dim noteId As Long
    Dim matchedRow As Long
    Dim aqarId As Long
    Dim unitTypeId As Long
    Dim unitId As Long
    Dim reasonText As String
    Dim descText As String
    Dim patchedCount As Long
    Dim matchStage As Integer
    Dim voucherValue As Double

    noteId = val(Me.TxtNoteID.Text)

    If noteId = 0 Then
        reportText = reportText & "Movement " & Me.TxtTransID.Text & ": ·« ÌÊÃœ —ﬁ„ ﬁÌœ „— »ÿ »«·Õ—ﬂ… «·„Õ„·…." & vbCrLf
        Exit Function
    End If

    If Me.GridInstallments.rows <= Me.GridInstallments.FixedRows Then
        reportText = reportText & "Movement " & Me.TxtTransID.Text & ": ·«  ÊÃœ Õ—ﬂ«  ›—⁄Ì… „⁄—Ê÷… ⁄·Ï «·‘«‘…." & vbCrLf
        Exit Function
    End If

    sql = "SELECT Double_Entry_Vouchers_ID, DEV_ID_Line_No, Account_Code, Value, " & _
          "Double_Entry_Vouchers_Description, Aqarid, unittype, unitno, uintid " & _
          "FROM dbo.DOUBLE_ENTREY_VOUCHERS " & _
          "WHERE Notes_ID = " & noteId & " " & _
          "AND (ISNULL(Aqarid,0)=0 OR ISNULL(unittype,0)=0 OR ISNULL(unitno,0)=0 OR ISNULL(uintid,0)=0) " & _
          "ORDER BY DEV_ID_Line_No"

    Set rsV = New ADODB.Recordset
    rsV.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rsV.EOF Then
        rsV.Close
        Set rsV = Nothing
        Exit Function
    End If

    Cn.BeginTrans
    BeginTrans = True

    Do While Not rsV.EOF
        matchedRow = 0
        aqarId = 0
        unitTypeId = 0
        unitId = 0
        reasonText = ""

        descText = "" & rsV.Fields("Double_Entry_Vouchers_Description").value
        voucherValue = GetVoucherLineValue(rsV)

        matchStage = MatchVoucherLineToSingleMovement(descText, voucherValue, matchedRow, aqarId, reasonText, rsV)

        If matchStage = 1 Or matchStage = 2 Then
            If TryGetMovementIds(matchedRow, aqarId, unitTypeId, unitId, reasonText) Then
                If ApplySafeVoucherPatch(rsV, aqarId, unitTypeId, unitId) Then
                    patchedCount = patchedCount + 1
                End If
            Else
                reportText = reportText & "Movement " & Me.TxtTransID.Text & " / Line " & val("" & rsV.Fields("DEV_ID_Line_No").value) & ": " & reasonText & vbCrLf
            End If
        ElseIf matchStage = 3 Or matchStage = 4 Then
            If aqarId > 0 Then
                If ApplyAqarOnlyPatch(rsV, aqarId) Then
                    patchedCount = patchedCount + 1
                End If
            Else
                reportText = reportText & "Movement " & Me.TxtTransID.Text & " / Line " & val("" & rsV.Fields("DEV_ID_Line_No").value) & ": " & reasonText & vbCrLf
            End If
        Else
            reportText = reportText & "Movement " & Me.TxtTransID.Text & " / Line " & val("" & rsV.Fields("DEV_ID_Line_No").value) & ": " & reasonText & vbCrLf
        End If

        rsV.MoveNext
    Loop

    Cn.CommitTrans
    BeginTrans = False

    rsV.Close
    Set rsV = Nothing

    PatchCurrentLoadedMovementSafely = patchedCount
    Exit Function

ErrTrap:
    If BeginTrans Then
        Cn.RollbackTrans
    End If

    reportText = reportText & "Movement " & Me.TxtTransID.Text & ": Œÿ√ √À‰«¡ SAFE patch. " & Err.Description & vbCrLf
End Function

Private Function MatchVoucherLineToSingleMovement(ByVal voucherDesc As String, ByVal voucherValue As Double, ByRef matchedRow As Long, ByRef aqarId As Long, ByRef reasonText As String, ByRef rsCurrentVoucherLine As ADODB.Recordset) As Integer
    Dim primaryRow As Long
    Dim primaryCount As Long
    Dim fallbackRow As Long
    Dim fallbackCount As Long
    Dim descNorm As String
    Dim descSerial As String
    Dim propertyName As String
    Dim resolvedAqarId As Long
    Dim i As Long

    matchedRow = 0
    aqarId = 0
    reasonText = ""

    descNorm = NormalizeArabicText(voucherDesc)
    descSerial = ExtractDigits(descNorm)

    If descSerial <> "" Then
        With Me.GridInstallments
            For i = .FixedRows To .rows - 1
                If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                    If IsPrimaryMovementMatch(i, descNorm, descSerial) Then
                        primaryCount = primaryCount + 1
                        primaryRow = i
                    End If
                End If
            Next i
        End With
    End If

    If primaryCount = 1 Then
        matchedRow = primaryRow
        MatchVoucherLineToSingleMovement = 1
        Exit Function
    End If

    If primaryCount > 1 Then
        reasonText = " „ «·⁄ÀÊ— ⁄·Ï √ﬂÀ— „‰ Õ—ﬂ… „ÿ«»ﬁ… »«·—»ÿ «·√”«”Ì° ·–·ﬂ  „ «· ŒÿÌ Õ›«Ÿ« ⁄·Ï ”·«„… «·»Ì«‰« ."
        Exit Function
    End If

    With Me.GridInstallments
        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                If IsFallbackTenantValueMatch(i, descNorm, voucherValue) Then
                    fallbackCount = fallbackCount + 1
                    fallbackRow = i
                End If
            End If
        Next i
    End With

    If fallbackCount = 1 Then
        matchedRow = fallbackRow
        MatchVoucherLineToSingleMovement = 2
        Exit Function
    End If

    propertyName = ExtractAqarNameFromDescription(voucherDesc)
    If propertyName <> "" Then
        If ResolveAqarIDByName(propertyName, resolvedAqarId, reasonText) Then
            aqarId = resolvedAqarId
            MatchVoucherLineToSingleMovement = 3
            Exit Function
        End If
    End If

    If ResolveAqarIDFromVoucherAccount(rsCurrentVoucherLine, resolvedAqarId, reasonText) Then
        aqarId = resolvedAqarId
        MatchVoucherLineToSingleMovement = 4
        Exit Function
    End If

    If reasonText = "" Then
        reasonText = "·« ÌÊÃœ  ÿ«»ﬁ „ÊÀÊﬁ ›Ì √Ì „—Õ·… „‰ „—«Õ· «·—»ÿ."
    End If
End Function

Private Function IsPrimaryMovementMatch(ByVal Row As Long, ByVal descNorm As String, ByVal descSerial As String) As Boolean
    Dim rowSerial As String
    Dim dueDateG As String
    Dim dueDateH As String
    Dim serialHit As Boolean
    Dim dateHit As Boolean

    With Me.GridInstallments
        rowSerial = ExtractDigits("" & .TextMatrix(Row, .ColIndex("NoteSerial1")))
        dueDateG = NormalizeArabicText("" & .TextMatrix(Row, .ColIndex("Due_Date")))
        dueDateH = NormalizeArabicText("" & .TextMatrix(Row, .ColIndex("Due_DateH")))
    End With

    serialHit = (rowSerial <> "" And rowSerial = descSerial)

    dateHit = False
    If dueDateG <> "" Then dateHit = ContainsToken(descNorm, dueDateG)
    If Not dateHit And dueDateH <> "" Then dateHit = ContainsToken(descNorm, dueDateH)

    IsPrimaryMovementMatch = (serialHit And dateHit)
End Function

Private Function IsFallbackTenantValueMatch(ByVal Row As Long, ByVal descNorm As String, ByVal voucherValue As Double) As Boolean
    Dim cusName As String
    Dim customerHit As Boolean
    Dim rowValue As Double

    With Me.GridInstallments
        cusName = NormalizeArabicText("" & .TextMatrix(Row, .ColIndex("CusName")))
        rowValue = GetGridRowComparableValue(Row, descNorm)
    End With

    customerHit = (cusName <> "" And ContainsToken(descNorm, cusName))
    IsFallbackTenantValueMatch = (customerHit And AmountsAreEqual(voucherValue, rowValue))
End Function

Private Function GetGridRowComparableValue(ByVal Row As Long, ByVal descNorm As String) As Double
    With Me.GridInstallments
        If ContainsToken(descNorm, "⁄„Ê·…") Or ContainsToken(descNorm, "⁄„Ê·« ") Then
            GetGridRowComparableValue = CDbl(val("" & .TextMatrix(Row, .ColIndex("Commissions"))))
        ElseIf ContainsToken(descNorm, " √„Ì‰") Then
            GetGridRowComparableValue = CDbl(val("" & .TextMatrix(Row, .ColIndex("Insurance"))))
        ElseIf ContainsToken(descNorm, "„Ì«Â") Or ContainsToken(descNorm, "„Ì«…") Then
            GetGridRowComparableValue = CDbl(val("" & .TextMatrix(Row, .ColIndex("Water"))))
        ElseIf ContainsToken(descNorm, "ﬂÂ—»«¡") Then
            GetGridRowComparableValue = CDbl(val("" & .TextMatrix(Row, .ColIndex("Electric"))))
        ElseIf ContainsToken(descNorm, "Œœ„« ") Or ContainsToken(descNorm, "Â« ›") Or ContainsToken(descNorm, "‰ ") Then
            GetGridRowComparableValue = CDbl(val("" & .TextMatrix(Row, .ColIndex("TelandNet"))))
        ElseIf ContainsToken(descNorm, "ﬁÌ„… „÷«›…") Or ContainsToken(descNorm, "«·ﬁÌ„Â «·„÷«›…") Or ContainsToken(descNorm, "«·ﬁÌ„… «·„÷«›…") Then
            GetGridRowComparableValue = CDbl(val("" & .TextMatrix(Row, .ColIndex("VATValue"))))
        Else
            GetGridRowComparableValue = CDbl(val("" & .TextMatrix(Row, .ColIndex("RentValue"))))
        End If
    End With
End Function

Private Function AmountsAreEqual(ByVal value1 As Double, ByVal value2 As Double, Optional ByVal tolerance As Double = 0.01) As Boolean
    AmountsAreEqual = (Abs(value1 - value2) <= tolerance)
End Function

Private Function GetVoucherLineValue(ByRef rsV As ADODB.Recordset) As Double
    GetVoucherLineValue = CDbl(val("" & rsV.Fields("Value").value))
End Function

Private Function ExtractAqarNameFromDescription(ByVal voucherDesc As String) As String
    Dim txt As String
    Dim p As Long
    Dim s As String
    Dim stopPos As Long
    Dim marker As Variant
    Dim markers As Variant

    txt = Replace(voucherDesc, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, ":", " ")
    txt = Replace(txt, "°", " ")

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    p = InStr(1, txt, "«·⁄ﬁ«—", vbTextCompare)
    If p = 0 Then Exit Function

    s = Trim$(mId$(txt, p + Len("«·⁄ﬁ«—")))

    markers = Array("«·ÊÕœ…", "«·„” √Ã—", "«·„«·ﬂ", "«·ﬁÌ„Â", "«·ﬁÌ„…", "«ÌÃ«—", "≈ÌÃ«—")

    stopPos = 0
    For Each marker In markers
        p = InStr(1, s, CStr(marker), vbTextCompare)
        If p > 0 Then
            If stopPos = 0 Or p < stopPos Then stopPos = p
        End If
    Next marker

    If stopPos > 0 Then
        s = left$(s, stopPos - 1)
    End If

    s = NormalizeArabicText(s)
    ExtractAqarNameFromDescription = Trim$(s)
End Function

Private Function ResolveAqarIDByName(ByVal aqarName As String, ByRef aqarId As Long, ByRef reasonText As String) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim sql As String
    Dim matchCount As Long
    Dim gridCount As Long
    Dim gridAqarId As Long
    Dim i As Long
    Dim rowAqarName As String

    aqarId = 0
    reasonText = ""

    If Trim$(aqarName) = "" Then
        reasonText = " ⁄–— «” Œ—«Ã «”„ «·⁄ﬁ«— „‰ Ê’› «·”ÿ—."
        Exit Function
    End If

    With Me.GridInstallments
        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                rowAqarName = NormalizeArabicText("" & .TextMatrix(i, .ColIndex("aqarname")))
                If rowAqarName <> "" Then
                    If rowAqarName = NormalizeArabicText(aqarName) Or ContainsToken(rowAqarName, aqarName) Or ContainsToken(aqarName, rowAqarName) Then
                        gridCount = gridCount + 1
                        If gridCount = 1 Then
                            gridAqarId = val("" & .TextMatrix(i, .ColIndex("Iqar")))
                        End If
                    End If
                End If
            End If
        Next i
    End With

    If gridCount = 1 And gridAqarId > 0 Then
        aqarId = gridAqarId
        ResolveAqarIDByName = True
        Exit Function
    ElseIf gridCount > 1 Then
        reasonText = "«”„ «·⁄ﬁ«— „‰ Ê’› «·”ÿ— Ìÿ«»ﬁ √ﬂÀ— „‰ ⁄ﬁ«— œ«Œ· «·Õ—ﬂ… «·Õ«·Ì…."
        Exit Function
    End If

    sql = "SELECT Aqarid, aqarname " & _
          "FROM dbo.TblAqar " & _
          "WHERE LTRIM(RTRIM(ISNULL(aqarname,''))) = N'" & Replace(Trim$(aqarName), "'", "''") & "'"

    Set rsTmp = New ADODB.Recordset
    rsTmp.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsTmp.EOF Then
        rsTmp.Close
        sql = "SELECT Aqarid, aqarname " & _
              "FROM dbo.TblAqar " & _
              "WHERE ISNULL(aqarname,'') LIKE N'%" & Replace(Trim$(aqarName), "'", "''") & "%'"
        rsTmp.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    End If

    If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            matchCount = matchCount + 1
            If matchCount = 1 Then
                aqarId = CLng(val("" & rsTmp.Fields("Aqarid").value))
            End If
            rsTmp.MoveNext
        Loop
    End If

    rsTmp.Close
    Set rsTmp = Nothing

    If matchCount = 1 Then
        ResolveAqarIDByName = True
    ElseIf matchCount = 0 Then
        reasonText = "·„ Ì „ «·⁄ÀÊ— ⁄·Ï ⁄ﬁ«— „ÿ«»ﬁ „‰ Ê’› «·”ÿ—."
    Else
        reasonText = "«”„ «·⁄ﬁ«— „‰ Ê’› «·”ÿ— Ìÿ«»ﬁ √ﬂÀ— „‰ ⁄ﬁ«—° ·–·ﬂ  „ «· ŒÿÌ."
    End If
End Function

Private Function ResolveAqarIDFromVoucherAccount(ByRef rsCurrentVoucherLine As ADODB.Recordset, ByRef aqarId As Long, ByRef reasonText As String) As Boolean
    Dim rsCus As ADODB.Recordset
    Dim rsAqar As ADODB.Recordset
    Dim sql As String
    Dim accountCode As String
    Dim cusId As Long
    Dim cusCount As Long
    Dim distinctCount As Long

    aqarId = 0
    reasonText = ""

    accountCode = Trim$("" & rsCurrentVoucherLine.Fields("Account_Code").value)
    If accountCode = "" Then
        reasonText = "Account_Code €Ì— „ÊÃÊœ ›Ì ”ÿ— «·ﬁÌœ."
        Exit Function
    End If

    sql = "SELECT DISTINCT CusID " & _
          "FROM dbo.TblCustemers " & _
          "WHERE LTRIM(RTRIM(ISNULL(Account_Code,''))) = N'" & Replace(accountCode, "'", "''") & "'"

    Set rsCus = New ADODB.Recordset
    rsCus.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Do While Not rsCus.EOF
        cusCount = cusCount + 1
        If cusCount = 1 Then
            cusId = CLng(val("" & rsCus.Fields("CusID").value))
        End If
        rsCus.MoveNext
    Loop

    rsCus.Close
    Set rsCus = Nothing

    If cusCount = 0 Then
        reasonText = "·„ Ì „ «·⁄ÀÊ— ⁄·Ï ⁄„Ì· „ÿ«»ﬁ „‰ Account_Code."
        Exit Function
    End If

    If cusCount > 1 Then
        reasonText = "Account_Code Ìÿ«»ﬁ √ﬂÀ— „‰ ⁄„Ì·° ·–·ﬂ  „ «· ŒÿÌ."
        Exit Function
    End If

    If cusId = 0 Then
        reasonText = " ⁄–—  ÕœÌœ CusID „‰ Account_Code."
        Exit Function
    End If

    sql = "SELECT DISTINCT Iqar " & _
          "FROM dbo.TblContract " & _
          "WHERE CusID = " & cusId & " " & _
          "AND ISNULL(Iqar,0) <> 0"

    Set rsAqar = New ADODB.Recordset
    rsAqar.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Do While Not rsAqar.EOF
        distinctCount = distinctCount + 1
        If distinctCount = 1 Then
            aqarId = CLng(val("" & rsAqar.Fields("Iqar").value))
        End If
        rsAqar.MoveNext
    Loop

    rsAqar.Close
    Set rsAqar = Nothing

    If distinctCount = 1 And aqarId > 0 Then
        ResolveAqarIDFromVoucherAccount = True
    ElseIf distinctCount = 0 Then
        reasonText = "«·⁄„Ì· «·„” Œ—Ã „‰ Account_Code ·Ì” ·Â ⁄ﬁ«— Ê«Õœ ’«·Õ Ì„ﬂ‰ «⁄ „«œÂ."
    Else
        reasonText = "«·⁄„Ì· «·„” Œ—Ã „‰ Account_Code „— »ÿ »√ﬂÀ— „‰ ⁄ﬁ«—° ·–·ﬂ  „ «· ŒÿÌ."
    End If
End Function

Private Function TryGetMovementIds(ByVal Row As Long, ByRef aqarId As Long, ByRef unitTypeId As Long, ByRef unitId As Long, ByRef reasonText As String) As Boolean
    Dim contNo As Long
    Dim gridAqarId As Long

    aqarId = 0
    unitTypeId = 0
    unitId = 0

    With Me.GridInstallments
        contNo = val("" & .TextMatrix(Row, .ColIndex("ContNo")))
        gridAqarId = val("" & .TextMatrix(Row, .ColIndex("Iqar")))
    End With

    If contNo = 0 Then
        reasonText = "«·Õ—ﬂ… «·„ÿ«»ﬁ… ·«  Õ ÊÌ ⁄·Ï —ﬁ„ ⁄ﬁœ."
        Exit Function
    End If

    GetContractUnitIds contNo, aqarId, unitTypeId, unitId

    If aqarId = 0 Then aqarId = gridAqarId

    If aqarId = 0 Then
        reasonText = " ⁄–—  ÕœÌœ «·⁄ﬁ«— „‰ «·Õ—ﬂ… √Ê «·⁄ﬁœ."
        Exit Function
    End If

    If unitId = 0 Then
        reasonText = " ⁄–—  ÕœÌœ —ﬁ„  ⁄—Ì› «·ÊÕœ… „‰ «·⁄ﬁœ."
        Exit Function
    End If

    TryGetMovementIds = True
End Function

Private Sub GetContractUnitIds(ByVal contNo As Long, ByRef aqarId As Long, ByRef unitTypeId As Long, ByRef unitId As Long)
    Dim rsTmp As ADODB.Recordset
    Dim sql As String

    aqarId = 0
    unitTypeId = 0
    unitId = 0

    sql = "SELECT TOP 1 Iqar, UnitType, UnitNo " & _
          "FROM dbo.TblContract WHERE ContNo = " & contNo

    Set rsTmp = New ADODB.Recordset
    rsTmp.Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rsTmp.EOF Then
        aqarId = val("" & rsTmp.Fields("Iqar").value)
        unitTypeId = val("" & rsTmp.Fields("UnitType").value)
        unitId = val("" & rsTmp.Fields("UnitNo").value)
    End If

    rsTmp.Close
    Set rsTmp = Nothing
End Sub

Private Function ApplySafeVoucherPatch(ByRef rsV As ADODB.Recordset, ByVal aqarId As Long, ByVal unitTypeId As Long, ByVal unitId As Long) As Boolean
    Dim changed As Boolean

    changed = False

    If val("" & rsV.Fields("Aqarid").value) = 0 And aqarId > 0 Then
        rsV.Fields("Aqarid").value = aqarId
        changed = True
    End If

    If val("" & rsV.Fields("unittype").value) = 0 And unitTypeId > 0 Then
        rsV.Fields("unittype").value = unitTypeId
        changed = True
    End If

    If val("" & rsV.Fields("unitno").value) = 0 And unitId > 0 Then
        rsV.Fields("unitno").value = unitId
        changed = True
    End If

    If val("" & rsV.Fields("uintid").value) = 0 And unitId > 0 Then
        rsV.Fields("uintid").value = unitId
        changed = True
    End If

    If changed Then
        rsV.update
        ApplySafeVoucherPatch = True
    End If
End Function

Private Function ApplyAqarOnlyPatch(ByRef rsV As ADODB.Recordset, ByVal aqarId As Long) As Boolean
    Dim changed As Boolean

    changed = False

    If val("" & rsV.Fields("Aqarid").value) = 0 And aqarId > 0 Then
        rsV.Fields("Aqarid").value = aqarId
        changed = True
    End If

    If changed Then
        rsV.update
        ApplyAqarOnlyPatch = True
    End If
End Function

Private Function NormalizeArabicText(ByVal txt As String) As String
    txt = UCase$(Trim$(txt))
    txt = Replace(txt, "⁄ﬁœ —ﬁ„", "")
    txt = Replace(txt, "⁄ﬁœ", "")
    txt = Replace(txt, "—ﬁ„", "")
    txt = Replace(txt, ":", " ")
    txt = Replace(txt, "°", " ")
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, vbTab, " ")

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    NormalizeArabicText = Trim$(txt)
End Function

Private Function ExtractDigits(ByVal txt As String) As String
    Dim i As Long
    Dim ch As String
    Dim resultText As String

    txt = NormalizeArabicText(txt)

    For i = 1 To Len(txt)
        ch = mId$(txt, i, 1)
        If ch >= "0" And ch <= "9" Then
            resultText = resultText & ch
        End If
    Next i

    ExtractDigits = resultText
End Function

Private Function ContainsToken(ByVal sourceText As String, ByVal tokenText As String) As Boolean
    sourceText = NormalizeArabicText(sourceText)
    tokenText = NormalizeArabicText(tokenText)

    If Trim$(tokenText) = "" Then Exit Function
    ContainsToken = (InStr(1, sourceText, tokenText, vbTextCompare) > 0)
End Function




