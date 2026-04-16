VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmAccountsSeetting 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«⁄œ«œ«    ﬂÊÌœ «·Õ”«»«  "
   ClientHeight    =   8955
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10035
   HelpContextID   =   580
   Icon            =   "frmAccountsSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   10035
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
      Height          =   8985
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10125
      _cx             =   17859
      _cy             =   15849
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
         Left            =   -90
         TabIndex        =   1
         Top             =   750
         Width           =   10185
         _cx             =   17965
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
         Caption         =   " ﬂÊÌœ «·Õ”«»« |«⁄œ«œ  «·Õ”«»« |‰ﬁ· «·Õ”«»« "
         Align           =   0
         CurrTab         =   1
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
            Left            =   -10740
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   10095
            _cx             =   17806
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6795
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   15225
               _cx             =   26855
               _cy             =   11986
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
               Begin VB.TextBox TXTRemark1 
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
                  Height          =   870
                  Left            =   10320
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   66
                  Top             =   2850
                  Visible         =   0   'False
                  Width           =   2760
               End
               Begin VB.TextBox TXTNO 
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
                  Left            =   6120
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1200
               End
               Begin VB.Frame Frame1 
                  Caption         =   "„⁄·Ê„« "
                  Height          =   2805
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   1395
                  Width           =   4695
                  Begin MSDataListLib.DataCombo DcBranch 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   53
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
                     TabIndex        =   54
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
                     TabIndex        =   57
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
                     TabIndex        =   56
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
                     TabIndex        =   55
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
                     TabIndex        =   52
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
                     TabIndex        =   51
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
                     TabIndex        =   50
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
                     TabIndex        =   49
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
                     TabIndex        =   48
                     Top             =   840
                     Width           =   1200
                  End
               End
               Begin VB.TextBox txtRemarks 
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
                  Left            =   6000
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   42
                  Top             =   45
                  Width           =   2400
               End
               Begin VB.CheckBox ChkLocked 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  Height          =   810
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   2910
                  Width           =   2310
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
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   3825
                  Value           =   -1  'True
                  Width           =   1095
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
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   3825
                  Width           =   1695
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   6255
                  Width           =   1590
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
                  Height          =   870
                  Left            =   10800
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Text            =   "0"
                  Top             =   2910
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox TxtyearsdataId 
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
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   -15
                  Visible         =   0   'False
                  Width           =   1200
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   255
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   6375
                  Width           =   2310
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
                  Left            =   -3930
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   16575
                  Width           =   2175
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
                  Left            =   2235
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   6090
                  Left            =   135
                  TabIndex        =   7
                  Top             =   585
                  Width           =   9825
                  _cx             =   17330
                  _cy             =   10742
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
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmAccountsSettings.frx":038A
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
               Begin MSComCtl2.DTPicker DbFromDate 
                  Height          =   285
                  Left            =   7560
                  TabIndex        =   12
                  Top             =   2340
                  Visible         =   0   'False
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   244318209
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   10440
                  TabIndex        =   13
                  Top             =   3705
                  Width           =   4365
                  _ExtentX        =   7699
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
                  Left            =   11760
                  TabIndex        =   15
                  Top             =   2670
                  Width           =   1605
                  _ExtentX        =   2831
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
                  Left            =   10680
                  TabIndex        =   30
                  Top             =   1395
                  Width           =   3285
                  _ExtentX        =   5794
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
                  Height          =   870
                  Left            =   10440
                  TabIndex        =   37
                  Top             =   2790
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   1535
                  _Version        =   393216
                  Format          =   244318209
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   10560
                  TabIndex        =   43
                  Top             =   3825
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
                  ButtonImage     =   "frmAccountsSettings.frx":04EA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   11280
                  TabIndex        =   44
                  Top             =   3825
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
                  ButtonImage     =   "frmAccountsSettings.frx":0884
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DCEmp 
                  Height          =   315
                  Left            =   10320
                  TabIndex        =   45
                  Top             =   2430
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
               Begin MSDataListLib.DataCombo dcitems 
                  Height          =   315
                  Left            =   10560
                  TabIndex        =   46
                  Top             =   3825
                  Width           =   4365
                  _ExtentX        =   7699
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
               Begin MSComCtl2.DTPicker DbTodate1 
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   60
                  Top             =   2370
                  Visible         =   0   'False
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   244318209
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   7
                  Left            =   5160
                  TabIndex        =   63
                  Top             =   225
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
                  ButtonImage     =   "frmAccountsSettings.frx":0E1E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   8
                  Left            =   4440
                  TabIndex        =   64
                  Top             =   225
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
                  ButtonImage     =   "frmAccountsSettings.frx":11B8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   315
                  Index           =   13
                  Left            =   12960
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   2730
                  Visible         =   0   'False
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ— › —…  ‰ ÂÌ"
                  Height          =   315
                  Index           =   15
                  Left            =   6360
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   2370
                  Visible         =   0   'False
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «Ê· › —…  »œ√"
                  Height          =   315
                  Index           =   14
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   2370
                  Visible         =   0   'False
                  Width           =   960
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·› —« "
                  Height          =   660
                  Index           =   12
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   1665
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ⁄—Ì› «·”‰Â"
                  Height          =   315
                  Index           =   3
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   345
                  Visible         =   0   'False
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   870
                  Index           =   2
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   2790
                  Width           =   360
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰œÊ»"
                  Height          =   315
                  Index           =   0
                  Left            =   10485
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   2430
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·„” ÊÌ« "
                  Height          =   285
                  Index           =   5
                  Left            =   8445
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   75
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  Height          =   270
                  Index           =   8
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   4860
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   825
                  Index           =   7
                  Left            =   -1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   345
                  Visible         =   0   'False
                  Width           =   1785
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
                  Height          =   930
                  Left            =   13800
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   1515
                  Width           =   855
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6810
            Index           =   0
            Left            =   45
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   45
            Width           =   10095
            _cx             =   17806
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
            Begin XtremeSuiteControls.CheckBox ChEntries 
               Height          =   195
               Left            =   6480
               TabIndex        =   117
               Top             =   600
               Width           =   1335
               _Version        =   786432
               _ExtentX        =   2355
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "«·ﬁÌÊœ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.TextBox TxtAccount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   120
               Width           =   1185
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
               Index           =   1
               Left            =   -3930
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   16575
               Width           =   2175
            End
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ "
               Height          =   255
               Left            =   10440
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   6375
               Width           =   2310
            End
            Begin VB.TextBox Text4 
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
               Height          =   870
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Text            =   "0"
               Top             =   2910
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ì"
               Enabled         =   0   'False
               Height          =   255
               Left            =   10440
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   6255
               Width           =   1590
            End
            Begin VB.OptionButton Option4 
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
               Left            =   10680
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   3825
               Width           =   1695
            End
            Begin VB.OptionButton Option3 
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
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   3825
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«› «· ⁄«„·"
               Height          =   810
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   2910
               Width           =   2310
            End
            Begin VB.Frame Frame2 
               Caption         =   "„⁄·Ê„« "
               Height          =   2805
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   1395
               Width           =   4695
               Begin MSDataListLib.DataCombo DataCombo1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   75
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
               Begin MSDataListLib.DataCombo DataCombo2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   76
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·„»Ì⁄« "
                  Height          =   315
                  Index           =   20
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «· Õ’Ì·« "
                  Height          =   195
                  Index           =   19
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   1150
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·„ √Œ—« "
                  Height          =   195
                  Index           =   18
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   1440
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ì »⁄ ›—⁄"
                  Height          =   315
                  Index           =   17
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ì »⁄ „Ã„Ê⁄Â"
                  Height          =   315
                  Index           =   16
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   480
                  Width           =   1200
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   1155
                  Width           =   1200
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   1440
                  Width           =   1200
               End
            End
            Begin VB.TextBox Text1 
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
               Height          =   870
               Left            =   10320
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   73
               Top             =   2850
               Visible         =   0   'False
               Width           =   2760
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   5370
               Left            =   120
               TabIndex        =   92
               Top             =   960
               Width           =   9825
               _cx             =   17330
               _cy             =   9472
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
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmAccountsSettings.frx":1752
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
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   285
               Left            =   7560
               TabIndex        =   93
               Top             =   2340
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   243793921
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo3 
               Height          =   315
               Left            =   10440
               TabIndex        =   94
               Top             =   3705
               Width           =   4365
               _ExtentX        =   7699
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
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   11760
               TabIndex        =   95
               Top             =   2670
               Width           =   1605
               _ExtentX        =   2831
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
            Begin MSDataListLib.DataCombo DataCombo5 
               Height          =   315
               Left            =   10680
               TabIndex        =   96
               Top             =   1395
               Width           =   3285
               _ExtentX        =   5794
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
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   870
               Left            =   10440
               TabIndex        =   97
               Top             =   2790
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   1535
               _Version        =   393216
               Format          =   243793921
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   9
               Left            =   10560
               TabIndex        =   98
               Top             =   3825
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
               ButtonImage     =   "frmAccountsSettings.frx":185C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   10
               Left            =   11280
               TabIndex        =   99
               Top             =   3825
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
               ButtonImage     =   "frmAccountsSettings.frx":1BF6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DataCombo6 
               Height          =   315
               Left            =   10320
               TabIndex        =   100
               Top             =   2430
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
            Begin MSDataListLib.DataCombo DataCombo7 
               Height          =   315
               Left            =   10560
               TabIndex        =   101
               Top             =   3825
               Width           =   4365
               _ExtentX        =   7699
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
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   285
               Left            =   4920
               TabIndex        =   102
               Top             =   2370
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   243793921
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton CmdDelete 
               Height          =   390
               Left            =   8880
               TabIndex        =   112
               Top             =   6360
               Width           =   1050
               _ExtentX        =   1852
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
               ButtonImage     =   "frmAccountsSettings.frx":2190
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DcbAccount 
               Height          =   315
               Left            =   240
               TabIndex        =   114
               Top             =   120
               Width           =   6375
               _ExtentX        =   11245
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   435
               Left            =   300
               TabIndex        =   116
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   480
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   767
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
               ButtonImage     =   "frmAccountsSettings.frx":272A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin XtremeSuiteControls.CheckBox ChTrialBalance 
               Height          =   195
               Left            =   5040
               TabIndex        =   118
               Top             =   600
               Width           =   1335
               _Version        =   786432
               _ExtentX        =   2355
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "«·„Ì“«‰"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChTreeAccount 
               Height          =   255
               Left            =   3360
               TabIndex        =   119
               Top             =   600
               Width           =   1335
               _Version        =   786432
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«·œ·Ì·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdDeleteAll 
               Height          =   390
               Left            =   6840
               TabIndex        =   120
               Top             =   6360
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·ﬂ·"
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
               ButtonImage     =   "frmAccountsSettings.frx":8F8C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ì"
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   25
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   600
               Width           =   1440
            End
            Begin VB.Label Label4 
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
               Height          =   930
               Left            =   13800
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   1515
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œ«Ì… «· Œ’Ì’"
               Height          =   270
               Index           =   29
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   4860
               Width           =   1785
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈Œ›«¡ «·Õ”«»«  «· «»⁄Â ·"
               Height          =   285
               Index           =   28
               Left            =   7845
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   195
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„‰œÊ»"
               Height          =   315
               Index           =   27
               Left            =   10485
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   2430
               Width           =   720
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï"
               Height          =   870
               Index           =   26
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   2790
               Width           =   360
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·› —« "
               Height          =   660
               Index           =   24
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   1665
               Width           =   1440
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «Ê· › —…  »œ√"
               Height          =   315
               Index           =   23
               Left            =   9000
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   2370
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ— › —…  ‰ ÂÌ"
               Height          =   315
               Index           =   22
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   2370
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   315
               Index           =   21
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   2730
               Visible         =   0   'False
               Width           =   840
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6810
            Index           =   3
            Left            =   10830
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   45
            Width           =   10095
            _cx             =   17806
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
            Begin VB.TextBox TxtFilter 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   3270
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   930
               Width           =   1215
            End
            Begin VB.CommandButton Command3 
               Caption         =   "‰ﬁ· «·»Ì«‰«  Sql 2"
               Height          =   375
               Left            =   4470
               RightToLeft     =   -1  'True
               TabIndex        =   173
               Top             =   930
               Width           =   1845
            End
            Begin VB.CommandButton Command2 
               Caption         =   "‰ﬁ· «·»Ì«‰«  Sql"
               Height          =   375
               Left            =   6420
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   930
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.CheckBox chkAll 
               Alignment       =   1  'Right Justify
               Caption         =   "All"
               Height          =   225
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   930
               Width           =   705
            End
            Begin VB.CommandButton Command1 
               Caption         =   "‰ﬁ· «·»Ì«‰« "
               Height          =   375
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   930
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.TextBox TxtAccount2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   1
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   450
               Width           =   1185
            End
            Begin VB.TextBox Text5 
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
               Height          =   870
               Left            =   10320
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   141
               Top             =   2850
               Visible         =   0   'False
               Width           =   2760
            End
            Begin VB.Frame Frame3 
               Caption         =   "„⁄·Ê„« "
               Height          =   2805
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   1395
               Width           =   4695
               Begin MSDataListLib.DataCombo DataCombo8 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   131
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
               Begin MSDataListLib.DataCombo DataCombo9 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   132
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
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   1440
                  Width           =   1200
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   1155
                  Width           =   1200
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ì »⁄ „Ã„Ê⁄Â"
                  Height          =   315
                  Index           =   34
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   480
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ì »⁄ ›—⁄"
                  Height          =   315
                  Index           =   33
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·„ √Œ—« "
                  Height          =   195
                  Index           =   32
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   1440
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «· Õ’Ì·« "
                  Height          =   195
                  Index           =   31
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   1150
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·„»Ì⁄« "
                  Height          =   315
                  Index           =   30
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   840
                  Width           =   1200
               End
            End
            Begin VB.CheckBox Check7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«› «· ⁄«„·"
               Height          =   810
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   2910
               Width           =   2310
            End
            Begin VB.OptionButton Option6 
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
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   3825
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton Option5 
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
               Left            =   10680
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   3825
               Width           =   1695
            End
            Begin VB.CheckBox Check6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ì"
               Enabled         =   0   'False
               Height          =   255
               Left            =   10440
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   6255
               Width           =   1590
            End
            Begin VB.TextBox Text3 
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
               Height          =   870
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Text            =   "0"
               Top             =   2910
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CheckBox Check5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ "
               Height          =   255
               Left            =   10440
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   6375
               Width           =   2310
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
               Index           =   2
               Left            =   -3930
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   16575
               Width           =   2175
            End
            Begin VB.TextBox TxtAccount2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   0
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   30
               Width           =   1185
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
               Height          =   4950
               Left            =   120
               TabIndex        =   142
               Top             =   1470
               Width           =   9825
               _cx             =   17330
               _cy             =   8731
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmAccountsSettings.frx":9526
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
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   285
               Left            =   7560
               TabIndex        =   143
               Top             =   2340
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   243793921
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo10 
               Height          =   315
               Left            =   10440
               TabIndex        =   144
               Top             =   3705
               Width           =   4365
               _ExtentX        =   7699
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
            Begin MSDataListLib.DataCombo DataCombo11 
               Height          =   315
               Left            =   11760
               TabIndex        =   145
               Top             =   2670
               Width           =   1605
               _ExtentX        =   2831
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
            Begin MSDataListLib.DataCombo DataCombo12 
               Height          =   315
               Left            =   10680
               TabIndex        =   146
               Top             =   1395
               Width           =   3285
               _ExtentX        =   5794
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
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   870
               Left            =   10440
               TabIndex        =   147
               Top             =   2790
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   1535
               _Version        =   393216
               Format          =   243793921
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   11
               Left            =   10560
               TabIndex        =   148
               Top             =   3825
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
               ButtonImage     =   "frmAccountsSettings.frx":9657
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   12
               Left            =   11280
               TabIndex        =   149
               Top             =   3825
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
               ButtonImage     =   "frmAccountsSettings.frx":99F1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DataCombo13 
               Height          =   315
               Left            =   10320
               TabIndex        =   150
               Top             =   2430
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
            Begin MSDataListLib.DataCombo DataCombo14 
               Height          =   315
               Left            =   10560
               TabIndex        =   151
               Top             =   3825
               Width           =   4365
               _ExtentX        =   7699
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
            Begin MSComCtl2.DTPicker DTPicker6 
               Height          =   285
               Left            =   4920
               TabIndex        =   152
               Top             =   2370
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   241303553
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   390
               Left            =   8880
               TabIndex        =   153
               Top             =   6360
               Width           =   1050
               _ExtentX        =   1852
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
               ButtonImage     =   "frmAccountsSettings.frx":9F8B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DcbAccount2 
               Height          =   315
               Index           =   0
               Left            =   240
               TabIndex        =   154
               Top             =   30
               Width           =   6375
               _ExtentX        =   11245
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton4 
               Height          =   435
               Left            =   240
               TabIndex        =   155
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   840
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   767
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
               ButtonImage     =   "frmAccountsSettings.frx":A525
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton5 
               Height          =   390
               Left            =   6840
               TabIndex        =   156
               Top             =   6360
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·ﬂ·"
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
               ButtonImage     =   "frmAccountsSettings.frx":10D87
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DcbAccount2 
               Height          =   315
               Index           =   1
               Left            =   240
               TabIndex        =   168
               Top             =   450
               Width           =   6375
               _ExtentX        =   11245
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ”«» «·ÃœÌœ"
               Height          =   285
               Index           =   44
               Left            =   7845
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   525
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   315
               Index           =   43
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   2730
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ— › —…  ‰ ÂÌ"
               Height          =   315
               Index           =   42
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   2370
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «Ê· › —…  »œ√"
               Height          =   315
               Index           =   41
               Left            =   9000
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   2370
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·› —« "
               Height          =   660
               Index           =   40
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   1665
               Width           =   1440
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï"
               Height          =   870
               Index           =   39
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   2790
               Width           =   360
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„‰œÊ»"
               Height          =   315
               Index           =   38
               Left            =   10485
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   2430
               Width           =   720
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰ﬁ· «·Õ”«»«  «· «»⁄Â ·"
               Height          =   285
               Index           =   37
               Left            =   7845
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   105
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œ«Ì… «· Œ’Ì’"
               Height          =   270
               Index           =   36
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   4860
               Width           =   1785
            End
            Begin VB.Label Label9 
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
               Height          =   930
               Left            =   13800
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   1515
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ì"
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   35
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   960
               Width           =   1440
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   7995
         Width           =   10065
         _cx             =   17754
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   11880
            TabIndex        =   17
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
            ButtonImage     =   "frmAccountsSettings.frx":11321
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
            ButtonImage     =   "frmAccountsSettings.frx":116BB
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
            ButtonImage     =   "frmAccountsSettings.frx":11A55
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   7140
            TabIndex        =   23
            Top             =   270
            Visible         =   0   'False
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
            Left            =   6240
            TabIndex        =   24
            Top             =   270
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
            Left            =   5400
            TabIndex        =   25
            Top             =   270
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
            Left            =   4395
            TabIndex        =   26
            Top             =   270
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
            Left            =   3360
            TabIndex        =   27
            Top             =   270
            Visible         =   0   'False
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
            Left            =   480
            TabIndex        =   28
            Top             =   270
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
            Left            =   2430
            TabIndex        =   29
            Top             =   270
            Visible         =   0   'False
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
            Left            =   9000
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
            MICON           =   "frmAccountsSettings.frx":11DEF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   765
         Index           =   5
         Left            =   0
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   0
         Width           =   10035
         _cx             =   17701
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
         Picture         =   "frmAccountsSettings.frx":11E0B
         Caption         =   "«⁄œ«œ«    ﬂÊÌœ «·Õ”«»« /≈⁄œ«œ«  «·Õ”«»«  "
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1695
            TabIndex        =   68
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
            ButtonImage     =   "frmAccountsSettings.frx":12AE5
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
            TabIndex        =   69
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
            ButtonImage     =   "frmAccountsSettings.frx":12E7F
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
            TabIndex        =   70
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
            ButtonImage     =   "frmAccountsSettings.frx":13219
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
            TabIndex        =   71
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
            ButtonImage     =   "frmAccountsSettings.frx":135B3
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
      ButtonImage     =   "frmAccountsSettings.frx":1394D
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmAccountsSeetting"
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







Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "

    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    rs.AddNew
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = Null
    rs("Remark").value = Msg

    rs("NoteType").value = 66
    rs("NoteDate").value = Date
    rs("UserID").value = user_id
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    'create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Function Create_dev1()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    'StrAccountCode = Account_Code_dynamic
        
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With

    Set rs = New ADODB.Recordset
    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
 
    rs("voucher_id").value = LngDevID
  
    rs.update
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    'create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function







Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        get_all_employee
    Else

        With Me.Grid
            .rows = 2
            .Clear flexClearScrollable
        End With

    End If

End Sub






Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
 
        '       If Trim(Me.TXTNO.text) = "" Then
        '            Msg = "ÌÃ»    «œŒ«· «·”‰Â..!!"
        '            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '            TXTNO.SetFocus
        '
        '            Exit Sub
        '        End If
 
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.text = "E" Then
    
        Cn.Execute "delete AccountsLevelsDetails where AccountsLevelsId=" & val(Me.TxtyearsdataId.text)
        Cn.Execute "delete from  AccountSetting where AccLevelID =" & val(Me.TxtyearsdataId.text)
       
    End If
    
    rs("AccountsLevelsid").value = TxtyearsdataId.text
    
    rs("no").value = IIf(val(Me.TxtNo.text) = 0, 0, val(Me.TxtNo.text))
   
    rs("Remarks").value = IIf(Me.txtRemarks.text = "", "", Me.txtRemarks.text)
    'rs("datesatrt").value = DbFromDate.value
    'rs("dateend").value = DbTodate1.value
 
    rs.update
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "AccountsLevelsDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer
    Dim lastIntervalid As Long
    'lastIntervalid = CStr(new_id("AccountsLevelsDetails", "AccountIntervalID", "", True))
    
    With Me.Grid

        For i = 1 To .rows - 1

            If .TextMatrix(i, .ColIndex("Level")) <> "" Then
         
                RsDev.AddNew
                RsDev("Level").value = i ' lastIntervalid
            
                RsDev("AccountsLevelsid").value = val(Me.TxtyearsdataId.text)
            
                ' RsDev("Level").value = (.TextMatrix(i, .ColIndex("Level")))
                RsDev("NoOfDigits").value = val(.TextMatrix(i, .ColIndex("NoOfDigits")))
                RsDev("Sample").value = .TextMatrix(i, .ColIndex("Sample"))
 
                If .cell(flexcpChecked, i, .ColIndex("Zeros")) = flexChecked Then
                    RsDev("Zeros").value = 1
                Else
                    RsDev("Zeros").value = 0
                End If
             
                RsDev.update

                lastIntervalid = lastIntervalid + 1
                    
            End If
            
            '
        Next i

    End With
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim sql As String
sql = "Select * from AccountSetting  where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.VSFlexGrid1
For i = 1 To .rows - 1
If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
rs2.AddNew
rs2("AccLevelID").value = val(TxtyearsdataId.text)
rs2("AccountCode").value = IIf(.TextMatrix(i, .ColIndex("AccountCode")) = "", Null, .TextMatrix(i, .ColIndex("AccountCode")))
If .cell(flexcpChecked, i, .ColIndex("Entries")) = flexChecked Then
rs2("Entries").value = 1
Else
rs2("Entries").value = 0
End If
If .cell(flexcpChecked, i, .ColIndex("TrialBalance")) = flexChecked Then
rs2("TrialBalance").value = 1
Else
rs2("TrialBalance").value = 0
End If
If .cell(flexcpChecked, i, .ColIndex("TreeAccount")) = flexChecked Then
rs2("TreeAccount").value = 1
Else
rs2("TreeAccount").value = 0
End If
rs2.update
End If
Next i
End With
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.text

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

    TxtModFlg.text = "R"
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

Private Sub ChkALL_Click()
Dim i As Long
For i = 1 To VSFlexGrid2.rows - 1
    VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("Select")) = ChkAll.value
Next
End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0
 
            TxtModFlg.text = "N"
            clear_all Me
            Me.TxtyearsdataId.text = CStr(new_id("AccountsLevels", "AccountsLevelsid", "", True))
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
             VSFlexGrid1.rows = 2
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
          
            Grid.Enabled = True
            Grid.rows = 1
            dbFromDate.value = Date
            Me.DbTodate1.value = Date
         
        Case 1
 
            TxtModFlg.text = "E"
    
            Grid.Enabled = True
     
        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            ' Del_Trans
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 3
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            '   ViewDataList
  
            addrow

        Case 8
            RemoveGridRow
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub RemoveGridRow()
    Grid.Clear flexClearScrollable, flexClearEverything
          
    Grid.Enabled = True
    Grid.rows = 1
    'With Me.Grid
    '    If .Row <= 0 Then Exit Sub
    '    .RemoveItem .Row
    'End With
    'ReLineGrid
End Sub

Function addrow()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
    Dim i As Integer
    Dim FromDate As Date
    Dim ToDate As Date
    Dim Remark As String
    Me.Grid.rows = val(txtRemarks.text) + 1

    For i = 1 To val(txtRemarks.text)

        If i = 1 Then
          '  FromDate = Me.dbFromDate.value

            If SystemOptions.UserInterface = ArabicInterface Then
                Remark = "«Ê· „” ÊÏ"
            Else
                Remark = " First Period"
            End If

          '  ToDate = MonthLastDay(FromDate)
        ElseIf i = val(txtRemarks.text) Then
          '  FromDate = DateAdd("d", 1, ToDate)
          '  ToDate = MonthLastDay(FromDate)

            'todate = Me.DbTodate1.value
            If SystemOptions.UserInterface = ArabicInterface Then
                Remark = "«Œ— „” ÊÏ"
            Else
                Remark = " last Period"
            End If

        Else
          '  FromDate = DateAdd("d", 1, ToDate)
          '  ToDate = MonthLastDay(FromDate)
            Remark = ""
        End If

        '   Me.Grid.Rows = Me.Grid.Rows + 1
        '   LngRow = Me.Grid.Rows - 1
        LngRow = i
 
        With Me.Grid
  
            .TextMatrix(LngRow, .ColIndex("Sample")) = Remark
    
            .TextMatrix(LngRow, .ColIndex("Level")) = i
    
            '  .TextMatrix(LngRow, .ColIndex("EndDate")) = todate
    
            .TextMatrix(LngRow, .ColIndex("Zeros")) = 0
     
            .AutoSize 0, .Cols - 1, False
        End With
 
    Next i
 
    Me.TxtRemark1.text = ""
    ReLineGrid

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub




Private Sub CmdDelete_Click()
If Me.TxtModFlg.text <> "R" Then

    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    If VSFlexGrid1.rows > 1 Then
        If VSFlexGrid1.rows = 1 Then
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.VSFlexGrid1.rows > 1 Then
                If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                    Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                End If
            End If
        End If
    End If
  End If
End Sub

Private Sub CmdDeleteAll_Click()

     If Me.TxtModFlg.text <> "R" Then
       Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 1
   End If
End Sub

Private Sub Command1_Click()
MoveSelectedAccounts DcbAccount2(1).BoundText

End Sub

Private Sub Command2_Click()
MoveSelectedAccounts_WithLog DcbAccount2(1).BoundText, ""
End Sub

Private Sub Command3_Click()
BulkMove DcbAccount2(1).BoundText, ""

VSFlexGrid2.rows = 1
DcbAccount2(0).BoundText = DcbAccount2(1).BoundText


End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click 0
'ISButton4_Click
End Sub

Function SHow_grig_col()
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If rs2("s18").value = True Then
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


End Sub

Private Sub DcbAccount2_Click(Index As Integer, Area As Integer)
TxtAccount2(Index).text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount2(Index).BoundText)
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
     Dcombos.GetAccountingCodes Me.DcbAccount, , True
     Dcombos.GetAccountingCodes Me.DcbAccount2(0), False, True
     Dcombos.GetAccountingCodes Me.DcbAccount2(1), False, True
    Dcombos.GetSalesRepData Me.DcEmp
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetSalesRepGroups Me.DCGroup

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
    StrSQL = "select * From AccountsLevels  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    lbl(28).Caption = "Hide Account"
    lbl(25).Caption = "In"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"
    CmdRemove.Caption = "Delete Row"
    ISButton2.Caption = "Add"
    lbl(5).Caption = "year Des"
    lbl(12).Caption = "Periods "
    lbl(14).Caption = "Start "
    lbl(15).Caption = "End "
    Me.CmdDelete.Caption = "Delete"
    Me.CmdDeleteAll.Caption = "Delete All"
    C1Tab1.Caption = "Coding|Settings"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Accounts Coding "
    Ele(5).Caption = Me.Caption
    lbl(5).Caption = "No Of Levels"
    Cmd(7).Caption = "Add"
    Cmd(8).Caption = "Delete"
    ChTreeAccount.RightToLeft = False
    ChTreeAccount.Caption = "Chart"
    ChTrialBalance.RightToLeft = False
    ChTrialBalance.Caption = "Trial Balance"
    ChEntries.RightToLeft = False
    ChEntries.Caption = "JL"
    With Me.VSFlexGrid1
       ' .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("Account_Serial")) = "Account Code"
        .TextMatrix(0, .ColIndex("Account_Name")) = "Account Name"
        .TextMatrix(0, .ColIndex("Entries")) = "JL"
        .TextMatrix(0, .ColIndex("TrialBalance")) = "Trial Balance"
        .TextMatrix(0, .ColIndex("TreeAccount")) = "Chart of Account"
    End With
    
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("Level")) = "Level"
        .TextMatrix(0, .ColIndex("NoOfDigits")) = "NoOfDigits"
        .TextMatrix(0, .ColIndex("Zeros")) = "Zeros"
        .TextMatrix(0, .ColIndex("Sample")) = "remarks"
 
    End With

End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .rows = Rs3.RecordCount + 1
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

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    On Error GoTo ErrTrap
 
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                rs.MoveNext
            
            Next

            rs.Close
        End If

        .rows = .rows + 1
        .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows - 1, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows - 1, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

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
 
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("Level")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtyearsdataId.text = IIf(IsNull(rs("AccountsLevelsid").value), "", rs("AccountsLevelsid").value)
   
    TxtNo.text = IIf(IsNull(rs("no").value), 0, rs("no").value)
    txtRemarks.text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
 
    'DbFromDate.value = IIf(IsNull(rs("datesatrt").value), Date, rs("datesatrt").value)
    'DbTodate1.value = IIf(IsNull(rs("dateend").value), Date, rs("dateend").value)

    StrSQL = " SELECT   * FROM         dbo.AccountsLevelsDetails  "
    StrSQL = StrSQL & "  where AccountsLevelsId=" & val(Me.TxtyearsdataId.text)
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
      
                If RsDev("Zeros").value = True Then
                    .cell(flexcpChecked, i, .ColIndex("Zeros")) = flexChecked
                    
                Else
                    .cell(flexcpChecked, i, .ColIndex("Zeros")) = flexUnchecked
                End If
            
                .TextMatrix(i, .ColIndex("Level")) = IIf(IsNull(RsDev("Level").value), 0, (RsDev("Level").value))
            
                .TextMatrix(i, .ColIndex("NoOfDigits")) = IIf(IsNull(RsDev("NoOfDigits").value), 0, (RsDev("NoOfDigits").value))
  
                .TextMatrix(i, .ColIndex("Sample")) = IIf(IsNull(RsDev("Sample").value), "", RsDev("Sample").value)
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 FillGridSetting
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 Sub FillGridSetting()
 Dim sql As String
 Dim i As Integer
   VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
   VSFlexGrid1.rows = 1
 Dim rs2 As ADODB.Recordset
 Set rs2 = New ADODB.Recordset
 sql = "SELECT     dbo.AccountSetting.ID, dbo.AccountSetting.AccLevelID, dbo.AccountSetting.TrialBalance, dbo.AccountSetting.TreeAccount, dbo.AccountSetting.Entries, "
 sql = sql & "                       dbo.AccountSetting.AccountCode , dbo.Accounts.account_name, dbo.Accounts.account_serial, dbo.Accounts.Account_NameEng"
 sql = sql & "   FROM         dbo.AccountSetting LEFT OUTER JOIN"
 sql = sql & "                      dbo.ACCOUNTS ON dbo.AccountSetting.AccountCode = dbo.ACCOUNTS.Account_Code"
 sql = sql & "  WHERE     (dbo.AccountSetting.AccLevelID = " & val(TxtyearsdataId.text) & ") "
 rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If rs2.RecordCount > 0 Then
 rs2.MoveFirst
 With Me.VSFlexGrid1
 .rows = rs2.RecordCount + 1
 For i = 1 To rs2.RecordCount
 .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs2("AccountCode").value), "", rs2("AccountCode").value)
 .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs2("Account_Serial").value), "", rs2("Account_Serial").value)
 If SystemOptions.UserInterface = ArabicInterface Then
 .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_Name").value), "", rs2("Account_Name").value)
 Else
 .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_NameEng").value), "", rs2("Account_NameEng").value)
 End If
 .TextMatrix(i, .ColIndex("Entries")) = IIf(IsNull(rs2("Entries").value), 0, rs2("Entries").value)
 .TextMatrix(i, .ColIndex("TrialBalance")) = IIf(IsNull(rs2("TrialBalance").value), 0, rs2("TrialBalance").value)
 .TextMatrix(i, .ColIndex("TreeAccount")) = IIf(IsNull(rs2("TreeAccount").value), 0, rs2("TreeAccount").value)
 rs2.MoveNext
 Next i
 End With
 End If
 End Sub


Sub FillGridSetting2()
 Dim sql As String
 Dim i As Integer
 
 
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
 VSFlexGrid1.rows = 1
 
 Dim rs2 As ADODB.Recordset
 Set rs2 = New ADODB.Recordset
 sql = "SELECT     dbo.AccountSetting.ID, dbo.AccountSetting.AccLevelID, dbo.AccountSetting.TrialBalance, dbo.AccountSetting.TreeAccount, dbo.AccountSetting.Entries, "
 sql = sql & "                       dbo.AccountSetting.AccountCode , dbo.Accounts.account_name, dbo.Accounts.account_serial, dbo.Accounts.Account_NameEng"
 sql = sql & "   FROM         dbo.AccountSetting LEFT OUTER JOIN"
 sql = sql & "                      dbo.ACCOUNTS ON dbo.AccountSetting.AccountCode = dbo.ACCOUNTS.Account_Code"
 sql = sql & "  WHERE     (dbo.AccountSetting.AccLevelID = " & val(TxtyearsdataId.text) & ") "
 
 
 sql = " Select Account_Code,Account_Name,Parent_Account_Code,Account_Serial,Account_NameEng,Account_NameEng,last_account from ACCOUNTS where IsNull(last_account,0) = 1 and Parent_Account_Code = N'" & Trim(DcbAccount2(0).BoundText) & "'"
 
 rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If rs2.RecordCount > 0 Then
 rs2.MoveFirst
 With Me.VSFlexGrid2
 .rows = rs2.RecordCount + 1
 For i = 1 To rs2.RecordCount
 .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs2("Account_Code").value), "", rs2("Account_Code").value)
 .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs2("Account_Serial").value), "", rs2("Account_Serial").value)
 If SystemOptions.UserInterface = ArabicInterface Then
 .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_Name").value), "", rs2("Account_Name").value)
 Else
 .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_NameEng").value), "", rs2("Account_NameEng").value)
 End If
 .TextMatrix(i, .ColIndex("last_account")) = IIf(IsNull(rs2("last_account").value), 0, rs2("last_account").value)
 
' .TextMatrix(i, .ColIndex("Entries")) = IIf(IsNull(rs2("Entries").value), 0, rs2("Entries").value)
' .TextMatrix(i, .ColIndex("TrialBalance")) = IIf(IsNull(rs2("TrialBalance").value), 0, rs2("TrialBalance").value)
' .TextMatrix(i, .ColIndex("TreeAccount")) = IIf(IsNull(rs2("TreeAccount").value), 0, rs2("TreeAccount").value)
 rs2.MoveNext
 Next i
 End With
 End If
 End Sub
 
 Sub MoveSelectedAccounts_WithLog(ByVal newParentCode As String, ByVal reason As String)
    Dim i As Long, AccCode As String, newSerial As String
    Dim Cmd As ADODB.Command
    Dim batchId As String
    batchId = CreateObject("Scriptlet.TypeLib").GUID  ' GUID œ›⁄…
    If newParentCode = "" Then Exit Sub
    Cn.BeginTrans
    On Error GoTo eh

    For i = 1 To VSFlexGrid2.rows - 1
        If VSFlexGrid2.ValueMatrix(i, VSFlexGrid2.ColIndex("Select")) <> "0" Then
            AccCode = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("AccountCode"))
            
            Set Cmd = New ADODB.Command
            With Cmd
                .ActiveConnection = Cn
                .CommandText = "dbo.Account_MoveUnderParent"
                .CommandType = adCmdStoredProc
                .CommandTimeout = 500
                .Parameters.Append .CreateParameter("@AccountCode", adVarWChar, adParamInput, 50, AccCode)
                .Parameters.Append .CreateParameter("@NewParentCode", adVarWChar, adParamInput, 50, newParentCode)
                .Parameters.Append .CreateParameter("@Reason", adVarWChar, adParamInput, 400, reason)
                .Parameters.Append .CreateParameter("@BatchId", adGUID, adParamInput, , batchId)
                .Parameters.Append .CreateParameter("@MovedByUserId", adInteger, adParamInput, , user_id)      ' Õ”» ‰Ÿ«„ﬂ
                .Parameters.Append .CreateParameter("@MovedByUser", adVarWChar, adParamInput, 100, user_name)
                .Parameters.Append .CreateParameter("@AppName", adVarWChar, adParamInput, 100, "Dynamic")
                .Parameters.Append .CreateParameter("@NewSerial", adVarWChar, adParamOutput, 200, Null)

                .Execute
                newSerial = .Parameters("@NewSerial").value
            End With

            VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("Account_Serial")) = newSerial
        End If
    Next

    Cn.CommitTrans
    MsgBox " „ «·‰ﬁ· Ê«· ÊÀÌﬁ ›Ì «·”Ã· (BatchId=" & batchId & ")", vbInformation
    Exit Sub
eh:
    Cn.RollbackTrans
    MsgBox "›‘· «·‰ﬁ·: " & Err.Description, vbCritical
End Sub

 
  
 Private Function IsCheckedVal(v As Variant) As Boolean
    On Error Resume Next
    ' Ìœ⁄„ 1/-1/True/Yes/X
    Select Case LCase$(Trim$(CStr(v)))
        Case "1", "-1", "true", "yes", "x", "checked": IsCheckedVal = True
        Case Else: IsCheckedVal = False
    End Select
    On Error GoTo 0
End Function

Sub BulkMove(ByVal newParentCode As String, ByVal reason As String)
    Dim i As Long, List As String, cnt As Long, batchId As String
    Dim colSel As Integer, colCode As Integer

    batchId = CreateObject("Scriptlet.TypeLib").GUID

    ' ÃÂ¯“ √—ﬁ«„ «·√⁄„œ… „—… Ê«Õœ…
    On Error Resume Next
    colSel = VSFlexGrid2.ColIndex("Select")          ' «”„ ⁄„Êœ «·«Œ Ì«— ⁄‰œﬂ
    colCode = VSFlexGrid2.ColIndex("AccountCode")    ' ﬂÊœ «·Õ”«»
    On Error GoTo 0
    If colSel < 0 Or colCode < 0 Then
        MsgBox "√⁄„œ… Select/AccountCode „‘ „ÊÃÊœ… ›Ì «·Ã—Ìœ.", vbExclamation
        Exit Sub
    End If

    Cn.CommandTimeout = 500

    For i = 1 To VSFlexGrid2.rows - 1

        '  Ã«Â· «·’›Ê› «·„Œ›Ì…
        If Not VSFlexGrid2.RowHidden(i) Then

            ' „ ⁄·„ + ·Â ﬂÊœ
            If IsCheckedVal(VSFlexGrid2.ValueMatrix(i, colSel)) Then
                Dim acc As String
                acc = Trim$(VSFlexGrid2.TextMatrix(i, colCode))
                If acc <> "" Then
                    If List <> "" Then List = List & ","
                    List = List & "'" & Replace(acc, "'", "''") & "'"
                    cnt = cnt + 1
                End If
            End If

            ' ‰›¯– «·œ›⁄… ﬂ· 400 √Ê ⁄‰œ ¬Œ— ’›
            If cnt >= 400 Then
                Call ExecBulkMove(newParentCode, reason, batchId, List)
                List = "": cnt = 0
            End If
        End If
    Next i

    ' ‰›¯– √Ï »Ê«ﬁÌ √ﬁ· „‰ 400
    If List <> "" Then Call ExecBulkMove(newParentCode, reason, batchId, List)

    MsgBox " „ «·‰ﬁ· »«·œı›⁄«  Ê ”ÃÌ· «··ÊÃ. BatchId=" & batchId, vbInformation
End Sub

Private Sub ExecBulkMove(ByVal newParentCode As String, ByVal reason As String, _
                         ByVal batchId As String, ByVal csvList As String)
    Dim Cmd As New ADODB.Command
    With Cmd
        .ActiveConnection = Cn
        .CommandText = "[dbo].[Account_BulkMoveUnderParent_Alt]"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 500

        ' Œıœ  ⁄—Ì› «·»«—«„ —«  „‰ «·”Ì—›—
        .Parameters.Refresh

        .Parameters("@NewParentCode").value = newParentCode
        .Parameters("@AccountCodesCsv").value = csvList
        If .Parameters.count > 2 Then .Parameters("@Reason").value = reason
        If .Parameters.count > 3 Then .Parameters("@BatchId").value = batchId
        If .Parameters.count > 4 Then .Parameters("@MovedByUserId").value = user_id
        If .Parameters.count > 5 Then .Parameters("@MovedByUser").value = user_name
        If .Parameters.count > 6 Then .Parameters("@AppName").value = "Dynamic"

        .Execute , , adExecuteNoRecords
    End With
End Sub
 
  ' Ì⁄Ìœ True ·Ê «·”Ì—Ì«· „ÊÃÊœ »«·›⁄·
Private Function AccountSerialExists(ByVal serial As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT 1 FROM ACCOUNTS WHERE Account_Serial = N'" & Replace(serial, "'", "''") & "'"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    AccountSerialExists = Not rs.EOF
    rs.Close
End Function

' Ì—Ã¯⁄ ⁄œœ «·Œ«‰«  (NoOfDigits) ·„” ÊÏ „⁄Ì‰
Private Function GetAccountsLevelDigits(ByVal levelNo As Integer) As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT NoOfDigits FROM AccountsLevelsDetails WHERE [Level] = " & CStr(levelNo), _
            Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Or IsNull(rs!NoOfDigits) Then
        GetAccountsLevelDigits = 0
    Else
        GetAccountsLevelDigits = rs!NoOfDigits
    End If
    rs.Close
End Function

' Ì—Ã¯⁄ «·”Ì—Ì«· «· «·Ì «·„ «Õ  Õ  √»¯ „⁄Ì¯‰
Public Function NextChildSerial(ByVal parentCOde As String) As String
    Dim parentSerial As String
    parentSerial = Get_Account_Serial(parentCOde)
    If Len(parentSerial) = 0 Then
        NextChildSerial = ""  ' „›Ì‘ ”Ì—Ì«· √»
        Exit Function
    End If

    ' „” ÊÏ «·«»‰ = ⁄œœ «·‹ a ›Ì ﬂÊœ «·√» + 1 (Õ”» ÿ—Ìﬁ ﬂ)
    Dim levelNo As Integer
    levelNo = CountAs(parentCOde) + 1

    Dim Digits As Integer
    Digits = GetAccountsLevelDigits(levelNo)
    If Digits <= 0 Then
        NextChildSerial = ""
        Exit Function
    End If

    ' √ﬂ»— –Ì· ··√» «·Õ«·Ì
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT MAX(CAST(RIGHT(Account_Serial," & CStr(Digits) & ") AS INT)) AS maxTail " & _
          "FROM ACCOUNTS " & _
          "WHERE Parent_Account_Code = N'" & Replace(parentCOde, "'", "''") & "' " & _
          "AND Account_Serial LIKE N'" & Replace(parentSerial, "'", "''") & "%' " & _
          "AND LEN(Account_Serial) = " & CStr(Len(parentSerial) + Digits)

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Dim nextNum As Long: nextNum = 1
    If Not rs.EOF Then
        If Not IsNull(rs!maxTail) Then nextNum = CLng(rs!maxTail) + 1
    End If
    rs.Close

    Dim tail As String
    tail = right$(String$(Digits, "0") & CStr(nextNum), Digits)
    NextChildSerial = parentSerial & tail

    ' ÷„«‰ ⁄œ„ «· ’«œ„
    Do While AccountSerialExists(NextChildSerial)
        nextNum = nextNum + 1
        tail = right$(String$(Digits, "0") & CStr(nextNum), Digits)
        NextChildSerial = parentSerial & tail
    Loop
End Function
Sub MoveSelectedAccounts(ByVal newParentCode As String)
    Dim i As Long, AccCode As String, newSerial As String
If newParentCode = "" Then Exit Sub
    Cn.BeginTrans
    On Error GoTo eh

    For i = 1 To VSFlexGrid2.rows - 1
        If VSFlexGrid2.ValueMatrix(i, VSFlexGrid2.ColIndex("Select")) <> "0" Then
            AccCode = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("AccountCode"))
            newSerial = NextChildSerial(newParentCode)
            If Len(newSerial) = 0 Then Err.Raise vbObjectError + 1, , " ⁄–¯—  Ê·Ìœ «·”Ì—Ì«· «·ÃœÌœ."

            '  ÕœÌÀ «·«» + «·”Ì—Ì«·
            Dim sql As String
            sql = "UPDATE ACCOUNTS SET Parent_Account_Code = N'" & Replace(newParentCode, "'", "''") & _
                  "', Account_Serial = N'" & Replace(newSerial, "'", "''") & _
                  "' WHERE Account_Code = N'" & Replace(AccCode, "'", "''") & "'"
            Cn.Execute sql

            '  ÕœÌÀ «·Ã—Ìœ ‘ﬂ·Ì«
            VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("Account_Serial")) = newSerial
        End If
    Next

    Cn.CommitTrans
    MsgBox " „ ‰ﬁ· «·Õ”«»«  Ê ÕœÌÀ «·”Ì—Ì«· »‰Ã«Õ.", vbInformation
    Exit Sub
eh:
    Cn.RollbackTrans
    MsgBox "›‘· «·‰ﬁ·: " & Err.Description, vbCritical
End Sub


Sub AddRowSitting()
    Dim k As Integer
    Dim i As Integer
    With VSFlexGrid1
      k = .rows
      .rows = .rows + 1
     For i = k To .rows - 1
      .TextMatrix(i, .ColIndex("Account_Serial")) = TxtAccount.text
      .TextMatrix(i, .ColIndex("AccountCode")) = DcbAccount.BoundText
      .TextMatrix(i, .ColIndex("Account_Name")) = DcbAccount.text
      If ChEntries.value = vbChecked Then
      .TextMatrix(i, .ColIndex("Entries")) = 1
      Else
      .TextMatrix(i, .ColIndex("Entries")) = 0
      End If
        If ChTrialBalance.value = vbChecked Then
      .TextMatrix(i, .ColIndex("TrialBalance")) = 1
      Else
      .TextMatrix(i, .ColIndex("TrialBalance")) = 0
      End If
        If ChTreeAccount.value = vbChecked Then
      .TextMatrix(i, .ColIndex("TreeAccount")) = 1
      Else
      .TextMatrix(i, .ColIndex("TreeAccount")) = 0
      End If
      .TextMatrix(i, .ColIndex("Account_Serial")) = 1
      .TextMatrix(i, .ColIndex("Account_Serial")) = 1
     Next i
      
    End With
End Sub


Private Sub ISButton2_Click()
If Me.TxtModFlg.text <> "R" Then
If DcbAccount.text = "" Or DcbAccount.BoundText = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«»"
Else
MsgBox "Please Sect Acount"
End If
DcbAccount.SetFocus
Exit Sub
End If
AddRowSitting
End If
End Sub



Private Sub ISButton4_Click()



VSFlexGrid2.rows = 1
If DcbAccount2(0).text = "" Or DcbAccount2(0).BoundText = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«»"
    Else
        MsgBox "Please Sect Acount"
    End If
    DcbAccount2(0).SetFocus
    Exit Sub
End If

FillGridSetting2





End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.text)
End If
End Sub
Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
End Sub


Private Sub TxtAccount2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DcbAccount2(Index).BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount2(Index).text)
End If
End Sub


Private Sub TxtFilter_Change()
    Dim key As String, ColName As String, Col As Integer
    Dim i As Long, Txt As String

    key = LCase$(Trim$(TXtFilter.text))
    ColName = "Account_Name"  ' ·Ê ⁄‰œﬂ ≈‰Ã·Ì“Ì: "Account_NameEng"

    On Error Resume Next
    Col = VSFlexGrid2.ColIndex(ColName)
    If Col < 0 Then
        ' fallback: ·Ê «·⁄„Êœ „‘ „ÊÃÊœ Ã—¯» «·≈‰Ã·Ì“Ì
        Col = VSFlexGrid2.ColIndex("Account_Name")
    End If
    On Error GoTo 0

    If Col < 0 Then Exit Sub  ' „›Ì‘ ⁄„Êœ „ÿ«»ﬁ

    VSFlexGrid2.Redraw = False
    For i = 1 To VSFlexGrid2.rows - 1
        Txt = LCase$(VSFlexGrid2.TextMatrix(i, Col))
        If key = "" Then
            VSFlexGrid2.RowHidden(i) = False
        Else
            VSFlexGrid2.RowHidden(i) = (InStr(1, Txt, key, vbTextCompare) = 0)
        End If
    Next i
    VSFlexGrid2.Redraw = True
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.VSFlexGrid1
Select Case .ColKey(Col)
Case "Account_Serial"
Cancel = True
Case "Account_Name"
Cancel = True
End Select
End With
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
