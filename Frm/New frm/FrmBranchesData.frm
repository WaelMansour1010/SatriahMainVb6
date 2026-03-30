VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmBranchesData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·«‰‘ÿÂ Ê «·›—Ê⁄ "
   ClientHeight    =   10005
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14835
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
   Icon            =   "FrmBranchesData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   14835
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   14805
      _cx             =   26114
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
      _GridInfo       =   $"FrmBranchesData.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8220
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14745
         _cx             =   26009
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
         Caption         =   ".|«·›Ê —… «·«·ﬂ —Ê‰Ì"
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
            Height          =   7800
            Index           =   2
            Left            =   -15300
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   14655
            _cx             =   25850
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5955
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   15225
               _cx             =   26855
               _cy             =   10504
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
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
               Begin VB.TextBox TxtVATNO 
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
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   1800
                  Width           =   1800
               End
               Begin VB.TextBox txtbranch_name 
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
                  Left            =   8040
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   1200
                  Width           =   4080
               End
               Begin VB.TextBox txtbranch_namee 
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
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   1200
                  Width           =   4080
               End
               Begin VB.TextBox txtnamee 
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
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   240
                  Width           =   4080
               End
               Begin MSComDlg.CommonDialog Cdg 
                  Left            =   960
                  Top             =   2760
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin VB.TextBox Txtbranch_Code 
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
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   1200
                  Width           =   1800
               End
               Begin VB.TextBox txtTel 
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
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   1800
                  Width           =   1800
               End
               Begin VB.TextBox txtmanger 
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
                  Left            =   8040
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   1800
                  Width           =   4080
               End
               Begin VB.TextBox txtname 
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
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   240
                  Width           =   3960
               End
               Begin VB.CheckBox ChkLocked 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   14400
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   1440
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
                  Left            =   15240
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   2160
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
                  Left            =   15000
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   1560
                  Width           =   1695
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   14880
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   3480
                  Width           =   1590
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   15840
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Text            =   "0"
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   12240
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   14760
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   3600
                  Width           =   2310
               End
               Begin VB.TextBox xxxxxxxxxx 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   -3930
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   9360
                  Width           =   2175
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4995
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   3120
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3555
                  Left            =   0
                  TabIndex        =   7
                  Top             =   2370
                  Width           =   14625
                  _cx             =   25797
                  _cy             =   6271
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
                  FormatString    =   $"FrmBranchesData.frx":040F
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
               Begin MSComCtl2.DTPicker dbFromDate 
                  Height          =   315
                  Left            =   15840
                  TabIndex        =   12
                  Top             =   960
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   235864065
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   16320
                  TabIndex        =   13
                  Top             =   2040
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
                  Left            =   14880
                  TabIndex        =   15
                  Top             =   1560
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
                  Left            =   15720
                  TabIndex        =   30
                  Top             =   840
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
                  Height          =   315
                  Left            =   15240
                  TabIndex        =   37
                  Top             =   960
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   243269633
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   1080
                  TabIndex        =   43
                  Top             =   1920
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
                  ButtonImage     =   "FrmBranchesData.frx":08B2
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   120
                  TabIndex        =   44
                  Top             =   1920
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
                  ButtonImage     =   "FrmBranchesData.frx":0C4C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DBCboClientName 
                  Height          =   315
                  Left            =   15360
                  TabIndex        =   45
                  Top             =   0
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
               Begin MSDataListLib.DataCombo dcitems 
                  Height          =   315
                  Left            =   15600
                  TabIndex        =   46
                  Top             =   2160
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
               Begin MSDataListLib.DataCombo DCRegionID 
                  Height          =   315
                  Left            =   3720
                  TabIndex        =   65
                  Top             =   1800
                  Width           =   4080
                  _ExtentX        =   7197
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
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
                  Caption         =   "«÷€ÿ · ⁄—Ì› «·„‰«ÿﬁ>>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Index           =   15
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   2160
                  Width           =   2280
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰ÿﬁ…"
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
                  Index           =   14
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   1560
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„  ”ÃÌ· «·›« "
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
                  Index           =   13
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   1560
                  Width           =   1440
               End
               Begin VB.Image ImgPic 
                  Height          =   495
                  Left            =   120
                  Stretch         =   -1  'True
                  Top             =   600
                  Width           =   3135
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·‰‘«ÿ «‰Ã·Ì“Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   12
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   240
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·›—⁄"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   11
                  Left            =   12840
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   960
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·›—Ê⁄ «· «»⁄Â ··‰‘«ÿ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Index           =   10
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   720
                  Width           =   2040
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ·Ì›Ê‰"
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
                  Index           =   9
                  Left            =   12360
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   1560
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄‰Ê«‰ «·›—⁄"
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
                  Index           =   6
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   1560
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·›—⁄ «‰Ã·Ì“Ì"
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
                  Index           =   4
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   960
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·‰‘«ÿ ⁄—»Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Index           =   3
                  Left            =   10560
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   240
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·›—⁄ ⁄—»Ì"
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
                  Index           =   2
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   960
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ê—œ"
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
                  Index           =   0
                  Left            =   15045
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   360
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ Â« „‰"
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
                  Index           =   5
                  Left            =   15405
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   1200
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   8
                  Left            =   14520
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   2640
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·‰‘«ÿ "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   7
                  Left            =   13260
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   375
                  Left            =   13680
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   960
                  Width           =   855
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
            Height          =   7800
            Index           =   30
            Left            =   45
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   45
            Width           =   14655
            _cx             =   25850
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
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Index           =   10
               Left            =   10650
               MaxLength       =   2
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   3270
               Width           =   2865
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Index           =   2
               Left            =   10650
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   2790
               Width           =   2865
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Index           =   7
               Left            =   10650
               MaxLength       =   5
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Tag             =   "5 digit at least"
               Top             =   2340
               Width           =   2865
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Index           =   4
               Left            =   10650
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Tag             =   "4 digit at least"
               Top             =   1920
               Width           =   2865
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   3
               Left            =   10650
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   1500
               Width           =   2865
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   0
               Left            =   10650
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   1170
               Width           =   2865
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   8
               Left            =   1410
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   86
               Top             =   120
               Width           =   9075
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   7
               Left            =   1410
               MaxLength       =   150
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   85
               Top             =   570
               Width           =   9075
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   5
               Left            =   1410
               MaxLength       =   150
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   84
               Top             =   1530
               Width           =   9075
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   4
               Left            =   1410
               MaxLength       =   150
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   83
               Top             =   1050
               Width           =   9075
            End
            Begin VB.ComboBox DefaultInvoicetype 
               Height          =   330
               ItemData        =   "FrmBranchesData.frx":11E6
               Left            =   1410
               List            =   "FrmBranchesData.frx":11E8
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Top             =   2580
               Width           =   2625
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   735
               Index           =   9
               Left            =   1410
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   81
               Top             =   4080
               Width           =   9075
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   11
               Left            =   1410
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   80
               Top             =   3600
               Width           =   9075
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   10
               Left            =   1410
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   79
               Top             =   3120
               Width           =   9075
            End
            Begin VB.ComboBox Invoicetype 
               Height          =   330
               ItemData        =   "FrmBranchesData.frx":11EA
               Left            =   5490
               List            =   "FrmBranchesData.frx":11EC
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   78
               Top             =   2580
               Width           =   3465
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   735
               Index           =   14
               Left            =   1410
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   77
               Top             =   5685
               Width           =   9075
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   735
               Index           =   13
               Left            =   1410
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   76
               Top             =   4920
               Width           =   9075
            End
            Begin VB.ComboBox SendingMode 
               Height          =   330
               ItemData        =   "FrmBranchesData.frx":11EE
               Left            =   1140
               List            =   "FrmBranchesData.frx":11F0
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   75
               Top             =   7425
               Width           =   2535
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   735
               Index           =   15
               Left            =   1410
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               Top             =   6510
               Width           =   9075
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›« Ê—…   —”· «·Ì« »„Ã—œ «·Õ›Ÿ"
               Height          =   285
               Index           =   212
               Left            =   7890
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   7410
               Width           =   2325
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ›⁄Ì· «·„—Õ·… «·À«‰Ì…"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   213
               Left            =   5370
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   7410
               Width           =   1545
            End
            Begin VB.TextBox XPTxtComment 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   6
               Left            =   1410
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   71
               Top             =   2040
               Width           =   9075
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   7800
               Index           =   0
               Left            =   60
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   90
               Width           =   14655
               _cx             =   25850
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
               Begin VB.TextBox XPTxtCompany 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   10590
                  LinkTimeout     =   255
                  MaxLength       =   255
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   330
                  Width           =   2865
               End
               Begin VB.TextBox XPTxtCompanye 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   10590
                  LinkTimeout     =   255
                  MaxLength       =   255
                  RightToLeft     =   -1  'True
                  TabIndex        =   151
                  Top             =   690
                  Width           =   2865
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "‰”Œ ··›—Ê⁄ «· «»⁄… ··‰‘«ÿ"
                  Height          =   285
                  Left            =   12660
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   7050
                  Width           =   1845
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Õ›Ÿ «·›—⁄"
                  Height          =   285
                  Left            =   10620
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   7050
                  Width           =   1995
               End
               Begin VB.TextBox txtNoOFDigitUser 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   8
                  Left            =   10620
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   6060
                  Width           =   2865
               End
               Begin VB.TextBox txtNoOFDigitUser 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   3
                  Left            =   10620
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   5490
                  Width           =   2865
               End
               Begin VB.TextBox txtNoOFDigitUser 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   5
                  Left            =   10620
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   4740
                  Width           =   2865
               End
               Begin VB.TextBox txtNoOFDigitUser 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0FF&
                  Height          =   315
                  Index           =   6
                  Left            =   10620
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   4230
                  Width           =   2865
               End
               Begin VB.TextBox txtNoOFDigitUser 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0FF&
                  Height          =   315
                  Index           =   9
                  Left            =   10590
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   3720
                  Width           =   2865
               End
               Begin MSDataListLib.DataCombo dcBranch 
                  Height          =   330
                  Left            =   10560
                  TabIndex        =   148
                  Top             =   6540
                  Width           =   2925
                  _ExtentX        =   5159
                  _ExtentY        =   582
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·‘—ﬂ… «‰Ã·Ì“Ì"
                  Height          =   375
                  Index           =   42
                  Left            =   12840
                  RightToLeft     =   -1  'True
                  TabIndex        =   156
                  Top             =   750
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·‘—ﬂ… ⁄—»Ì"
                  Height          =   375
                  Index           =   41
                  Left            =   12780
                  RightToLeft     =   -1  'True
                  TabIndex        =   155
                  Top             =   360
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·‘—ﬂ… ⁄—»Ì"
                  Height          =   375
                  Index           =   40
                  Left            =   15780
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   330
                  Width           =   315
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·‘—ﬂ… «‰Ã·Ì“Ì"
                  Height          =   375
                  Index           =   39
                  Left            =   15780
                  RightToLeft     =   -1  'True
                  TabIndex        =   153
                  Top             =   780
                  Width           =   315
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›—⁄"
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Left            =   13575
                  TabIndex        =   149
                  Top             =   6540
                  Width           =   555
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„œÌ‰… «·›—⁄Ì…"
                  Height          =   375
                  Index           =   51
                  Left            =   13350
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   6180
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·‘«—⁄2"
                  Height          =   375
                  Index           =   46
                  Left            =   13350
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   5610
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·„Œÿÿ"
                  Height          =   375
                  Index           =   48
                  Left            =   13350
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   4860
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„œÌ‰…*"
                  Height          =   375
                  Index           =   49
                  Left            =   12630
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   4230
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Org name"
                  Height          =   375
                  Index           =   38
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   2115
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Secret Key"
                  Height          =   375
                  Index           =   37
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   6585
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ê÷⁄ «·Õ«·Ì"
                  Height          =   375
                  Index           =   36
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   7470
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Private Key"
                  Height          =   375
                  Index           =   35
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   5160
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Public key certpem"
                  Height          =   495
                  Index           =   34
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   5880
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‰Ê⁄ «·«› —«÷Ì"
                  Height          =   375
                  Index           =   33
                  Left            =   3810
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   2640
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "CSR"
                  Height          =   375
                  Index           =   32
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   4200
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Invoice type"
                  Height          =   375
                  Index           =   30
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   2640
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Location"
                  Height          =   375
                  Index           =   29
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   3120
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Industry"
                  Height          =   255
                  Index           =   28
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   3600
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Common Name"
                  Height          =   375
                  Index           =   27
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   90
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Serial Number"
                  Height          =   375
                  Index           =   26
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   570
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Org Identifier"
                  Height          =   255
                  Index           =   25
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   1050
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Org Unit name"
                  Height          =   375
                  Index           =   24
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   1605
                  Width           =   1725
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„—Õ·… «·À«‰·Ì…"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Index           =   0
                  Left            =   11640
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   30
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”Ã·"
                  Height          =   375
                  Index           =   23
                  Left            =   12630
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   1080
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· ”ÃÌ· VAT"
                  Height          =   375
                  Index           =   22
                  Left            =   12810
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   1470
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·„»‰Ï*"
                  Height          =   375
                  Index           =   20
                  Left            =   12630
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   2010
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—„“ «·»—ÌœÌ*"
                  Height          =   375
                  Index           =   19
                  Left            =   12630
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   2460
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·‘«—⁄*"
                  Height          =   375
                  Index           =   18
                  Left            =   12630
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   2880
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·œÊ·…*"
                  Height          =   345
                  Index           =   17
                  Left            =   12630
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   3300
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÕÌ*"
                  Height          =   375
                  Index           =   16
                  Left            =   12630
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   3690
                  Width           =   1725
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÕÌ*"
               Height          =   375
               Index           =   52
               Left            =   13170
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   3690
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·œÊ·…*"
               Height          =   465
               Index           =   53
               Left            =   13260
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   3300
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·‘«—⁄*"
               Height          =   375
               Index           =   45
               Left            =   13140
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   2880
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—„“ «·»—ÌœÌ*"
               Height          =   375
               Index           =   50
               Left            =   12990
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   2460
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·„»‰Ï*"
               Height          =   375
               Index           =   47
               Left            =   13170
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   2010
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «· ”ÃÌ· VAT"
               Height          =   375
               Index           =   31
               Left            =   12690
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   1560
               Width           =   1725
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·”Ã·"
               Height          =   375
               Index           =   21
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   780
               Width           =   1725
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„—Õ·… «·À«‰·Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   2
               Left            =   10890
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   240
               Width           =   1350
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Org Unit name"
               Height          =   375
               Index           =   83
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   1605
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Org Identifier"
               Height          =   255
               Index           =   69
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   1050
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Serial Number"
               Height          =   375
               Index           =   68
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   570
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Common Name"
               Height          =   375
               Index           =   67
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   90
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Industry"
               Height          =   255
               Index           =   88
               Left            =   210
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   3600
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Location"
               Height          =   375
               Index           =   87
               Left            =   210
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   3120
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Invoice type"
               Height          =   375
               Index           =   86
               Left            =   210
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   2640
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "CSR"
               Height          =   375
               Index           =   79
               Left            =   210
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   4200
               Width           =   1725
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‰Ê⁄ «·«› —«÷Ì"
               Height          =   375
               Index           =   90
               Left            =   4170
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   2580
               Width           =   1485
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Public key certpem"
               Height          =   495
               Index           =   81
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   5880
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Private Key"
               Height          =   375
               Index           =   80
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   5160
               Width           =   1725
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ê÷⁄ «·Õ«·Ì"
               Height          =   375
               Index           =   89
               Left            =   3540
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   7470
               Width           =   1485
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Secret Key"
               Height          =   375
               Index           =   82
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   6585
               Width           =   1725
            End
            Begin VB.Label lbl 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Org name"
               Height          =   375
               Index           =   84
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   2115
               Width           =   1725
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   8265
         Visible         =   0   'False
         Width           =   14745
         _cx             =   26009
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
            ButtonImage     =   "FrmBranchesData.frx":11F2
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
            ButtonImage     =   "FrmBranchesData.frx":158C
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
            ButtonImage     =   "FrmBranchesData.frx":1926
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   7980
            TabIndex        =   23
            Top             =   480
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
            Left            =   7080
            TabIndex        =   24
            Top             =   510
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
            Left            =   6240
            TabIndex        =   25
            Top             =   510
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
            Left            =   5235
            TabIndex        =   26
            Top             =   510
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
            Left            =   4200
            TabIndex        =   27
            Top             =   510
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
            Top             =   510
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
            Left            =   3270
            TabIndex        =   29
            Top             =   510
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
            Left            =   9120
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
            MICON           =   "FrmBranchesData.frx":1CC0
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
      ButtonImage     =   "FrmBranchesData.frx":1CDC
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Index           =   5
      Left            =   30
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   0
      Width           =   14775
      _cx             =   26061
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
      Picture         =   "FrmBranchesData.frx":2076
      Caption         =   "»Ì«‰«  «·«‰‘ÿÂ Ê «·›—Ê⁄  "
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
      Begin VB.TextBox Text1 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   210
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   61
         Top             =   480
         Width           =   2055
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   57
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
         ButtonImage     =   "FrmBranchesData.frx":2D50
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
         TabIndex        =   58
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
         ButtonImage     =   "FrmBranchesData.frx":30EA
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
         TabIndex        =   59
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
         ButtonImage     =   "FrmBranchesData.frx":3484
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
         TabIndex        =   60
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
         ButtonImage     =   "FrmBranchesData.frx":381E
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
Attribute VB_Name = "FrmBranchesData"
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
Dim lastBrandh_id As Integer
Dim Account_Code_dynamic As String

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Public Sub YearMonth()

End Sub

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub

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
    create_report_data

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
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Private Sub ALLButton2_Click()
    'Dcemp.text = ""

    dcproject.text = ""
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
            .rows = 2
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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If txtid.text <> "" Then
        'StrSQL = "select * From Notes where BoxID=" & Trim(XPTxtBoxID.text)
        'RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        'If Not (RsTemp.EOF Or RsTemp.BOF) Then
        '    Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Chr(13)
        '    Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «·Œ“‰…"
        '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    Exit Sub
        'End If
        Msg = "”Ì „ Õ–› »Ì«‰«  «·‰‘«ÿ —ﬁ„ " & CHR(13)
        Msg = Msg + (txtid.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Dim StrAccountCode As String
             
                StrSQL = "delete TblBranchesData where ActivityTypeId=" & val(txtid.text)
                Cn.Execute StrSQL

                rs.delete
             
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                 
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
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
        Msg = Msg & CHR(13) & Err.Description
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

   On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
 
        If Trim(Me.TxtName.text) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»    «œŒ«· «”„ «·‰‘«ÿ ..!!"
                        Else
                            Msg = "Enter Activity Name"
                        End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             TxtName.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
 
 
         If Trim(Me.TxtNameE.text) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»    «œŒ«· «”„ «·‰‘«ÿ  »«··€… «·«‰Ã·Ì“Ì… ..!!"
                        Else
                            Msg = "Enter Activity Name in English"
                        End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             TxtNameE.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
        
        
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
        rs.AddNew
                    Me.txtid.text = CStr(new_id("tblActivitesType", "id", "", True))

    ElseIf Me.TxtModFlg.text = "E" Then
        Cn.Execute "delete TblBranchesData where ActivityTypeId=" & val(Me.txtid.text)
   Cn.Execute "delete TblUsersBranches where ActivityTypeId=" & val(Me.txtid.text)
   
    End If
    
    rs("id").value = val(Me.txtid.text)
    rs("name").value = Trim(Me.TxtName.text)
    rs("namee").value = Trim(Me.TxtNameE.text)
 
 
              
            
    rs("Company_Arabic_Name").value = XPTxtCompany.text
    rs("Company_Name_Eng").value = XPTxtCompanye.text
    
    
    rs("Commonname").value = XPTxtComment(8)
    rs("SerialNumber").value = XPTxtComment(7)
    rs("OrganizationName").value = XPTxtComment(6)
        
    rs("Invoicetype").value = Me.Invoicetype.ListIndex
    rs("DefaultInvoicetype").value = Me.DefaultInvoicetype.ListIndex
    rs("SendingMode").value = Me.SendingMode.ListIndex
    
     

    rs("industrey").value = XPTxtComment(11)
    rs("CSR").value = XPTxtComment(9)
    rs("Privatekey").value = XPTxtComment(13)
    rs("PublickeycertPem").value = XPTxtComment(14)
    rs("SecretKey").value = XPTxtComment(15)
 savepart2
 
 
    rs.update
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TblBranchesData", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer
 
    With Me.Grid

        For i = 1 To .rows - 1

            If .TextMatrix(i, .ColIndex("branch_id")) <> "" Then
          
                RsDev.AddNew
                RsDev("ActivityTypeId").value = Me.txtid.text
                RsDev("branch_id").value = val(.TextMatrix(i, .ColIndex("branch_id")))
                RsDev("branch_Code").value = (.TextMatrix(i, .ColIndex("branch_Code")))
                RsDev("branch_name").value = .TextMatrix(i, .ColIndex("branch_name"))
                RsDev("branch_namee").value = .TextMatrix(i, .ColIndex("branch_namee"))
                RsDev("StoreId").value = val(.TextMatrix(i, .ColIndex("StoreId")))
                
                RsDev("manger").value = .TextMatrix(i, .ColIndex("manger"))
                RsDev("Tel").value = .TextMatrix(i, .ColIndex("Tel"))
                RsDev("Users").value = .TextMatrix(i, .ColIndex("Users"))
                RsDev("RegionID").value = val(.TextMatrix(i, .ColIndex("RegionID")))
                RsDev("VATNO").value = (.TextMatrix(i, .ColIndex("VATNO")))
             RsDev("Beauty").value = val(.TextMatrix(i, .ColIndex("Beauty")))
             
                If .TextMatrix(i, .ColIndex("Account_Code")) = "" Then
                    RsDev("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, "Ã«—Ì  " & Trim$(.TextMatrix(i, .ColIndex("branch_name"))), True, False, Trim(.TextMatrix(i, .ColIndex("branch_namee"))) & "-  Current  ", , , , , , , , , , 1, 1, 1, 0, 0)
                Else
                    ModAccounts.EditAccount .TextMatrix(i, .ColIndex("Account_Code")), "Ã«—Ì  " & Trim$(.TextMatrix(i, .ColIndex("branch_name"))), Trim$(.TextMatrix(i, .ColIndex("branch_namee"))) & "-  Current  ", , , , , , , , , 1, 1, 1, 0, 0, , , , True
                    RsDev("Account_Code").value = .TextMatrix(i, .ColIndex("Account_Code"))
                End If
             
         
                    .Row = i
    .Col = 26
'On Error Resume Next
  ImgPic.Picture = Me.Grid.CellPicture
  
                 If Me.ImgPic.Picture <> 0 Then
                 ImgPic.Picture = Me.Grid.CellPicture
        SavePictureToDB ImgPic, RsDev, "branchLogo"
         RsDev("ShowlogoInReports").value = 1
         Else
          RsDev("ShowlogoInReports").value = 0
  End If
       RsDev.update

                LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«  «·«‰‘ÿÂ Ê «·›—Ê⁄ " & " ﬂÊœ «·›—⁄  " & (.TextMatrix(i, .ColIndex("branch_Code"))) & " Ê «”„… " & .TextMatrix(i, .ColIndex("branch_name")) & " Ê «·„œÌ— " & .TextMatrix(i, .ColIndex("manger")) & " Ê «· ·Ì›Ê‰ " & .TextMatrix(i, .ColIndex("Tel"))
 
                LogTexte = "    Save Window   " & "   Activity And Branches Data " & " Branch Code  " & (.TextMatrix(i, .ColIndex("branch_Code"))) & "Name " & .TextMatrix(i, .ColIndex("branch_namee")) & " Manger " & .TextMatrix(i, .ColIndex("manger")) & " Tel " & .TextMatrix(i, .ColIndex("Tel"))
               
                AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    
            End If
            
            '
        Next i

    End With
 
 RsDev.Close
 Set RsDev = Nothing
 
'branches users
    Dim RsDev1 As New ADODB.Recordset
    Dim j As Integer
Dim test_split() As String
Dim s As String
 
 
Dim sql As String
Dim nElements As Integer

    RsDev1.Open "TblUsersBranches", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
  
 
    With Me.Grid

        For i = 1 To .rows - 1

            If .TextMatrix(i, .ColIndex("Users")) <> "" Then
         .TextMatrix(i, .ColIndex("Updated")) = 0
                         Cn.Execute "delete TblUsersBranches where BranchID=" & val(.TextMatrix(i, .ColIndex("branch_id")))
                 
                 
                 
                 s = Grid.TextMatrix(i, Grid.ColIndex("Users"))
test_split = Split(s, ",")
 
nElements = UBound(test_split) - LBound(test_split) ' To UBound(cSearchDcbo)

'MsgBox nElements

           '     astrSplitItems = Split(AllDate, strFilterText)
       '  astrSplitItems1 = Split(Alliids, strFilterText)
         
            '    For intX = 0 To UBound(astrSplitItems)
                
                    For j = 0 To UBound(test_split)
                    

                              RsDev1.AddNew
                                  RsDev1("ActivityTypeId").value = Me.txtid.text
                                  RsDev1("BranchID").value = val(.TextMatrix(i, .ColIndex("branch_id")))
                                   RsDev1("userid").value = val(test_split(j))
                                   RsDev1.update
                                   
                    Next j
  
       
            End If
            
            '
        Next i

    End With
 
 
 RsDev1.Close
 Set RsDev1 = Nothing
  
  
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = "Operation Successfully Saved " & CHR(13)
                Msg = Msg + "Do You want add another Record"
            End If

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Successfully Update", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If

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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Can't  Save " & CHR(13)
            Msg = Msg + " Error In Entered values " & CHR(13)
            Msg = Msg + "Check Entered values"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry ... Error During Saving Data" & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub



Function savepart2()
        If Me.Chkbarcode(212).value = vbChecked Then
        rs("IsBluee").value = 1
    ElseIf Me.Chkbarcode(212).value = vbUnchecked Then
        rs("IsBluee").value = 0
    End If
    
        rs("Company_Comment").value = XPTxtComment(0).text

    rs("VATRegNo").value = XPTxtComment(3).text

             If Me.Chkbarcode(213).value = vbChecked Then
        rs("ApplyEinvoice").value = 1
    ElseIf Me.Chkbarcode(213).value = vbUnchecked Then
        rs("ApplyEinvoice").value = 0
    End If
    
       rs("StreetName").value = txtNoOFDigitUser(2).text
        rs("BuildingNumber").value = txtNoOFDigitUser(4).text
         rs("CitySubdivisionName").value = txtNoOFDigitUser(9).text
          rs("CityName").value = txtNoOFDigitUser(6).text
           rs("PostalZone").value = txtNoOFDigitUser(7).text
            rs("IdentificationCode").value = txtNoOFDigitUser(10).text
             rs("PlotIdentification").value = txtNoOFDigitUser(5).text
              rs("AdditionalStreetName").value = txtNoOFDigitUser(3).text
              rs("CountrySubentity").value = txtNoOFDigitUser(8).text
              
          
                  
    
        
End Function
Private Sub Cmd_Click(Index As Integer)

    On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
 
            lastBrandh_id = CStr(new_id("TblBranchesData", "branch_id", "", True))
       
            'Me.dbFromDate.value = Date
            'Me.dbTodate.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1
            Grid.Enabled = True

            '   Option2.value = True
        Case 1
                    
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 
            TxtModFlg.text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
 
            lastBrandh_id = CStr(new_id("TblBranchesData", "branch_id", "", True))

        Case 2

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 
            Account_Code_dynamic = get_account_code_branch(72, my_branch)
                    
            If Account_Code_dynamic = "NO account" Then
                MsgBox " ·„ Ì „  ÕœÌœ Õ”«» Ã«—Ì «·›—Ê⁄ ", vbCritical
                Exit Sub
                     
            End If
 
            LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«  «·«‰‘ÿÂ Ê «·›—Ê⁄ " & " «·‰‘«ÿ " & TxtName
            LogTexte = " Save Window " & "   Activity And Branches Data " & " Activity " & TxtNameE
            AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Company

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
        Case 20
            addrow

        Case 21
            RemoveGridRow
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

Function addrow()

    Dim wherestr As String

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim RsUnit As ADODB.Recordset
    Set RsUnit = New ADODB.Recordset

    Dim j As Integer

    Dim sql As String
    Dim i As Integer
    Dim Msg  As String
    Dim lastrow As Integer
    Dim LngItemID As Integer
 
    If TxtName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»          «œŒ«· «”„ «·›—⁄   ...!!!"
        Else
            Msg = "must Specify branch Name ...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If
 
    With Grid
 
        lastrow = .rows
        .rows = lastrow + 1
         
        .TextMatrix(lastrow, .ColIndex("branch_id")) = lastBrandh_id
        lastBrandh_id = lastBrandh_id + 1
        .TextMatrix(lastrow, .ColIndex("RegionID")) = val(Me.DCRegionID.BoundText)
        .TextMatrix(lastrow, .ColIndex("RegionName")) = DCRegionID.text
        .TextMatrix(lastrow, .ColIndex("VATNO")) = TxtVATNO.text
        
        .TextMatrix(lastrow, .ColIndex("branch_Code")) = txtBranch_Code.text
        .TextMatrix(lastrow, .ColIndex("branch_name")) = txtbranch_name.text
        .TextMatrix(lastrow, .ColIndex("branch_namee")) = txtbranch_namee.text
        .TextMatrix(lastrow, .ColIndex("manger")) = TxtManger.text
        .TextMatrix(lastrow, .ColIndex("Tel")) = txtTel.text
      
    End With
 
    txtbranch_name.text = ""
    txtbranch_namee.text = ""
    TxtManger.text = ""
    txtTel.text = ""
    txtBranch_Code.text = ""
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

Private Sub Dcdep_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcedara_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
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

    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim BranchID As Integer
 
    BranchID = val(Grid.TextMatrix(Grid.Row, Grid.ColIndex("branch_id")))
 
    StrSQL = "select * From Transaction_Details where BranchId=" & BranchID
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·›—⁄ ·ÊÃÊœ Õ—ﬂ«  ⁄·ÌÂ" & CHR(13)
        Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«   „— »ÿ… »Â–« «·›—⁄"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    RsTemp.Close
    StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where branch_id=" & BranchID
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·›—⁄ ·ÊÃÊœ Õ—ﬂ«  ⁄·ÌÂ" & CHR(13)
        Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«   „— »ÿ… »Â–« «·›—⁄"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

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

 

Private Sub Command1_Click()
 Dim s As String
 Dim rsDummy As New ADODB.Recordset
 
 s = "Select * from TblBranchesData  where branch_id = " & val(dcBranch.BoundText)
 Set rsDummy = New ADODB.Recordset
 
 rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
 If Not rsDummy.EOF Then
 
 
               
            
        
    rsDummy("Company_Arabic_Name").value = XPTxtCompany
    rsDummy("Company_Name_Eng").value = XPTxtCompanye
               
    rsDummy("Commonname").value = XPTxtComment(8)
    rsDummy("SerialNumber").value = XPTxtComment(7)
    rsDummy("OrganizationName").value = XPTxtComment(6)
        
    rsDummy("Invoicetype").value = Me.Invoicetype.ListIndex
    rsDummy("DefaultInvoicetype").value = Me.DefaultInvoicetype.ListIndex
    rsDummy("SendingMode").value = Me.SendingMode.ListIndex
    
     

    rsDummy("industrey").value = XPTxtComment(11)
    rsDummy("CSR").value = XPTxtComment(9)
    rsDummy("Privatekey").value = XPTxtComment(13)
    rsDummy("PublickeycertPem").value = XPTxtComment(14)
    rsDummy("SecretKey").value = XPTxtComment(15)
    
    If Me.Chkbarcode(212).value = vbChecked Then
        rsDummy("IsBluee").value = 1
    ElseIf Me.Chkbarcode(212).value = vbUnchecked Then
        rsDummy("IsBluee").value = 0
    End If
    
    
    rsDummy("Company_Comment").value = XPTxtComment(0).text

    rsDummy("VATRegNo").value = XPTxtComment(3).text

    If Me.Chkbarcode(213).value = vbChecked Then
        rsDummy("ApplyEinvoice").value = 1
    ElseIf Me.Chkbarcode(213).value = vbUnchecked Then
        rsDummy("ApplyEinvoice").value = 0
    End If
    
    rsDummy("StreetName").value = txtNoOFDigitUser(2).text
    rsDummy("BuildingNumber").value = txtNoOFDigitUser(4).text
    rsDummy("CitySubdivisionName").value = txtNoOFDigitUser(9).text
    rsDummy("CityName").value = txtNoOFDigitUser(6).text
    rsDummy("PostalZone").value = txtNoOFDigitUser(7).text
    rsDummy("IdentificationCode").value = txtNoOFDigitUser(10).text
    rsDummy("PlotIdentification").value = txtNoOFDigitUser(5).text
    rsDummy("AdditionalStreetName").value = txtNoOFDigitUser(3).text
    rsDummy("CountrySubentity").value = txtNoOFDigitUser(8).text
    rsDummy.update
    MsgBox "Data Saved"
          
 End If

End Sub

Private Sub Command2_Click()
 Dim s As String
 Dim rsDummy As New ADODB.Recordset
 Dim rsDummy2 As New ADODB.Recordset
 Dim mActivityTypeId As Long
 s = "Select TblBranchesData.ActivityTypeId from TblBranchesData  where branch_id = " & val(dcBranch.BoundText)
 Set rsDummy2 = New ADODB.Recordset
 rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
 If Not rsDummy2.EOF Then
    mActivityTypeId = val(rsDummy2!ActivityTypeId & "")
    
 End If
 
 
 Set rsDummy = New ADODB.Recordset
 s = "Select * from TblBranchesData  where TblBranchesData.ActivityTypeId= " & mActivityTypeId
 rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
 Do While Not rsDummy.EOF
 
 
               
               
               
            
        
    rsDummy("Company_Arabic_Name").value = XPTxtCompany
    rsDummy("Company_Name_Eng").value = XPTxtCompanye
    
               
    rsDummy("Commonname").value = XPTxtComment(8)
    rsDummy("SerialNumber").value = XPTxtComment(7)
    rsDummy("OrganizationName").value = XPTxtComment(6)
        
    rsDummy("Invoicetype").value = Me.Invoicetype.ListIndex
    rsDummy("DefaultInvoicetype").value = Me.DefaultInvoicetype.ListIndex
    rsDummy("SendingMode").value = Me.SendingMode.ListIndex
    
     

    rsDummy("industrey").value = XPTxtComment(11)
    rsDummy("CSR").value = XPTxtComment(9)
    rsDummy("Privatekey").value = XPTxtComment(13)
    rsDummy("PublickeycertPem").value = XPTxtComment(14)
    rsDummy("SecretKey").value = XPTxtComment(15)
    
    If Me.Chkbarcode(212).value = vbChecked Then
        rsDummy("IsBluee").value = 1
    ElseIf Me.Chkbarcode(212).value = vbUnchecked Then
        rsDummy("IsBluee").value = 0
    End If
    
    rsDummy("Company_Comment").value = XPTxtComment(0).text

    rsDummy("VATRegNo").value = XPTxtComment(3).text

    If Me.Chkbarcode(213).value = vbChecked Then
        rsDummy("ApplyEinvoice").value = 1
    ElseIf Me.Chkbarcode(213).value = vbUnchecked Then
        rsDummy("ApplyEinvoice").value = 0
    End If
    
    rsDummy("StreetName").value = txtNoOFDigitUser(2).text
    rsDummy("BuildingNumber").value = txtNoOFDigitUser(4).text
    rsDummy("CitySubdivisionName").value = txtNoOFDigitUser(9).text
    rsDummy("CityName").value = txtNoOFDigitUser(6).text
    rsDummy("PostalZone").value = txtNoOFDigitUser(7).text
    rsDummy("IdentificationCode").value = txtNoOFDigitUser(10).text
    rsDummy("PlotIdentification").value = txtNoOFDigitUser(5).text
    rsDummy("AdditionalStreetName").value = txtNoOFDigitUser(3).text
    rsDummy("CountrySubentity").value = txtNoOFDigitUser(8).text
    rsDummy.update
   
     rsDummy.MoveNext
 Loop
 MsgBox "Data Saved"

End Sub

Private Sub dcBranch_Validate(Cancel As Boolean)
 Dim s As String
 Dim rsDummy As New ADODB.Recordset
 s = "Select * from TblBranchesData  where branch_id = " & val(dcBranch.BoundText)
 Set rsDummy = New ADODB.Recordset
 
 rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
 
        XPTxtComment(10) = ""
        
        XPTxtComment(8) = ""
        XPTxtComment(7) = ""
        'XPTxtComment(6) = IIf(IsNull(rs("OrganizationName").value), "", rs("OrganizationName").value)
        
        'XPTxtComment(11) = IIf(IsNull(rs("industrey").value), "", rs("industrey").value)
        XPTxtComment(9) = ""
        XPTxtComment(13) = ""
        XPTxtComment(14) = ""
        XPTxtComment(15) = ""
        
        

        If Not rsDummy.EOF Then
         
         
                txtNoOFDigitUser(2).text = IIf(IsNull(rsDummy("StreetName").value), "", rsDummy("StreetName").value)
                txtNoOFDigitUser(4).text = IIf(IsNull(rsDummy("BuildingNumber").value), "", rsDummy("BuildingNumber").value)
                txtNoOFDigitUser(9).text = IIf(IsNull(rsDummy("CitySubdivisionName").value), "", rsDummy("CitySubdivisionName").value)
                txtNoOFDigitUser(6).text = IIf(IsNull(rsDummy("CityName").value), "", rsDummy("CityName").value)
                txtNoOFDigitUser(7).text = IIf(IsNull(rsDummy("PostalZone").value), "", rsDummy("PostalZone").value)
                txtNoOFDigitUser(10).text = IIf(IsNull(rsDummy("IdentificationCode").value), "SA", rsDummy("IdentificationCode").value)
                If txtNoOFDigitUser(10).text = "" Then txtNoOFDigitUser(10).text = "SA"
                txtNoOFDigitUser(5).text = IIf(IsNull(rsDummy("PlotIdentification").value), "", rsDummy("PlotIdentification").value)
                txtNoOFDigitUser(3).text = IIf(IsNull(rsDummy("AdditionalStreetName").value), "", rsDummy("AdditionalStreetName").value)
                txtNoOFDigitUser(8).text = IIf(IsNull(rsDummy("CountrySubentity").value), "", rsDummy("CountrySubentity").value)
                
                XPTxtComment(3).text = IIf(IsNull(rsDummy("VATRegNo").value), "", rsDummy("VATRegNo").value)
                XPTxtComment(0).text = IIf(IsNull(rsDummy("Company_Comment").value), "", rsDummy("Company_Comment").value)
                
                
                If rsDummy("IsBluee").value = True Then
                Me.Chkbarcode(212).value = vbChecked
                Else
                Me.Chkbarcode(212).value = Unchecked
                End If
                
                
                If rsDummy("ApplyEinvoice").value = 1 Then
                    Me.Chkbarcode(213).value = vbChecked
                Else
                    Me.Chkbarcode(213).value = Unchecked
                End If
                XPTxtComment(4) = XPTxtComment(3)
                XPTxtComment(5) = XPTxtComment(0)
                'Wael
                'XPTxtComment(12) = txtNoOFDigitUser(10)
                XPTxtComment(10) = txtNoOFDigitUser(6)
                
                               
            
                    

                XPTxtCompany = rsDummy("Company_Arabic_Name").value & ""
                XPTxtCompanye = rsDummy("Company_Name_Eng").value & ""
    
                
                XPTxtComment(8) = IIf(IsNull(rsDummy("Commonname").value), "", rsDummy("Commonname").value)
                XPTxtComment(7) = IIf(IsNull(rsDummy("SerialNumber").value), "", rsDummy("SerialNumber").value)
                XPTxtComment(6) = IIf(IsNull(rsDummy("OrganizationName").value), "", rsDummy("OrganizationName").value)
                
                XPTxtComment(11) = IIf(IsNull(rsDummy("industrey").value), "", rsDummy("industrey").value)
                XPTxtComment(9) = IIf(IsNull(rsDummy("CSR").value), "", rsDummy("CSR").value)
                XPTxtComment(13) = IIf(IsNull(rsDummy("Privatekey").value), "", rsDummy("Privatekey").value)
                XPTxtComment(14) = IIf(IsNull(rsDummy("PublickeycertPem").value), "", rsDummy("PublickeycertPem").value)
                XPTxtComment(15) = IIf(IsNull(rsDummy("SecretKey").value), "", rsDummy("SecretKey").value)
                
                
                Me.Invoicetype.ListIndex = IIf(IsNull(rsDummy("Invoicetype").value), 0, rsDummy("Invoicetype").value)
                Me.DefaultInvoicetype.ListIndex = IIf(IsNull(rsDummy("DefaultInvoicetype").value), 0, rsDummy("DefaultInvoicetype").value)
                
                Me.SendingMode.ListIndex = IIf(IsNull(rsDummy("SendingMode").value), 0, rsDummy("SendingMode").value)
                
        End If
        

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
    LogTextA = "     «·œŒÊ·  «·Ï  ‘«‘… " & " »Ì«‰«  «·«‰‘ÿÂ Ê «·›—Ê⁄ "
    LogTexte = " Open Window " & "   Activity And Branches Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

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
         .ColComboList(.ColIndex("Users1")) = "..."
         .ColComboList(.ColIndex("Logo")) = "..."
         
    End With
With Me.DefaultInvoicetype
            .Clear
            
             


            .AddItem " ›« Ê—… ÷—Ì»Ì…  "
            .ItemData(0) = 0
     
            .AddItem " ›« Ê—… „»”ÿ… "
            .ItemData(1) = 2
         
        End With
        
   With Me.Invoicetype
            .Clear
            
             


            .AddItem "standard  Invoices only  ›« Ê—… ÷—Ì»Ì…  ›ﬁÿ"
            .ItemData(0) = 0
            .AddItem "standard & Simplified Invoices  ›« Ê—… ÷—Ì»Ì… Ê„»”ÿ…"
            .ItemData(1) = 1
            .AddItem "Simplified Invoices only  ›« Ê—… „»”ÿ… ›ﬁÿ"
            .ItemData(2) = 2
         
        End With
        
        
                With Me.SendingMode
            .Clear
            
             


            .AddItem "dev"
            .ItemData(0) = 0
            .AddItem " Simulation"
            .ItemData(1) = 1
            .AddItem "production"
            .ItemData(2) = 2
         
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
     
    
     Dcombos.GetBranches Me.dcBranch
     
    Set BKGrndPic = New ClsBackGroundPic

    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetItemsNames dcitems
     Dcombos.GetSection Me.DCRegionID
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
    StrSQL = "select * From tblActivitesType  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

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
    Cmd(20).Caption = "ˆAdd"
    Cmd(21).Caption = "Remove"
lbl(15).Caption = "Click To Define Areas"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Activity And Branches Data"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(3).Caption = "Activity Name Arabic"
    lbl(12).Caption = "Activity Name English"
    'Ele(3).Caption = "Select Interval"
    'lbl(2).Caption = "Year"
    lbl(10).Caption = "Branches Data"
    lbl(11).Caption = "Branch Code"
    lbl(2).Caption = "Branch NameA"
    lbl(4).Caption = "Branch NameE"
    lbl(6).Caption = "BranchManger"
    lbl(9).Caption = "Branch Tel"

    CmdRemove.Caption = "Remove Line"
   lbl(14).Caption = "Region"
   lbl(13).Caption = "VAT NO."
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        '.TextMatrix(0, .ColIndex("Emp_code")) = "Emp_code"
        '.TextMatrix(0, .ColIndex("Emp_Name")) = "Emp_Name"
        '.TextMatrix(0, .ColIndex("JobTypeName")) = "Job"
        .TextMatrix(0, .ColIndex("branch_id")) = "Branch ID"
        .TextMatrix(0, .ColIndex("Logo")) = "Logo"
        
        .TextMatrix(0, .ColIndex("branch_Code")) = "Branch Code"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name A"
        .TextMatrix(0, .ColIndex("branch_namee")) = "Branch NameE "
        .TextMatrix(0, .ColIndex("manger")) = "Branch manger"
        .TextMatrix(0, .ColIndex("Tel")) = "Branch Tel"
        .TextMatrix(0, .ColIndex("Users1")) = "Users"
        .TextMatrix(0, .ColIndex("VATNO")) = "VAT NO."
        .TextMatrix(0, .ColIndex("RegionName")) = "Region"
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

Private Sub Form_Unload(Cancel As Integer)
    LogTextA = "     «·Œ—ÊÃ  «·Ï  ‘«‘… " & " »Ì«‰«  «·«‰‘ÿÂ Ê «·›—Ê⁄ "
    LogTexte = " Exit  Window " & "   Activity And Branches Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

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
  Case "RegionName"
                 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("RegionID"), False, True)
                .TextMatrix(Row, .ColIndex("RegionID")) = StrAccountCode
                
  Case "StoreName"
                 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("StoreID"), False, True)
                .TextMatrix(Row, .ColIndex("StoreID")) = StrAccountCode
                
                
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

    With Me.Grid

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("branch_id")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
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

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Grid

        Select Case .ColKey(Col)
Case "Logo"
                 With cdg
        '*.jpg,*.jpeg,*.jpe,*.jfif
        .CancelError = False
        .DialogTitle = " ≈Œ Ì«— ’Ê—…"
        'Set The Filter to show pictures only
        .filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|" & "GIF (*.gif)|*.gif|All Files|*.*" ' choose formats to include
        .ShowOpen

        If .FileName <> "" Then
           Set Me.ImgPic.Picture = LoadPicture(.FileName)
 'Me.Grid.CellPicture
  
   Set Me.Grid.CellPicture = Me.ImgPic.Picture
 
        
        End If

   End With
            Case "Users1"
 Unload FrmUsersBranches
FrmUsersBranches.Row = Grid.Row
If Me.TxtModFlg = "E" Or Me.TxtModFlg = "R" Then
                If val(Grid.TextMatrix(Grid.Row, Grid.ColIndex("Updated"))) = 1 Then
                FrmUsersBranches.branch_id = 0
                        
                Else
                  FrmUsersBranches.branch_id = val(Grid.TextMatrix(Grid.Row, Grid.ColIndex("branch_id")))
                End If


Else
        FrmUsersBranches.branch_id = 0
End If
FrmUsersBranches.branches = (Grid.TextMatrix(Grid.Row, Grid.ColIndex("Users")))

 FrmUsersBranches.show
 
        End Select
       End With
End Sub

Private Sub Grid_Click()
                 ImgPic.Picture = Me.Grid.CellPicture
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
        Case "StoreName"

                    StrSQL = "SELECT StoreID,StoreName,StoreNamee FROM TblStore"
                    
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If rs.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                    MyStrList = .BuildComboList(rs, "StoreName", "StoreID")
                    Else
                    MyStrList = .BuildComboList(rs, "StoreNamee", "StoreID")
                    End If
                    End If
                If MyStrList <> "" Then
                    MyStrList = "|" & MyStrList
                End If
                 .ComboList = MyStrList
        Case "RegionName"
            
                    StrSQL = "SELECT  Id,   name, namee"
                    StrSQL = StrSQL & "  From dbo.TblSection"
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If rs.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                    MyStrList = .BuildComboList(rs, "name", "Id")
                    Else
                    MyStrList = .BuildComboList(rs, "namee", "Id")
                    End If
                    End If
                If MyStrList <> "" Then
                    MyStrList = "|" & MyStrList
                End If
                 .ComboList = MyStrList
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
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.txtid.text = IIf(IsNull(rs("id").value), "", rs("id").value)
 
    TxtName.text = IIf(IsNull(rs("name").value), "", rs("name").value)
    TxtNameE.text = IIf(IsNull(rs("namee").value), "", rs("namee").value)
 
 
 
 
 txtNoOFDigitUser(2).text = IIf(IsNull(rs("StreetName").value), "", rs("StreetName").value)
txtNoOFDigitUser(4).text = IIf(IsNull(rs("BuildingNumber").value), "", rs("BuildingNumber").value)
txtNoOFDigitUser(9).text = IIf(IsNull(rs("CitySubdivisionName").value), "", rs("CitySubdivisionName").value)
txtNoOFDigitUser(6).text = IIf(IsNull(rs("CityName").value), "", rs("CityName").value)
txtNoOFDigitUser(7).text = IIf(IsNull(rs("PostalZone").value), "", rs("PostalZone").value)
txtNoOFDigitUser(10).text = IIf(IsNull(rs("IdentificationCode").value), "SA", rs("IdentificationCode").value)
If txtNoOFDigitUser(10).text = "" Then txtNoOFDigitUser(10).text = "SA"
txtNoOFDigitUser(5).text = IIf(IsNull(rs("PlotIdentification").value), "", rs("PlotIdentification").value)
txtNoOFDigitUser(3).text = IIf(IsNull(rs("AdditionalStreetName").value), "", rs("AdditionalStreetName").value)
txtNoOFDigitUser(8).text = IIf(IsNull(rs("CountrySubentity").value), "", rs("CountrySubentity").value)

XPTxtComment(3).text = IIf(IsNull(rs("VATRegNo").value), "", rs("VATRegNo").value)
XPTxtComment(0).text = IIf(IsNull(rs("Company_Comment").value), "", rs("Company_Comment").value)


  If rs("IsBluee").value = True Then
        Me.Chkbarcode(212).value = vbChecked
    Else
        Me.Chkbarcode(212).value = Unchecked
    End If
    
    
        If rs("ApplyEinvoice").value = 1 Then
        Me.Chkbarcode(213).value = vbChecked
    Else
        Me.Chkbarcode(213).value = Unchecked
    End If
XPTxtComment(4) = XPTxtComment(3)
XPTxtComment(5) = XPTxtComment(0)
'Wael
'XPTxtComment(12) = txtNoOFDigitUser(10)

XPTxtComment(10) = txtNoOFDigitUser(6)

        

    XPTxtCompany = rs("Company_Arabic_Name").value & ""
    XPTxtCompanye = rs("Company_Name_Eng").value & ""



XPTxtComment(8) = IIf(IsNull(rs("Commonname").value), "", rs("Commonname").value)
XPTxtComment(7) = IIf(IsNull(rs("SerialNumber").value), "", rs("SerialNumber").value)
XPTxtComment(6) = IIf(IsNull(rs("OrganizationName").value), "", rs("OrganizationName").value)
 
XPTxtComment(11) = IIf(IsNull(rs("industrey").value), "", rs("industrey").value)
XPTxtComment(9) = IIf(IsNull(rs("CSR").value), "", rs("CSR").value)
XPTxtComment(13) = IIf(IsNull(rs("Privatekey").value), "", rs("Privatekey").value)
XPTxtComment(14) = IIf(IsNull(rs("PublickeycertPem").value), "", rs("PublickeycertPem").value)
XPTxtComment(15) = IIf(IsNull(rs("SecretKey").value), "", rs("SecretKey").value)


        Me.Invoicetype.ListIndex = IIf(IsNull(rs("Invoicetype").value), 0, rs("Invoicetype").value)
            Me.DefaultInvoicetype.ListIndex = IIf(IsNull(rs("DefaultInvoicetype").value), 0, rs("DefaultInvoicetype").value)
            
            Me.SendingMode.ListIndex = IIf(IsNull(rs("SendingMode").value), 0, rs("SendingMode").value)
            

    'StrSQL = " select * from TblBranchesData "
    
   StrSQL = " SELECT    TblBranchesData.Beauty, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.Tel, dbo.TblBranchesData.manger,"
   StrSQL = StrSQL & "                   TblStore.StoreName,TblStore.StoreNamee,TblStore.StoreID,"
   StrSQL = StrSQL & "                   dbo.TblBranchesData.Remarks, dbo.TblBranchesData.branch_Code, dbo.TblBranchesData.Account_Code, dbo.TblBranchesData.branchLogo,"
   StrSQL = StrSQL & "                   dbo.TblBranchesData.ShowlogoInReports , dbo.TblBranchesData.VATNO, dbo.TblBranchesData.RegionID, dbo.TblSection.Name, dbo.TblSection.NameE ,dbo.TblBranchesData.Users"
   StrSQL = StrSQL & " FROM         dbo.TblBranchesData LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblSection ON dbo.TblBranchesData.RegionID = dbo.TblSection.Id"
   
   StrSQL = StrSQL & "                   LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblStore ON dbo.TblBranchesData.StoreID = dbo.TblStore.StoreID"
   
   
   StrSQL = StrSQL & "  where ActivityTypeId=" & val(Me.txtid.text)
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
 
                .TextMatrix(i, .ColIndex("branch_id")) = IIf(IsNull(RsDev("branch_id").value), "", RsDev("branch_id").value)
            
                .TextMatrix(i, .ColIndex("branch_Code")) = IIf(IsNull(RsDev("branch_Code").value), "", RsDev("branch_Code").value)
                         
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDev("branch_name").value), "", RsDev("branch_name").value)
            
                .TextMatrix(i, .ColIndex("branch_namee")) = IIf(IsNull(RsDev("branch_namee").value), "", RsDev("branch_namee").value)
                .TextMatrix(i, .ColIndex("manger")) = IIf(IsNull(RsDev("manger").value), "", RsDev("manger").value)
                .TextMatrix(i, .ColIndex("Tel")) = IIf(IsNull(RsDev("Tel").value), "", RsDev("Tel").value)
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                .TextMatrix(i, .ColIndex("Users")) = IIf(IsNull(RsDev("Users").value), "", RsDev("Users").value)
                .TextMatrix(i, .ColIndex("VATNO")) = IIf(IsNull(RsDev("VATNO").value), "", RsDev("VATNO").value)
                .TextMatrix(i, .ColIndex("RegionID")) = IIf(IsNull(RsDev("RegionID").value), "", RsDev("RegionID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("RegionName")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
                .TextMatrix(i, .ColIndex("RegionName")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                End If
                .TextMatrix(i, .ColIndex("StoreId")) = IIf(IsNull(RsDev("StoreId").value), "", RsDev("StoreId").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsDev("StoreName").value), "", RsDev("StoreName").value)
                Else
                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsDev("StoreNamee").value), "", RsDev("StoreNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("Beauty")) = IIf(IsNull(RsDev("Beauty").value), 0, RsDev("Beauty").value)
                
                              
                             
              If Not (IsNull(RsDev("branchLogo").value)) Then
    .Row = i
    .Col = 26
        LoadPictureFromDB ImgPic, RsDev, "branchLogo"
        Me.Grid.CellPicture = ImgPic.Picture
       ImgPic.Visible = True
        Else
        ImgPic.Visible = False
    End If

                RsDev.MoveNext
            Next i
 
        End With

    End If
 
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub lbl_Click(Index As Integer)
FrmSection.show
End Sub

Private Sub Text1_Change()



If Text1.text = "Alex2025" Then

EltCont.Visible = True
'Text1.Visible = False
Else
  EltCont.Visible = False
End If

End Sub

Private Sub txtbranch_name_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtbranch_namee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
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
Ele(1).Enabled = True
    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False
Ele(1).Enabled = True
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

Private Sub TxtName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
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

