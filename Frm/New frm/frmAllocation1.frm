VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmAllocationToContract1 
   BackColor       =   &H00E2E9E9&
   Caption         =   "     ‘«‘… «À»«  «·«Ì—«œ   "
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
   HelpContextID   =   580
   Icon            =   "frmAllocation1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   18960
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
      Height          =   9525
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18960
      _cx             =   33443
      _cy             =   16801
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
         Height          =   7905
         Left            =   30
         TabIndex        =   1
         Top             =   735
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
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   45
            Width           =   18810
            _cx             =   33179
            _cy             =   13203
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
            Begin VB.CheckBox Check17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   195
               Left            =   16740
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   1680
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Frame Frame10 
               Caption         =   "»Ì«‰«  „Õ«”»Ì…"
               Height          =   780
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   6630
               Width           =   5415
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.CommandButton Command9 
                  Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                  Height          =   375
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   240
                  Width           =   1095
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
                  TabIndex        =   76
                  Top             =   360
                  Width           =   990
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "«Ã„«·Ì« "
               Height          =   795
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   6630
               Width           =   12600
               Begin VB.TextBox TxtPhone 
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
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   360
                  Width           =   945
               End
               Begin VB.TextBox TxtCommiValue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   8280
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtElectricity 
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
                  Left            =   2160
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   360
                  Width           =   945
               End
               Begin VB.TextBox TxtWater 
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
                  Left            =   4080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtInsuranceValue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   6240
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtTotalContract 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   10320
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   360
                  Width           =   1065
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
                  TabIndex        =   72
                  Top             =   480
                  Width           =   990
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
                  TabIndex        =   71
                  Top             =   360
                  Width           =   810
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
                  TabIndex        =   70
                  Top             =   480
                  Width           =   750
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
                  TabIndex        =   69
                  Top             =   480
                  Width           =   750
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
                  TabIndex        =   68
                  Top             =   360
                  Width           =   510
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·«ÌÃ«—"
                  Height          =   195
                  Index           =   6
                  Left            =   11505
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   480
                  Width           =   870
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Õœœ «· «—ÌŒ"
               Height          =   1020
               Left            =   12255
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   525
               Width           =   6120
               Begin MSComCtl2.DTPicker Fromdate 
                  Height          =   330
                  Left            =   3135
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   193724417
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker todate 
                  Height          =   330
                  Left            =   840
                  TabIndex        =   54
                  Top             =   240
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   32964609
                  CurrentDate     =   41640
               End
               Begin Dynamic_Byte.NourHijriCal Fromdate√H 
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   55
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   510
                  Index           =   9
                  Left            =   120
                  TabIndex        =   56
                  Top             =   480
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   900
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmAllocation1.frx":038A
                  DrawFocusRectangle=   0   'False
               End
               Begin Dynamic_Byte.NourHijriCal todateH 
                  Height          =   255
                  Left            =   840
                  TabIndex        =   57
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
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
                  TabIndex        =   59
                  Top             =   240
                  Width           =   945
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
                  TabIndex        =   58
                  Top             =   240
                  Width           =   540
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "«œŒ«· «·”‰Ê«  «·„«÷Ì…"
               Height          =   765
               Left            =   -5520
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1740
               Width           =   3885
               Begin VB.OptionButton OptActual 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·Ì"
                  Height          =   195
                  Index           =   1
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   240
                  Width           =   1305
               End
               Begin VB.OptionButton OptActual 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÌœÊÌ"
                  Height          =   195
                  Index           =   0
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   240
                  Width           =   1305
               End
            End
            Begin VB.OptionButton OptAlarms 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«› «·Õ”«»"
               Height          =   75
               Index           =   1
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   -4575
               Width           =   1890
            End
            Begin VB.OptionButton OptAlarms 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ–Ì— ›ﬁÿ"
               Height          =   75
               Index           =   0
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   -4575
               Width           =   1455
            End
            Begin VB.ComboBox OperatorsID 
               Height          =   315
               ItemData        =   "frmAllocation1.frx":0724
               Left            =   19200
               List            =   "frmAllocation1.frx":0734
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Text            =   " "
               Top             =   2280
               Width           =   1395
            End
            Begin VB.TextBox Percentage 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   19695
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Text            =   "0"
               Top             =   2280
               Width           =   1125
            End
            Begin VB.Frame Frame6 
               Caption         =   "Õœœ «·„Ê«“‰«  «·”«»ﬁ…"
               Height          =   1665
               Left            =   20055
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   495
               Width           =   8445
               Begin VSFlex8Ctl.VSFlexGrid GridOldEstimation 
                  Height          =   915
                  Left            =   120
                  TabIndex        =   44
                  Top             =   240
                  Width           =   8265
                  _cx             =   14579
                  _cy             =   1614
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
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmAllocation1.frx":0750
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
            Begin VB.Frame Frame5 
               Caption         =   "Õœœ ”‰Ê«  «·„ﬁ«—‰…"
               Height          =   1665
               Left            =   19515
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   495
               Width           =   5355
               Begin VSFlex8Ctl.VSFlexGrid GridIntervals1 
                  Height          =   915
                  Left            =   120
                  TabIndex        =   42
                  Top             =   240
                  Width           =   4545
                  _cx             =   8017
                  _cy             =   1614
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
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmAllocation1.frx":07EE
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
            Begin VB.Frame Frame1 
               Caption         =   "«· Ê“Ì⁄ ⁄·Ï «Õ”«»« "
               Height          =   1200
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   9570
               Width           =   17265
               Begin VB.TextBox TxtRemarks1 
                  Alignment       =   1  'Right Justify
                  Height          =   615
                  Left            =   2160
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   34
                  Top             =   120
                  Width           =   3615
               End
               Begin VB.TextBox TxtPercentage 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   240
                  Width           =   1215
               End
               Begin MSDataListLib.DataCombo DCAccountDist 
                  Height          =   315
                  Left            =   8760
                  TabIndex        =   35
                  Top             =   240
                  Width           =   3855
                  _ExtentX        =   6800
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   960
                  TabIndex        =   36
                  Top             =   240
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmAllocation1.frx":08D3
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   240
                  TabIndex        =   37
                  Top             =   240
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   688
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
                  ButtonImage     =   "frmAllocation1.frx":0C6D
                  DrawFocusRectangle=   0   'False
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
                  TabIndex        =   40
                  Top             =   240
                  Width           =   840
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
                  TabIndex        =   39
                  Top             =   240
                  Width           =   615
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
                  TabIndex        =   38
                  Top             =   240
                  Width           =   1080
               End
            End
            Begin VB.TextBox TxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   1005
               Left            =   4935
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   555
               Width           =   6345
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   19095
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   1185
               Width           =   2670
               Begin VB.OptionButton PercentagType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”» ÌœÊÌÂ"
                  Height          =   210
                  Index           =   1
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.OptionButton PercentagType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”» «·ÌÂ"
                  Height          =   210
                  Index           =   0
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.TextBox TxtTransID 
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
               Left            =   16410
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   105
               Width           =   1455
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
               Height          =   510
               Index           =   0
               Left            =   -4830
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   12495
               Width           =   2625
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
               Height          =   435
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   -495
               Visible         =   0   'False
               Width           =   2640
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1260
               Left            =   20670
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   495
               Width           =   2910
               Begin VB.OptionButton DistType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· Ê“Ì⁄ ⁄·Ï  «·›—Ê⁄"
                  Height          =   210
                  Index           =   2
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   720
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
                  TabIndex        =   23
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.OptionButton DistType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· Ê“Ì⁄ ⁄·Ï „—«ﬂ“  ﬂ·›…"
                  Height          =   210
                  Index           =   1
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   480
                  Width           =   2055
               End
            End
            Begin MSDataListLib.DataCombo DCAccountMaster 
               Height          =   315
               Left            =   22890
               TabIndex        =   77
               Top             =   555
               Width           =   6360
               _ExtentX        =   11218
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
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   5385
               Left            =   20730
               TabIndex        =   78
               Top             =   2595
               Width           =   18480
               _cx             =   32597
               _cy             =   9499
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
               Cols            =   28
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmAllocation1.frx":1207
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
               Height          =   345
               Left            =   12600
               TabIndex        =   79
               Top             =   105
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   609
               _Version        =   393216
               Format          =   232259585
               CurrentDate     =   41640
            End
            Begin MSDataListLib.DataCombo DcBranch 
               Height          =   315
               Left            =   4995
               TabIndex        =   80
               Top             =   0
               Width           =   6315
               _ExtentX        =   11139
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
            Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
               Height          =   4710
               Left            =   360
               TabIndex        =   81
               Top             =   1920
               Width           =   18390
               _cx             =   32438
               _cy             =   8308
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
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   12
               Cols            =   38
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmAllocation1.frx":163E
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
               Left            =   14115
               TabIndex        =   82
               Top             =   105
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰œ „Œ«·›… «· ﬁœÌ—Ï"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   16
               Left            =   8415
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   -4575
               Width           =   2370
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   210
               Left            =   11310
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   -2715
               Width           =   840
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‰”»…"
               Height          =   225
               Index           =   0
               Left            =   21030
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   2280
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «· ﬁœÌ— „ Ê”ÿ „«”»ﬁ"
               Height          =   300
               Index           =   15
               Left            =   19200
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   2280
               Width           =   2355
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   13
               Left            =   11235
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   0
               Width           =   720
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
               Height          =   420
               Index           =   37
               Left            =   2985
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   225
               Visible         =   0   'False
               Width           =   1650
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
               Height          =   1065
               Index           =   38
               Left            =   225
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   420
               Width           =   4290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   240
               Index           =   2
               Left            =   11010
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   555
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·› —… „‰ "
               Height          =   720
               Index           =   4
               Left            =   18420
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   555
               Width           =   1065
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «· Ê“Ì⁄"
               Height          =   495
               Index           =   3
               Left            =   18690
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   1530
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Õ—ﬂ…"
               Height          =   420
               Index           =   8
               Left            =   15345
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   105
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ—ﬂ…"
               Height          =   420
               Index           =   7
               Left            =   17115
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   105
               Width           =   1560
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
               Height          =   375
               Left            =   16995
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   1380
               Width           =   1170
            End
            Begin VB.Shape Shape1 
               BorderWidth     =   2
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   1350
               Left            =   195
               Top             =   195
               Width           =   4545
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   810
         Left            =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   8685
         Width           =   18900
         _cx             =   33338
         _cy             =   1429
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
         Begin VB.CommandButton Command2 
            Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
            Height          =   345
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   0
            Width           =   1665
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   11880
            TabIndex        =   4
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
            ButtonImage     =   "frmAllocation1.frx":1C2D
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   5
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
            ButtonImage     =   "frmAllocation1.frx":1FC7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   6
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
            ButtonImage     =   "frmAllocation1.frx":2361
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   11100
            TabIndex        =   9
            Top             =   510
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
            Left            =   10200
            TabIndex        =   10
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
            Left            =   9360
            TabIndex        =   11
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
            Left            =   8355
            TabIndex        =   12
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
            Left            =   7320
            TabIndex        =   13
            Top             =   510
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
            Left            =   3720
            TabIndex        =   14
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
            Left            =   6390
            TabIndex        =   15
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
            Left            =   16560
            TabIndex        =   16
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
            MICON           =   "frmAllocation1.frx":26FB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   495
            Left            =   5400
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            ButtonImage     =   "frmAllocation1.frx":2717
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì"
            Height          =   255
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
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
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   225
            Width           =   1020
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
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
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1155
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   765
         Index           =   5
         Left            =   0
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   0
         Width           =   18945
         _cx             =   33417
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
         Picture         =   "frmAllocation1.frx":2AB1
         Caption         =   "     ‘«‘… «À»«  «·«Ì—«œ   "
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
         Begin VB.TextBox TxtRowNumber 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4635
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Text            =   "Text4"
            Top             =   360
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7125
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Text            =   "Text1"
            Top             =   240
            Visible         =   0   'False
            Width           =   345
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1605
            TabIndex        =   99
            Top             =   90
            Width           =   450
            _ExtentX        =   794
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
            ButtonImage     =   "frmAllocation1.frx":378B
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
            Left            =   600
            TabIndex        =   100
            Top             =   90
            Width           =   450
            _ExtentX        =   794
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
            ButtonImage     =   "frmAllocation1.frx":3B25
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
            Left            =   2085
            TabIndex        =   101
            Top             =   90
            Width           =   465
            _ExtentX        =   820
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
            ButtonImage     =   "frmAllocation1.frx":3EBF
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
            Left            =   1080
            TabIndex        =   102
            Top             =   90
            Width           =   480
            _ExtentX        =   847
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
            ButtonImage     =   "frmAllocation1.frx":4259
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin MSComDlg.CommonDialog cd 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   8160
            Picture         =   "frmAllocation1.frx":45F3
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmAllocation1.frx":825B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmAllocationToContract1"
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
Dim Account_Code_dynamic80 As String
Dim Account_Code_dynamic81 As String
Dim Account_Code_dynamic82 As String
Dim Account_Code_dynamic83 As String
Dim Account_Code_dynamic84 As String
Dim Account_Code_dynamic85 As String
Dim Account_Code_dynamic86 As String
Dim hijriorJerojian As Integer
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
    If TxtTransID.text <> "" Then
     Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtTransID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                 
         StrSQL = "Delete From Notes Where NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
                   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
    ' Cn.Execute " update  TblContractInstallments set  allocations=0 where id in( " & " select installid from tblContractInsAllocations2 where transid=" & TxtTransID & ")"
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

Public Sub YearMonth()

End Sub

Private Sub ChkDetails_Click()
     
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub









 

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

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

    If TxtModFlg.text = "N" Then
        rs.AddNew
           Me.TxtTransID.text = CStr(new_id("tblContractInsAllocations1", "transID", "", True))
    ElseIf Me.TxtModFlg.text = "E" Then
          Cn.Execute "delete tblContractInsAllocationsDetails2  where transID=" & val(Me.TxtTransID.text)
      StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords


    End If
    
  rs("transID").value = TxtTransID.text
    rs("recordDate").value = XPDtbTrans.value
    rs("RecorddateH").value = RecorddateH.value
     rs("Fromdate").value = FromDate.value
           rs("todate").value = ToDate.value
       rs("Fromdateh").value = ToHijriDate(FromDate.value)
           rs("todateh").value = ToHijriDate(ToDate.value)
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", Null, val(Me.dcBranch.BoundText))
  
      
    rs("Remarks").value = IIf(Me.txtRemarks.text = "", "", Me.txtRemarks.text)
   

    rs.update
    
 
    Set RsDetails1 = New ADODB.Recordset
 
           StrSQL = "SELECT  *  from dbo.tblContractInsAllocationsDetails2  Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Dim i As Integer
      
    With Me.GridInstallments
'Selected
        For i = 1 To .rows - 1
   
        If val(.TextMatrix(i, .ColIndex("value"))) <> 0 And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
       RsDetails1.AddNew
      RsDetails1("transID").value = Me.TxtTransID.text
      
   RsDetails1("NoteSerial").value = (.TextMatrix(i, .ColIndex("NoteSerial1")))
   
    RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
     RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
           RsDetails1("Installdate").value = .TextMatrix(i, .ColIndex("Due_Date"))
           RsDetails1("InstalldateH").value = .TextMatrix(i, .ColIndex("Due_DateH"))
          RsDetails1("installValue").value = val(.TextMatrix(i, .ColIndex("value")))
    RsDetails1("RentValue").value = val(.TextMatrix(i, .ColIndex("RentValue")))
          RsDetails1("Commissions").value = val(.TextMatrix(i, .ColIndex("Commissions")))
          RsDetails1("Insurance").value = val(.TextMatrix(i, .ColIndex("Insurance")))
          RsDetails1("Water").value = val(.TextMatrix(i, .ColIndex("Water")))
          RsDetails1("Electric").value = val(.TextMatrix(i, .ColIndex("Electric")))
        RsDetails1("TelandNet").value = val(.TextMatrix(i, .ColIndex("TelandNet")))
        RsDetails1("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
                                    RsDetails1("allocations").value = 1
                         '   Cn.Execute " update  TblContractInstallments set  allocations=1 where id=" & val(.TextMatrix(i, .ColIndex("Installid")))

                        '     If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
'                            RsDetails1("Select").value = 1
                   '         RsDetails1("allocations").value = 1
                         '   Cn.Execute " update  TblContractInstallments set  allocations=1 where id=" & val(.TextMatrix(i, .ColIndex("Installid")))
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
    
 
 
 
 createVoucher
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
Case 9
If Me.TxtModFlg.text = "E" Then
'x = MsgBox("”Ì „ «·€«¡ «·«Ì—«œ «·Õ«·Ì", vbCritical + vbOKCancel)
'            If x = vbOK Then
'                 Cn.Execute " update  TblContractInstallments set  allocations=0 where id in( " & " select installid from tblContractInsAllocations2 where transid=" & TxtTransID & ")"
'        Else
'
'        Exit Sub
'            End If
End If



FillGrid

        Case 0
 
            TxtModFlg.text = "N"
            clear_all Me
        OperatorsID.ListIndex = 0
       OptAlarms(0).value = True
       OptActual(1).value = True
            Me.XPDtbTrans.value = Date
            RecorddateH.value = ToHijriDate(Date)
            
            Me.FromDate.value = Date
            Me.ToDate.value = Date
            Check17.value = vbChecked
            Me.Fromdate√H.value = ToHijriDate(Date)
todateH.value = ToHijriDate(Date)

   Me.dcBranch.BoundText = Current_branch
       
            'XPDtbTrans.SetFocus
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 1
            GridInstallments.Enabled = True
      
 
 

        Case 1
                    If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            TxtModFlg.text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True

        Case 2
                           If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
       If val(Me.dcBranch.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
            Else
                MsgBox "Select Branch Firstly    ", vbCritical
            End If

            dcBranch.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
If CheckAcconts = False Then Exit Sub
If TxtNoteSerial.text = "" Then     'ÃœÌœ ›ﬁÿ
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

    If val(Me.TxtRowNumber.text) <> 0 Then
        LngRow = val(Me.TxtRowNumber.text)
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
    
        .TextMatrix(LngRow, .ColIndex("AName")) = Me.DCAccountDist.text
    
        .TextMatrix(LngRow, .ColIndex("Percentage")) = val(Me.TxtPercentage.text)
    
        .TextMatrix(LngRow, .ColIndex("Remarks")) = (Me.TxtRemarks1.text)
     
        .AutoSize 0, .Cols - 1, False
    End With

    Me.DCAccountDist.BoundText = ""
    Me.TxtPercentage.text = ""
    Me.TxtRemarks1.text = ""
  
    ReLineGrid
 
End Sub

Private Sub Undo()
   ' On Error GoTo ErrTrap

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

 

Function SHow_grig_col()
 
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
     cd.FileName = "File1"
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
        ShowGL_cc Me.TxtNoteSerial.text, , 200, , val(Me.TXTNoteID.text)
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
            Account_Code_dynamic86 = get_account_code_branch(86, my_branch)
            
            If Account_Code_dynamic86 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·«Ì—œ««   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
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
                 
              
              
                    If (val(TxtWater)) > 0 Then
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
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   Œœ„«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
              
                
End If



   CheckAcconts = True
   Exit Function
ErrTrap:
      CheckAcconts = False
End Function


Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = "«À»«  «Ì—«œ«   ⁄‰ «·› —… „‰  " '& Fromdate√H.value & "  Õ Ï  " & TodateH.value & Chr(13)
des = des & " „‰ " & FromDate.value & "  Õ Ï  " & ToDate.value & CHR(13)
des = des & " ··›—⁄ " & dcBranch.text & "     " & txtRemarks

Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "tblContractInsAllocations1"
Filedname = "transID"
ContNo = TxtTransID
Notevalue = 0

'If SystemOptions.WorkWithFirstInstallOnly = False Then
Notevalue = val(TxtTotalContract) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone)
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
CreateNotes NoteID, (XPDtbTrans.value), val(dcBranch.BoundText), 61, Notevalue, NoteSerial, TxtTransID, tablename, Filedname, ContNo, des, RecorddateH.value
 TXTNoteID.text = NoteID
TxtNoteSerial.text = NoteSerial
Else
sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
sql = sql & ",NoteSerial1='" & Me.TxtTransID & "',remark='" & des & "'"
  sql = sql & " where NoteID=" & val(TXTNoteID.text)
   Cn.Execute sql
End If

CREATE_VOUCHER_GE val(TXTNoteID.text), val(dcBranch.BoundText), user_id, XPDtbTrans.value
rs.Resync adAffectCurrent
 

End Function



Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Double
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Long
 
 Dim OtherInformation As New ClsGLOther
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        
 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    
    
   
    my_branch = BranchID
 
            '
    StrTempDes = "«À»«  «Ì—«œ ⁄‰ «·› —… „‰  " '& Fromdate√H.value & "  Õ Ï  " & TodateH.value & Chr(13)
    StrTempDes = StrTempDes & " „‰ " & FromDate.value & "  Õ Ï  " & ToDate.value & CHR(13)
    StrTempDes = StrTempDes & " ··›—⁄ " & dcBranch.text & "     " & txtRemarks.text
    
                 
                 
 If val(TxtTotalContract.text) > 0 Then
        
        
        Notevalue = val(TxtTotalContract.text)
        
        StrTempAccountCode = Account_Code_dynamic80
        
   
                         LngDevNO = LngDevNO + 1
                     If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «· ⁄«ﬁœ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                         GoTo ErrTrap
                     End If

  End If
  
  
 If (val(TxtCommiValue.text)) > 0 Then
       StrTempAccountCode = Account_Code_dynamic81
       
        
             Notevalue = (val(TxtCommiValue.text))
  

   
   
   
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "      ⁄„Ê·«  Ê—”Ê„ «œ«—Ì… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
  
  
'   If val(TxtInsuranceValue.text) > 0 Then
'       StrTempAccountCode = Account_Code_dynamic82
'
'
'      Notevalue = val(TxtInsuranceValue.text)
'
'
'
'
'
'   LngDevNO = LngDevNO + 1
'            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "     √„Ì‰ „” —œ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'                GoTo ErrTrap
'            End If
'  End If
  
  
     If val(TxtWater.text) > 0 Then
       StrTempAccountCode = Account_Code_dynamic83
'
    
    Notevalue = val(TxtWater.text)
  
   
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "    „Ì«Â ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
  
  
       If val(TxtElectricity.text) > 0 Then
       StrTempAccountCode = Account_Code_dynamic84
     '  Notevalue = val(TxtElectricity.text)
   
             
    Notevalue = val(TxtElectricity.text)
  
     
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "      ﬂÂ—»«¡ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
  
  
       If (val(TxtPhone.text)) > 0 Then
       StrTempAccountCode = Account_Code_dynamic85
'       Notevalue = (val(TxtPhone.text) + val(TxtEnternet.text))
   
            
    Notevalue = val(TxtPhone.text)
     
     
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "    Œœ„«    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
    
    
        
     
    With Me.GridInstallments

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("Value")) <> "" Then
            
            
            
                                                   
                    '     OtherInformation.Unitss = IIf(.TextMatrix(i, .ColIndex("Unitss")) = "", "", .TextMatrix(i, .ColIndex("Unitss")))
                    '  OtherInformation.StrUnit = IIf(.TextMatrix(i, .ColIndex("StrUnit")) = "", "", .TextMatrix(i, .ColIndex("StrUnit")))
                      OtherInformation.uintid = val(.TextMatrix(i, .ColIndex("UnitNo")))
                    '  OtherInformation.mType = IIf(.TextMatrix(i, .ColIndex("type")) = "", 0, val(.TextMatrix(i, .ColIndex("type"))))
                      OtherInformation.iqarid = IIf(.TextMatrix(i, .ColIndex("Iqar")) = "", 0, val(.TextMatrix(i, .ColIndex("Iqar"))))

                                                      
                                  
'                IntCounter = IntCounter + 1
'                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                
                  If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                        ' Me.TxtTotalContract.text = val(Me.TxtTotalContract.text) + .TextMatrix(i, .ColIndex("RentValue"))
                         Notevalue = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("Water"))) + val(.TextMatrix(i, .ColIndex("Electric")))

            LngDevNO = LngDevNO + 1
 'Notevalue = val(TxtTotalContract.text) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone)
         StrTempAccountCode = Account_Code_dynamic86
                                      If Notevalue > 0 Then
                                      '«·ÿ—› «·œ«∆‰
                                   
                                             If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID, , , , , , , , , , val(.TextMatrix(i, .ColIndex("Iqar"))), val(.TextMatrix(i, .ColIndex("UnitType"))), val(.TextMatrix(i, .ColIndex("UnitNo"))), , , , , , , , , , , , , , , OtherInformation) = False Then
                                     GoTo ErrTrap
                                 End If
                                 
                                     End If
                                     
                                                      End If
            End If
        Next
    End With
                                     
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

 
Dcombos.GetBranches dcBranch

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

StrSQL = "select * From tblContractInsAllocations1  "
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
        .TextMatrix(0, .ColIndex("Percentage")) = "Percentage"
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


    sql = "Select * from tblContractInsAllocations2 "
 
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

Public Sub FillGrid()

  '  On Error GoTo ErrTrap

    Dim i As Double
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
 
Dim notpayed As Double
notpayed = 0
 
My_SQL = " SELECT     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.tblContractInsAllocationsDetails1.*"
My_SQL = My_SQL + " FROM         dbo.tblContractInsAllocationsDetails1 INNER JOIN"
My_SQL = My_SQL + " dbo.TblCustemers ON dbo.tblContractInsAllocationsDetails1.CusID = dbo.TblCustemers.CusID"
My_SQL = My_SQL + " WHERE     (dbo.tblContractInsAllocationsDetails1.allocations = 0 OR"
My_SQL = My_SQL + " dbo.tblContractInsAllocationsDetails1.allocations IS NULL) " 'WHERE     (dbo.tblContractInsAllocations1.allocations = 0 or dbo.tblContractInsAllocations1.allocations IS NULL)  AND (dbo.tblContractInsAllocations1.Status = 0 OR dbo.tblContractInsAllocations1.Status IS NULL)"


        My_SQL = My_SQL + " and (Installdate >=" & SQLDate(Me.FromDate, True) & ""
        My_SQL = My_SQL + " and Installdate <=" & SQLDate(ToDate, True) & " )"
         My_SQL = My_SQL + " and  (commission=0 or commission is null)"
'        My_SQL = My_SQL + " and  tblContractInsAllocationsDetails1.ContractFlag  In (Select tc.ContNo from tblcontract as tc where tc.NoteSerial1 Not In (Select  VV.ContractNo from TblFiterWaiver VV where FilterDate <=" & SQLDate(ToDate, True) & ") )"
        
        
        My_SQL = My_SQL & " and     dbo.tblContractInsAllocationsDetails1.NoteSerial Not In (Select  VV.ContractNo from TblFiterWaiver VV where FilterDate <=" & SQLDate(ToDate, True) & ")"
'
   
       ' My_SQL = My_SQL + " and (Installdate >=" & SQLDate(Me.FromDate, True) & ""
     

 
       ' My_SQL = My_SQL + " and Installdate <=" & SQLDate(ToDate, True) & " )"
        
        
        
    
        My_SQL = My_SQL + "   order by Installdate "
 My_SQL = " SELECT DISTINCT "
My_SQL = My_SQL & "     C.CusName, "
My_SQL = My_SQL & "     C.CusNamee, "
My_SQL = My_SQL & "     D.Installid, "
My_SQL = My_SQL & "     D.InstallNo, "
My_SQL = My_SQL & "    TblContract.NoteSerial1 NoteSerial, "
My_SQL = My_SQL & "     D.Installdateh, "
My_SQL = My_SQL & "     D.Installdate, "
My_SQL = My_SQL & "     D.installValue, "
My_SQL = My_SQL & "     D.CusID, "
My_SQL = My_SQL & "     D.RentValue, "
My_SQL = My_SQL & "     D.commission AS Commissions, "
My_SQL = My_SQL & "     D.Insurance, "
My_SQL = My_SQL & "     D.Water, "
My_SQL = My_SQL & "     D.Electric, "
My_SQL = My_SQL & "     D.TelandNet, "
My_SQL = My_SQL & "     D.allocations, "
My_SQL = My_SQL & "     D.Countsofall, "
My_SQL = My_SQL & "     D.Doneofall ,TblContractInstallments.ContNo,TblContract.NoteSerial1,"
My_SQL = My_SQL & "     TblContract.UnitNo ,TblContract.unittype,TblContract.Iqar"

My_SQL = My_SQL & " FROM dbo.tblContractInsAllocationsDetails1 AS D "

My_SQL = My_SQL & " inner join TblContractInstallments on D.Installid = TblContractInstallments.id"
My_SQL = My_SQL & " inner join TblContract on TblContract.ContNo = TblContractInstallments.ContNo"
My_SQL = My_SQL & " INNER JOIN dbo.TblCustemers AS C ON D.CusID = C.CusID "

My_SQL = My_SQL & " WHERE (D.allocations = 0 OR D.allocations IS NULL) "

'  «—ÌŒ „‰ ñ ≈·Ï
My_SQL = My_SQL & " AND (D.Installdate >= " & SQLDate(Me.FromDate.value, True)
My_SQL = My_SQL & " AND D.Installdate <= " & SQLDate(ToDate.value, True) & ") "

' ⁄„Ê·… ’›—
My_SQL = My_SQL & " AND (D.commission = 0 OR D.commission IS NULL) "
'My_SQL = My_SQL + " and  d.ContractFlag  In (Select tc.ContNo from tblcontract as tc where tc.NoteSerial1 Not In (Select  VV.ContractNo from TblFiterWaiver VV where FilterDate <=" & SQLDate(ToDate, True) & ") )"
' «” »⁄«œ «·⁄ﬁÊœ «·„›· —…
My_SQL = My_SQL & " AND TblContract.NoteSerial1 NOT IN ( "
My_SQL = My_SQL & "       SELECT VV.ContractNo "
My_SQL = My_SQL & "       FROM TblFiterWaiver VV "
My_SQL = My_SQL & "       WHERE VV.FilterDate <= " & SQLDate(ToDate.value, True) & " "
My_SQL = My_SQL & " ) "

' «· — Ì»
My_SQL = My_SQL & " ORDER BY D.Installdate "


    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
   
   Dim noOfRemaindays1 As Integer
      Dim noOfRemaindays2 As Integer

      With Me.GridInstallments
       .rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .rows - 1
            
           '**************************************************************************
                   If hijriorJerojian = 0 Then 'jorjian
                      
                         VBA.Calendar = vbCalHijri
                     noOfRemaindays1 = DateDiff("D", Fromdate√H.value, rs.Fields("Installdateh").value)
                      noOfRemaindays2 = DateDiff("D", rs.Fields("Installdateh").value, todateH.value)
                      
                   Else
                       VBA.Calendar = vbCalGreg
                       
                   End If
                   
                   If noOfRemaindays1 > 0 And noOfRemaindays2 > 0 Then
              Else
                 'GoTo ll
                   End If
                       VBA.Calendar = vbCalGreg
          '**************************************************************************
            
            
              .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(rs.Fields("Installid").value), 0, rs.Fields("Installid").value))
               .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("InstallNo").value), 0, rs.Fields("InstallNo").value))
 .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value))
 
 
 
 .TextMatrix(i, .ColIndex("UnitNo")) = (IIf(IsNull(rs.Fields("UnitNo").value), "", rs.Fields("UnitNo").value))
 .TextMatrix(i, .ColIndex("unittype")) = (IIf(IsNull(rs.Fields("unittype").value), "", rs.Fields("unittype").value))
 .TextMatrix(i, .ColIndex("Iqar")) = (IIf(IsNull(rs.Fields("Iqar").value), "", rs.Fields("Iqar").value))
                
 .TextMatrix(i, .ColIndex("Due_DateH")) = (IIf(IsNull(rs.Fields("Installdateh").value), ToHijriDate(Date), rs.Fields("Installdateh").value))
  .TextMatrix(i, .ColIndex("Due_Date")) = IIf(IsNull(rs.Fields("Installdate").value), Date, rs.Fields("Installdate").value)
        
    .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("installValue").value), 0, rs.Fields("installValue").value))
     .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value))
   
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
   Else
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
   End If
 
   .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
 
    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
If SystemOptions.amlaketbatrentOnly = False Then
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = 0 ' (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("TelandNet").value), 0, rs.Fields("TelandNet").value))
    .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("Insurance"))) + val(.TextMatrix(i, .ColIndex("Electric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("Water"))) + val(.TextMatrix(i, .ColIndex("RentValue")))
 Else
 
    .TextMatrix(i, .ColIndex("Commissions")) = 0
    .TextMatrix(i, .ColIndex("Insurance")) = 0 ' (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = 0
    .TextMatrix(i, .ColIndex("Electric")) = 0
    .TextMatrix(i, .ColIndex("TelandNet")) = 0
 .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue")))
 End If
 
    
       .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(rs.Fields("allocations").value), 0, rs.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(rs.Fields("Countsofall").value), 0, rs.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))


        rs.MoveNext
ll:
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

Private Sub FromDate_Change()
If Me.TxtModFlg.text <> "R" Then
     hijriorJerojian = 1
         Me.Fromdate√H.value = ToHijriDate(FromDate.value)
       
End If
End Sub

Private Sub Fromdate√H_LostFocus()
     If Me.TxtModFlg.text <> "R" Then
             
            FromDate.value = ToGregorianDate(Fromdate√H.value)
               hijriorJerojian = 0
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
    Dim IntCounter As Double
    IntCounter = 0
    Dim i As Double
 
    Dim Percenrage As Double
 
 
    IntCounter = 0
  Me.TxtTotalContract.text = 0
  Me.TxtCommiValue.text = 0
    Me.TxtInsuranceValue.text = 0
      Me.TxtWater.text = 0
      Me.TxtElectricity.text = 0
        Me.TxtPhone.text = 0
     
    With Me.GridInstallments

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("Value")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                
                
                
                     If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
  Me.TxtTotalContract.text = val(Me.TxtTotalContract.text) + .TextMatrix(i, .ColIndex("RentValue"))
  Me.TxtCommiValue.text = val(Me.TxtCommiValue.text) + .TextMatrix(i, .ColIndex("Commissions"))
  Me.TxtInsuranceValue.text = val(Me.TxtInsuranceValue.text) + .TextMatrix(i, .ColIndex("Insurance"))
  Me.TxtWater.text = val(Me.TxtWater.text) + .TextMatrix(i, .ColIndex("Water"))
  Me.TxtElectricity.text = val(Me.TxtElectricity.text) + .TextMatrix(i, .ColIndex("Electric"))
  Me.TxtPhone.text = val(Me.TxtPhone.text) + .TextMatrix(i, .ColIndex("TelandNet"))
  
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
 
    Me.TxtTransID.text = IIf(IsNull(rs("transID").value), "", rs("transID").value)
 
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
  RecorddateH.value = IIf(IsNull(rs("recordDateH").value), ToHijriDate(Date), rs("recordDateH").value)
  dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    FromDate.value = IIf(IsNull(rs("Fromdate").value), Date, rs("Fromdate").value)
 Me.Fromdate√H.value = IIf(IsNull(rs("FromDateh").value), ToHijriDate(Date), rs("FromDateh").value)
    
        ToDate.value = IIf(IsNull(rs("todate").value), Date, rs("todate").value)
  todateH.value = IIf(IsNull(rs("todateH").value), ToHijriDate(Date), rs("todateH").value)
    
    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
   Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    txtRemarks.text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
 
 
 

    StrSQL = "   SELECT     dbo.tblContractInsAllocations1.transID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.tblContractInsAllocationsDetails2.*, "
StrSQL = StrSQL & "   dbo.tblContractInsAllocationsDetails2.id, dbo.tblContractInsAllocationsDetails2.transID, dbo.tblContractInsAllocationsDetails2.CusID,"
StrSQL = StrSQL & "  dbo.tblContractInsAllocationsDetails2.InstallNo, dbo.tblContractInsAllocationsDetails2.Installdate, dbo.tblContractInsAllocationsDetails2.InstalldateH,"
StrSQL = StrSQL & "   dbo.tblContractInsAllocationsDetails2.installValue, dbo.tblContractInsAllocationsDetails2.RentValue, dbo.tblContractInsAllocationsDetails2.Commissions,"
StrSQL = StrSQL & "   dbo.tblContractInsAllocationsDetails2.Insurance, dbo.tblContractInsAllocationsDetails2.Water, dbo.tblContractInsAllocationsDetails2.Electric,"
StrSQL = StrSQL & "   dbo.tblContractInsAllocationsDetails2.TelandNet, dbo.tblContractInsAllocationsDetails2.allocations, dbo.tblContractInsAllocationsDetails2.Countsofall,"
StrSQL = StrSQL & "     dbo.tblContractInsAllocationsDetails2.Doneofall, dbo.tblContractInsAllocationsDetails2.Installid, dbo.tblContractInsAllocationsDetails2.hijri,"
StrSQL = StrSQL & "     dbo.tblContractInsAllocationsDetails2.NoteSerial,"
StrSQL = StrSQL & "     TblContract.UnitNo ,TblContract.unittype,TblContract.Iqar"
StrSQL = StrSQL & "  FROM         dbo.tblContractInsAllocations1 INNER JOIN"
 StrSQL = StrSQL & "   dbo.tblContractInsAllocationsDetails2 ON dbo.tblContractInsAllocations1.transID = dbo.tblContractInsAllocationsDetails2.transID INNER JOIN"
StrSQL = StrSQL & "     dbo.TblCustemers ON dbo.tblContractInsAllocationsDetails2.CusID = dbo.TblCustemers.CusID"

StrSQL = StrSQL & " inner join TblContractInstallments on tblContractInsAllocationsDetails2.Installid = TblContractInstallments.id"
StrSQL = StrSQL & " inner join TblContract on TblContract.ContNo = TblContractInstallments.ContNo"

StrSQL = StrSQL & "   Where (dbo.tblContractInsAllocations1.TransID = " & val(Me.TxtTransID.text) & ")"


 

'StrSQL = StrSQL & "  WHERE     (dbo.tblContractInsAllocations2.transID = " & val(Me.TxtTransID.text) & ") "
    'StrSQL = StrSQL & "  where transID=" & val(Me.TxtTransID.text)
  
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.GridInstallments
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
           .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(RsDev.Fields("Installid").value), 0, RsDev.Fields("Installid").value))
               .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(RsDev.Fields("InstallNo").value), 0, RsDev.Fields("InstallNo").value))
 .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(RsDev.Fields("NoteSerial").value), "", RsDev.Fields("NoteSerial").value))
                
 .TextMatrix(i, .ColIndex("Due_DateH")) = (IIf(IsNull(RsDev.Fields("Installdateh").value), ToHijriDate(Date), RsDev.Fields("Installdateh").value))
  .TextMatrix(i, .ColIndex("Due_Date")) = IIf(IsNull(RsDev.Fields("Installdate").value), Date, RsDev.Fields("Installdate").value)
        
        
    .TextMatrix(i, .ColIndex("UnitNo")) = (IIf(IsNull(RsDev.Fields("UnitNo").value), "", RsDev.Fields("UnitNo").value))
 .TextMatrix(i, .ColIndex("unittype")) = (IIf(IsNull(RsDev.Fields("unittype").value), "", RsDev.Fields("unittype").value))
 .TextMatrix(i, .ColIndex("Iqar")) = (IIf(IsNull(RsDev.Fields("Iqar").value), "", RsDev.Fields("Iqar").value))
 
 
 
    .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(RsDev.Fields("installValue").value), 0, RsDev.Fields("installValue").value))
     .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(RsDev.Fields("CusID").value), "", RsDev.Fields("CusID").value))
   
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(RsDev.Fields("CusName").value), "", RsDev.Fields("CusName").value))
   Else
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(RsDev.Fields("CusNamee").value), "", RsDev.Fields("CusNamee").value))
   End If
 
   .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
 
    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(RsDev.Fields("RentValue").value), 0, RsDev.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(RsDev.Fields("Commissions").value), 0, RsDev.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(RsDev.Fields("Insurance").value), 0, RsDev.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(RsDev.Fields("Water").value), 0, RsDev.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(RsDev.Fields("Electric").value), 0, RsDev.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(RsDev.Fields("TelandNet").value), 0, RsDev.Fields("TelandNet").value))
 
    
       .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(RsDev.Fields("allocations").value), 0, RsDev.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(RsDev.Fields("Countsofall").value), 0, RsDev.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(RsDev.Fields("Doneofall").value), 0, RsDev.Fields("Doneofall").value))
.TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("Insurance"))) + val(.TextMatrix(i, .ColIndex("Electric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("Water"))) + val(.TextMatrix(i, .ColIndex("RentValue")))
             
             
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
 
         If .ColKey(Col) <> "Select" Then
   
        
        Cancel = True
        
        End If
 
        
    End With
End Sub

Private Sub GridIntervals_Click()

  

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption


End Sub

Private Sub ISButton2_Click()
print_report
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
MySQL = " SELECT     dbo.tblContractInsAllocationsDetails2.Installid, dbo.tblContractInsAllocationsDetails2.InstallNo, dbo.tblContractInsAllocationsDetails2.Installdate, "
MySQL = MySQL & "                      dbo.tblContractInsAllocationsDetails2.InstalldateH, dbo.tblContractInsAllocationsDetails2.installValue, dbo.TblContractInstallments.installValue AS total,"
MySQL = MySQL & "                      dbo.tblContractInsAllocationsDetails2.Doneofall, dbo.TblContract.Iqar, dbo.TblContract.UnitNo, dbo.TblContract.NoteSerial1 as NoteSerial, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname,"
MySQL = MySQL & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.tblContractInsAllocationsDetails2.transID, dbo.tblContractInsAllocationsDetails2.Insurance,"
MySQL = MySQL & "                      dbo.tblContractInsAllocationsDetails2.Commissions, dbo.tblContractInsAllocationsDetails2.Water, dbo.tblContractInsAllocationsDetails2.Electric,"
MySQL = MySQL & "                      dbo.tblContractInsAllocationsDetails2.TelandNet , dbo.tblContractInsAllocationsDetails2.RentValue"
MySQL = MySQL & " FROM         dbo.TblContract INNER JOIN     dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN     dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN      dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid RIGHT OUTER JOIN     dbo.tblContractInsAllocationsDetails2 ON dbo.TblContractInstallments.id = dbo.tblContractInsAllocationsDetails2.Installid  "

Dim X As Integer
X = MsgBox("ÿ»«⁄Â «·”‰œ «·„Õœœ ›ﬁÿ", vbInformation + vbYesNo)

If X = vbNo Then
MySQL = MySQL & "  Where dbo.tblContractInsAllocationsDetails2.installid"
 MySQL = MySQL & "   in ("


MySQL = MySQL & "   select installid from  dbo.tblContractInsAllocationsDetails2"
MySQL = MySQL & "   Where (dbo.tblContractInsAllocationsDetails2.TransID = " & val(Me.TxtTransID.text) & " ) )"
Else
MySQL = MySQL & "  Where dbo.tblContractInsAllocationsDetails2.transid= " & val(Me.TxtTransID.text)
  
End If
MySQL = MySQL & " order by  tblContractInsAllocationsDetails2.id"


'MySQL = MySQL & "  Where (dbo.tblContractInsAllocationsDetails2.TransID = " & val(Me.TxtTransID.text) & ")"

 

If X = vbNo Then
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "aqarRevenue.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "aqarRevenue.rpt"
        End If

Else
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "aqarRevenue1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "aqarRevenue1.rpt"
        End If

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
        Msg = "?CE??I E?C?CE ?????"
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " EIC?E ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(lbl(23).Caption), "0.00"), 0, True, ".")
'        xReport.ParameterFields(11).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(11).AddCurrentValue WriteNo(Format(val(lbl(23).Caption), "0.00"), 0, True, ".")
'  xReport.ParameterFields(12).AddCurrentValue WriteNo(Format(val(lbl(31).Caption), "0.00"), 0, True, ".")
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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

Private Sub PercentagType_Click(Index As Integer)

    Select Case Index
        
        Case 0
            TxtPercentage.locked = True
            TxtPercentage.text = ""

        Case 1
            TxtPercentage.locked = False
            TxtPercentage.text = ""

    End Select

End Sub

Private Sub RecordDateH_LostFocus()
     If Me.TxtModFlg.text <> "R" Then
             
            XPDtbTrans.value = ToGregorianDate(RecorddateH.value)
               
        End If
End Sub

Private Sub ToDate_Change()
If Me.TxtModFlg.text <> "R" Then
     hijriorJerojian = 1
         todateH.value = ToHijriDate(ToDate.value)
       
End If
End Sub

Private Sub ToDateH_LostFocus()
     If Me.TxtModFlg.text <> "R" Then
             
            ToDate.value = ToGregorianDate(todateH.value)
               hijriorJerojian = 0
        End If
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

 
Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.text <> "R" Then
     
         RecorddateH.value = ToHijriDate(XPDtbTrans.value)
       
End If
End Sub
