VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBuyBillInvestment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   Icon            =   "FrmBuyBillInvestment.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   14235
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   53
      Top             =   720
      Width           =   14295
      _cx             =   25215
      _cy             =   13361
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   14871017
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "«·—∆Ì”Ì…"
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
      Begin VB.Frame Frm2 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   7200
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   45
         Width           =   14205
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   4095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   3240
            Width           =   14055
            Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
               Height          =   3195
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Width           =   13845
               _cx             =   24421
               _cy             =   5636
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
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmBuyBillInvestment.frx":6852
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
            Begin ImpulseButton.ISButton ISButton6 
               Height          =   330
               Left            =   12000
               TabIndex        =   81
               ToolTipText     =   "Õ–› «·ﬂ·"
               Top             =   3600
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ’›"
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
               ButtonImage     =   "FrmBuyBillInvestment.frx":6A05
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton4 
               Height          =   330
               Left            =   10080
               TabIndex        =   82
               ToolTipText     =   "Õ–› «·ﬂ·"
               Top             =   3600
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
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
               ButtonImage     =   "FrmBuyBillInvestment.frx":D267
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«—»«Õ «·»Ì⁄ ··„”«Â„ «·„ ‰«“·"
               Height          =   285
               Index           =   3
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   3600
               Width           =   2235
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   0
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   3600
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   6
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   3600
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   285
               Index           =   5
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   3600
               Width           =   1515
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   3735
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   14055
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   3855
               Left            =   0
               TabIndex        =   60
               Top             =   0
               Width           =   14055
               Begin VB.TextBox TxtValueCom1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   945
               End
               Begin VB.TextBox TxtComm1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   360
                  Width           =   945
               End
               Begin VB.ComboBox DcbTypeCom1 
                  Height          =   315
                  ItemData        =   "FrmBuyBillInvestment.frx":13AC9
                  Left            =   1890
                  List            =   "FrmBuyBillInvestment.frx":13ACB
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.TextBox TxtShrNetValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   -120
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin VB.TextBox TxtNetShare 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   -120
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin VB.Frame Frame9 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·„”«Â„…"
                  Height          =   735
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1800
                  Width           =   13935
                  Begin VB.TextBox TxtCode 
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
                     Left            =   11880
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   240
                     Width           =   945
                  End
                  Begin VB.TextBox TxtRemarks 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1320
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   14
                     Top             =   240
                     Width           =   1785
                  End
                  Begin VB.TextBox TxtSharValue 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     Height          =   315
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   13
                     Top             =   240
                     Width           =   1305
                  End
                  Begin VB.TextBox TxtSharNo 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     Height          =   315
                     Left            =   6240
                     RightToLeft     =   -1  'True
                     TabIndex        =   12
                     Top             =   240
                     Width           =   1305
                  End
                  Begin MSDataListLib.DataCombo DcbInvise 
                     Height          =   315
                     Left            =   8640
                     TabIndex        =   11
                     Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                     Top             =   240
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin ImpulseButton.ISButton ISButton3 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   15
                     ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                     Top             =   240
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   556
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
                     ButtonImage     =   "FrmBuyBillInvestment.frx":13ACD
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   285
                     Index           =   5
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„”«Â„…"
                     Height          =   285
                     Index           =   3
                     Left            =   12600
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·”Â„"
                     Height          =   285
                     Index           =   12
                     Left            =   4950
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·«”Â„"
                     Height          =   285
                     Index           =   11
                     Left            =   7200
                     RightToLeft     =   -1  'True
                     TabIndex        =   75
                     Top             =   240
                     Width           =   1515
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·„‘ —Ì"
                  Height          =   1215
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   720
                  Width           =   13935
                  Begin VB.TextBox TxtRemarks2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   120
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   102
                     Top             =   600
                     Width           =   3105
                  End
                  Begin VB.TextBox TxtValueCom2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFC0&
                     Height          =   315
                     Left            =   6960
                     RightToLeft     =   -1  'True
                     TabIndex        =   96
                     Top             =   960
                     Visible         =   0   'False
                     Width           =   945
                  End
                  Begin VB.TextBox TxtComm2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFC0&
                     Height          =   315
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   600
                     Width           =   2025
                  End
                  Begin VB.ComboBox DcbTypeCom2 
                     Height          =   315
                     ItemData        =   "FrmBuyBillInvestment.frx":1A32F
                     Left            =   7080
                     List            =   "FrmBuyBillInvestment.frx":1A331
                     RightToLeft     =   -1  'True
                     TabIndex        =   91
                     Top             =   600
                     Width           =   2385
                  End
                  Begin VB.TextBox TxtCusID 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   10
                     Top             =   240
                     Width           =   3105
                  End
                  Begin VB.TextBox TxtRecordNo 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   240
                     Width           =   2025
                  End
                  Begin VB.TextBox Text9 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   11760
                     TabIndex        =   6
                     Top             =   240
                     Width           =   1035
                  End
                  Begin MSDataListLib.DataCombo DcCustomerType 
                     Height          =   315
                     Left            =   10770
                     TabIndex        =   8
                     Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
                     Top             =   600
                     Width           =   2025
                     _ExtentX        =   3572
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbCus 
                     Height          =   315
                     Left            =   7080
                     TabIndex        =   7
                     Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                     Top             =   240
                     Width           =   4515
                     _ExtentX        =   7964
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   285
                     Index           =   6
                     Left            =   2880
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   600
                     Width           =   1515
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   285
                     Index           =   13
                     Left            =   5760
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   600
                     Width           =   1515
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄„Ê·…"
                     Height          =   285
                     Index           =   10
                     Left            =   9240
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   600
                     Width           =   1515
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·”Ã·"
                     Height          =   285
                     Index           =   16
                     Left            =   5760
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÂÊÌ…"
                     Height          =   285
                     Index           =   17
                     Left            =   2910
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄„Ì·"
                     Height          =   285
                     Index           =   1
                     Left            =   12480
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·⁄„Ì·"
                     Height          =   285
                     Index           =   0
                     Left            =   11970
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   600
                     Width           =   1890
                  End
               End
               Begin VB.TextBox Text12 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   11670
                  TabIndex        =   3
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.ComboBox CboPayMentType 
                  Height          =   315
                  Left            =   4230
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   360
                  Width           =   1545
               End
               Begin MSDataListLib.DataCombo DcbSales 
                  Height          =   315
                  Left            =   6960
                  TabIndex        =   4
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   360
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„Ê·…"
                  Height          =   285
                  Index           =   9
                  Left            =   3030
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   360
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   285
                  Index           =   19
                  Left            =   750
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   360
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                  Height          =   285
                  Index           =   4
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   360
                  Width           =   1890
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ«∆„ »«·»Ì⁄"
                  Height          =   285
                  Index           =   1
                  Left            =   12480
                  TabIndex        =   61
                  Top             =   360
                  Width           =   1515
               End
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   120
            TabIndex        =   55
            Top             =   0
            Width           =   14055
            Begin VB.TextBox TxtSerial1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   11400
               RightToLeft     =   -1  'True
               TabIndex        =   0
               Top             =   240
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker XPDtbTrans 
               Height          =   315
               Left            =   8400
               TabIndex        =   1
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Format          =   94961665
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Bindings        =   "FrmBuyBillInvestment.frx":1A333
               Height          =   315
               Left            =   2280
               TabIndex        =   2
               Top             =   240
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
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
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   7
               Left            =   6480
               TabIndex        =   58
               Top             =   240
               Width           =   1605
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·›« Ê—…"
               Height          =   285
               Index           =   4
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· «—ÌŒ"
               Height          =   285
               Index           =   2
               Left            =   10050
               TabIndex        =   56
               Top             =   255
               Width           =   885
            End
         End
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmBuyBillInvestment.frx":1A348
      Left            =   15480
      List            =   "FrmBuyBillInvestment.frx":1A358
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   14505
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox tXTRootAccount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   26
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuyBillInvestment.frx":1A371
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   27
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuyBillInvestment.frx":1A70B
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   28
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuyBillInvestment.frx":1AAA5
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   29
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuyBillInvestment.frx":1AE3F
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "≈‘⁄«—  ‰«“·/»Ì⁄ «”Â„"
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
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmBuyBillInvestment.frx":1B1D9
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   33
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   15480
      TabIndex        =   34
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   2145
      Left            =   0
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   8280
      Width           =   14235
      _cx             =   25109
      _cy             =   3784
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
      Begin VB.Frame Frame7 
         Caption         =   "»Ì«‰«  „Õ«”»Ì…"
         Height          =   735
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   97
         Top             =   600
         Width           =   5895
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   98
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
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   36
         Top             =   1440
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   17
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyBillInvestment.frx":1C5DE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   19
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyBillInvestment.frx":22E40
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   18
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
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
            ButtonImage     =   "FrmBuyBillInvestment.frx":231DA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   20
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyBillInvestment.frx":29A3C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   21
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyBillInvestment.frx":29DD6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyBillInvestment.frx":2A370
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   3960
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmBuyBillInvestment.frx":2A70A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1920
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyBillInvestment.frx":30F6C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   10200
         TabIndex        =   42
         Top             =   120
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   -840
         Width           =   13965
         _cx             =   24633
         _cy             =   1005
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
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   405
         Left            =   4320
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
         Top             =   840
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   714
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… ⁄ﬁœ  ‰«“· ··„”«Â„"
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
         ButtonImage     =   "FrmBuyBillInvestment.frx":31306
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton7 
         Height          =   405
         Left            =   1560
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
         Top             =   840
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   714
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… ⁄ﬁœ ÃœÌœ ··„ ‰«“·"
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
         ButtonImage     =   "FrmBuyBillInvestment.frx":37B68
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton9 
         Height          =   330
         Left            =   5520
         TabIndex        =   86
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonStyle     =   1
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
         ButtonImage     =   "FrmBuyBillInvestment.frx":3E3CA
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13080
         TabIndex        =   43
         Top             =   120
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15600
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyBillInvestment.frx":44C2C
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyBillInvestment.frx":44FC6
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyBillInvestment.frx":45360
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyBillInvestment.frx":456FA
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyBillInvestment.frx":45A94
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyBillInvestment.frx":45E2E
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyBillInvestment.frx":461C8
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyBillInvestment.frx":46762
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
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
      ButtonImage     =   "FrmBuyBillInvestment.frx":46AFC
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
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
      ButtonImage     =   "FrmBuyBillInvestment.frx":4D35E
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
      ButtonImage     =   "FrmBuyBillInvestment.frx":53BC0
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   15480
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmBuyBillInvestment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
                Dim RootAccount1 As String
                        Dim RootAccount2 As String
                        Dim RootAccount3 As String
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecID As String
 Dim Account_Code_dynamic As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "  ≈‘⁄«—  ‰«“· /»Ì⁄ «”Â„ " & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 
Dim Sql As String
tablename = "TblBuyBilllInvestment"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1.Text)
Notevalue = 0
 notytype = 9053
Notevalue = val(lbl(6).Caption)
 BranchID = val(Dcbranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TxtNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 Sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                Sql = Sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   Sql = Sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                   Cn.Execute Sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function

Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim StrTempAccountCode As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
Dim i As Integer
Dim Sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
Dim LngDevID As Long
Dim Msg As String
Dim NotValue As Double
Dim Account_Code As String
Dim X As Integer
Dim rs As New ADODB.Recordset
Dim notes_serial As String
Dim notes_id As String
If TxtRemarks2.Text = "" Then
        Msg = " —ﬁ„ «·”‰œ " & TxtSerial1.Text & " » «—ÌŒ" & XPDtbTrans.value & "«‘⁄«—  ‰«“· Ê»Ì⁄ «”Â„   "
   Else
 Msg = TxtRemarks2.Text
End If
          Account_Code = GetMyAccountCode("TblCustemers", "CusID", val(DcbSales.BoundText))
          StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(DcbCus.BoundText))
    notes_id = general_noteid
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
   
    Dim Branch As Integer
    BranchID = val(Dcbranch.BoundText)
NotValue = val(lbl(6).Caption)
If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code, Round(NotValue, 2), 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
           If ModAccounts.AddNewDev(LngDevID, line_no, StrTempAccountCode, Round(NotValue, 2), 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If
''//////

Dim Account_Code_dynamic As String
Account_Code_dynamic = get_account_code_branch(131, my_branch)
NotValue = val(TxtValueCom1.Text)
If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code, Round(NotValue, 2), 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
               If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, Round(NotValue, 2), 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
                
End If
NotValue = val(TxtValueCom2.Text)
If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, StrTempAccountCode, Round(NotValue, 2), 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
               If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, Round(NotValue, 2), 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
                
End If

updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function



Sub loadInvisment(Optional CusID As Double)
Dim Sql As String
If SystemOptions.UserInterface = ArabicInterface Then
Sql = "SELECT     dbo.Tblinvestment.ID, dbo.Tblinvestment.Name"
Sql = Sql & " FROM         dbo.TblIPOSharer LEFT OUTER JOIN"
Sql = Sql & "                      dbo.Tblinvestment ON dbo.TblIPOSharer.OrderInvse = dbo.Tblinvestment.ID"
Sql = Sql & " Where (dbo.TblIPOSharer.SharID = " & CusID & ")"
Sql = Sql & " GROUP BY dbo.Tblinvestment.ID, dbo.Tblinvestment.Name"
Else
Sql = "SELECT     dbo.Tblinvestment.ID, dbo.Tblinvestment.NameE"
Sql = Sql & " FROM         dbo.TblIPOSharer LEFT OUTER JOIN"
Sql = Sql & "                      dbo.Tblinvestment ON dbo.TblIPOSharer.OrderInvse = dbo.Tblinvestment.ID"
Sql = Sql & " Where (dbo.TblIPOSharer.SharID = " & CusID & ")"
Sql = Sql & " GROUP BY dbo.Tblinvestment.ID, dbo.Tblinvestment.NameE"
End If
fill_combo DcbInvise, Sql
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DcbInvise_Change()
DcbInvise_Click (0)
End Sub
Private Sub DcbInvise_Click(Area As Integer)
Dim ShareValue As Double
If Me.TxtModFlg.Text <> "R" Then
GetInvestInformation val(DcbInvise.BoundText), , , ShareValue
TxtShrNetValue.Text = ShareValue
TxtNetShare.Text = GetTotalSharOfCustomer(val(DcbSales.BoundText), val(DcbInvise.BoundText))
End If
TxtCode.Text = Me.DcbInvise.BoundText
End Sub

Private Sub DcbInvise_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
If KeyCode = vbKeyF3 Then
FrmSearchinvestment.inde = 18
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End If
End If
End Sub

Private Sub Dcbranch_Change()
Dcbranch_Click (0)
End Sub

Private Sub Dcbranch_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub

Private Sub DcbSales_Change()
DcbSales_Click (0)
End Sub



Private Sub DcbSales_Click(Area As Integer)
  If val(DcbSales.BoundText) = 0 Then Exit Sub
    If val(DcbSales.BoundText) = val(Me.DcbCus.BoundText) Then
  If SystemOptions.UserInterface = ArabicInterface Then
  MsgBox "·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·»«∆⁄ ÂÊ «·„‘ —Ì "
  Else
  MsgBox "Can not be a Seller is the Buyer"
  End If
  DcbCus.BoundText = 0
  Exit Sub
  End If
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbSales.BoundText, EmpCode
    Me.Text12.Text = EmpCode

If val(DcbSales.BoundText) <> 0 Then
loadInvisment val(DcbSales.BoundText)

End If

End Sub

Private Sub DcbTypeCom1_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTypeCom1.ListIndex) = 1 Then
lbl(19).Caption = "‰”»…"
TxtValueCom1.Text = Round((val(TxtComm1.Text) * val(lbl(6).Caption)) / 100, 2)
Else
TxtValueCom1.Text = TxtComm1.Text
lbl(19).Caption = "ﬁÌ„…"
End If
End If
End Sub

Private Sub DcbTypeCom1_Click()
DcbTypeCom1_Change
End Sub

Private Sub DcbTypeCom2_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTypeCom2.ListIndex) = 1 Then
lbl(13).Caption = "‰”»…"
TxtValueCom2.Text = Round((val(TxtComm2.Text) * val(lbl(6).Caption)) / 100, 2)
Else
TxtValueCom2.Text = TxtComm2.Text
lbl(13).Caption = "ﬁÌ„…"
End If
End If
End Sub

Private Sub DcbTypeCom2_Click()
DcbTypeCom2_Change
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblBuyBilllInvestment order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
If SystemOptions.UserInterface = ArabicInterface Then
With CboPayMentType
.Clear
.AddItem "‰ﬁœ«"
.AddItem "«Ã·"
End With
Else
With CboPayMentType
.Clear
.AddItem "Cash"
.AddItem "Credit"
End With
End If
  If SystemOptions.UserInterface = ArabicInterface Then
     With DcbTypeCom1
     .AddItem "ﬁÌ„…"
     .AddItem "‰”»…"
     End With
          With DcbTypeCom2
     .AddItem "ﬁÌ„…"
     .AddItem "‰”»…"
     End With
     Else
      With DcbTypeCom1
     .AddItem "Value"
     .AddItem "Percentage"
     End With
          With DcbTypeCom2
     .AddItem "Value"
     .AddItem "Percentage"
     End With
     End If
     
     If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=20 and Flg=1  order by CusName"
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=20 and Flg=1  order by CusNamee"
    End If
    fill_combo DcbSales, My_SQL
    fill_combo DcbCus, My_SQL
   
    Dim Dcombos As New ClsDataCombos
   ' Dcombos.GetCustomersSuppliers 1, Me.DcbSales, True
   ' Dcombos.GetCustomersSuppliers 1, Me.DcbCus, True
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
   ' Dcombos.GetInvestmentActive Me.DcbInvise, 1
   Dcombos.GetInvStoreType Me.DcCustomerType
  '  Dcombos.GetCustomerType Me.DcCustomerType
    BtnLast_Click

    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
   Me.Refresh
   FiLLTXT
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim Sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                 StrSQL = "Delete From TblBuyBilllInvestmentDet Where BuyBilInvsID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
               StrSQL = "Delete From TblTransactionInvest Where BuyBilID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
              End If
    RsSavRec.Fields("Remarks2").value = TxtRemarks2.Text
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("InvesID").value = val(Me.DcbInvise.BoundText)
    RsSavRec.Fields("SellerID").value = val(Me.DcbSales.BoundText)
    RsSavRec.Fields("Payment").value = val(Me.CboPayMentType.ListIndex)
    RsSavRec.Fields("Cus_Type").value = val(Me.DcCustomerType.BoundText)
    RsSavRec.Fields("Cus_ID").value = val(Me.DcbCus.BoundText)
    RsSavRec.Fields("RecordNo").value = (Me.TxtRecordNo.Text)
    RsSavRec.Fields("CusID").value = (Me.TxtCusID.Text)
    RsSavRec.Fields("Remarks").value = (Me.TxtRemarks.Text)
    RsSavRec.Fields("SharValue").value = val((Me.TxtSharValue.Text))
    RsSavRec.Fields("SharNo").value = val((Me.TxtSharNo.Text))
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.Fields("NetShare").value = val((Me.TxtNetShare.Text))
    RsSavRec.Fields("ShrNetValue").value = val((Me.TxtShrNetValue.Text))
    ''/////////
    RsSavRec.Fields("Comm1").value = val((Me.TxtComm1.Text))
    RsSavRec.Fields("Comm2").value = val((Me.TxtComm2.Text))
    RsSavRec.Fields("ValueCom1").value = val((Me.TxtValueCom1.Text))
    RsSavRec.Fields("ValueCom2").value = val((Me.TxtValueCom2.Text))
    RsSavRec.Fields("TypeCom1").value = val((Me.DcbTypeCom1.ListIndex))
    RsSavRec.Fields("TypeCom2").value = val((Me.DcbTypeCom2.ListIndex))
    
    RsSavRec.update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblBuyBilllInvestmentDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
    Msg = Msg & " ‰«“·"
    Else
    Msg = Msg & "Waiver/Sale of Shares"
    End If
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("InvesID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("BuyBilInvsID").value = val(Me.TxtSerial1.Text)
                RsDevsub("InvesID").value = IIf((.TextMatrix(i, .ColIndex("InvesID"))) = "", Null, val(.TextMatrix(i, .ColIndex("InvesID"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(i, .ColIndex("Remarks"))))
                RsDevsub("SharNo").value = IIf((.TextMatrix(i, .ColIndex("SharNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("SharNo"))))
                RsDevsub("SharValue").value = IIf((.TextMatrix(i, .ColIndex("SharValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("SharValue"))))
                RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", Null, val(.TextMatrix(i, .ColIndex("Total"))))
                RsDevsub("BeforTotal").value = IIf((.TextMatrix(i, .ColIndex("BeforTotal"))) = "", Null, val(.TextMatrix(i, .ColIndex("BeforTotal"))))
                RsDevsub("Profit").value = IIf((.TextMatrix(i, .ColIndex("Profit"))) = "", Null, (.TextMatrix(i, .ColIndex("Profit"))))
                RsDevsub("SharValueBefor").value = IIf((.TextMatrix(i, .ColIndex("SharValueBefor"))) = "", Null, (.TextMatrix(i, .ColIndex("SharValueBefor"))))
       RsDevsub.update
       SavedTranInvest , val(TxtSerial1.Text), Msg, val(.TextMatrix(i, .ColIndex("SharNo"))), val(.TextMatrix(i, .ColIndex("SharValue"))), val(.TextMatrix(i, .ColIndex("InvesID"))), val(DcbSales.BoundText), -1
       SavedTranInvest , val(TxtSerial1.Text), Msg, val(.TextMatrix(i, .ColIndex("SharNo"))), val(.TextMatrix(i, .ColIndex("SharValue"))), val(.TextMatrix(i, .ColIndex("InvesID"))), val(DcbCus.BoundText), 1
        Sql = "Update TblDivInvesment  set BuyPayed=1 where InvesID= " & val(.TextMatrix(i, .ColIndex("InvesID"))) & ""
       Cn.Execute Sql
      End If
     Next i
    End With
 createVoucher
    
'''///////////////
  
      Select Case Me.TxtModFlg.Text
        Case "N"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & Chr(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbSales.BoundText = IIf(IsNull(RsSavRec.Fields("SellerID").value), "", RsSavRec.Fields("SellerID").value)
    CboPayMentType.ListIndex = IIf(IsNull(RsSavRec.Fields("Payment").value), -1, RsSavRec.Fields("Payment").value)
    DcCustomerType.BoundText = IIf(IsNull(RsSavRec.Fields("Cus_Type").value), "", RsSavRec.Fields("Cus_Type").value)
    DcbCus.BoundText = IIf(IsNull(RsSavRec.Fields("Cus_ID").value), "", RsSavRec.Fields("Cus_ID").value)
    TxtRecordNo.Text = IIf(IsNull(RsSavRec.Fields("RecordNo").value), "", RsSavRec.Fields("RecordNo").value)
    TxtCusID.Text = IIf(IsNull(RsSavRec.Fields("CusID").value), "", RsSavRec.Fields("CusID").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtSharNo.Text = IIf(IsNull(RsSavRec.Fields("SharNo").value), 0, RsSavRec.Fields("SharNo").value)
    TxtSharValue.Text = IIf(IsNull(RsSavRec.Fields("SharValue").value), 0, RsSavRec.Fields("SharValue").value)
     Me.DcbInvise.BoundText = IIf(IsNull(RsSavRec.Fields("InvesID").value), "", RsSavRec.Fields("InvesID").value)
     Me.TxtNetShare.Text = IIf(IsNull(RsSavRec.Fields("NetShare").value), 0, RsSavRec.Fields("NetShare").value)
     Me.TxtShrNetValue.Text = IIf(IsNull(RsSavRec.Fields("ShrNetValue").value), "", RsSavRec.Fields("ShrNetValue").value)
     Me.TxtRemarks2.Text = IIf(IsNull(RsSavRec.Fields("Remarks2").value), "", RsSavRec.Fields("Remarks2").value)
     '''//////////
     DcbTypeCom2.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeCom2").value), -1, RsSavRec.Fields("TypeCom2").value)
     DcbTypeCom1.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeCom1").value), -1, RsSavRec.Fields("TypeCom1").value)
     Me.TxtValueCom2.Text = IIf(IsNull(RsSavRec.Fields("ValueCom2").value), 0, RsSavRec.Fields("ValueCom2").value)
     Me.TxtValueCom1.Text = IIf(IsNull(RsSavRec.Fields("ValueCom1").value), 0, RsSavRec.Fields("ValueCom1").value)
     Me.TxtComm1.Text = IIf(IsNull(RsSavRec.Fields("Comm1").value), 0, RsSavRec.Fields("Comm1").value)
     Me.TxtComm2.Text = IIf(IsNull(RsSavRec.Fields("Comm2").value), 0, RsSavRec.Fields("Comm2").value)
    Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
     Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
     
    ''//////////
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
RelinGrid
ErrTrap:
End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
RelinGrid
End Sub
Sub GetInformationCustomer(Optional Cus_ID As Double)
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim Sql As String
If Cus_ID <> 0 Then
Sql = "select TypeInvestor,CustomerTypeID ,CustGID ,RecordNo from TblCustemers where CusID =" & Cus_ID & " "
Rs6.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
DcCustomerType.BoundText = IIf(IsNull(Rs6("TypeInvestor").value), "", Rs6("TypeInvestor").value)
TxtRecordNo.Text = IIf(IsNull(Rs6("CustGID").value), "", Rs6("CustGID").value)
TxtCusID.Text = IIf(IsNull(Rs6("CustGID").value), "", Rs6("CustGID").value)
Else
TxtCusID = ""
TxtRecordNo = ""
DcCustomerType.BoundText = ""
End If
End If
End Sub


Private Sub DcbCus_Change()
DcbCus_Click (0)
End Sub

Private Sub DcbCus_Click(Area As Integer)
  If val(DcbCus.BoundText) = 0 Then Exit Sub
  If val(DcbSales.BoundText) = val(Me.DcbCus.BoundText) Then
  If SystemOptions.UserInterface = ArabicInterface Then
  MsgBox "·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·»«∆⁄ ÂÊ «·„‘ —Ì "
  Else
  MsgBox "Can not be a Seller is the Buyer"
  End If
  DcbCus.BoundText = 0
  Exit Sub
  End If
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCus.BoundText, EmpCode
    Me.Text9.Text = EmpCode

If Me.TxtModFlg.Text <> "R" Then
If val(DcbCus.BoundText) <> 0 Then
GetInformationCustomer val(DcbCus.BoundText)

End If
End If
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub

Private Sub ISButton3_Click()

If val(DcbInvise.BoundText) = 0 Or DcbInvise.BoundText = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ   ≈Œ Ì«— «·„”«Â„… "
Else
MsgBox "Please Select Contributing"
End If
DcbInvise.SetFocus
Exit Sub
End If
If val(TxtSharNo.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· ⁄œœ «·«”Â„"
Else
MsgBox "Please Enter No of  Share"
End If
TxtSharNo.SetFocus
Exit Sub
End If
If val(TxtSharValue.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· ﬁÌ„… «·”Â„ "
Else
MsgBox "Please Enter Value"
End If
TxtSharValue.SetFocus
Exit Sub
End If
filgrid1
RelinGrid
End Sub

Private Sub ISButton4_Click()
Dim i As Integer
Dim Sql As String
If Me.TxtModFlg.Text <> "R" Then
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("InvesID"))) <> 0 Then
        Sql = "Update TblDivInvesment  set BuyPayed=null where InvesID= " & val(.TextMatrix(i, .ColIndex("InvesID"))) & ""
       Cn.Execute Sql
      End If
     Next i
    End With
    
 GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
End If
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double

    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.Text = "" And val(Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            Dcbranch.SetFocus
            Exit Sub
     End If
           If DcbSales.Text = "" And val(DcbSales.BoundText) = 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡≈Œ Ì«— «·»«∆⁄  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
           Else
            MsgBox "Please Select The Seller ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbSales.SetFocus
          Exit Sub
        End If
     If val(DcbCus.BoundText) = 0 And DcbCus.Text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·⁄„Ì·  "
     Else
     MsgBox "Please Select Customer"
     End If
     DcbCus.SetFocus
     Exit Sub
     End If

    
    With Me.GridInstallments
           If .Rows >= 2 Then
           If val(.TextMatrix(1, .ColIndex("InvesID"))) = 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "Ì—ÃÏ «Œ Ì«— «·„”«Â„…„⁄ «· ›«’Ì·"
           Else
           MsgBox "Please Enter Investment"
           End If
           Exit Sub
           End If
           End If
           If .Rows < 2 Then
           If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox "Ì—ÃÏ «Œ Ì«— «·„”«Â„…„⁄ «· ›«’Ì·"
           Else
           MsgBox "Please Enter Investment"
           End If
           Exit Sub
           End If
    End With
    '------------------------------ check if Empcode exist ----------------------
'   StrVacName = IsRecExist("TblEmploymentModel", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtSerial1.text) & "'")
  ' If StrVacName <> "" Then
 '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·«”„ „‰ ﬁ»·"
  '     MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
  '    TxtVacName.SetFocus
 '     Exit Sub
'   End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblBuyBilllInvestment", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim Sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
Sql = "SELECT     dbo.TblBuyBilllInvestmentDet.ID, dbo.TblBuyBilllInvestmentDet.BuyBilInvsID, dbo.TblBuyBilllInvestmentDet.Remarks, dbo.TblBuyBilllInvestmentDet.SharNo, "
Sql = Sql & "                       dbo.TblBuyBilllInvestmentDet.SharValue, dbo.TblBuyBilllInvestmentDet.Total, dbo.TblBuyBilllInvestmentDet.BeforTotal, dbo.TblBuyBilllInvestmentDet.Profit,"
Sql = Sql & "                       dbo.TblBuyBilllInvestmentDet.SharValueBefor , dbo.TblBuyBilllInvestmentDet.InvesID, dbo.Tblinvestment.name, dbo.Tblinvestment.NameE"
Sql = Sql & "  FROM         dbo.TblBuyBilllInvestmentDet LEFT OUTER JOIN"
Sql = Sql & "                       dbo.Tblinvestment ON dbo.TblBuyBilllInvestmentDet.InvesID = dbo.Tblinvestment.ID"
Sql = Sql & "  Where (dbo.TblBuyBilllInvestmentDet.BuyBilInvsID =" & val(TxtSerial1.Text) & ") "

  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   .TextMatrix(i, .ColIndex("SharNo")) = IIf(IsNull(Rs1("SharNo").value), 0, Rs1("SharNo").value)
                   .TextMatrix(i, .ColIndex("SharValue")) = IIf(IsNull(Rs1("SharValue").value), 0, Rs1("SharValue").value)
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), 0, Rs1("Total").value)
                   .TextMatrix(i, .ColIndex("BeforTotal")) = IIf(IsNull(Rs1("BeforTotal").value), 0, Rs1("BeforTotal").value)
                   .TextMatrix(i, .ColIndex("Profit")) = IIf(IsNull(Rs1("Profit").value), 0, Rs1("Profit").value)
                   .TextMatrix(i, .ColIndex("SharValueBefor")) = IIf(IsNull(Rs1("SharValueBefor").value), 0, Rs1("SharValueBefor").value)
                   .TextMatrix(i, .ColIndex("InvesID")) = IIf(IsNull(Rs1("InvesID").value), 0, Rs1("InvesID").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub



Private Sub ISButton6_Click()
Dim Sql As String
If Me.TxtModFlg.Text <> "R" Then
With Me.GridInstallments
If .Rows < 2 Then
Exit Sub
Else
 Sql = "Update TblDivInvesment  set BuyPayed=null where InvesID= " & val(.TextMatrix(.Row, .ColIndex("InvesID"))) & ""
       Cn.Execute Sql
.RemoveItem .Row
End If
End With
End If
End Sub

Private Sub ISButton8_Click()
FrmSearchinvestment.inde = 13
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End Sub





Private Sub ISButton9_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.Text, "170420168"
ErrTrap:
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text12.Text, EmpID
        DcbSales.BoundText = EmpID
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text9.Text, EmpID
        DcbCus.BoundText = EmpID
    End If
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
Me.DcbInvise.BoundText = TxtCode.Text
End Sub

Private Sub TxtComm1_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTypeCom1.ListIndex) = 1 Then
TxtValueCom1.Text = Round((val(TxtComm1.Text) * val(lbl(6).Caption)) / 100, 2)
Else
TxtValueCom1.Text = TxtComm1.Text
End If
End If
End Sub

Private Sub TxtComm1_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtComm1.Text, 0)
End Sub

Private Sub TxtComm2_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTypeCom2.ListIndex) = 1 Then
TxtValueCom2.Text = Round((val(TxtComm2.Text) * val(lbl(6).Caption)) / 100, 2)
Else
TxtValueCom2.Text = TxtComm2.Text
End If
End If
End Sub

Private Sub TxtComm2_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtComm2.Text, 0)
End Sub

Private Sub TxtNetShare_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtSharNo.Text = Me.TxtNetShare.Text
End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecID, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim Sql As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
               With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
        Sql = "Update TblDivInvesment  set BuyPayed=Null where InvesID= " & val(.TextMatrix(i, .ColIndex("InvesID"))) & ""
       Cn.Execute Sql
     Next i
    End With
          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
                  StrSQL = "Delete From TblBuyBilllInvestmentDet Where BuyBilInvsID =" & val(TxtSerial1.Text) & ""
                 Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "Delete From TblTransactionInvest Where BuyBilID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                                          
                                
                                    
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
             
               '''''''''''''''''''''''''''''''

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
     LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select

End Sub

' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
        Select Case Me.TxtModFlg.Text
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
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecID As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    XPDtbTrans.Enabled = True
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
     XPDtbTrans.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
   XPDtbTrans.Enabled = True
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
 
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
            GridInstallments.Rows = GridInstallments.Rows + 1
        Me.DCboUserName.BoundText = user_id
        Frm2.Enabled = True
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & Chr(13)
            Msg = Msg & " You can not edit this the record now" & Chr(13)
            Msg = Msg & "It was being edited by another user on the network"
           
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    lbl(6).Caption = 0
    TxtModFlg.Text = "N"
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
  
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      
        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
       
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
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
  MySQL = "SELECT     dbo.TblBuyBilllInvestment.ID, dbo.TblBuyBilllInvestment.RecordDate, dbo.TblBuyBilllInvestment.Payment, dbo.TblBuyBilllInvestment.RecordNo, "
  MySQL = MySQL & "                    dbo.TblBuyBilllInvestment.CusID, dbo.TblBuyBilllInvestment.SharNo, dbo.TblBuyBilllInvestment.Remarks, dbo.TblBuyBilllInvestment.SharValue,"
  MySQL = MySQL & "                    dbo.TblBuyBilllInvestment.Cus_ID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblBuyBilllInvestment.BranchID,"
  MySQL = MySQL & "                    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBuyBilllInvestment.UserID, dbo.TblBuyBilllInvestment.SellerID,"
  MySQL = MySQL & "                    TblCustemers_1.CusName AS SalerCusName, TblCustemers_1.CusNamee AS SalerCusNameE, TblCustemers_1.Fullcode AS SalFullcode,"
  MySQL = MySQL & "                    dbo.TblBuyBilllInvestment.Cus_Type, dbo.TblInvestorType.Name, dbo.TblInvestorType.NameE, dbo.TblInvestorType.Code, dbo.TblBuyBilllInvestment.InvesID,"
  MySQL = MySQL & "                    dbo.Tblinvestment.Name AS InvName, dbo.Tblinvestment.NameE AS InvNameE, dbo.TblBuyBilllInvestmentDet.Remarks AS DetRemarks,"
  MySQL = MySQL & "                    dbo.TblBuyBilllInvestmentDet.SharNo AS DetSharNo, dbo.TblBuyBilllInvestmentDet.SharValue AS DetSharValue, dbo.TblBuyBilllInvestmentDet.Total,"
  MySQL = MySQL & "                    dbo.TblBuyBilllInvestmentDet.BeforTotal, dbo.TblBuyBilllInvestmentDet.Profit, dbo.TblBuyBilllInvestmentDet.SharValueBefor,"
  MySQL = MySQL & "                    dbo.TblBuyBilllInvestmentDet.InvesID AS DetInvesID, Tblinvestment_1.Name AS DetInvName, Tblinvestment_1.NameE AS DetInvNameE"
  MySQL = MySQL & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBuyBilllInvestment LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBuyBilllInvestmentDet ON dbo.TblBuyBilllInvestment.ID = dbo.TblBuyBilllInvestmentDet.BuyBilInvsID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.Tblinvestment Tblinvestment_1 ON dbo.TblBuyBilllInvestmentDet.InvesID = Tblinvestment_1.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.Tblinvestment ON dbo.TblBuyBilllInvestment.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblInvestorType ON dbo.TblBuyBilllInvestment.Cus_Type = dbo.TblInvestorType.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCustemers TblCustemers_1 ON dbo.TblBuyBilllInvestment.SellerID = TblCustemers_1.CusID ON"
  MySQL = MySQL & "                    dbo.TblBranchesData.branch_id = dbo.TblBuyBilllInvestment.BranchID ON dbo.TblCustemers.CusID = dbo.TblBuyBilllInvestment.Cus_ID"
  MySQL = MySQL & "  Where (dbo.TblBuyBilllInvestment.ID =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBuyBillInvestment.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBuyBillInvestmentE.rpt"
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
        Msg = "No Data"
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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

'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = Chr(13) + Chr(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub


Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
   lbl(19).Caption = "Value"
   lbl(13).Caption = "Value"
   lbl(10).Caption = "Commission"
   lbl(9).Caption = "Commission"
   ISButton4.Caption = "Delete All"
   ISButton6.Caption = "Delete"
   ISButton2.Caption = "Print  contract"
   ISButton7.Caption = "Print a new contract"
    Me.Caption = "Waiver/Sale of Shares  "
    ISButton9.Caption = "Attachments"
    Me.lbl(4).Caption = "No"
    Me.lbl(2).Caption = "Date"
    Me.lbl(7).Caption = "Branch"
    lbl(1).Caption = "Employee"
    lbl(5).Caption = "Total"
    Me.Label1(2).Caption = Me.Caption
    Label1(4).Caption = "Payment"
    Frame8.Caption = "Buyer Data"
    Label1(1).Caption = "Name"
    Label1(0).Caption = "Type"
    lbl(3).Caption = "Profit"
    lbl(16).Caption = "Record No."
    lbl(17).Caption = "ID"
    Frame9.Caption = "Investment Data"
    Label1(3).Caption = "Investment"
    lbl(11).Caption = "Share No."
    lbl(12).Caption = "Share Value"
    Label1(5).Caption = "Remarks"
    ISButton3.Caption = "Add"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    C1Tab1.Caption = "Data"
    '''''''''''''' next
      Label1(35).Caption = "No.GL"
  '  Command8.Caption = "Acc.Statement"
    Frame7.Caption = "Accounting"
    Command9.Caption = "Print GL"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    
  With Me.GridInstallments
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("InvesID")) = "Investment No."
  .TextMatrix(0, .ColIndex("Name")) = "Investment"
  .TextMatrix(0, .ColIndex("SharNo")) = "Share No."
  .TextMatrix(0, .ColIndex("SharValue")) = "Value"
  .TextMatrix(0, .ColIndex("Total")) = "Total"
  .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
  End With
ErrTrap:
End Sub
Sub filgrid1()
Dim i, k As Integer

Dim Shareval As Double
With GridInstallments
k = .Rows
.Rows = .Rows + 1
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("InvesID")) = val(DcbInvise.BoundText)
.TextMatrix(i, .ColIndex("Name")) = DcbInvise.Text
.TextMatrix(i, .ColIndex("SharNo")) = val(TxtSharNo.Text)
.TextMatrix(i, .ColIndex("SharValue")) = val(TxtSharValue.Text)
.TextMatrix(i, .ColIndex("Total")) = val(TxtSharValue.Text) * val(TxtSharNo.Text)
.TextMatrix(i, .ColIndex("Remarks")) = TxtRemarks.Text
GetInvestInformation val(DcbInvise.BoundText), , , Shareval
     .TextMatrix(i, .ColIndex("SharValueBefor")) = Shareval
    .TextMatrix(i, .ColIndex("BeforTotal")) = Shareval * val(TxtSharNo.Text)
    .TextMatrix(i, .ColIndex("Profit")) = val(.TextMatrix(i, .ColIndex("Total"))) - val(.TextMatrix(i, .ColIndex("BeforTotal")))

Next i
'.AutoSize 0, .Cols - 1, False
End With
End Sub
Sub RelinGrid()
Dim Sm, summation As Double
Dim Counter As Integer
Dim i As Integer
Counter = 0
Sm = 0
summation = 0
lbl(6).Caption = 0
With Me.GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("Total"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("Total")))
Sm = Sm + val(.TextMatrix(i, .ColIndex("Profit")))
End If
Next i
lbl(6).Caption = summation
lbl(0).Caption = Sm
End With
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblBuyBilllInvestment"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub TxtSharNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSharNo.Text, 0)
End Sub

Private Sub TxtSharNo_LostFocus()
If val(Me.TxtSharNo.Text) > val(Me.TxtNetShare.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ ⁄œœ «·«”Â„ «ﬂ»— „‰ «·«”Â„ «·„ «Õ…"
Else
MsgBox "It can not be the number of shares larger than permissible shares"
End If
TxtSharNo.Text = 0
TxtSharNo.SetFocus
Exit Sub
End If
End Sub

Private Sub TxtSharValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSharValue.Text, 0)
End Sub

Private Sub TxtSharValue_LostFocus()
If Round(val(TxtShrNetValue.Text), 0) > Round(val(TxtSharValue.Text), 0) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ ”⁄— »Ì⁄ «·”Â„ «ﬁ· „‰ ”⁄— «·‘—«¡"
Else
MsgBox "Can not be a sales price lower than the Purchase"
End If
TxtSharValue.Text = 0
TxtSharValue.SetFocus
Exit Sub
End If
End Sub

Private Sub TxtShrNetValue_Change()
If Me.TxtModFlg <> "R" Then
TxtSharValue.Text = val(TxtShrNetValue.Text)
End If
End Sub
Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub
