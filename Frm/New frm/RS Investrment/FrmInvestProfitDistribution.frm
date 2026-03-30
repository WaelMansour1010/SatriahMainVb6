VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmInvestProfitDistribution 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14295
   Icon            =   "FrmInvestProfitDistribution.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10110
   ScaleWidth      =   14295
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
         Begin VB.Frame Frame8 
            BackColor       =   &H0080FFFF&
            Caption         =   "›Ê« Ì— «·„»Ì⁄« "
            Height          =   6375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   840
            Visible         =   0   'False
            Width           =   12735
            Begin VB.CheckBox Check18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   195
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   300
               Width           =   1200
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid1 
               Height          =   5100
               Left            =   120
               TabIndex        =   77
               Top             =   600
               Width           =   12360
               _cx             =   21802
               _cy             =   8996
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
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmInvestProfitDistribution.frx":6852
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
               ExplorerBar     =   1
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
            Begin ImpulseButton.ISButton ISButton7 
               Height          =   315
               Left            =   7080
               TabIndex        =   81
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   240
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   556
               Caption         =   "„Ê«›ﬁ"
               BackColor       =   14871017
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmInvestProfitDistribution.frx":6A05
               ColorButton     =   14871017
               ColorHoverText  =   16711680
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               ColorToggledText=   16711680
               ColorToggledHoverText=   16711680
               LowerToggledContent=   0   'False
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   82
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   5880
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·›Ê« Ì—"
               Height          =   255
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   80
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   5880
               Width           =   1575
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   79
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   5880
               Width           =   3975
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   12360
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   240
               Visible         =   0   'False
               Width           =   135
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   4695
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   2520
            Width           =   14055
            Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
               Height          =   4275
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   13845
               _cx             =   24421
               _cy             =   7541
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
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmInvestProfitDistribution.frx":D267
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
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   6
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   3720
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   285
               Index           =   5
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   3720
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
               Begin VB.TextBox TxtRemarks2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   90
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   98
                  Top             =   1440
                  Width           =   11865
               End
               Begin VB.TextBox TxtNeProfit 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin VB.TextBox TxtNetComm 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   945
               End
               Begin VB.ComboBox DcbTyp 
                  Height          =   315
                  ItemData        =   "FrmInvestProfitDistribution.frx":D3E3
                  Left            =   3120
                  List            =   "FrmInvestProfitDistribution.frx":D3E5
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   360
                  Width           =   1545
               End
               Begin VB.TextBox TxtCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   10890
                  TabIndex        =   86
                  Top             =   720
                  Width           =   1065
               End
               Begin VB.TextBox TotalShare 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   1080
                  Width           =   1545
               End
               Begin VB.TextBox TxtPorfetValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   1080
                  Width           =   1545
               End
               Begin VB.TextBox TxtSalValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   10410
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   1080
                  Width           =   1545
               End
               Begin VB.TextBox TxtInvestValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   720
                  Width           =   1545
               End
               Begin VB.TextBox Text3 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   10890
                  TabIndex        =   3
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtComm 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   360
                  Width           =   1545
               End
               Begin VB.Frame Frame9 
                  BackColor       =   &H00E2E9E9&
                  Height          =   735
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   2640
                  Visible         =   0   'False
                  Width           =   13935
                  Begin VB.TextBox Text12 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   8790
                     TabIndex        =   11
                     Top             =   240
                     Width           =   1065
                  End
                  Begin VB.TextBox TxtRemarks 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1320
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   14
                     Top             =   240
                     Width           =   2265
                  End
                  Begin VB.TextBox TxtSharNo 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     Height          =   315
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   13
                     Top             =   360
                     Width           =   1545
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
                     ButtonImage     =   "FrmInvestProfitDistribution.frx":D3E7
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin MSDataListLib.DataCombo DcbSales 
                     Height          =   315
                     Left            =   4560
                     TabIndex        =   12
                     Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                     Top             =   240
                     Width           =   4155
                     _ExtentX        =   7329
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   255
                     Index           =   0
                     Left            =   12240
                     TabIndex        =   83
                     Top             =   240
                     Width           =   1575
                     _Version        =   786432
                     _ExtentX        =   2778
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·ﬂ· «·„”«Â„Ì‰"
                     ForeColor       =   12582912
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   255
                     Index           =   1
                     Left            =   10920
                     TabIndex        =   84
                     Top             =   240
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·„”«Â„ „Õœœ"
                     ForeColor       =   12582912
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„”«Â„"
                     Height          =   285
                     Index           =   1
                     Left            =   9600
                     TabIndex        =   70
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   285
                     Index           =   5
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   67
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·«”Â„"
                     Height          =   285
                     Index           =   11
                     Left            =   1800
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1515
                  End
               End
               Begin MSDataListLib.DataCombo DcbInvise 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   6
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   720
                  Width           =   4035
                  _ExtentX        =   7117
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbEmp 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   4
                  Top             =   360
                  Width           =   4035
                  _ExtentX        =   7117
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   675
                  Left            =   120
                  TabIndex        =   8
                  ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                  Top             =   720
                  Width           =   2865
                  _ExtentX        =   5054
                  _ExtentY        =   1191
                  Caption         =   "›Ê« Ì— «·„»Ì⁄« "
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
                  ButtonImage     =   "FrmInvestProfitDistribution.frx":13C49
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
                  Index           =   1
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   1440
                  Width           =   1995
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„Ê·…"
                  Height          =   285
                  Index           =   10
                  Left            =   4740
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   360
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·«”Â„"
                  Height          =   285
                  Index           =   9
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   1080
                  Width           =   1995
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—»Õ «·ﬁ«»· ·· Ê“Ì⁄"
                  Height          =   285
                  Index           =   12
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1080
                  Width           =   1995
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã„«·Ì «·„»Ì⁄«  «·€Ì— „Ê“⁄…"
                  Height          =   285
                  Index           =   3
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   1080
                  Width           =   1995
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„… «·«Ã„«·Ì… ··„”«Â„…"
                  Height          =   285
                  Index           =   0
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   720
                  Width           =   1995
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁ«∆„ »«· Ê“Ì⁄"
                  Height          =   285
                  Index           =   0
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   360
                  Width           =   1995
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„”«Â„…"
                  Height          =   285
                  Index           =   3
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   720
                  Width           =   1995
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„…"
                  Height          =   285
                  Index           =   13
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   360
                  Width           =   1275
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
               Format          =   93388801
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Bindings        =   "FrmInvestProfitDistribution.frx":1A4AB
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
               Caption         =   "—ﬁ„ "
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
      ItemData        =   "FrmInvestProfitDistribution.frx":1A4C0
      Left            =   15480
      List            =   "FrmInvestProfitDistribution.frx":1A4D0
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
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   360
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
         ButtonImage     =   "FrmInvestProfitDistribution.frx":1A4E9
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
         ButtonImage     =   "FrmInvestProfitDistribution.frx":1A883
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
         ButtonImage     =   "FrmInvestProfitDistribution.frx":1AC1D
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
         ButtonImage     =   "FrmInvestProfitDistribution.frx":1AFB7
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ≈À»«   Ê“Ì⁄ «—»«Õ «·„”«Â„Ì‰"
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   7200
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmInvestProfitDistribution.frx":1B351
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
      Height          =   1905
      Left            =   0
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   8280
      Width           =   14235
      _cx             =   25109
      _cy             =   3360
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   480
         Width           =   6015
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   120
            Visible         =   0   'False
            Width           =   855
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
            TabIndex        =   97
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
         Top             =   1080
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
            ButtonImage     =   "FrmInvestProfitDistribution.frx":1C756
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
            ButtonImage     =   "FrmInvestProfitDistribution.frx":22FB8
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
            ButtonImage     =   "FrmInvestProfitDistribution.frx":23352
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
            ButtonImage     =   "FrmInvestProfitDistribution.frx":29BB4
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
            ButtonImage     =   "FrmInvestProfitDistribution.frx":29F4E
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
            ButtonImage     =   "FrmInvestProfitDistribution.frx":2A4E8
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
            ButtonImage     =   "FrmInvestProfitDistribution.frx":2A882
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
            ButtonImage     =   "FrmInvestProfitDistribution.frx":310E4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9840
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
         Height          =   330
         Left            =   1200
         TabIndex        =   88
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   720
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
         ButtonImage     =   "FrmInvestProfitDistribution.frx":3147E
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
            Picture         =   "FrmInvestProfitDistribution.frx":37CE0
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInvestProfitDistribution.frx":3807A
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInvestProfitDistribution.frx":38414
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInvestProfitDistribution.frx":387AE
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInvestProfitDistribution.frx":38B48
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInvestProfitDistribution.frx":38EE2
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInvestProfitDistribution.frx":3927C
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInvestProfitDistribution.frx":39816
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
      ButtonImage     =   "FrmInvestProfitDistribution.frx":39BB0
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
      ButtonImage     =   "FrmInvestProfitDistribution.frx":40412
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
      ButtonImage     =   "FrmInvestProfitDistribution.frx":46C74
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
Attribute VB_Name = "FrmInvestProfitDistribution"
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
 Dim RecId As String
 Dim Account_Code_dynamic As String
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double



Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DcbInvise_Change()
DcbInvise_Click (0)
End Sub

Private Sub DcbInvise_Click(Area As Integer)
Dim InvesTotal As Double
Dim CounterShare As Double
TxtCode.Text = DcbInvise.BoundText
If Me.TxtModFlg.Text <> "R" Then
If val(DcbInvise.BoundText) <> 0 Then
GetInvestInformation val(DcbInvise.BoundText), InvesTotal, CounterShare
TxtInvestValue.Text = InvesTotal
TotalShare.Text = CounterShare
End If
End If
End Sub

Private Sub DcbInvise_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
If KeyCode = vbKeyF3 Then
FrmSearchinvestment.inde = 19
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
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbSales.BoundText, EmpCode
    Me.Text12.Text = EmpCode
End Sub

Private Sub DcbTyp_Click()
DcbTyp_Change
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblInvestProfitDistri order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
     If SystemOptions.UserInterface = ArabicInterface Then
     With DcbTyp
     .AddItem "ﬁÌ„…"
     .AddItem "‰”»…"
     End With
     Else
      With DcbTyp
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
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetEmployees DcbEmp
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetInvestmentActive Me.DcbInvise
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
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                 StrSQL = "Delete From TblInvestProfitDistriDet Where InvProID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
              End If
    RsSavRec.Fields("Remarks2").value = TxtRemarks2.Text
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("InvesID").value = val(Me.DcbInvise.BoundText)
    RsSavRec.Fields("ShareID").value = val(Me.DcbSales.BoundText)
    RsSavRec.Fields("EmpID").value = val(Me.DcbEmp.BoundText)
    RsSavRec.Fields("Comm").value = val(Me.TxtComm.Text)
    RsSavRec.Fields("InvestValue").value = val(Me.TxtInvestValue.Text)
    RsSavRec.Fields("Remarks").value = (Me.TxtRemarks.Text)
    RsSavRec.Fields("SalValue").value = val((Me.TxtSalValue.Text))
    RsSavRec.Fields("PorfetValue").value = val((Me.TxtPorfetValue.Text))
    RsSavRec.Fields("SharNo").value = val((Me.TxtSharNo.Text))
    RsSavRec.Fields("TotalShare").value = val((Me.TotalShare.Text))
    If Rd(0).value = True Then
    RsSavRec.Fields("TypeShere").value = 1
    Else
    RsSavRec.Fields("TypeShere").value = 0
    End If
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.Fields("TypeCom").value = val(Me.DcbTyp.ListIndex)
    RsSavRec.Fields("NetComm").value = val((Me.TxtNetComm.Text))
    RsSavRec.Fields("NeProfit").value = (Me.TxtNeProfit.Text)
    
    RsSavRec.update

      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblInvestProfitDistriDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("ShareID"))) <> 0 Then
       RsDevsub.AddNew
       
                RsDevsub("InvProID").value = val(Me.TxtSerial1.Text)
                RsDevsub("TypeTrans").value = 0
                RsDevsub("ShareID").value = IIf((.TextMatrix(i, .ColIndex("ShareID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ShareID"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(i, .ColIndex("Remarks"))))
                RsDevsub("SharNo").value = IIf((.TextMatrix(i, .ColIndex("SharNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("SharNo"))))
                RsDevsub("Profit").value = IIf((.TextMatrix(i, .ColIndex("Profit"))) = "", Null, val(.TextMatrix(i, .ColIndex("Profit"))))
                RsDevsub("SharNoDes").value = IIf((.TextMatrix(i, .ColIndex("SharNoDes"))) = "", Null, (.TextMatrix(i, .ColIndex("SharNoDes"))))
       RsDevsub.update
      End If
     Next i
    End With
'''///////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = " SELECT  *  from TblInvestProfitDistriDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    With Me.Grid1
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("Cus_ID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("InvProID").value = val(Me.TxtSerial1.Text)
                RsDevsub("TypeTrans").value = 1
                RsDevsub("ShareID").value = IIf((.TextMatrix(i, .ColIndex("Cus_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Cus_ID"))))
                RsDevsub("SBINVID").value = IIf((.TextMatrix(i, .ColIndex("SBINVID"))) = "", Null, val(.TextMatrix(i, .ColIndex("SBINVID"))))
                RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
                RsDevsub("BilValue").value = IIf((.TextMatrix(i, .ColIndex("BilValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("BilValue"))))
                RsDevsub("Profit").value = IIf((.TextMatrix(i, .ColIndex("SumProfit"))) = "", Null, val(.TextMatrix(i, .ColIndex("SumProfit"))))
                
       RsDevsub.update
       If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
       sql = "Update TblSaleBilllInvestmentDet set Payed=1 where InvesID=" & val(Me.DcbInvise.BoundText) & " and SBINVID=" & val((.TextMatrix(i, .ColIndex("SBINVID")))) & ""
       Cn.Execute sql
       Else
        sql = "Update TblSaleBilllInvestmentDet set Payed=Null where InvesID=" & val(Me.DcbInvise.BoundText) & " and SBINVID=" & val((.TextMatrix(i, .ColIndex("SBINVID")))) & ""
       Cn.Execute sql
      End If
      End If
     Next i
    End With
  createVoucher
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
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
       TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
       XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
       Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
       DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
       Me.DcbInvise.BoundText = IIf(IsNull(RsSavRec.Fields("InvesID").value), "", RsSavRec.Fields("InvesID").value)
       Me.DcbEmp.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
       Me.TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
       Me.TxtComm.Text = IIf(IsNull(RsSavRec.Fields("Comm").value), 0, RsSavRec.Fields("Comm").value)
       Me.TxtInvestValue.Text = IIf(IsNull(RsSavRec.Fields("InvestValue").value), 0, RsSavRec.Fields("InvestValue").value)
       Me.TxtSalValue.Text = IIf(IsNull(RsSavRec.Fields("SalValue").value), 0, RsSavRec.Fields("SalValue").value)
       Me.TxtPorfetValue.Text = IIf(IsNull(RsSavRec.Fields("PorfetValue").value), 0, RsSavRec.Fields("PorfetValue").value)
       DcbSales.BoundText = IIf(IsNull(RsSavRec.Fields("ShareID").value), "", RsSavRec.Fields("ShareID").value)
       Me.TxtSharNo.Text = IIf(IsNull(RsSavRec.Fields("SharNo").value), 0, RsSavRec.Fields("SharNo").value)
       Me.TotalShare.Text = IIf(IsNull(RsSavRec.Fields("TotalShare").value), 0, RsSavRec.Fields("TotalShare").value)
       DcbTyp.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeCom").value), -1, RsSavRec.Fields("TypeCom").value)
       TxtNetComm.Text = IIf(IsNull(RsSavRec.Fields("NetComm").value), 0, RsSavRec.Fields("NetComm").value)
       TxtNeProfit.Text = IIf(IsNull(RsSavRec.Fields("NeProfit").value), 0, RsSavRec.Fields("NeProfit").value)
       TxtRemarks2.Text = IIf(IsNull(RsSavRec.Fields("Remarks2").value), "", RsSavRec.Fields("Remarks2").value)
       If Not (IsNull(RsSavRec.Fields("SharNo").value)) Then
       If RsSavRec.Fields("SharNo").value = 1 Then
       Me.Rd(0).value = True
       Else
       Me.Rd(1).value = True
       End If
       End If
      Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
     Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
RelinGrid
ErrTrap:
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Reline
End Sub
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "  «À»«   Ê“Ì⁄ «·«—»«Õ  " & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 
Dim sql As String
tablename = "TblInvestProfitDistri"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1.Text)
Notevalue = 0
 notytype = 9052
Notevalue = val(TxtPorfetValue.Text)
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
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                   Cn.Execute sql
                                                               
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
Dim sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
Dim LngDevID As Long
Dim Msg As String
Dim NotValue As Double
Dim Account_code As String
Dim X As Integer
Dim rs As New ADODB.Recordset
Dim notes_serial As String
Dim notes_id As String
If TxtRemarks2.Text = "" Then
        Msg = " —ﬁ„ «·”‰œ " & TxtSerial1.Text & " » «—ÌŒ" & XPDtbTrans.value & "«À»«   Ê“Ì⁄ «·«—»«Õ   "
    Else
    Msg = TxtRemarks2.Text
 End If
          Account_code = GetActiveInvestmenAccound(val(DcbInvise.BoundText))
    notes_id = general_noteid
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
   
    Dim Branch As Integer
    BranchID = val(Dcbranch.BoundText)
NotValue = val(TxtPorfetValue.Text)
If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, Account_code, Round(NotValue, 2), 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If

''//////////////////

With GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ShareID"))) <> 0 Then
NotValue = .TextMatrix(i, .ColIndex("Profit"))
StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("ShareID"))))
If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, StrTempAccountCode, Round(NotValue, 2), 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If
End If
Next i
End With

updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function
Private Sub Grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
With Me.Grid1
Select Case .ColKey(Col)
Case "payed"
Cancel = False
End Select
End With
End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
RelinGrid
End Sub
Private Sub DcbEmp_Change()
DcbEmp_Click (0)
End Sub
Private Sub Check18_Click()
    Dim i As Integer

    If Check18.value = vbChecked Then

        With Me.Grid1
 
            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.Grid1

            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    Reline
End Sub
Sub Reline()
    Dim IntCounter As Integer
    Dim Sm As Double
    Dim SumProfit As Double
    Sm = 0
    IntCounter = 0
    SumProfit = 0
    Dim i As Integer
    With Me.Grid1
        For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("BilValue")))
           SumProfit = SumProfit + val(.TextMatrix(i, .ColIndex("SumProfit"))) - val(.TextMatrix(i, .ColIndex("commission")))
           End If
           Next i
  
    End With
    Label3.Caption = SumProfit
   Label16.Caption = Sm
End Sub
Sub GetExpenseInfo(Optional ID As Double = 0)
If ID <> 0 Then
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = "SELECT  *"
sql = sql & " From dbo.TblExpensesInvesment"
sql = sql & " Where (ID = " & ID & ") "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
TxtInvestValue.Text = IIf(IsNull(Rs4("AfterDevlopValue").value), 0, Rs4("AfterDevlopValue").value)
Else
TxtInvestValue.Text = 0
End If
End If
End Sub

Private Sub DcbEmp_Click(Area As Integer)
 If val(DcbEmp.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcbEmp.BoundText, EmpCode
    Me.Text3.Text = EmpCode
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

Cancel = True
With Me.GridInstallments
Select Case .ColKey(Col)
Case "Remarks"
Cancel = False
End Select
End With
End Sub

Private Sub ISButton2_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtSerial1.Text, "170420169"
ErrTrap:
End Sub

Private Sub ISButton4_Click()
Frame8.Visible = True
Label16.Caption = 0
Label3.Caption = 0
If Me.TxtModFlg.Text = "N" Then
If val(DcbInvise.BoundText) <> 0 Then
FulBills val(DcbInvise.BoundText)
End If
Else
FullGridDataBill
Reline
RelinGrid
End If
End Sub

Private Sub ISButton7_Click()
TxtSalValue.Text = val(Label16.Caption)
TxtPorfetValue.Text = val(Label3.Caption) - val(TxtNetComm.Text)
TxtNeProfit.Text = val(Label3.Caption)
Frame8.Visible = False
FulGridInformation val(DcbInvise.BoundText)
RelinGrid
End Sub

Private Sub ISButton8_Click()
FrmSearchinvestment.inde = 15
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End Sub

Private Sub Label17_Click()
Frame8.Visible = False
End Sub

Private Sub DcbTyp_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTyp.ListIndex) = 1 Then
lbl(13).Caption = "‰”»…"
TxtNetComm.Text = Round((val(TxtComm.Text) * val(TxtNeProfit.Text)) / 100, 2)
Else
TxtNetComm.Text = TxtComm.Text
lbl(13).Caption = "ﬁÌ„…"
End If
End If
End Sub

Private Sub Rd_Click(Index As Integer)
If Rd(0).value = True Then
Text12.Enabled = False
Text12.Text = ""
DcbSales.Enabled = False
DcbSales.BoundText = 0
Else
Text12.Enabled = True
DcbSales.Enabled = True
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text12.Text, EmpID
        DcbSales.BoundText = EmpID
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text3.Text, EmpID
        DcbEmp.BoundText = EmpID
    End If
End Sub

Private Sub ISButton3_Click()
RelinGrid
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim total As Double
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
           If DcbEmp.Text = "" And val(DcbEmp.BoundText) = 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡≈Œ Ì«— «·„ÊŸ›  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
           Else
            MsgBox "Please Select Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbEmp.SetFocus
          Exit Sub
        End If
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
    StrRecID = new_id("TblInvestProfitDistri", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
 Sub FulGridInformation(Optional InvesID As Double = 0)
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1

sql = " SELECT     SUM(dbo.TblTransactionInvest.SharCount * dbo.TblTransactionInvest.Effict) AS Totalshar, dbo.TblTransactionInvest.CusID, dbo.TblTransactionInvest.InvesID, "
sql = sql & "                       dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode"
sql = sql & " FROM         dbo.TblTransactionInvest LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblTransactionInvest.CusID = dbo.TblCustemers.CusID"
sql = sql & " Where   (dbo.TblTransactionInvest.InvesID = " & InvesID & ")"
sql = sql & " GROUP BY dbo.TblTransactionInvest.CusID, dbo.TblTransactionInvest.InvesID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode"
sql = sql & " Having (SUM(dbo.TblTransactionInvest.SharCount * dbo.TblTransactionInvest.Effict) <> 0)"
Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
   
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("ShareID")) = IIf(IsNull(Rs1("CusID").value), 0, Rs1("CusID").value)
                   .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("Totalshar")) = IIf(IsNull(Rs1("Totalshar").value), 0, Rs1("Totalshar").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("CusName").value), "", Rs1("CusName").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("CusNamee").value), "", Rs1("CusNamee").value)
                   End If
                   If val(TotalShare.Text) <> 0 Then
                   .TextMatrix(i, .ColIndex("SharNo")) = Round(val(.TextMatrix(i, .ColIndex("Totalshar"))) / val(TotalShare.Text), 20)
                    .TextMatrix(i, .ColIndex("SharNo")) = Round(val(.TextMatrix(i, .ColIndex("Totalshar"))) / val(TotalShare.Text), 20)
                    .TextMatrix(i, .ColIndex("SharNo")) = Round(val(.TextMatrix(i, .ColIndex("SharNo"))), 20)
                    .TextMatrix(i, .ColIndex("SharNoDes")) = val(.TextMatrix(i, .ColIndex("SharNo"))) * 100 & "%"
                    End If
                      .TextMatrix(i, .ColIndex("Profit")) = Round((val(.TextMatrix(i, .ColIndex("SharNo"))) * val(TxtPorfetValue.Text)), 20)
                    .TextMatrix(i, .ColIndex("Profit")) = Round(val(.TextMatrix(i, .ColIndex("Profit"))), 20)
                    
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
    
 Sub FulBills(Optional InvesID As Double = 0)
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    Grid1.Clear flexClearScrollable, flexClearEverything
    Grid1.Rows = 1

sql = " SELECT     TOP 100 PERCENT dbo.TblSaleBilllInvestmentDet.InvesID, SUM(dbo.TblSaleBilllInvestmentDet.Net) AS BilValue, dbo.TblSaleBilllInvestment.Cus_ID,"
sql = sql & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblSaleBilllInvestment.BranchID, dbo.TblSaleBilllInvestmentDet.SBINVID,"
sql = sql & "                      dbo.TblSaleBilllInvestment.RecordDate, SUM(dbo.TblSaleBilllInvestmentDet.Profit) AS SumProfit, SUM(dbo.TblSaleBilllInvestmentDet.TotalCost) AS SumTotalCost,"
sql = sql & "                      SUM(dbo.TblSaleBilllInvestmentDet.MeterValue) As SumMeterValue , dbo.TblCustemers.Fullcode , dbo.TblSaleBilllInvestment.NetComm"
sql = sql & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSaleBilllInvestment ON dbo.TblCustemers.CusID = dbo.TblSaleBilllInvestment.Cus_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblSaleBilllInvestmentDet ON dbo.TblSaleBilllInvestment.ID = dbo.TblSaleBilllInvestmentDet.SBINVID"
sql = sql & " WHERE     (dbo.TblSaleBilllInvestmentDet.Payed IS NULL) AND (NOT (dbo.TblSaleBilllInvestmentDet.ID IS NULL)) AND (dbo.TblSaleBilllInvestmentDet.InvesID = " & InvesID & ") AND"
sql = sql & "                      (NOT (dbo.TblSaleBilllInvestmentDet.SBINVID IS NULL)) and (dbo.TblSaleBilllInvestmentDet.ReturnSal IS NULL) "

sql = sql & " AND (NOT (dbo.TblSaleBilllInvestment.RecordDate IS NULL)) AND"
sql = sql & "    (dbo.TblSaleBilllInvestmentDet.Payed IS NULL)and (dbo.TblSaleBilllInvestment.TypeRetSal is null or dbo.TblSaleBilllInvestment.TypeRetSal=0)"
sql = sql & " GROUP BY dbo.TblSaleBilllInvestmentDet.InvesID, dbo.TblSaleBilllInvestment.Cus_ID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
sql = sql & "                     dbo.TblSaleBilllInvestment.BranchId , dbo.TblSaleBilllInvestmentDet.SBINVID, dbo.TblSaleBilllInvestment.recorddate , dbo.TblCustemers.Fullcode , dbo.TblSaleBilllInvestment.NetComm"
Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   
                   .TextMatrix(i, .ColIndex("commission")) = IIf(IsNull(Rs1("NetComm").value), 0, Rs1("NetComm").value)
                   .TextMatrix(i, .ColIndex("Cus_ID")) = IIf(IsNull(Rs1("Cus_ID").value), 0, Rs1("Cus_ID").value)
                   .TextMatrix(i, .ColIndex("BilValue")) = IIf(IsNull(Rs1("BilValue").value), 0, Rs1("BilValue").value)
                   .TextMatrix(i, .ColIndex("SumProfit")) = IIf(IsNull(Rs1("SumProfit").value), 0, Rs1("SumProfit").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
                   .TextMatrix(i, .ColIndex("SBINVID")) = IIf(IsNull(Rs1("SBINVID").value), 0, Rs1("SBINVID").value)
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs1("CusName").value), "", Rs1("CusName").value)
                   Else
                   .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs1("CusNamee").value), "", Rs1("CusNamee").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
     Sub FullGridDataBill()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    Grid1.Clear flexClearScrollable, flexClearEverything
            Grid1.Rows = 1
sql = "SELECT     dbo.TblInvestProfitDistriDet.Profit, dbo.TblInvestProfitDistriDet.SharNo, dbo.TblInvestProfitDistriDet.Remarks, dbo.TblInvestProfitDistriDet.ID, "
sql = sql & "                      dbo.TblInvestProfitDistriDet.InvProID, dbo.TblInvestProfitDistriDet.ShareID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
sql = sql & "                      dbo.TblInvestProfitDistriDet.TypeTrans , dbo.TblInvestProfitDistriDet.SBINVID, dbo.TblInvestProfitDistriDet.BilValue, dbo.TblInvestProfitDistriDet.recorddate ,dbo.TblInvestProfitDistriDet.SharNoDes ,dbo.TblInvestProfitDistriDet.commission"
sql = sql & " FROM         dbo.TblInvestProfitDistriDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblInvestProfitDistriDet.ShareID = dbo.TblCustemers.CusID"
sql = sql & " Where (dbo.TblInvestProfitDistriDet.InvProID =" & val(TxtSerial1.Text) & ") and  dbo.TblInvestProfitDistriDet.TypeTrans=1"
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("BilValue")) = IIf(IsNull(Rs1("BilValue").value), 0, Rs1("BilValue").value)
                   .TextMatrix(i, .ColIndex("SBINVID")) = IIf(IsNull(Rs1("SBINVID").value), 0, Rs1("SBINVID").value)
                   .TextMatrix(i, .ColIndex("Cus_ID")) = IIf(IsNull(Rs1("ShareID").value), 0, Rs1("ShareID").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("SumProfit")) = IIf(IsNull(Rs1("Profit").value), 0, Rs1("Profit").value)
                   .TextMatrix(i, .ColIndex("commission")) = IIf(IsNull(Rs1("commission").value), 0, Rs1("commission").value)
                   .TextMatrix(i, .ColIndex("payed")) = True
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs1("CusName").value), "", Rs1("CusName").value)
                   Else
                   .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs1("CusNamee").value), "", Rs1("CusNamee").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
    
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
sql = "SELECT     dbo.TblInvestProfitDistriDet.Profit, dbo.TblInvestProfitDistriDet.SharNo, dbo.TblInvestProfitDistriDet.Remarks, dbo.TblInvestProfitDistriDet.ID, "
sql = sql & "                      dbo.TblInvestProfitDistriDet.InvProID, dbo.TblInvestProfitDistriDet.ShareID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
sql = sql & "                      dbo.TblCustemers.fullcode ,dbo.TblInvestProfitDistriDet.SharNoDes"
sql = sql & " FROM         dbo.TblInvestProfitDistriDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblInvestProfitDistriDet.ShareID = dbo.TblCustemers.CusID"
sql = sql & " Where (dbo.TblInvestProfitDistriDet.InvProID =" & val(TxtSerial1.Text) & ") and  dbo.TblInvestProfitDistriDet.TypeTrans=0"
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("SharNoDes")) = IIf(IsNull(Rs1("SharNoDes").value), "", Rs1("SharNoDes").value)
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   .TextMatrix(i, .ColIndex("SharNo")) = IIf(IsNull(Rs1("SharNo").value), 0, Rs1("SharNo").value)
                   .TextMatrix(i, .ColIndex("ShareID")) = IIf(IsNull(Rs1("ShareID").value), 0, Rs1("ShareID").value)
                   .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("Profit")) = IIf(IsNull(Rs1("Profit").value), 0, Rs1("Profit").value)
                   .TextMatrix(i, .ColIndex("Profit")) = IIf(IsNull(Rs1("Profit").value), 0, Rs1("Profit").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("CusName").value), "", Rs1("CusName").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("CusNamee").value), "", Rs1("CusNamee").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
DcbInvise.BoundText = TxtCode.Text
End Sub

Private Sub TxtComm_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtNeProfit.Text) <> 0 Then
If val(DcbTyp.ListIndex) = 1 Then
TxtNetComm.Text = Round((val(TxtComm.Text) * val(TxtNeProfit.Text)) / 100, 2)
Else
TxtNetComm.Text = TxtComm.Text
End If
If val(TxtNetComm.Text) <> 0 Then
ISButton7_Click
End If
End If
End If
End Sub



Private Sub TxtComm_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtComm.Text, 0)
End Sub

Private Sub TxtComm_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtNeProfit.Text) = 0 Then

    If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...  ·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·⁄„Ê·… «ﬂ»— „‰ «·—»Õ  "
     Else
     MsgBox "Can not commission greater than the profit "
     End If
     TxtComm.Text = 0
  Exit Sub
End If

End If


End Sub

Private Sub TxtNetComm_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtNeProfit.Text) < val(TxtNetComm.Text) Then
    If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...  ·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·⁄„Ê·… «ﬂ»— „‰ «·—»Õ  "
     Else
     MsgBox "Can not commission greater than the profit "
     End If
TxtComm.Text = 0
TxtNetComm.Text = 0
Exit Sub
End If

End If
End Sub

Private Sub TxtPorfetValue_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTyp.ListIndex) = 1 Then
TxtNetComm.Text = Round((val(TxtComm.Text) * val(TxtNeProfit.Text)) / 100, 2)
Else
TxtNetComm.Text = TxtComm.Text
End If
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
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
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
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
    Dim sql As String
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
               FullGridDataBill
                  With Me.Grid1
       For i = .FixedRows To .Rows - 1
        sql = "Update TblSaleBilllInvestmentDet set Payed=Null where InvesID=" & val(Me.DcbInvise.BoundText) & " and SBINVID=" & val((.TextMatrix(i, .ColIndex("SBINVID")))) & ""
       Cn.Execute sql
    
     Next i
    End With
                  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                  StrSQL = "Delete From TblInvestProfitDistriDet Where InvProID =" & val(TxtSerial1.Text) & ""
                 Cn.Execute StrSQL, , adExecuteNoRecords
                                          RsSavRec.delete
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
                   RecId As String)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
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
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
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
                Grid1.Clear flexClearScrollable, flexClearEverything
Grid1.Rows = 1
  
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
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
  MySQL = "SELECT     dbo.TblInvestProfitDistri.ID, dbo.TblInvestProfitDistri.RecordDate, dbo.TblInvestProfitDistri.Remarks, dbo.TblInvestProfitDistri.Comm, "
  MySQL = MySQL & "                    dbo.TblInvestProfitDistri.InvestValue, dbo.TblInvestProfitDistri.SalValue, dbo.TblInvestProfitDistri.PorfetValue, dbo.TblInvestProfitDistri.SharNo,"
  MySQL = MySQL & "                    dbo.TblInvestProfitDistri.TotalShare, dbo.TblInvestProfitDistri.UserID, dbo.TblInvestProfitDistri.BranchID, dbo.TblBranchesData.branch_name,"
  MySQL = MySQL & "                    dbo.TblBranchesData.branch_namee, dbo.TblInvestProfitDistri.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
  MySQL = MySQL & "                    dbo.TblInvestProfitDistri.InvesID, dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.TblInvestProfitDistri.TypeShere, dbo.TblInvestProfitDistri.ShareID,"
  MySQL = MySQL & "                    dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblInvestProfitDistriDet.Remarks AS DetRemarks,"
  MySQL = MySQL & "                    dbo.TblInvestProfitDistriDet.SharNo AS DetSharNo, dbo.TblInvestProfitDistriDet.Profit, dbo.TblInvestProfitDistriDet.TypeTrans, dbo.TblInvestProfitDistriDet.SBINVID,"
  MySQL = MySQL & "                    dbo.TblInvestProfitDistriDet.BilValue, dbo.TblInvestProfitDistriDet.RecordDate AS DetRecordDate, dbo.TblInvestProfitDistriDet.ShareID AS DetShareID,"
  MySQL = MySQL & "                    TblCustemers_1.CusName AS DetCusName, TblCustemers_1.CusNamee AS DetCusNamee, TblCustemers_1.Fullcode AS DetFullcode"
  MySQL = MySQL & "  FROM         dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblInvestProfitDistriDet ON TblCustemers_1.CusID = dbo.TblInvestProfitDistriDet.ShareID RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblInvestProfitDistri ON dbo.TblInvestProfitDistriDet.InvProID = dbo.TblInvestProfitDistri.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCustemers ON dbo.TblInvestProfitDistri.ShareID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.Tblinvestment ON dbo.TblInvestProfitDistri.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmployee ON dbo.TblInvestProfitDistri.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblInvestProfitDistri.BranchID = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & "  Where (dbo.TblInvestProfitDistri.id =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepInvsetProfitDistrbution.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepInvsetProfitDistrbutionE.rpt"
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
    Wrap = CHR(13) + CHR(10)
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
    Me.Caption = "Profit Distribution to Shareholders"
    Me.lbl(4).Caption = "No"
    Me.lbl(2).Caption = "Date"
    Me.lbl(7).Caption = "Branch"
    Label1(0).Caption = "Employee"
    lbl(13).Caption = "Commission"
    Label1(3).Caption = "Investment"
    lbl(0).Caption = "Investment Value"
    Me.Label1(2).Caption = Me.Caption
    ISButton4.Caption = "Sales Bills"
    lbl(3).Caption = "Total Sales"
    lbl(12).Caption = "Profit"
    Rd(0).RightToLeft = True
    Rd(1).RightToLeft = True
    Rd(0).Caption = "All"
    Rd(1).Caption = "Select "
    lbl(1).Caption = "Shareholder"
    Label1(5).Caption = "Remark"
    lbl(9).Caption = "No.Share"
    lbl(10).Caption = "Commission"
    ISButton2.Caption = "Attachments"
''///////
  Label1(35).Caption = "No.GL"
  '  Command8.Caption = "Acc.Statement"
    Frame7.Caption = "Accounting"
    Command9.Caption = "Print GL"
    
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
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
  .TextMatrix(0, .ColIndex("Code")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Shareholder"
  .TextMatrix(0, .ColIndex("SharNo")) = "Percentage"
  .TextMatrix(0, .ColIndex("Profit")) = "Profit"
  .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
  End With
  ISButton3.Caption = "Add"
  Frame8.Caption = "Sales Bills"
  Check18.RightToLeft = True
  Check18.Caption = "Select All"
  ISButton7.Caption = "Accept"
  Label15.Caption = "Total"
  C1Tab1.Caption = "Data"
  lbl(5).Caption = "Total"
    With Me.Grid1
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("payed")) = "Select"
  .TextMatrix(0, .ColIndex("SBINVID")) = "Bill No."
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("BilValue")) = "Value"
  .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
  .TextMatrix(0, .ColIndex("CusName")) = "Shareholder"
  End With
ErrTrap:
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
If val(.TextMatrix(i, .ColIndex("Profit"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("Profit")))
End If
Next i
lbl(6).Caption = summation

End With
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblInvestProfitDistri"
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

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub
